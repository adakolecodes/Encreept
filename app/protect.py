"""The Encreept pipeline: document -> PDF -> watermarked page images -> locked PDF.

Every page is rasterised, so the output contains no text, fonts, links or
original metadata. Permission flags are added on top, but the rasterising is
what actually stops copying: PDF permissions are only honoured by viewers.
"""

import io
import os
import shutil
import signal
import subprocess
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import BinaryIO

import pikepdf
import pypdfium2 as pdfium
from lxml import etree
from PIL import Image, ImageDraw, ImageFont

APP_NAME = "Encreept"

MAX_PAGES = int(os.environ.get("ENCREEPT_MAX_PAGES", "400"))
CONVERT_TIMEOUT_SECONDS = int(os.environ.get("ENCREEPT_CONVERT_TIMEOUT", "900"))
# Caps the memory one page can use, however large its physical size is.
MAX_PAGE_SIDE_PX = 3000
# LibreOffice loads one piece at a time, which keeps long decks within 512 MB of memory.
SLIDES_PER_CONVERSION = 20

_PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_XML_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)

QUALITY_PRESETS = {
    "standard": {"dpi": 150, "jpeg_quality": 80},
    "high": {"dpi": 220, "jpeg_quality": 85},
}

PDF_MAGIC = b"%PDF-"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
# Extension -> folder that must exist inside the Office zip package.
OOXML_FOLDERS = {".docx": "word/", ".pptx": "ppt/"}
SUPPORTED_EXTENSIONS = {".pdf", ".doc", ".ppt", *OOXML_FOLDERS}


class ProtectionError(ValueError):
    """A problem with the uploaded file; the message is shown to the user."""


class ConverterError(RuntimeError):
    """The Office converter is missing or failed; the message is shown to the user."""


def _validate_file(path: Path, extension: str) -> None:
    with path.open("rb") as handle:
        head = handle.read(8)

    if extension == ".pdf":
        valid = head.startswith(PDF_MAGIC)
    elif extension in OOXML_FOLDERS:
        valid = _is_ooxml(path, OOXML_FOLDERS[extension])
    else:
        valid = head == OLE_MAGIC

    if not valid:
        raise ProtectionError(
            f"This file isn't a valid {extension[1:].upper()} document. "
            "It may be damaged or have the wrong extension."
        )


def _is_ooxml(path: Path, folder: str) -> bool:
    try:
        with zipfile.ZipFile(path) as package:
            return any(name.startswith(folder) for name in package.namelist())
    except zipfile.BadZipFile:
        return False


def _soffice_command() -> list[str] | None:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return [found]
    windows_default = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
    return [str(windows_default)] if windows_default.exists() else None


def _convert_to_pdf(source: Path, workdir: Path) -> Path:
    command = _soffice_command()
    if command is None:
        raise ConverterError("Word and PowerPoint conversion isn't available on this server.")

    outdir = workdir / "converted"
    profile = workdir / "lo-profile"
    # LibreOffice runs on untrusted files, so it doesn't inherit the app's secrets.
    env = {k: os.environ[k] for k in ("PATH", "SystemRoot", "TEMP", "TMP", "LANG") if k in os.environ}
    env["HOME"] = str(workdir)

    process = subprocess.Popen(
        [
            *command,
            "--headless",
            "--norestore",
            "--nolockcheck",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(source),
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        env=env,
        # soffice spawns child processes, so a timeout has to kill the whole group.
        start_new_session=os.name == "posix",
    )
    try:
        process.wait(timeout=CONVERT_TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired:
        if os.name == "posix":
            os.killpg(process.pid, signal.SIGKILL)
        else:
            process.kill()
        process.wait()
        raise ConverterError("The document took too long to convert. Try a smaller file.") from None

    converted = outdir / f"{source.stem}.pdf"
    if not converted.exists():
        raise ProtectionError(
            "This document couldn't be converted. It may be damaged or password-protected."
        )
    return converted


def _slide_part_name(target: str) -> str:
    return target.lstrip("/") if target.startswith("/") else f"ppt/{target}"


def _slide_is_hidden(package: zipfile.ZipFile, target: str | None) -> bool:
    if target is None:
        return False
    try:
        return etree.fromstring(package.read(_slide_part_name(target)), _XML_PARSER).get("show") == "0"
    except (KeyError, etree.XMLSyntaxError):
        return False


def _fix_slide_number(slide_xml: bytes, number: int) -> bytes:
    """Replace slide number fields with fixed text, as a piece would otherwise restart from 1."""
    slide = etree.fromstring(slide_xml, _XML_PARSER)
    fields = slide.xpath("//a:fld[@type='slidenum']", namespaces={"a": _DML_NS})
    if not fields:
        return slide_xml
    for field in fields:
        field.tag = f"{{{_DML_NS}}}r"
        field.attrib.clear()
        for child in list(field):
            if child.tag == f"{{{_DML_NS}}}pPr":
                field.remove(child)
            elif child.tag == f"{{{_DML_NS}}}t":
                child.text = str(number)
    return etree.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True)


def _split_deck(source: Path, workdir: Path) -> list[Path]:
    """Split a long deck into smaller decks, in slide order. Short decks come back unchanged.

    LibreOffice's memory grows with the slides in the file it loads, whatever it exports.
    Each piece keeps every package file and only lists its own slides in presentation.xml,
    so links between slides still resolve and LibreOffice doesn't load the rest.
    """
    try:
        with zipfile.ZipFile(source) as package:
            presentation = etree.fromstring(package.read("ppt/presentation.xml"), _XML_PARSER)
            slide_list = presentation.find(f"{{{_PML_NS}}}sldIdLst")
            slide_ids = [] if slide_list is None else list(slide_list)
            if len(slide_ids) <= SLIDES_PER_CONVERSION:
                return [source]

            relationships = etree.fromstring(package.read("ppt/_rels/presentation.xml.rels"), _XML_PARSER)
            targets = {rel.get("Id"): rel.get("Target") for rel in relationships}
            first_number = int(presentation.get("firstSlideNum", "1"))

            # Custom shows and section lists can name slides that a piece no longer lists.
            for tag in ("custShowLst", "extLst"):
                element = presentation.find(f"{{{_PML_NS}}}{tag}")
                if element is not None:
                    presentation.remove(element)

            pieces = []
            for start in range(0, len(slide_ids), SLIDES_PER_CONVERSION):
                chunk = slide_ids[start : start + SLIDES_PER_CONVERSION]
                # LibreOffice skips hidden slides, and a piece with nothing to export makes no PDF.
                if all(_slide_is_hidden(package, targets.get(s.get(_REL_ID))) for s in chunk):
                    continue

                slide_list[:] = chunk
                changed = {
                    "ppt/presentation.xml": etree.tostring(
                        presentation, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                }
                for offset, slide_id in enumerate(chunk):
                    name = _slide_part_name(targets[slide_id.get(_REL_ID)])
                    changed[name] = _fix_slide_number(package.read(name), first_number + start + offset)

                piece = workdir / f"piece-{start}.pptx"
                with zipfile.ZipFile(piece, "w", zipfile.ZIP_STORED) as output:
                    for info in package.infolist():
                        output.writestr(info.filename, changed.get(info.filename) or package.read(info))
                pieces.append(piece)
            return pieces or [source]
    except (KeyError, ValueError, etree.XMLSyntaxError, zipfile.BadZipFile):
        # An unusual package: let LibreOffice try the whole file.
        return [source]


def _convert_document(source: Path, extension: str, workdir: Path) -> list[Path]:
    pieces = _split_deck(source, workdir) if extension == ".pptx" else [source]
    converted = []
    for piece in pieces:
        converted.append(_convert_to_pdf(piece, workdir))
        if piece != source:
            piece.unlink()
    return converted


@lru_cache(maxsize=8)
def _watermark_tile(font_px: int) -> Image.Image:
    """One rotated logo-and-name stamp. Dark ink with a light edge shows on any page colour."""
    font = ImageFont.load_default(size=font_px)
    ink = (20, 38, 79, 62)
    edge = (255, 255, 255, 62)
    stroke = max(1, font_px // 24)

    measure = ImageDraw.Draw(Image.new("L", (1, 1)))
    left, top, right, bottom = measure.textbbox((0, 0), APP_NAME, font=font, stroke_width=stroke)
    text_w, text_h = right - left, bottom - top

    icon_w = int(font_px * 0.8)
    gap = font_px // 3
    pad = font_px // 4
    stamp = Image.new("RGBA", (icon_w + gap + text_w + pad * 2, max(text_h, font_px) + pad * 2), (0, 0, 0, 0))
    draw = ImageDraw.Draw(stamp)

    # Padlock logo: shackle over a solid body.
    body_top = pad + int(font_px * 0.42)
    body_bottom = pad + int(font_px * 0.98)
    line = max(2, font_px // 10)
    shackle = (pad + icon_w * 0.18, pad, pad + icon_w * 0.82, body_top + font_px * 0.3)
    draw.arc(shackle, 180, 360, fill=edge, width=line + stroke * 2)
    draw.arc(shackle, 180, 360, fill=ink, width=line)
    draw.rounded_rectangle(
        (pad, body_top, pad + icon_w, body_bottom),
        radius=font_px // 10,
        fill=ink,
        outline=edge,
        width=stroke,
    )
    draw.text(
        (pad + icon_w + gap - left, pad - top),
        APP_NAME,
        font=font,
        fill=ink,
        stroke_width=stroke,
        stroke_fill=edge,
    )
    return stamp.rotate(30, expand=True, resample=Image.Resampling.BICUBIC)


def _apply_watermark(page: Image.Image) -> Image.Image:
    width, height = page.size
    tile = _watermark_tile(max(16, min(width, height) // 20))
    step_x, step_y = int(tile.width * 1.25), int(tile.height * 1.1)

    overlay = Image.new("RGBA", page.size, (0, 0, 0, 0))
    for row, y in enumerate(range(-tile.height // 2, height, step_y)):
        offset = step_x // 2 if row % 2 else 0
        for x in range(-tile.width // 2 - offset, width, step_x):
            overlay.paste(tile, (x, y))
    return Image.alpha_composite(page.convert("RGBA"), overlay).convert("RGB")


class _ImagePdfWriter:
    """Writes image-only pages straight to disk, so long documents don't build up in memory.

    Building the file with pikepdf holds every page image in memory until it saves.
    """

    def __init__(self, handle: BinaryIO) -> None:
        self._handle = handle
        self._offsets: dict[int, int] = {}
        self._page_numbers: list[int] = []
        # Objects 1 and 2 (catalog and page tree) are written last, once the pages are known.
        self._next_number = 3
        handle.write(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")

    def _write_object(self, number: int, body: bytes) -> None:
        self._offsets[number] = self._handle.tell()
        self._handle.write(b"%d 0 obj\n%s\nendobj\n" % (number, body))

    def _write_stream(self, number: int, dictionary: bytes, data: bytes) -> None:
        self._write_object(number, b"<< %s /Length %d >>\nstream\n%s\nendstream" % (dictionary, len(data), data))

    def add_page(self, jpeg: bytes, size_px: tuple[int, int], size_pt: tuple[float, float]) -> None:
        page, contents, image = self._next_number, self._next_number + 1, self._next_number + 2
        self._next_number += 3
        width_pt, height_pt = size_pt

        self._write_object(
            page,
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width_pt:.2f} {height_pt:.2f}] "
            f"/Resources << /XObject << /Im0 {image} 0 R >> >> /Contents {contents} 0 R >>".encode(),
        )
        self._write_stream(contents, b"", f"q {width_pt:.2f} 0 0 {height_pt:.2f} 0 0 cm /Im0 Do Q".encode())
        self._write_stream(
            image,
            f"/Type /XObject /Subtype /Image /Width {size_px[0]} /Height {size_px[1]} "
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode".encode(),
            jpeg,
        )
        self._page_numbers.append(page)

    def finish(self) -> None:
        kids = " ".join(f"{number} 0 R" for number in self._page_numbers)
        self._write_object(1, b"<< /Type /Catalog /Pages 2 0 R >>")
        self._write_object(2, f"<< /Type /Pages /Kids [{kids}] /Count {len(self._page_numbers)} >>".encode())

        xref_offset = self._handle.tell()
        self._handle.write(b"xref\n0 %d\n0000000000 65535 f \n" % self._next_number)
        for number in range(1, self._next_number):
            self._handle.write(b"%010d 00000 n \n" % self._offsets[number])
        self._handle.write(
            b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (self._next_number, xref_offset)
        )


def _open_pdf(path: Path) -> pdfium.PdfDocument:
    try:
        return pdfium.PdfDocument(path)
    except pdfium.PdfiumError:
        raise ProtectionError(
            "This PDF couldn't be opened. It may be damaged or password-protected."
        ) from None


def _rasterise_to_pdf(sources: list[Path], preset: dict, destination: Path) -> None:
    page_count = 0
    for source in sources:
        document = _open_pdf(source)
        page_count += len(document)
        document.close()
    if page_count == 0:
        raise ProtectionError("This document has no pages.")
    if page_count > MAX_PAGES:
        raise ProtectionError(
            f"This document has {page_count} pages. The limit is {MAX_PAGES}."
        )

    with destination.open("wb") as handle:
        writer = _ImagePdfWriter(handle)
        for source in sources:
            document = _open_pdf(source)
            try:
                for index in range(len(document)):
                    page = document[index]
                    try:
                        size_pt = page.get_size()
                        scale = min(preset["dpi"] / 72, MAX_PAGE_SIDE_PX / max(size_pt))
                        bitmap = page.render(scale=scale, may_draw_forms=True)
                        image = _apply_watermark(bitmap.to_pil())
                        bitmap.close()
                    finally:
                        page.close()

                    buffer = io.BytesIO()
                    image.save(buffer, format="JPEG", quality=preset["jpeg_quality"], optimize=True)
                    writer.add_page(buffer.getvalue(), image.size, size_pt)
                    image.close()
            finally:
                document.close()
        writer.finish()


def _save_locked(plain_pdf: Path, owner_password: str, destination: Path) -> None:
    """Anyone can open and read the file; everything else needs the owner password."""
    with pikepdf.open(plain_pdf, access_mode=pikepdf.AccessMode.stream) as pdf:
        _write_locked(pdf, owner_password, destination)


def _write_locked(pdf: pikepdf.Pdf, owner_password: str, destination: Path) -> None:
    with pdf.open_metadata(set_pikepdf_as_editor=False) as metadata:
        metadata["pdf:Producer"] = APP_NAME
    pdf.docinfo["/Producer"] = APP_NAME
    pdf.docinfo["/Creator"] = APP_NAME

    # The accessibility flag isn't set: AES-256 PDFs always allow it, and there's no text to read.
    permissions = pikepdf.Permissions(
        extract=False,
        modify_annotation=False,
        modify_assembly=False,
        modify_form=False,
        modify_other=False,
        print_lowres=False,
        print_highres=False,
    )
    pdf.save(destination, encryption=pikepdf.Encryption(owner=owner_password, user="", R=6, allow=permissions))


def protect_document(upload: BinaryIO, filename: str, quality: str, owner_password: str, workdir: Path) -> Path:
    """Run the pipeline and return the finished PDF. The caller owns and deletes workdir."""
    extension = Path(filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ProtectionError("Upload a PDF, Word (.doc, .docx) or PowerPoint (.ppt, .pptx) file.")
    preset = QUALITY_PRESETS.get(quality)
    if preset is None:
        raise ProtectionError("Choose standard or high quality.")

    # A fixed name keeps the user's filename away from the file system and command line.
    source = workdir / f"input{extension}"
    with source.open("wb") as handle:
        shutil.copyfileobj(upload, handle)

    _validate_file(source, extension)
    pdf_paths = [source] if extension == ".pdf" else _convert_document(source, extension, workdir)
    plain_pdf = workdir / "pages.pdf"
    _rasterise_to_pdf(pdf_paths, preset, plain_pdf)
    destination = workdir / "output.pdf"
    _save_locked(plain_pdf, owner_password, destination)
    return destination
