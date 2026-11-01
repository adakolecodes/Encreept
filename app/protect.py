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
import tempfile
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import BinaryIO

import pikepdf
import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont

APP_NAME = "Encreept"

MAX_PAGES = int(os.environ.get("ENCREEPT_MAX_PAGES", "100"))
CONVERT_TIMEOUT_SECONDS = int(os.environ.get("ENCREEPT_CONVERT_TIMEOUT", "240"))
# Caps the memory one page can use, however large its physical size is.
MAX_PAGE_SIDE_PX = 3000

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


def _add_image_page(pdf: pikepdf.Pdf, jpeg: bytes, size_px: tuple[int, int], size_pt: tuple[float, float]) -> None:
    width_pt, height_pt = size_pt
    page = pdf.add_blank_page(page_size=size_pt)
    image = pikepdf.Stream(
        pdf,
        jpeg,
        Type=pikepdf.Name.XObject,
        Subtype=pikepdf.Name.Image,
        Width=size_px[0],
        Height=size_px[1],
        ColorSpace=pikepdf.Name.DeviceRGB,
        BitsPerComponent=8,
        Filter=pikepdf.Name.DCTDecode,
    )
    page.Resources = pikepdf.Dictionary(XObject=pikepdf.Dictionary(Im0=image))
    page.Contents = pdf.make_stream(f"q {width_pt:.2f} 0 0 {height_pt:.2f} 0 0 cm /Im0 Do Q".encode())


def _rasterise_to_pdf(source_pdf: Path, preset: dict) -> pikepdf.Pdf:
    try:
        document = pdfium.PdfDocument(source_pdf)
    except pdfium.PdfiumError:
        raise ProtectionError(
            "This PDF couldn't be opened. It may be damaged or password-protected."
        ) from None

    output = pikepdf.new()
    try:
        page_count = len(document)
        if page_count == 0:
            raise ProtectionError("This document has no pages.")
        if page_count > MAX_PAGES:
            raise ProtectionError(
                f"This document has {page_count} pages. The limit is {MAX_PAGES}."
            )

        for index in range(page_count):
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
            _add_image_page(output, buffer.getvalue(), image.size, size_pt)
            image.close()
    except Exception:
        output.close()
        raise
    finally:
        document.close()
    return output


def _save_locked(pdf: pikepdf.Pdf, owner_password: str) -> bytes:
    """Anyone can open and read the file; everything else needs the owner password."""
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
    buffer = io.BytesIO()
    pdf.save(buffer, encryption=pikepdf.Encryption(owner=owner_password, user="", R=6, allow=permissions))
    return buffer.getvalue()


def protect_document(upload: BinaryIO, filename: str, quality: str, owner_password: str) -> bytes:
    """Run the pipeline. All working files are deleted before this returns."""
    extension = Path(filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ProtectionError("Upload a PDF, Word (.doc, .docx) or PowerPoint (.ppt, .pptx) file.")
    preset = QUALITY_PRESETS.get(quality)
    if preset is None:
        raise ProtectionError("Choose standard or high quality.")

    with tempfile.TemporaryDirectory(prefix="encreept-", ignore_cleanup_errors=True) as tmp:
        workdir = Path(tmp)
        # A fixed name keeps the user's filename away from the file system and command line.
        source = workdir / f"input{extension}"
        with source.open("wb") as handle:
            shutil.copyfileobj(upload, handle)

        _validate_file(source, extension)
        pdf_path = source if extension == ".pdf" else _convert_to_pdf(source, workdir)
        with _rasterise_to_pdf(pdf_path, preset) as output:
            return _save_locked(output, owner_password)
