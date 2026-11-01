import io
import sys
import textwrap
import zipfile

import pikepdf
import pypdfium2 as pdfium
import pytest
from docx import Document
from fastapi.testclient import TestClient
from lxml import etree
from pptx import Presentation

from app import main, protect

SECRET_TEXT = "Confidential paragraph"

needs_libreoffice = pytest.mark.skipif(protect._soffice_command() is None, reason="LibreOffice isn't installed")


def make_pdf(pages: int = 2, size=(612, 792), rotate: int = 0) -> bytes:
    pdf = pikepdf.new()
    pdf.docinfo["/Author"] = "Original Author"
    font = pikepdf.Dictionary(Type=pikepdf.Name.Font, Subtype=pikepdf.Name.Type1, BaseFont=pikepdf.Name.Helvetica)
    for _ in range(pages):
        page = pdf.add_blank_page(page_size=size)
        page.Resources = pikepdf.Dictionary(Font=pikepdf.Dictionary(F1=font))
        page.Contents = pdf.make_stream(f"BT /F1 24 Tf 72 400 Td ({SECRET_TEXT}) Tj ET".encode())
        if rotate:
            page.Rotate = rotate
    buffer = io.BytesIO()
    pdf.save(buffer)
    return buffer.getvalue()


def extract_text(data: bytes) -> str:
    document = pdfium.PdfDocument(data)
    return "".join(document[i].get_textpage().get_text_range() for i in range(len(document)))


def make_office(folder: str) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr(f"{folder}main.xml", "<xml/>")
    return buffer.getvalue()


def upload(client, data: bytes, name: str = "report.pdf", **form):
    return client.post("/api/protect", files={"file": (name, data)}, data=form)


@pytest.fixture
def client():
    return TestClient(main.app)


def test_output_has_no_text_and_is_locked(client):
    source = make_pdf()
    assert SECRET_TEXT in extract_text(source)

    response = upload(client, source)

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/pdf"
    assert 'filename="report-encreept.pdf"' in response.headers["content-disposition"]
    assert response.headers["cache-control"] == "no-store"
    assert SECRET_TEXT not in extract_text(response.content)
    assert extract_text(response.content).strip() == ""

    # Opens without a password, with every permission off.
    with pikepdf.open(io.BytesIO(response.content)) as pdf:
        assert pdf.is_encrypted
        assert pdf.encryption.R == 6
        allowed = pdf.allow
        assert not any(
            [
                allowed.extract,
                allowed.modify_other,
                allowed.modify_assembly,
                allowed.modify_annotation,
                allowed.modify_form,
                allowed.print_lowres,
                allowed.print_highres,
            ]
        )
        assert len(pdf.pages) == 2
        assert "/Font" not in pdf.pages[0].Resources
        assert "/Author" not in pdf.docinfo

    with pikepdf.open(io.BytesIO(response.content), password=main.OWNER_PASSWORD) as pdf:
        assert pdf.owner_password_matched


def test_page_size_and_rotation_are_kept(client):
    response = upload(client, make_pdf(pages=1, size=(800, 400), rotate=90))
    document = pdfium.PdfDocument(response.content)
    assert document[0].get_size() == pytest.approx((400, 800), abs=0.1)


def test_watermark_marks_the_page(client):
    response = upload(client, make_pdf(pages=1))
    image = pdfium.PdfDocument(response.content)[0].render(scale=1).to_pil().convert("L")
    # The top of the source page is blank, so anything darker than paper is the watermark.
    top = image.crop((0, 0, image.width, 200))
    assert top.getextrema()[0] < 225


def test_high_quality_produces_more_pixels(client):
    source = make_pdf(pages=1)
    standard = upload(client, source, quality="standard")
    high = upload(client, source, quality="high")
    assert len(high.content) > len(standard.content)


def test_huge_page_is_capped(client):
    response = upload(client, make_pdf(pages=1, size=(5000, 5000)))
    with pikepdf.open(io.BytesIO(response.content)) as pdf:
        image = pdf.pages[0].Resources.XObject.Im0
        assert max(int(image.Width), int(image.Height)) <= protect.MAX_PAGE_SIDE_PX


@pytest.mark.parametrize(
    ("data", "name", "status", "message"),
    [
        (b"hello", "notes.txt", 400, "Upload a PDF"),
        (b"", "empty.pdf", 400, "empty"),
        (b"not a pdf", "fake.pdf", 400, "valid PDF"),
        (b"%PDF-1.7 broken", "broken.pdf", 400, "couldn't be opened"),
        (b"PK not a zip", "fake.docx", 400, "valid DOCX"),
        (b"not ole", "fake.ppt", 400, "valid PPT"),
    ],
)
def test_bad_uploads_get_clear_errors(client, data, name, status, message):
    response = upload(client, data, name)
    assert response.status_code == status
    assert message in response.json()["detail"]


def test_password_protected_pdf_is_rejected(client):
    buffer = io.BytesIO()
    with pikepdf.open(io.BytesIO(make_pdf())) as pdf:
        pdf.save(buffer, encryption=pikepdf.Encryption(owner="o", user="u"))
    response = upload(client, buffer.getvalue())
    assert response.status_code == 400
    assert "password-protected" in response.json()["detail"]


def test_unknown_quality_is_rejected(client):
    assert upload(client, make_pdf(), quality="ultra").status_code == 400


def test_page_limit(client, monkeypatch):
    monkeypatch.setattr(protect, "MAX_PAGES", 1)
    response = upload(client, make_pdf(pages=2))
    assert response.status_code == 400
    assert "limit is 1" in response.json()["detail"]


def test_multi_page_document_keeps_page_order_and_sizes(client):
    sizes = [(300, 400), (500, 200), (612, 792), (200, 200), (400, 300)]
    source = pikepdf.new()
    for size in sizes:
        source.add_blank_page(page_size=size)
    buffer = io.BytesIO()
    source.save(buffer)

    response = upload(client, buffer.getvalue())

    assert response.status_code == 200
    document = pdfium.PdfDocument(response.content)
    assert [document[i].get_size() for i in range(len(document))] == sizes
    with pikepdf.open(io.BytesIO(response.content)) as pdf:
        assert pdf.is_encrypted and not pdf.allow.extract
        assert pdf.check_pdf_syntax() == []


def test_size_limit(client, monkeypatch):
    monkeypatch.setattr(main, "MAX_REQUEST_BYTES", 10)
    assert upload(client, make_pdf()).status_code == 413


def test_unicode_filename_download_header(client):
    response = upload(client, make_pdf(pages=1), "résumé \"final\".pdf")
    disposition = response.headers["content-disposition"]
    assert "filename*=UTF-8''r%C3%A9sum%C3%A9" in disposition


def test_no_working_files_left_behind(client, tmp_path, monkeypatch):
    monkeypatch.setattr(main.tempfile, "tempdir", str(tmp_path))
    assert upload(client, make_pdf(pages=1)).status_code == 200
    assert upload(client, b"bad", "bad.pdf").status_code == 400
    assert list(tmp_path.iterdir()) == []


def test_office_file_without_converter_reports_it(client, monkeypatch):
    monkeypatch.setattr(protect, "_soffice_command", lambda: None)
    response = upload(client, make_office("word/"), "paper.docx")
    assert response.status_code == 503


@pytest.fixture
def fake_soffice(tmp_path, monkeypatch):
    """Stands in for LibreOffice: checks the command line, then writes a PDF like soffice does."""
    pdf_path = tmp_path / "converted-source.pdf"
    pdf_path.write_bytes(make_pdf(pages=3))
    env_log = tmp_path / "env.txt"
    script = tmp_path / "fake_soffice.py"
    script.write_text(
        textwrap.dedent(
            f"""
            import os, pathlib, shutil, sys
            args = sys.argv[1:]
            pathlib.Path({str(env_log)!r}).write_text("\\n".join(os.environ))
            assert "--headless" in args and args[args.index("--convert-to") + 1] == "pdf"
            outdir = pathlib.Path(args[args.index("--outdir") + 1])
            outdir.mkdir()
            shutil.copy({str(pdf_path)!r}, outdir / (pathlib.Path(args[-1]).stem + ".pdf"))
            """
        )
    )
    monkeypatch.setattr(protect, "_soffice_command", lambda: [sys.executable, str(script)])
    monkeypatch.setenv("ENCREEPT_OWNER_PASSWORD", "must-not-reach-the-converter")
    return env_log


@pytest.mark.parametrize(("name", "folder"), [("paper.docx", "word/"), ("slides.pptx", "ppt/")])
def test_office_files_go_through_the_converter(client, fake_soffice, name, folder):
    response = upload(client, make_office(folder), name)
    assert response.status_code == 200
    assert len(pdfium.PdfDocument(response.content)) == 3
    assert extract_text(response.content).strip() == ""
    assert "ENCREEPT_OWNER_PASSWORD" not in fake_soffice.read_text().split("\n")


def make_deck(slides: int, hidden: range = range(0)) -> bytes:
    deck = Presentation()
    for number in range(slides):
        slide = deck.slides.add_slide(deck.slide_layouts[1])
        slide.shapes.title.text = f"Slide {number + 1}"
        if number in hidden:
            slide._element.set("show", "0")
    buffer = io.BytesIO()
    deck.save(buffer)
    return buffer.getvalue()


def test_long_deck_is_split_in_slide_order(tmp_path, monkeypatch):
    monkeypatch.setattr(protect, "SLIDES_PER_CONVERSION", 4)
    source = tmp_path / "deck.pptx"
    source.write_bytes(make_deck(10))

    pieces = protect._split_deck(source, tmp_path)

    titles = [[s.shapes.title.text for s in Presentation(p).slides] for p in pieces]
    assert titles == [[f"Slide {n}" for n in group] for group in ([1, 2, 3, 4], [5, 6, 7, 8], [9, 10])]


@needs_libreoffice
def test_slide_numbers_continue_across_pieces(client, tmp_path, monkeypatch):
    monkeypatch.setattr(protect, "SLIDES_PER_CONVERSION", 2)
    deck = Presentation()
    for _ in range(5):
        slide = deck.slides.add_slide(deck.slide_layouts[1])
        paragraph = slide.shapes.title.text_frame.paragraphs[0]._p
        paragraph.append(
            etree.fromstring(
                f'<a:fld xmlns:a="{protect._DML_NS}" id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
                '<a:rPr lang="en-US"/><a:t>0</a:t></a:fld>'
            )
        )
    source = tmp_path / "numbered.pptx"
    deck.save(source)

    (tmp_path / "job").mkdir()
    pieces = protect._split_deck(source, tmp_path / "job")
    pdfs = [protect._convert_to_pdf(piece, tmp_path / "job") for piece in pieces]

    pages = []
    for path in pdfs:
        document = pdfium.PdfDocument(path)
        pages += [document[i].get_textpage().get_text_range().strip() for i in range(len(document))]
        document.close()
    assert pages == ["1", "2", "3", "4", "5"]


def test_short_or_unusual_decks_are_not_split(tmp_path):
    source = tmp_path / "deck.pptx"
    source.write_bytes(make_deck(protect.SLIDES_PER_CONVERSION))
    assert protect._split_deck(source, tmp_path) == [source]

    unusual = tmp_path / "unusual.pptx"
    unusual.write_bytes(make_office("ppt/"))
    assert protect._split_deck(unusual, tmp_path) == [unusual]


@needs_libreoffice
def test_real_conversion_of_long_deck_with_hidden_slides(client, monkeypatch):
    monkeypatch.setattr(protect, "SLIDES_PER_CONVERSION", 3)
    # Slides 4-6 are a whole hidden piece; slide 8 is hidden inside a piece.
    response = upload(client, make_deck(9, hidden=[3, 4, 5, 7]), "deck.pptx")
    assert response.status_code == 200
    assert len(pdfium.PdfDocument(response.content)) == 5
    assert extract_text(response.content).strip() == ""


@needs_libreoffice
def test_real_conversion_of_word_document(client):
    document = Document()
    document.add_paragraph(SECRET_TEXT)
    buffer = io.BytesIO()
    document.save(buffer)
    response = upload(client, buffer.getvalue(), "paper.docx")
    assert response.status_code == 200
    assert extract_text(response.content).strip() == ""


def test_converter_producing_nothing_is_an_error(client, tmp_path, monkeypatch):
    monkeypatch.setattr(protect, "_soffice_command", lambda: [sys.executable, "-c", "pass"])
    response = upload(client, make_office("word/"), "paper.docx")
    assert response.status_code == 400
    assert "couldn't be converted" in response.json()["detail"]


def test_converter_timeout(client, monkeypatch):
    monkeypatch.setattr(protect, "CONVERT_TIMEOUT_SECONDS", 1)
    monkeypatch.setattr(
        protect, "_soffice_command", lambda: [sys.executable, "-c", "import time; time.sleep(30)"]
    )
    response = upload(client, make_office("word/"), "paper.docx")
    assert response.status_code == 503
    assert "too long" in response.json()["detail"]
