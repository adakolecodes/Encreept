import asyncio
import logging
import os
import secrets
import shutil
import tempfile
from collections.abc import Callable
from functools import partial
from pathlib import Path
from urllib.parse import quote

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from app.protect import ConverterError, ImageInsert, ProtectionError, insert_images, protect_document

logger = logging.getLogger("encreept")

MAX_UPLOAD_MB = int(os.environ.get("ENCREEPT_MAX_UPLOAD_MB", "60"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
# Room for the multipart boundaries and form fields around the file.
MAX_REQUEST_BYTES = MAX_UPLOAD_BYTES + 64 * 1024

OWNER_PASSWORD = os.environ.get("ENCREEPT_OWNER_PASSWORD") or ""
if not OWNER_PASSWORD:
    OWNER_PASSWORD = secrets.token_urlsafe(32)
    logger.warning("ENCREEPT_OWNER_PASSWORD is not set; using a random one for this run.")

STATIC_DIR = Path(__file__).parent / "static"

app = FastAPI(title="Encreept", docs_url=None, redoc_url=None, openapi_url=None)

# One document at a time: the Render free plan has 512 MB of memory and pdfium isn't thread-safe.
processing_slot = asyncio.Semaphore(1)


@app.middleware("http")
async def limit_upload_size(request: Request, call_next):
    if request.method == "POST":
        length = request.headers.get("content-length", "")
        if not length.isdigit():
            return JSONResponse({"detail": "The upload is missing its size."}, status_code=411)
        if int(length) > MAX_REQUEST_BYTES:
            return JSONResponse({"detail": f"Uploads are limited to {MAX_UPLOAD_MB} MB in total."}, status_code=413)
    return await call_next(request)


def _download_name(filename: str) -> str:
    stem = Path(filename.replace("\\", "/")).stem or "document"
    return f"{stem}-encreept.pdf"


@app.get("/healthz", include_in_schema=False)
def healthz() -> dict:
    return {"status": "ok"}


@app.exception_handler(RequestValidationError)
async def missing_form_fields(request: Request, error: RequestValidationError) -> JSONResponse:
    return JSONResponse({"detail": "Part of the upload is missing. Reload the page and try again."}, status_code=400)


@app.post("/api/protect")
async def protect(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    quality: str = Form("standard"),
    watermark: str = Form(""),
) -> Response:
    if file.size == 0:
        raise HTTPException(status_code=400, detail="This file is empty.")
    job = partial(protect_document, file.file, file.filename or "", quality, watermark, OWNER_PASSWORD)
    return await _run_job(background_tasks, file.filename or "", job)


@app.post("/api/insert")
async def insert(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    image1: UploadFile = File(...),
    position1: str = Form(""),
    page1: str = Form(""),
    image2: UploadFile | None = File(None),
    position2: str = Form(""),
    page2: str = Form(""),
    quality: str = Form("standard"),
    watermark: str = Form(""),
) -> Response:
    images = [ImageInsert(image1.file, image1.filename or "", position1, page1)]
    if image2 is not None and image2.filename:
        images.append(ImageInsert(image2.file, image2.filename, position2, page2))
    job = partial(insert_images, file.file, images, quality, watermark, OWNER_PASSWORD)
    return await _run_job(background_tasks, file.filename or "", job)


async def _run_job(background_tasks: BackgroundTasks, filename: str, job: Callable[[Path], Path]) -> Response:
    """Run a pipeline in its own working folder and send the PDF it makes."""
    # Removed once the download has been sent, or straight away if anything fails.
    workdir = tempfile.mkdtemp(prefix="encreept-")
    try:
        async with processing_slot:
            try:
                result = await run_in_threadpool(job, Path(workdir))
            except ProtectionError as error:
                raise HTTPException(status_code=400, detail=str(error)) from None
            except ConverterError as error:
                raise HTTPException(status_code=503, detail=str(error)) from None
            except Exception:
                # The traceback goes to the logs, but never the document's name or contents.
                logger.exception("Protecting a document failed")
                raise HTTPException(
                    status_code=500, detail="This document couldn't be protected. Try a different file."
                ) from None
    except BaseException:
        shutil.rmtree(workdir, ignore_errors=True)
        raise

    background_tasks.add_task(shutil.rmtree, workdir, ignore_errors=True)
    name = _download_name(filename)
    ascii_name = name.encode("ascii", "ignore").decode().replace('"', "") or "document-encreept.pdf"
    return FileResponse(
        result,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(name)}",
            "Cache-Control": "no-store",
            "X-Content-Type-Options": "nosniff",
        },
    )


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
