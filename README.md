# Encreept

A small FastAPI app that turns a PDF, Word or PowerPoint file into a PDF that can be read but not copied. Nothing is kept: working files are deleted as soon as the download has been sent.

## How it works

1. Word and PowerPoint files are converted to PDF with LibreOffice.
2. Each page is rendered to an image and stamped with a tiled Encreept watermark. This removes all text, fonts, links and original metadata.
3. The images are assembled into a new PDF at the original page size.
4. The PDF is encrypted with AES-256. It opens without a password, but copying, printing and editing are turned off and require the owner password.

Step 2 does the real work. PDF permissions are only honoured by well-behaved viewers, and some tools ignore them. Screenshots and text recognition remain possible, which is why the watermark is part of each image.

## Run locally

Requires [uv](https://docs.astral.sh/uv/). Word and PowerPoint uploads also need [LibreOffice](https://www.libreoffice.org/) installed; PDFs work without it.

```bash
uv sync
```

```bash
uv run uvicorn app.main:app --reload
```

Then open http://127.0.0.1:8000.

```bash
uv run pytest
```

## Deploy to Render

1. Push this folder to a GitHub repository.
2. In Render, choose **New > Blueprint** and select the repository. `render.yaml` sets up a free Docker web service and generates the owner password.

The free plan has 512 MB of memory and a fraction of a CPU, and it sleeps after 15 idle minutes. Expect slow first requests and slow Office conversions. Documents are processed one at a time.

## Settings

| Environment variable | Default | Purpose |
| --- | --- | --- |
| `ENCREEPT_OWNER_PASSWORD` | Random for each run | Master password for the output PDFs |
| `ENCREEPT_MAX_UPLOAD_MB` | `60` | Upload limit. The page's limit in `app/static/app.js` and its hint text need to match. |
| `ENCREEPT_MAX_PAGES` | `400` | Page limit. The page's hint text needs to match. |
| `ENCREEPT_CONVERT_TIMEOUT` | `900` | Seconds allowed for an Office conversion |
