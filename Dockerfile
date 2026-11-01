FROM python:3.12-slim

# LibreOffice converts Word and PowerPoint files. The extra fonts match the
# metrics of Calibri, Cambria, Arial and Times, so layouts don't shift.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-impress \
        fonts-liberation \
        fonts-crosextra-carlito \
        fonts-crosextra-caladea \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

ENV UV_NO_CACHE=1
RUN pip install --no-cache-dir uv

WORKDIR /srv
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev
COPY app ./app

RUN useradd --create-home encreept
USER encreept

EXPOSE 10000
CMD ["sh", "-c", "exec .venv/bin/uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-10000} --workers 1"]
