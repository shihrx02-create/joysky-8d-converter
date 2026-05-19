FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8787

WORKDIR /app

# LibreOffice is used only to convert legacy .doc uploads into .docx before parsing.
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

COPY . /app

RUN mkdir -p /app/uploads /app/outputs

EXPOSE 8787

CMD ["sh", "-c", "python app.py --serve --host 0.0.0.0 --port ${PORT:-8787}"]
