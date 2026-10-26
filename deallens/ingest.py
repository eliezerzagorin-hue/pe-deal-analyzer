"""Ephemeral, text-only PDF ingestion for CIMs."""
from typing import Any

MAX_PDF_BYTES = 25 * 1024 * 1024
MAX_PAGES = 160


def extract_pdf_pages(file_bytes: bytes) -> list[dict[str, Any]]:
    import fitz

    if not file_bytes:
        raise ValueError("Upload a nonempty PDF.")
    if len(file_bytes) > MAX_PDF_BYTES:
        raise ValueError("PDF exceeds the 25 MB prototype limit.")
    try:
        pdf = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as exc:
        raise ValueError("Unable to read this PDF.") from exc
    try:
        if pdf.needs_pass:
            raise ValueError("Password-protected PDFs are not supported.")
        if len(pdf) > MAX_PAGES:
            raise ValueError(f"PDF exceeds the {MAX_PAGES}-page prototype limit.")
        pages = [
            {"page": i + 1, "text": pdf.load_page(i).get_text("text", sort=True)}
            for i in range(len(pdf))
        ]
    finally:
        pdf.close()
    if sum(len(p["text"].strip()) for p in pages) < 250:
        raise ValueError("Little searchable text found. Scanned PDFs need OCR, not yet supported.")
    return pages
