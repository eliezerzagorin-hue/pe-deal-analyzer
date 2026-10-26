"""Private, reusable LBO template loader for Streamlit Community Cloud.

Never commit the underlying workbook or its encoded data to the public repo.
Set LBO_TEMPLATE_GZIP_BASE64 once in the Streamlit app's Secrets configuration.
"""
import base64
import gzip
import io
import os
import zipfile

MAX_COMPRESSED_BYTES = 8_000_000
MAX_WORKBOOK_BYTES = 15_000_000
KEY = "LBO_TEMPLATE_GZIP_BASE64"


def decode_template(encoded):
    """Validate and extract a user-managed encrypted-at-rest Streamlit secret.

    The payload is gzip-compressed for convenience, NOT cryptographically
    encrypted; only Streamlit's secret storage protects it.
    """
    if not isinstance(encoded, str) or not encoded.strip():
        raise ValueError("Saved LBO template secret is empty.")
    if len(encoded) > MAX_COMPRESSED_BYTES * 2:
        raise ValueError("Saved LBO template setting exceeds the size limit.")
    try:
        compressed = base64.b64decode(encoded.strip(), validate=True)
        if len(compressed) > MAX_COMPRESSED_BYTES:
            raise ValueError("Compressed LBO template is too large.")
        with gzip.GzipFile(fileobj=io.BytesIO(compressed), mode="rb") as stream:
            content = stream.read(MAX_WORKBOOK_BYTES + 1)
    except (OSError, ValueError) as exc:
        raise ValueError("Saved LBO template is not valid gzip/base64 data.") from exc
    if len(content) > MAX_WORKBOOK_BYTES:
        raise ValueError("Saved LBO workbook is too large.")
    if not zipfile.is_zipfile(io.BytesIO(content)):
        raise ValueError("Saved template does not contain an Excel .xlsx workbook.")
    return content


def get_saved_template(secrets=None):
    """Use Streamlit Secrets or a runtime environment secret, not repository files."""
    data = os.environ.get(KEY)
    if not data and secrets is not None:
        try:
            data = secrets.get(KEY)
        except (KeyError, FileNotFoundError, RuntimeError):
            data = None
    if not data:
        return None
    return decode_template(data)
