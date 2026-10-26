"""Test private template storage without publishing the actual deal workbook."""
import base64
import gzip
from io import BytesIO

import pytest
from openpyxl import Workbook
from deallens.template_storage import decode_template, get_saved_template, KEY


def example_encoded_template():
    wb = Workbook()
    wb.active["A1"] = "Fictional LBO"
    out = BytesIO()
    wb.save(out)
    return base64.b64encode(gzip.compress(out.getvalue())).decode()


def test_saved_template_decodes_only_when_configured(monkeypatch):
    monkeypatch.delenv(KEY, raising=False)
    assert get_saved_template({}) is None
    encoded = example_encoded_template()
    result = get_saved_template({KEY: encoded})
    assert result.startswith(b"PK")
    assert result == decode_template(encoded)


def test_environment_template_configuration(monkeypatch):
    encoded = example_encoded_template()
    monkeypatch.setenv(KEY, encoded)
    assert get_saved_template({KEY: "incorrect"}).startswith(b"PK")


def test_reject_invalid_or_truncated_template():
    with pytest.raises(ValueError):
        decode_template("garbage")
    bad = base64.b64encode(gzip.compress(b"not an XLSX")).decode()
    with pytest.raises(ValueError, match="Excel"):
        decode_template(bad)
