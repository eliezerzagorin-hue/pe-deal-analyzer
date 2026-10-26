"""Test private template storage without publishing the actual deal workbook."""
import base64
import gzip
from io import BytesIO

import pytest
from openpyxl import Workbook
from deallens.template_storage import decode_template, get_saved_template, validate_integrated_template, REQUIRED_SHEETS, KEY


def example_encoded_template():
    wb = Workbook()
    wb.active.title = REQUIRED_SHEETS[0]
    for sheet in REQUIRED_SHEETS[1:]:
        wb.create_sheet(sheet)
    wb["ASSUMPTIONS"]["F18"] = 7.0
    wb["REV"]["E9"] = 10000
    out = BytesIO()
    wb.save(out)
    return base64.b64encode(gzip.compress(out.getvalue())).decode()


def test_saved_template_decodes_only_when_configured(monkeypatch):
    monkeypatch.delenv(KEY, raising=False)
    assert get_saved_template({}) is None
    encoded = example_encoded_template()
    result = get_saved_template({"lbo_template": {"gzip_base64": encoded}})
    assert result.startswith(b"PK")
    assert result == decode_template(encoded)
    assert validate_integrated_template(result)["sheet_count"] == 14


def test_environment_template_configuration(monkeypatch):
    encoded = example_encoded_template()
    monkeypatch.setenv(KEY, encoded)
    assert get_saved_template({"lbo_template": {"gzip_base64": "incorrect"}}).startswith(b"PK")


def test_reject_invalid_or_truncated_template():
    with pytest.raises(ValueError):
        decode_template("garbage")
    bad = base64.b64encode(gzip.compress(b"not an XLSX")).decode()
    with pytest.raises(ValueError, match="Excel"):
        decode_template(bad)

def test_reject_other_excel_workbook_instead_of_claiming_template_recognized():
    wb=Workbook()
    wb.active["A1"]="Not the 14-tab template"
    b=BytesIO()
    wb.save(b)
    with pytest.raises(ValueError, match="not the original 14-tab model"):
        validate_integrated_template(b.getvalue())


def test_reject_correct_tabs_when_input_cells_are_formulas():
    from openpyxl import load_workbook
    b=BytesIO(example_decoded_template())
    wb=load_workbook(b)
    wb["REV"]["E9"]="=SUM(A1:A4)"
    o=BytesIO()
    wb.save(o)
    with pytest.raises(ValueError, match="REV!E9"):
        validate_integrated_template(o.getvalue())


def example_decoded_template():
    return gzip.decompress(base64.b64decode(example_encoded_template()))
