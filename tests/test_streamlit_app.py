"""Smoke test the actual deployed Streamlit entry point, not only backend modules."""
from streamlit.testing.v1 import AppTest
from pathlib import Path


def test_streamlit_homepage_loads_without_exception():
    at=AppTest.from_file(Path(__file__).resolve().parents[1] / "app.py",default_timeout=30).run()
    assert not at.exception, "\n".join(str(e.message) for e in at.exception)
    assert at.title
    assert "DealLens" in at.title[0].value


def test_deployment_imports_all_lbo_entrypoints():
    from deallens.full_model import (
        extract_financials, make_integrated_model, read_recalculated_returns,
        recommendation, HOLD_YEARS, EXACT_MOIC_FOR_TARGET_IRR
    )
    assert callable(extract_financials)
    assert callable(make_integrated_model)
    assert callable(read_recalculated_returns)
    assert callable(recommendation)
    assert HOLD_YEARS == 5
    assert EXACT_MOIC_FOR_TARGET_IRR > 3.0
