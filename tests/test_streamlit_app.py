"""Smoke test the actual deployed Streamlit entry point, not only backend modules."""
from streamlit.testing.v1 import AppTest
from pathlib import Path


def test_streamlit_homepage_loads_without_exception():
    at=AppTest.from_file(Path(__file__).resolve().parents[1] / "app.py",default_timeout=30).run()
    assert not at.exception, "\n".join(str(e.message) for e in at.exception)
    assert at.title
    assert "Underwriting" in at.title[0].value
    assert any("DealLens" in item.value for item in at.sidebar.title)
    assert len(at.tabs) == 0  # Workflow appears only after a CIM/demo is loaded

def test_branded_demo_workspace_loads_and_preserves_three_steps():
    at=AppTest.from_file(Path(__file__).resolve().parents[1] / "app.py",default_timeout=30).run()
    demo=next(b for b in at.sidebar.button if "fictional demo" in b.label.lower())
    demo.click().run()
    assert not at.exception, "\\n".join(str(e.message) for e in at.exception)
    assert any("Northstar" in item.value for item in at.title)
    assert len(at.tabs)==3
    assert "Build LBO" in at.tabs[0].label
    assert "Verify returns" in at.tabs[1].label
    assert "IC memo" in at.tabs[2].label


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

def test_provider_selector_changes_key_field_without_making_requests():
    at=AppTest.from_file(Path(__file__).resolve().parents[1] / "app.py",default_timeout=30).run()
    provider=next(item for item in at.sidebar.selectbox if item.label=="AI provider")
    assert tuple(provider.options)==("OpenAI","Claude","Gemini")
    provider.set_value("Claude").run()
    assert not at.exception
    assert any(i.label=="Claude API key" for i in at.sidebar.text_input)
    provider=next(item for item in at.sidebar.selectbox if item.label=="AI provider")
    provider.set_value("Gemini").run()
    assert not at.exception
    assert any(i.label=="Gemini API key" for i in at.sidebar.text_input)
