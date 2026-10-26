"""DealLens: institutional-style Streamlit underwriting workspace."""
import os

import streamlit as st

from deallens.demo import DEMO_COMPANY, demo_memo
from deallens.ingest import extract_pdf_pages
from deallens.memo import analyze_cim, blank_memo
from deallens.full_model import (
    extract_financials,
    make_integrated_model,
    read_recalculated_returns,
    recommendation,
    HOLD_YEARS,
    EXACT_MOIC_FOR_TARGET_IRR,
)
from deallens.ic_brief import export_two_page_ic

st.set_page_config(
    page_title="DealLens | Private Markets Intelligence",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

defaults = {
    "company": "",
    "memo": blank_memo(),
    "financials": {"periods": [], "warnings": []},
    "issues": [],
    "model_bytes": None,
    "returns": None,
    "reviewed": False,
    "resolved_missing": False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

with st.sidebar:
    st.title("◈ DealLens")
    st.caption("PRIVATE MARKETS INTELLIGENCE")
    st.divider()

    st.subheader("Source document")
    doc = st.file_uploader(
        "Upload CIM",
        type=["pdf"],
        max_upload_size=25,
        help="Searchable PDF, maximum 25 MB. Use public, fictional, or expressly authorized documents only.",
    )

    try:
        secret = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        secret = ""
    key = st.text_input(
        "OpenAI API key",
        type="password",
        value=os.environ.get("OPENAI_API_KEY", "") or secret,
        help="Your key is used for CIM extraction. Never commit it to GitHub.",
    )
    if st.button("Analyze CIM", type="primary", use_container_width=True, icon=":material/analytics:"):
        if not doc or not key:
            st.error("Upload a searchable CIM and provide an OpenAI API key.")
        else:
            with st.spinner("Extracting the business profile and five-year management plan..."):
                try:
                    pages = extract_pdf_pages(doc.getvalue())
                    narrative = analyze_cim(pages, key)
                    data = extract_financials(pages, key)
                    st.session_state.update(
                        company=narrative["company_name"],
                        memo=narrative["memo"],
                        financials=data,
                        issues=[],
                        model_bytes=None,
                        returns=None,
                        reviewed=False,
                        resolved_missing=False,
                    )
                    st.rerun()
                except Exception as exc:
                    st.error(f"CIM analysis failed: {exc}")

    st.divider()
    st.caption("PREVIEW WORKSPACE")
    if st.button("Explore fictional demo", use_container_width=True, icon=":material/visibility:"):
        st.session_state.update(
            company=DEMO_COMPANY,
            memo=demo_memo(),
            financials={
                "periods": [],
                "warnings": [
                    "Narrative demo only. A searchable CIM with detailed financial statements is required to build a 14-tab LBO."
                ],
            },
            issues=[],
            model_bytes=None,
            returns=None,
            reviewed=False,
            resolved_missing=False,
        )
        st.rerun()
    st.caption("This is a prototype, not a secure data room. Do not submit restricted deal materials without authorization.")

# Brand line is intentionally not dependent on loading a CIM.
header_left, header_right = st.columns([4, 1], vertical_alignment="center")
with header_left:
    st.caption("DEALLENS  /  INVESTMENT UNDERWRITING")
with header_right:
    st.caption("MANAGEMENT CASE  ·  5-YEAR HOLD")

if not st.session_state.company:
    st.title("Underwriting, from first read to IC.")
    st.write(
        "Turn a confidential information memorandum into a structured investment screen, "
        "a fully linked Excel underwriting model, and a two-page investment committee brief."
    )
    st.space("small")
    with st.container(border=True):
        st.subheader("Start a new investment review")
        st.write(
            "Upload a searchable CIM using the left panel. DealLens extracts business facts and "
            "management forecasts, then prepares source-linked inputs for your 14-tab Excel model."
        )
        st.caption("The original LBO Excel template remains yours to upload. Its figures are never bundled in the public code.")

    st.subheader("A focused investment workflow")
    a, b, c = st.columns(3, gap="medium")
    with a:
        with st.container(border=True):
            st.caption("01  /  INGEST")
            st.markdown("#### Analyze the CIM")
            st.write("Business profile, management's historicals, and year-by-year projections.")
    with b:
        with st.container(border=True):
            st.caption("02  /  UNDERWRITE")
            st.markdown("#### Build the LBO")
            st.write("Your integrated 14-sheet Excel template, populated with management-plan inputs.")
    with c:
        with st.container(border=True):
            st.caption("03  /  DECIDE")
            st.markdown("#### Verify and export")
            st.write("Check five-year sponsor returns and download a concise IC memorandum.")

    st.space("small")
    p1, p2, p3 = st.columns(3)
    with p1:
        st.metric("Investment hold", f"{HOLD_YEARS} years")
    with p2:
        st.metric("IRR underwriting hurdle", "25.0%")
    with p3:
        st.metric("Equivalent five-year MOIC", f"{EXACT_MOIC_FOR_TARGET_IRR:.2f}x")
    st.divider()
    st.caption(
        "DEALLENS IS AN EXPERIMENTAL ANALYTICAL TOOL. Generated figures, valuations and "
        "recommendations require independent source verification. Public Community Cloud "
        "is not approved for confidential transaction processing."
    )
    st.stop()

st.title(st.session_state.company)
st.caption("ACTIVE INVESTMENT REVIEW  /  FIVE-YEAR MANAGEMENT PLAN")
with st.container(border=True):
    m1, m2, m3, m4 = st.columns(4, gap="medium")
    with m1:
        st.metric("CIM status", "Analyzed")
    with m2:
        st.metric("Financial periods", str(len(st.session_state.financials["periods"])))
    with m3:
        st.metric("Hold period", f"{HOLD_YEARS} years")
    with m4:
        st.metric("Return target", f"{EXACT_MOIC_FOR_TARGET_IRR:.2f}x / 25%")
st.caption("Single management case only. Historical figures and source-page references are kept in the Excel workbook.")

model_tab, review_tab, memo_tab = st.tabs(
    ["01  Build LBO", "02  Verify returns", "03  IC memo"]
)

with model_tab:
    left, right = st.columns([3, 2], gap="large")
    with left:
        st.subheader("Integrated LBO model")
        st.write(
            "Populate your original 14-tab Excel workbook with CIM historicals "
            "and management's five-year forecasts. The three-case toggle is removed."
        )
        original = st.file_uploader(
            "Excel model template (.xlsx)",
            type=["xlsx"],
            key="integrated_template",
        )
        periods = st.session_state.financials["periods"]
        future = sorted({p["year"] for p in periods if p["kind"] in ("estimate", "projection")})
        actual = sorted({p["year"] for p in periods if p["kind"] == "actual"})
        initial = (actual[-1] if actual else (future[0] - 1 if future else 2025))
        base = st.number_input(
            "Final historical / entry fiscal year",
            min_value=1990,
            max_value=2100,
            value=int(initial),
            help="The next five fiscal years must be CIM-supported management projections.",
        )
        if original and periods:
            if st.button("Generate integrated LBO", type="primary", icon=":material/table_chart:"):
                with st.spinner("Mapping source financials into the Excel schedules..."):
                    try:
                        data, issues = make_integrated_model(
                            original.getvalue(),
                            st.session_state.company,
                            st.session_state.financials,
                            base_year=int(base),
                        )
                        st.session_state.model_bytes = data
                        st.session_state.issues = issues
                        st.session_state.returns = None
                        st.session_state.reviewed = False
                        st.session_state.resolved_missing = False
                        st.success("All 14 model sheets have been generated from your template.")
                    except Exception as exc:
                        st.error(f"Unable to build model: {exc}")
        elif not periods:
            st.info("Detailed annual financial data is required. Analyze a searchable CIM to enable model generation.")
        else:
            st.caption("Upload the Excel template to enable model generation.")
    with right:
        with st.container(border=True):
            st.caption("MODEL STANDARD")
            st.markdown("#### CIM management plan")
            st.write("Five projected years, one operating case, full financial schedules.")
            st.divider()
            st.markdown("**14 integrated worksheets**")
            st.caption("Cover · Assumptions · Summary · Revenue · Expenses · D&A · NWC · Debt · IS · BS · CF · Returns · Comps · Cap table")
            st.markdown("**Five-year sponsor hurdle**")
            st.write(f"25% IRR / approximately {EXACT_MOIC_FOR_TARGET_IRR:.2f}x MOIC")
            st.caption("Financial figures are unverified until the completed Excel workbook is independently checked.")

    if st.session_state.model_bytes:
        with st.container(border=True):
            st.markdown("#### Workbook ready for analyst review")
            st.download_button(
                "Download 14-tab Excel LBO",
                data=st.session_state.model_bytes,
                file_name="DealLens_Integrated_LBO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                icon=":material/download:",
            )
            st.caption(
                "Open in Microsoft Excel, correct source-data gaps, recalculate all formulas, "
                "save, and return to the Verify returns tab."
            )
    if st.session_state.issues:
        with st.expander(f"Model review items ({len(st.session_state.issues)})", expanded=False):
            for issue in st.session_state.issues:
                st.write("• " + issue)

with review_tab:
    st.subheader("Returns and model verification")
    st.write(
        "Check the recalculated Excel workbook before using any returns "
        "in an investment recommendation or committee memorandum."
    )
    with st.container(border=True):
        completed = st.file_uploader(
            "Recalculated 14-tab Excel workbook (.xlsx)",
            type=["xlsx"],
            key="completed_model",
        )
        if completed:
            try:
                moic, irr = read_recalculated_returns(
                    completed.getvalue(), st.session_state.company
                )
                st.session_state.returns = (moic, irr)
                r1, r2 = st.columns(2)
                with r1:
                    st.metric("Five-year sponsor MOIC", f"{moic:.2f}x")
                with r2:
                    st.metric("Five-year sponsor IRR", f"{irr:.1%}")
            except Exception as exc:
                st.session_state.returns = None
                st.error(str(exc))
        else:
            st.caption("Recalculate and save the downloaded LBO in Microsoft Excel, then upload it here.")

    st.markdown("#### Analyst sign-off")
    st.session_state.reviewed = st.checkbox(
        "I independently reconciled CIM figures, financing assumptions, debt and capitalization "
        "schedules, and the Excel returns calculations.",
        value=st.session_state.reviewed,
    )
    numeric_blockers = [
        x for x in st.session_state.issues
        if x.startswith("Missing ") or "no CIM forecast revenue" in x or "needs COGS" in x
    ]
    resolved = st.checkbox(
        "I resolved every previously missing financial input in the recalculated workbook "
        "and verified that all financial reconciliation checks pass.",
        key="resolved_missing",
    )
    if numeric_blockers and not resolved:
        st.error(
            "Missing inputs still block a positive recommendation. Correct the workbook, "
            "recalculate it in Excel, and verify the resolved items."
        )

    effective_blockers = (
        [] if resolved and st.session_state.returns and st.session_state.reviewed
        else numeric_blockers
    )
    moic, irr = st.session_state.returns or (None, None)
    state, explanation = recommendation(
        moic, irr, st.session_state.reviewed, effective_blockers
    )
    with st.container(border=True):
        st.caption("INVESTMENT SCREENING STATUS")
        st.markdown("#### " + state)
        st.write(explanation)
        st.caption(
            f"Fixed five-year hold. Both independently verified returns must reconcile "
            f"to at least 25% IRR (approximately {EXACT_MOIC_FOR_TARGET_IRR:.2f}x MOIC)."
        )

with memo_tab:
    st.subheader("Investment committee brief")
    st.write(
        "A concise, two-page Word memorandum covering the business, historical financials, "
        "industry, growth, merits and risks, investment thesis, potential exit paths, and recommendation."
    )
    moic, irr = st.session_state.returns or (None, None)
    blockers = [
        x for x in st.session_state.issues
        if x.startswith("Missing ") or "no CIM forecast revenue" in x or "needs COGS" in x
    ]
    if (
        st.session_state.get("resolved_missing")
        and st.session_state.reviewed
        and st.session_state.returns
    ):
        blockers = []
    status, _ = recommendation(moic, irr, st.session_state.reviewed, blockers)
    with st.container(border=True):
        st.caption("WORD DELIVERABLE  /  TWO PAGES")
        st.markdown("#### " + st.session_state.company + " | Initial investment screen")
        st.write("**Recommendation:** " + status)
        st.caption(
            "The memo uses CIM-derived narrative. A conditional proceed recommendation "
            "requires verified five-year sponsor returns."
        )
        memo_bytes = export_two_page_ic(
            st.session_state.company,
            st.session_state.memo,
            moic,
            irr,
            st.session_state.reviewed,
            blockers,
        )
        st.download_button(
            "Download investment committee memo",
            data=memo_bytes,
            file_name="DealLens_Investment_Committee_Memo.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            icon=":material/download:",
        )

st.divider()
st.caption(
    "DEALLENS  /  RESEARCH PROTOTYPE  •  Documents and financial outputs require "
    "analyst validation. This publicly hosted application is not a confidential data room."
)
