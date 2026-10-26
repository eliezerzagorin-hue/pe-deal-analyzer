"""DealLens: upload CIM + user's 14-tab template, export linked workbook and 2-page IC brief."""
import os

import streamlit as st
from deallens.demo import DEMO_COMPANY, DEMO_HISTORY, demo_memo
from deallens.ingest import extract_pdf_pages
from deallens.memo import analyze_cim, blank_memo
from deallens.full_model import extract_financials, make_integrated_model, read_recalculated_returns, recommendation, HOLD_YEARS, EXACT_MOIC_FOR_TARGET_IRR
from deallens.ic_brief import export_two_page_ic

st.set_page_config(page_title="DealLens | Buyout Underwriting", page_icon="📊", layout="wide")
st.title("DealLens")
st.caption("Full 14-tab financial model + two-page investment committee memo")
st.warning("Prototype for fictional/public or expressly authorized documents only. This is NOT a secure data room. CIM extractions need source verification.", icon="⚠️")

defaults={"company":"","memo":blank_memo(),"financials":{"periods":[],"warnings":[]},"issues":[],
          "model_bytes":None,"returns":None,"reviewed":False}
for k,v in defaults.items():
    if k not in st.session_state:st.session_state[k]=v

with st.sidebar:
    st.subheader("Step 1 · Source CIM")
    doc=st.file_uploader("Searchable CIM (.pdf)",type=["pdf"],max_upload_size=25)
    try: secret=st.secrets.get("OPENAI_API_KEY","")
    except Exception: secret=""
    key=st.text_input("OpenAI API key",type="password",value=os.environ.get("OPENAI_API_KEY","") or secret,
                      help="Keep credentials out of GitHub. Only your submitted CIM text is sent for extraction.")
    if st.button("Analyze CIM",type="primary",use_container_width=True):
        if not doc or not key:
            st.error("Upload a searchable PDF and enter an API key.")
        else:
            with st.spinner("Reading CIM and extracting financial periods and memo contents..."):
                try:
                    pages=extract_pdf_pages(doc.getvalue())
                    narrative=analyze_cim(pages,key)
                    data=extract_financials(pages,key)
                    st.session_state.update(company=narrative["company_name"],memo=narrative["memo"],
                        financials=data,issues=[],model_bytes=None,returns=None,reviewed=False,resolved_missing=False)
                    st.rerun()
                except Exception as exc: st.error(f"Could not analyze the CIM: {exc}")
    st.divider()
    if st.button("Load fictional narrative demo",use_container_width=True):
        st.session_state.update(company=DEMO_COMPANY,memo=demo_memo(),
            financials={"periods":[],"warnings":["Illustrative narrative only. A CIM with annual financial statements is required for a complete 14-tab underwriting model."]},
            issues=[],model_bytes=None,returns=None,reviewed=False,resolved_missing=False)
        st.rerun()

if not st.session_state.company:
    st.info("Upload a searchable CIM in the sidebar, then analyze it to prepare a full integrated model and Word IC memo.")
    st.stop()

st.subheader(st.session_state.company)
c1,c2,c3,c4=st.columns(4)
c1.metric("Document","CIM analyzed")
c2.metric("Financial periods",len(st.session_state.financials["periods"]))
c3.metric("Hold period",f"{HOLD_YEARS} years")
c4.metric("Return hurdle",f"≥{EXACT_MOIC_FOR_TARGET_IRR:.2f}x / ≥25%")
st.caption("One management case only. No scenario toggle or analyst-modified operating projections.")
st.caption("Historical sources are written into the Excel model as source-page comments and a SUMMARY source-audit log, not displayed as a separate screen.")
model_tab,memo_tab,review_tab=st.tabs(["1 · Full 14-tab LBO","2 · Two-page IC memo","3 · Verification"])

with model_tab:
    st.markdown("### Five-year management-case LBO")
    st.write("Upload your **14-tab .xlsx template** each time. DealLens preserves the financial statement, debt and returns schedules, but deletes the management/base/downside case toggle. All operating projections use only CIM management figures.")
    st.caption("Only the five CIM management projection years feed the model. Select the final actual year before those projections. If any management forecast year is missing, model generation stops rather than extrapolating or inheriting old template figures. Five-year hold is fixed; 25% IRR requires about 3.05x MOIC.")
    original=st.file_uploader("Integrated 14-tab Excel LBO template (.xlsx)",type=["xlsx"],key="integrated_template")
    periods=st.session_state.financials["periods"]
    future=sorted({p["year"] for p in periods if p["kind"] in ("estimate","projection")})
    actual=sorted({p["year"] for p in periods if p["kind"]=="actual"})
    initial=(actual[-1] if actual else (future[0]-1 if future else 2025))
    base=st.number_input("Base / entry financial year (FY)",min_value=1990,max_value=2100,value=int(initial),
                         help="Use the last actual fiscal year. The next five fiscal years must be directly supported by the CIM management plan.")
    if not periods:st.warning("No full CIM financial periods extracted. Analyze a CIM containing detailed financials to build a usable underwriting model.")
    if original and periods and st.button("Generate 14-tab LBO",type="primary"):
        try:
            data,issues=make_integrated_model(original.getvalue(),st.session_state.company,
                                           st.session_state.financials,base_year=int(base))
            st.session_state.model_bytes=data
            st.session_state.issues=issues
            st.session_state.returns=None
            st.session_state.reviewed=False
            st.session_state.resolved_missing=False
            st.success("Generated all 14 worksheets using the uploaded workbook's structure.")
        except Exception as exc:
            st.error(f"Unable to build model: {exc}")
    if st.session_state.model_bytes:
        st.download_button("Download full 14-tab LBO (.xlsx)",data=st.session_state.model_bytes,
           file_name="DealLens_Integrated_LBO.xlsx",
           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
           type="primary")
        st.info("Open this file in Microsoft Excel, review the highlighted/source-commented inputs, fill outstanding diligence items, recalculate (Ctrl+Alt+F9), then SAVE. Reupload the recalculated workbook under Verification.")
    if st.session_state.issues:
        with st.expander(f"Model review items ({len(st.session_state.issues)})",expanded=True):
            for issue in st.session_state.issues:st.write("• "+issue)

with review_tab:
    st.markdown("### Verify the completed model")
    completed=st.file_uploader("Recalculated and saved 14-tab Excel workbook (.xlsx)",type=["xlsx"],key="completed_model")
    if completed:
        try:
            moic,irr=read_recalculated_returns(completed.getvalue(),st.session_state.company)
            st.session_state.returns=(moic,irr)
            r1,r2=st.columns(2)
            r1.metric("Sponsor MOIC",f"{moic:.2f}x")
            r2.metric("Sponsor IRR",f"{irr:.1%}")
        except Exception as exc:
            st.session_state.returns=None
            st.error(str(exc))
    st.session_state.reviewed=st.checkbox(
        "I independently verified the CIM figures, completed all missing model assumptions, reviewed debt and capitalization terms, and checked the Excel returns calculations.",
        value=st.session_state.reviewed)
    numeric_blockers=[x for x in st.session_state.issues if
        x.startswith("Missing ") or "no CIM forecast revenue" in x or "needs COGS" in x]
    resolved=st.checkbox("I also corrected or supplied every previously missing numerical input in the recalculated Excel workbook and confirmed all reconciliation checks pass.",key="resolved_missing")
    if numeric_blockers and not resolved:
        st.error("Unresolved missing/unsupported inputs block a proceed recommendation. Update, recalculate and reupload Excel, then explicitly confirm you resolved them.")
    effective_blockers=[] if (resolved and st.session_state.returns and st.session_state.reviewed) else numeric_blockers
    moic,irr=st.session_state.returns or (None,None)
    state,explanation=recommendation(moic,irr,st.session_state.reviewed,effective_blockers)
    st.subheader("Decision: "+state)
    st.write(explanation)
    st.caption(f"Strict five-year gate: ≥25% sponsor IRR, requiring ≈{EXACT_MOIC_FOR_TARGET_IRR:.2f}x MOIC. Returns must reconcile to the five-year hold. No threshold override.")

with memo_tab:
    st.markdown("### Download the two-page IC memo")
    st.write("Professional, decision-first Word memorandum. Includes your eight sections, concise business overview, merits/risks, thesis, exit strategy, returns and a gated recommendation.")
    st.caption("Long CIM sections are condensed to fit two pages. Original claims remain subject to verification.")
    moic,irr=st.session_state.returns or (None,None)
    blockers=[x for x in st.session_state.issues if
        x.startswith("Missing ") or "no CIM forecast revenue" in x or "needs COGS" in x]
    if st.session_state.get("resolved_missing") and st.session_state.reviewed and st.session_state.returns:
        blockers=[]
    status,_=recommendation(moic,irr,st.session_state.reviewed,blockers)
    st.write("**Current recommendation:** "+status)
    memo_bytes=export_two_page_ic(st.session_state.company,st.session_state.memo,moic,irr,
                                  st.session_state.reviewed,blockers)
    st.download_button("Download two-page IC memo (.docx)",data=memo_bytes,
         file_name="DealLens_Investment_Committee_Memo.docx",
         mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
         type="primary")

st.divider()
st.caption("Privacy: no CIM, generated workbook or template is committed to GitHub. Hosting and AI provider security must be approved by your organization before processing sensitive deals.")
