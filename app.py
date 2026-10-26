import os
import streamlit as st
import pandas as pd
from deallens.demo import DEMO_COMPANY, DEMO_HISTORY, demo_memo
from deallens.ingest import extract_pdf_pages
from deallens.memo import analyze_cim, blank_memo, FIELD_GROUPS
from deallens.lbo import Assumptions, calculate_lbo
from deallens.exports import export_memo, export_lbo

st.set_page_config(page_title="DealLens | Private Equity", page_icon="📊", layout="wide")
st.title("DealLens")
st.caption("CIM to investment committee memo and adjustable five-year LBO")
st.warning("Prototype. Use fictional/public documents only. Verify all extracted claims and financial figures.")
if "company" not in st.session_state:
    st.session_state.update(company="", memo=blank_memo(), history=[], notes=[], verified=False)
with st.sidebar:
    st.header("Get started")
    doc = st.file_uploader("Searchable CIM PDF", type="pdf")
    try: secret = st.secrets.get("OPENAI_API_KEY", "")
    except Exception: secret = ""
    key = st.text_input("OpenAI API key", type="password", value=os.environ.get("OPENAI_API_KEY", "") or secret)
    if st.button("Load fictional demo"):
        st.session_state.update(company=DEMO_COMPANY,memo=demo_memo(),history=[dict(x) for x in DEMO_HISTORY],notes=["All demonstration data are fictional."],verified=False)
        st.rerun()
    if st.button("Analyze CIM", type="primary"):
        if not doc or not key: st.error("Upload a searchable PDF and provide an API key.")
        else:
            try:
                with st.spinner("Analyzing CIM..."):
                    result=analyze_cim(extract_pdf_pages(doc.getvalue()), key)
                st.session_state.update(company=result["company_name"],memo=result["memo"],history=result["historical"],notes=result["extraction_notes"],verified=False)
                st.rerun()
            except Exception as e: st.error(f"Unable to analyze: {e}")
if not st.session_state.company:
    st.info("Load the fictional demo or analyze your own searchable PDF.")
    st.stop()
st.subheader(st.session_state.company)
memo_tab, history_tab, lbo_tab, notes_tab=st.tabs(["IC memo","Historical figures","LBO model","Review notes"])
with memo_tab:
    st.caption("All sections editable. AI-generated source references require manual checking.")
    for section, fields in FIELD_GROUPS:
        st.subheader(section)
        for name,label in fields:
            st.session_state.memo[name]=st.text_area(label, value=st.session_state.memo.get(name,""), height=135, key="memo_"+name)
    st.download_button("Download Word IC memo",export_memo(st.session_state.company,st.session_state.memo),"DealLens_IC_Memo.docx")
with history_tab:
    st.caption("USD millions. Check values, units, fiscal periods, and adjusted vs reported EBITDA.")
    frame=pd.DataFrame(st.session_state.history,columns=["year","revenue","ebitda","adjusted_ebitda","source_page","notes"])
    updated=st.data_editor(frame,num_rows="dynamic",hide_index=True,use_container_width=True)
    st.session_state.history=updated.where(pd.notnull(updated),None).to_dict("records")
    st.session_state.verified=st.checkbox("I checked model data against the original sources",value=st.session_state.verified)
with lbo_tab:
    valid=[x for x in st.session_state.history if x.get("year") is not None and x.get("revenue") is not None and x.get("ebitda") is not None]
    latest=max(valid,key=lambda x:float(x["year"])) if valid else None
    st.caption("USD millions. Analyst assumptions, not CIM-confirmed projections.")
    if latest: st.info(f"Most recent extracted year: {int(latest['year'])}. Reported EBITDA and revenue are UNVERIFIED until checked.")
    with st.form("assumptions"):
        c1,c2,c3=st.columns(3)
        with c1:
            rev=st.number_input("Entry revenue ($M)",0.01,value=float(latest["revenue"]) if latest else 24.5)
            ebitda=st.number_input("Entry reported EBITDA ($M)",0.01,value=float(latest["ebitda"]) if latest else 3.9)
            entry=st.number_input("Entry EV/EBITDA",0.5,30.0,7.0,step=0.25)
            debt=st.number_input("Debt/EBITDA",0.0,20.0,3.0,step=0.25)
            fee=st.number_input("Fees (% EV)",0.0,30.0,2.0)/100
        with c2:
            growth=st.number_input("Annual revenue growth (%)",-90.0,100.0,5.0)/100
            margin=st.number_input("Entry EBITDA margin (%)",0.1,100.0,round(ebitda/rev*100,2))/100
            marginchange=st.number_input("Annual margin change (pp)",-20.0,20.0,0.5)/100
            rate=st.number_input("Interest rate (%)",0.0,100.0,9.0)/100
            tax=st.number_input("Cash tax (%)",0.0,100.0,25.0)/100
        with c3:
            da=st.number_input("D&A / revenue (%)",0.0,100.0,2.0)/100
            capex=st.number_input("Capex / revenue (%)",0.0,100.0,3.0)/100
            nwc=st.number_input("NWC / change in revenue (%)",0.0,100.0,10.0)/100
            exitmult=st.number_input("Exit EV/EBITDA",0.5,30.0,7.0,step=0.25)
            years=st.number_input("Hold (years)",1,10,5)
        st.form_submit_button("Recalculate")
    try:
        a=Assumptions(entry_revenue=rev,entry_ebitda=ebitda,entry_multiple=entry,debt_multiple=debt,deal_fees_pct=fee,revenue_growth=growth,entry_margin=margin,annual_margin_change=marginchange,interest_rate=rate,tax_rate=tax,da_pct_revenue=da,capex_pct_revenue=capex,nwc_pct_revenue=nwc,exit_multiple=exitmult,hold_years=years)
        results=calculate_lbo(a)
        k1,k2,k3,k4=st.columns(4)
        k1.metric("Entry EV",f"${results['entry_ev']:.1f}M")
        k2.metric("Sponsor equity",f"${results['initial_equity']:.1f}M")
        k3.metric("MOIC",f"{results['moic']:.2f}x")
        k4.metric("IRR",f"{results['irr']:.1%}" if results['irr'] is not None else "N/A")
        st.dataframe(pd.DataFrame(results["projection"]).round(3),hide_index=True)
        st.download_button("Download Excel LBO",export_lbo(a,results),"DealLens_LBO.xlsx")
        if not st.session_state.verified: st.warning("Underlying data not marked verified.")
    except ValueError as exc: st.error(str(exc))
with notes_tab:
    for note in st.session_state.notes: st.write("- "+note)
    st.caption("No independent market research or independently verified buyer universe is included.")
