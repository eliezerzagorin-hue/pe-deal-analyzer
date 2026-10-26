# DealLens

**CIM-to-IC memo and adjustable five-year LBO prototype for lower-middle-market private equity.**

DealLens is an experimental Streamlit application which combines CIM extraction, a full uploaded 14-sheet LBO template and a concise two-page downloadable Word IC screening memorandum.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Windows: activate with `.venv\Scripts\activate`.

Without an OpenAI API key, click **Load fictional demo** to explore the model and exports. To analyze a real PDF, enter your own key in the app sidebar or configure `OPENAI_API_KEY` as a Streamlit secret.

## Deploy free

1. Sign in to [Streamlit Community Cloud](https://share.streamlit.io/).
2. Choose **Create app** and select `eliezerzagorin-hue/pe-deal-analyzer`, branch `main`, entrypoint `app.py`.
3. In advanced settings, add `OPENAI_API_KEY = "..." ` only if you want server-side AI enabled.
4. Deploy. Hosting is free on Community Cloud's eligible tier, while API usage is billed separately.


## Current 14-sheet integrated LBO and 2-page IC brief workflow

1. Upload a **searchable PDF CIM** and provide your own OpenAI API key.
2. Select **Analyze CIM**. The app extracts the eight IC memo narratives and a separate structured annual financial dataset in USD millions; missing and inconsistent fields remain unverified.
3. Upload your own compatible **14-tab `.xlsx` LBO template** in the **Full 14-tab LBO** screen. DealLens does not ship any customer's workbook or confidential information in the GitHub repository.
4. Choose the **latest actual financial year** preceding five CIM management forecast years. Generate and download a `.xlsx` containing the **same 14 sheets** as the template: COVER, ASSUMPTIONS, SUMMARY, REV, EXP, D&A, NWC, DEBT, IS, BS, CF, R, COMPS, CAP. The model populates a **single CIM management forecast case**, directly from five management-projected fiscal years, with the company name on the cover and no operating-case selector. Input-cell comments and SUMMARY source audit record CIM source pages; old financial data and comparables are cleared.
5. **Review all cells, complete unavailable inputs and transaction assumptions manually**, and open the workbook in Excel to recalculate. Keep the model's original integrated formulas, but confirm they make sense for the new business and capital structure. This app does not guarantee all financial details are extracted from a CIM.
6. Reupload the **Excel-recalculated workbook** in Verification to read sponsor MOIC from `R!E21` and IRR from `R!E22`, confirm assumptions, and download the **two-page formatted Word IC memo**.

**Five-year investment decision policy:** The model is locked to an exactly **5-year hold** from December 31 of the entry year to December 31 five years later. A 3.00x sponsor MOIC implies 24.57% annualized IRR. **25.00% IRR requires 3.05176x MOIC**, approximately 3.05x. A proceed recommendation is prohibited unless both verified sponsor returns meet these equivalent five-year hurdle rates with no unresolved core missing inputs, and they reconcile mathematically. Unverified or stale returns are put on hold. Reaching the thresholds does not equal investment approval.

The Word memo condenses the user's specified eight sections into two intentionally composed pages with a returns-and-recommendation header, business overview, financials, market, growth, risks, thesis and buyer hypotheses.

This is a **prototype that maps the supplied style of workbook**, not a generic model-mapping system. It will reject a workbook with different tab names/order or structural formula locations. If the CIM lacks any of the five management forecast years, forecast revenue, cost support, or adjusted EBITDA, model generation stops and identifies the gaps. No third-party base/downside projections or unprovided sixth-year extrapolation is inserted. Missing historical working capital and transaction assumptions still need analyst review. Comparable-company values are **not** independently researched.

## Important limitations

- **Prototype, not investment advice.** AI-extracted claims and financial values are explicitly unverified until an analyst checks them against original source documents.
- Supports searchable PDFs. Scanned images and complex tabular layouts may not extract correctly; OCR is not implemented.
- Memo narratives are based **only on the uploaded CIM**. External industry research and buyer validation are not performed.
- LBO uses a fixed five-year hold, a CIM management-only operating forecast, template-linked financing and a debt cash sweep. Valuation, leverage and financial structure defaults are illustrative until approved by the analyst. The sponsor-return verification checks the five-year period, forecast balance sheets, debt balances and exit equity waterfall before allowing a positive IC recommendation.
- **Do not upload confidential or nonpublic deal materials** without appropriate permissions, a data-handling agreement, secure deployment, and firm approval. Public Community Cloud is for **fictional or public demonstrations only**.
- Never commit API keys or uploaded deal files to GitHub. Uploads remain in ephemeral session memory; the selected AI provider receives submitted document text for analysis.

## Planned next versions

Real statement normalization and reconciliation, greater PDF/table fidelity, debt tranche and revolver modeling, management rollovers, sophisticated returns sensitivity, audit logs, enterprise access controls.

## Tests

```bash
pytest -q
```
