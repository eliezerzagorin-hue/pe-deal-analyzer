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
4. Choose the final pre-close financial year. Generate and download a `.xlsx` containing the **same 14 sheets** as the template: COVER, ASSUMPTIONS, SUMMARY, REV, EXP, D&A, NWC, DEBT, IS, BS, CF, R, COMPS, CAP. The model attempts to fill supported historical and management forecast inputs in its appropriate schedules. Input-cell comments and SUMMARY source audit record CIM source pages; old financial data and comparables are cleared.
5. **Review all cells, complete unavailable inputs and transaction assumptions manually**, and open the workbook in Excel to recalculate. Keep the model's original integrated formulas, but confirm they make sense for the new business and capital structure. This app does not guarantee all financial details are extracted from a CIM.
6. Reupload the **Excel-recalculated workbook** in Verification to read sponsor MOIC from `R!E21` and IRR from `R!E22`, confirm assumptions, and download the **two-page formatted Word IC memo**.

**Investment decision policy:** A proceed recommendation is prohibited unless the analyst confirms a credible model supports **both sponsor MOIC >= 3.0x and sponsor IRR >= 25%**, with no unresolved core missing inputs. Returns below either threshold are rejected; unverified or uncached results are placed on hold. Reaching both thresholds does not equal investment approval.

The Word memo condenses the user's specified eight sections into two intentionally composed pages with a returns-and-recommendation header, business overview, financials, market, growth, risks, thesis and buyer hypotheses.

This is a **prototype that maps the supplied style of workbook**, not a generic model-mapping system. It will reject a workbook with different tab names/order or structural formula locations. If a new CIM lacks complete historical statements, detailed cost categories, working capital or forecast data, the generated file will contain gaps, which must be resolved before investment use. Outlook beyond the CIM can be extended as a clearly identified analyst assumption. Comparable-company values are **not** independently researched.

## Important limitations

- **Prototype, not investment advice.** AI-extracted claims and financial values are explicitly unverified until an analyst checks them against original source documents.
- Supports searchable PDFs. Scanned images and complex tabular layouts may not extract correctly; OCR is not implemented.
- Memo narratives are based **only on the uploaded CIM**. External industry research and buyer validation are not performed.
- LBO assumes a debt-free/cash-free entry, constant annual operating assumptions, five-year hold, annual debt sweep, no refinancing or distributions, and explicitly tracked sponsor equity injections if free cash flow is negative.
- **Do not upload confidential or nonpublic deal materials** without appropriate permissions, a data-handling agreement, secure deployment, and firm approval. Public Community Cloud is for **fictional or public demonstrations only**.
- Never commit API keys or uploaded deal files to GitHub. Uploads remain in ephemeral session memory; the selected AI provider receives submitted document text for analysis.

## Planned next versions

Real statement normalization and reconciliation, greater PDF/table fidelity, debt tranche and revolver modeling, management rollovers, sophisticated returns sensitivity, audit logs, enterprise access controls.

## Tests

```bash
pytest -q
```
