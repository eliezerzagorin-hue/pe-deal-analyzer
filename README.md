# DealLens

**CIM-to-IC memo and adjustable five-year LBO prototype for lower-middle-market private equity.**

DealLens is an experimental Streamlit application. Its first version can read searchable PDFs, draft a source-cited IC memo with eight requested sections, extract proposed historical financial figures for analyst confirmation, run a five-year LBO, and export an editable Word memo and formula-linked Excel model.

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
