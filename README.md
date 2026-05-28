# 🧾 Invoice Data Extractor

**Process multiple PDF invoices at once. Get a clean consolidated Excel in seconds.**

Built for small businesses that waste hours every month manually copying invoice data into spreadsheets.

---

## What it does

- ✅ Upload **multiple PDF invoices** at once
- ✅ AI extracts: vendor, invoice #, dates, payment terms, line items, subtotal, tax, total
- ✅ **One consolidated Excel** with two sheets: Summary + all Line Items
- ✅ Validates the math in Python (`subtotal + tax ≈ total`) and flags mismatches
- ✅ Continues processing even if one file fails — one bad PDF never breaks the batch
- ✅ **Try it instantly** with the included sample invoice

## Live Demo

> **Deploy in ~2 minutes** (see below) and paste your live Streamlit Cloud URL here.

*No invoice handy? Click "Try with sample invoice" in the app to see it in action.*

## Tech Stack

| Tool | Purpose |
|---|---|
| Python 3.11 | Core language |
| pdfplumber | PDF text extraction |
| GPT-4o-mini | AI data structuring |
| Streamlit | Web interface |
| openpyxl | Excel generation |

## Cost Per Invoice

~$0.0005 USD per invoice using GPT-4o-mini.
$5 of OpenAI credit processes ~10,000 invoices.

Clients never interact with the API — the developer absorbs this cost as part of the service.

## Run Locally

```bash
git clone <your-repo-url>
cd invoice-extractor
pip install -r requirements.txt

# Add your API key
echo "OPENAI_API_KEY=sk-your-key" > .env

streamlit run app.py
```

## Deploy on Streamlit Cloud (Free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your repo
4. In **Settings → Secrets**, add:
   ```toml
   OPENAI_API_KEY = "sk-your-key"
   ```
5. Deploy — live URL in 2 minutes

## Reliability

Built to survive real-world inputs, not just the happy path:

- **Retry with exponential backoff** — transient OpenAI rate-limits/timeouts retry up to 3× (1s → 2s → 4s) instead of failing the invoice.
- **Math validated in code, not by the AI** — the model only extracts; Python checks the totals, so summation is always correct.
- **Cost guardrail** — oversized PDFs are truncated to a bounded character limit, with a note to the user.
- **One client per session** — the OpenAI client is reused across a batch rather than re-created per file.
- **Graceful per-file errors** — a failed file is reported in the Summary sheet; the rest of the batch still completes.

## Built by

**Alex** — Finance Automation Specialist
Specializing in automating manual invoice and reporting workflows for small businesses.
