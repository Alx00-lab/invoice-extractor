# 🧾 Invoice Data Extractor

**Process multiple PDF invoices at once. Get a clean consolidated Excel in seconds.**

Built for small businesses that waste hours every month manually copying invoice data into spreadsheets.

---

## What it does

- ✅ Upload **multiple PDF invoices** at once
- ✅ AI extracts: vendor, invoice #, dates, payment terms, line items, subtotal, tax, total
- ✅ **One consolidated Excel** with two sheets: Summary + all Line Items
- ✅ Detects total mismatches automatically (line items don't add up)
- ✅ Continues processing even if one file fails
- ✅ **Try it instantly** with the included sample invoice

## Live Demo

👉 [Try it here](https://your-app.streamlit.app)

*No invoice handy? Click "Try with sample invoice" to see it in action.*

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
git clone https://github.com/yourusername/invoice-extractor
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

## Built by

**Alex** — Finance Automation Specialist
Specializing in automating manual invoice and reporting workflows for small businesses.
