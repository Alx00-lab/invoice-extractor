# 🧾 Invoice Data Extractor

Automatically extract structured data from PDF invoices using AI — no manual entry, no errors.

**Upload a PDF → Get clean Excel in seconds.**

---

## What it does

- Reads any text-based PDF invoice
- Extracts: vendor, invoice number, dates, payment terms, line items, subtotal, tax, total
- Detects total mismatches automatically (e.g., line items don't add up to stated total)
- Delivers a formatted, ready-to-use Excel file

## Live Demo

👉 [Try it here](https://your-app.streamlit.app) *(replace with your Streamlit Cloud URL)*

## Tech Stack

| Tool | Purpose |
|---|---|
| Python 3.11 | Core language |
| pdfplumber | PDF text extraction |
| GPT-4o-mini | AI data structuring |
| Streamlit | Web interface |
| openpyxl | Excel generation |

## Run Locally

```bash
git clone https://github.com/yourusername/invoice-extractor
cd invoice-extractor
pip install -r requirements.txt

# Add your API key
echo "OPENAI_API_KEY=sk-your-key" > .env

streamlit run app.py
```

## Deploy on Streamlit Cloud (free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your repo
4. Add `OPENAI_API_KEY` in the secrets section
5. Done — live URL in 2 minutes

## Cost per invoice

Processing 100 invoices with GPT-4o-mini costs approximately **$0.10–0.20 total.**
Clients never interact with the API — this is fully managed on your end.

## Built by

Alex — Finance Automation Specialist  
Specializing in automating manual invoice and reporting workflows for small businesses.  
📧 your@email.com | [LinkedIn](https://linkedin.com/in/yourprofile) | [Upwork](https://upwork.com/yourprofile)
