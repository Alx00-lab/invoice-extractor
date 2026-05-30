# 🧾 AI Invoice Data Extractor

> Extract structured data from PDF invoices in seconds — powered by Python & GPT-4o-mini.

[![Python](https://img.shields.io/badge/Python-3.11+-blue?style=flat-square&logo=python)](https://python.org)
[![OpenAI](https://img.shields.io/badge/GPT--4o--mini-OpenAI-412991?style=flat-square&logo=openai)](https://openai.com)
[![Streamlit](https://img.shields.io/badge/Live%20Demo-Streamlit-FF4B4B?style=flat-square&logo=streamlit)](YOUR_STREAMLIT_URL)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)

---

## The Problem

Small businesses waste **2–6 hours per week** manually copying data from PDF invoices into spreadsheets — one field at a time, vendor by vendor. It's repetitive, error-prone, and completely unnecessary.

## The Solution

Upload your invoices. Get a clean, structured Excel file in under 30 seconds.

**[→ Try the live demo]([YOUR_STREAMLIT_URL](https://invoice-extractor-cmdkq7ehi57dhjqgqbzlj6.streamlit.app/))**

---

## What It Does

| Input | Output |
|---|---|
| One or multiple PDF invoices | Organized Excel workbook |
| Any vendor format | Two sheets: Summary + Line Items |
| Messy, inconsistent layouts | Validated, structured data |

### Extracted Fields

- Vendor name, invoice number, issue date, due date
- Line items: description, quantity, unit price, amount
- Subtotal, tax, total amount, payment terms

### Built-in Validation

If the extracted total doesn't match the sum of line items, the tool flags it automatically — so errors never slip through silently.

---

## Demo

![Invoice Extractor Demo]([assets/demo.gif](https://invoice-extractor-cmdkq7ehi57dhjqgqbzlj6.streamlit.app/))

*(Upload → Extract → Download — in under 30 seconds)*

---

## Tech Stack

| Layer | Tool | Why |
|---|---|---|
| PDF Extraction | `pdfplumber` | Handles complex invoice layouts reliably |
| AI Parsing | `GPT-4o-mini` | Structured JSON extraction with retry logic |
| Excel Output | `openpyxl` | Two-sheet workbook with formatted totals |
| Web Interface | `Streamlit` | Zero-friction demo — no setup required |

---

## Run Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/invoice-extractor
cd invoice-extractor

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add your OpenAI API key
echo "OPENAI_API_KEY=sk-your-key-here" > .env

# 4. Run via CLI
python main.py samples/sample_invoice.pdf

# 5. Or launch the web interface
streamlit run app.py
```

---

## Project Structure

```
invoice-extractor/
├── app.py                  # Streamlit web interface
├── main.py                 # CLI entry point
├── extractor/
│   ├── pdf_reader.py       # PDF text extraction (pdfplumber)
│   ├── ai_parser.py        # GPT-4o-mini parsing + retry logic
│   ├── validator.py        # Math validation — flags mismatches
│   └── excel_writer.py     # Formatted Excel output (openpyxl)
├── samples/
│   └── sample_invoice.pdf  # Test invoice included
├── requirements.txt
└── .gitignore
```

---

## Sample Output

**Sheet 1 — Summary**

| Vendor | Invoice # | Date | Due Date | Subtotal | Tax | Total | Status |
|---|---|---|---|---|---|---|---|
| Acme Corp | INV-2024-001 | 2024-01-15 | 2024-02-15 | $1,200.00 | $96.00 | $1,296.00 | ✅ Valid |

**Sheet 2 — Line Items**

| Invoice # | Description | Qty | Unit Price | Amount |
|---|---|---|---|---|
| INV-2024-001 | Consulting Services | 8 | $150.00 | $1,200.00 |

---

## Use Cases

- **Accounting teams** processing high volumes of vendor invoices
- **Small businesses** moving off manual data entry
- **Bookkeepers** who need clean data for QuickBooks / Xero import
- **Developers** looking for a foundation to extend with custom fields

---

## Roadmap

- [ ] Support for image-based invoices (OCR)
- [ ] Direct QuickBooks / Xero export
- [ ] Multi-language invoice support
- [ ] Batch email ingestion via Gmail/Outlook

---

## Author

Built by **Alex** — SMB Finance Automation Specialist.
I build Python + AI systems that eliminate manual financial workflows for small businesses.

**Available for freelance projects** → [Upwork Profile](YOUR_UPWORK_URL) · [LinkedIn](YOUR_LINKEDIN_URL)

---

*If this saved you time, give it a ⭐ — it helps other businesses find it.*
