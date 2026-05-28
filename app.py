from dotenv import load_dotenv
load_dotenv()

import streamlit as st
import logging
import os
from extractor import extract_text_from_pdf, parse_invoice, generate_excel

logging.basicConfig(level=logging.INFO)

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Invoice Extractor",
    page_icon="🧾",
    layout="centered",
)

# ── Styles ────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 720px; }
    .stAlert { border-radius: 8px; }
    .block-container { padding-top: 2rem; }
    div[data-testid="stDownloadButton"] button {
        background-color: #2E75B6;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────
st.title("🧾 Invoice Data Extractor")
st.markdown(
    "Upload a PDF invoice → get structured data and a clean Excel file. "
    "No manual entry. No errors."
)
st.divider()

# ── API Key ───────────────────────────────────────────────────
# Priority: environment variable → user input in sidebar
api_key = os.getenv("OPENAI_API_KEY", "")

with st.sidebar:
    st.header("⚙️ Settings")
    if not api_key:
        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            placeholder="sk-...",
            help="Your key is never stored. It lives in memory only for this session.",
        )
        st.caption("Get a key at [platform.openai.com](https://platform.openai.com)")
    else:
        st.success("API key loaded from environment.")

    st.divider()
    st.markdown("**How it works**")
    st.markdown("""
1. Upload your PDF invoice
2. AI extracts all key fields
3. Review the data on screen
4. Download clean Excel file
""")
    st.divider()
    st.caption("Built by Alex · Invoice Automation Specialist")

# ── Upload ────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload your invoice (PDF)",
    type=["pdf"],
    help="Text-based PDFs only. Scanned documents are not supported in this version.",
)

if uploaded_file and not api_key:
    st.warning("Please enter your OpenAI API key in the sidebar to continue.")
    st.stop()

if uploaded_file and api_key:
    st.info(f"📄 **{uploaded_file.name}** — ready to process.")

    if st.button("⚡ Extract Invoice Data", use_container_width=True):

        # ── Step 1: Extract text ──────────────────────────────
        with st.spinner("Reading PDF..."):
            try:
                raw_text = extract_text_from_pdf(uploaded_file)
            except (ValueError, RuntimeError) as e:
                st.error(str(e))
                st.stop()

        # ── Step 2: Parse with AI ─────────────────────────────
        with st.spinner("Extracting data with AI..."):
            try:
                data = parse_invoice(raw_text, api_key)
            except (ValueError, RuntimeError) as e:
                st.error(str(e))
                st.stop()

        st.success("✅ Extraction complete!")
        st.divider()

        # ── Step 3: Display results ───────────────────────────
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Vendor**")
            st.write(data.get("vendor_name") or "—")

            st.markdown("**Invoice #**")
            st.write(data.get("invoice_number") or "—")

            st.markdown("**Invoice Date**")
            st.write(data.get("invoice_date") or "—")

        with col2:
            st.markdown("**Due Date**")
            st.write(data.get("due_date") or "—")

            st.markdown("**Payment Terms**")
            st.write(data.get("payment_terms") or "—")

            st.markdown("**Total**")
            st.write(data.get("total") or "—")

        # ── Line items ────────────────────────────────────────
        line_items = data.get("line_items") or []
        if line_items:
            st.divider()
            st.markdown("**Line Items**")
            st.table([
                {
                    "Description": item.get("description") or "—",
                    "Qty":         item.get("quantity") or "—",
                    "Unit Price":  item.get("unit_price") or "—",
                    "Amount":      item.get("amount") or "—",
                }
                for item in line_items
            ])

        # ── Totals ────────────────────────────────────────────
        st.divider()
        t1, t2, t3 = st.columns(3)
        t1.metric("Subtotal", data.get("subtotal") or "—")
        t2.metric("Tax",      data.get("tax") or "—")
        t3.metric("Total",    data.get("total") or "—")

        # ── Validation warning ────────────────────────────────
        warning = data.get("validation_warning")
        if warning:
            st.warning(f"⚠️ {warning}")

        # ── Step 4: Download Excel ────────────────────────────
        st.divider()
        with st.spinner("Generating Excel file..."):
            try:
                excel_bytes = generate_excel(data)
            except Exception as e:
                st.error(f"Could not generate Excel: {e}")
                st.stop()

        filename = uploaded_file.name.replace(".pdf", "_extracted.xlsx")
        st.download_button(
            label="📥 Download Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
