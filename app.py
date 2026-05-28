from dotenv import load_dotenv
load_dotenv()

import streamlit as st
import logging
import os
import io
from pathlib import Path
from extractor import (
    extract_text_from_pdf,
    parse_invoice,
    generate_batch_excel,
)

logging.basicConfig(level=logging.INFO)

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Invoice Extractor — Batch",
    page_icon="🧾",
    layout="wide",
)

# ── Styles ────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 1100px; }
    .stAlert { border-radius: 8px; }
    div[data-testid="stDownloadButton"] button {
        background-color: #2E75B6;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        width: 100%;
        font-size: 16px;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #1F3864;
    }
    .metric-card {
        background-color: #F8F9FA;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #E0E0E0;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────
st.title("🧾 Invoice Data Extractor")
st.markdown(
    "**Upload multiple PDF invoices** → AI extracts all data → "
    "**download one consolidated Excel** with summary + line items. "
    "No manual entry. No errors."
)
st.divider()

# ── API Key handling ──────────────────────────────────────────
# Priority: Streamlit secrets → environment variable → user input
api_key = ""
try:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
except Exception:
    pass
if not api_key:
    api_key = os.getenv("OPENAI_API_KEY", "")

with st.sidebar:
    st.header("⚙️ Settings")

    if api_key:
        st.success("✓ API key loaded")
    else:
        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            placeholder="sk-...",
            help="Your key is never stored. It lives in memory only for this session.",
        )
        st.caption("Get a key at [platform.openai.com](https://platform.openai.com)")

    st.divider()
    st.markdown("**How it works**")
    st.markdown("""
1. Upload one or more PDF invoices
2. AI extracts all key fields
3. Review consolidated table
4. Download one clean Excel file
""")

    st.divider()
    st.markdown("**Cost per invoice**")
    st.caption("~$0.0005 USD using GPT-4o-mini")

    st.divider()
    st.caption("Built by Alex · Invoice Automation Specialist")

# ── Sample loader ─────────────────────────────────────────────
SAMPLE_PATH = Path(__file__).parent / "samples" / "sample_invoice.pdf"

if "uploaded_files_list" not in st.session_state:
    st.session_state.uploaded_files_list = None
if "use_sample" not in st.session_state:
    st.session_state.use_sample = False

# ── Try sample button ─────────────────────────────────────────
col_sample, col_info = st.columns([1, 3])
with col_sample:
    if st.button("🎯 Try with sample invoice", use_container_width=True):
        st.session_state.use_sample = True
with col_info:
    st.caption("No invoice handy? Click to test with a sample PDF.")

# ── Upload ────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "📁 Upload your invoices (PDF — multiple allowed)",
    type=["pdf"],
    accept_multiple_files=True,
    help="Text-based PDFs only. Scanned documents are not supported in this version.",
)

# Determine which files to process
files_to_process = []
if st.session_state.use_sample and SAMPLE_PATH.exists():
    files_to_process = [("sample_invoice.pdf", str(SAMPLE_PATH), "path")]
elif uploaded_files:
    files_to_process = [(f.name, f, "file") for f in uploaded_files]

# ── Guard: no API key ─────────────────────────────────────────
if files_to_process and not api_key:
    st.warning("⚠️ Please enter your OpenAI API key in the sidebar to continue.")
    st.stop()

# ── Show files queued ─────────────────────────────────────────
if files_to_process:
    st.info(f"📄 **{len(files_to_process)} file(s)** ready to process.")
    with st.expander("Files queued", expanded=False):
        for name, _, _ in files_to_process:
            st.markdown(f"- {name}")

    process_clicked = st.button(
        f"⚡ Extract Data from {len(files_to_process)} Invoice(s)",
        use_container_width=True,
        type="primary",
    )

    if process_clicked:
        # ── Process all invoices ──────────────────────────────
        progress = st.progress(0.0)
        status_text = st.empty()
        results = []

        for i, (filename, file_obj, file_type) in enumerate(files_to_process):
            status_text.markdown(f"**Processing {i+1}/{len(files_to_process)}** — `{filename}`")

            try:
                # Step 1: extract text
                if file_type == "path":
                    raw_text = extract_text_from_pdf(file_obj)
                else:
                    file_obj.seek(0)
                    raw_text = extract_text_from_pdf(file_obj)

                # Step 2: parse with AI
                data = parse_invoice(raw_text, api_key)

                results.append({
                    "filename": filename,
                    "status": "ok",
                    "data": data,
                })

            except (ValueError, RuntimeError) as e:
                logging.error(f"Failed on {filename}: {e}")
                results.append({
                    "filename": filename,
                    "status": "error",
                    "error": str(e),
                })

            progress.progress((i + 1) / len(files_to_process))

        status_text.empty()
        progress.empty()

        # ── Summary metrics ───────────────────────────────────
        ok_count = sum(1 for r in results if r["status"] == "ok")
        err_count = sum(1 for r in results if r["status"] == "error")

        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Processed", len(results))
        m2.metric("Successful", ok_count, delta=None)
        m3.metric("Failed", err_count, delta=None,
                  delta_color="inverse" if err_count > 0 else "off")

        if ok_count > 0:
            st.success(f"✅ Extracted data from {ok_count} invoice(s)")

        if err_count > 0:
            with st.expander(f"⚠️ {err_count} file(s) had errors", expanded=True):
                for r in results:
                    if r["status"] == "error":
                        st.markdown(f"**{r['filename']}** — {r['error']}")

        # ── Display consolidated results ──────────────────────
        if ok_count > 0:
            st.divider()
            st.markdown("### 📊 Extracted Data")

            table_data = []
            for r in results:
                if r["status"] == "ok":
                    d = r["data"]
                    table_data.append({
                        "File":         r["filename"],
                        "Vendor":       d.get("vendor_name") or "—",
                        "Invoice #":    d.get("invoice_number") or "—",
                        "Date":         d.get("invoice_date") or "—",
                        "Due":          d.get("due_date") or "—",
                        "Subtotal":     d.get("subtotal") or "—",
                        "Tax":          d.get("tax") or "—",
                        "Total":        d.get("total") or "—",
                    })
            st.dataframe(table_data, use_container_width=True, hide_index=True)

            # Warnings collected
            warnings = [
                (r["filename"], r["data"].get("validation_warning"))
                for r in results
                if r["status"] == "ok" and r["data"].get("validation_warning")
            ]
            if warnings:
                st.markdown("**⚠️ Validation warnings:**")
                for fname, w in warnings:
                    st.warning(f"`{fname}` — {w}")

        # ── Download Excel ────────────────────────────────────
        st.divider()
        with st.spinner("Generating Excel file..."):
            try:
                excel_bytes = generate_batch_excel(results)
            except Exception as e:
                st.error(f"Could not generate Excel: {e}")
                st.stop()

        download_name = (
            "invoices_extracted.xlsx"
            if len(results) > 1
            else f"{results[0]['filename'].replace('.pdf', '')}_extracted.xlsx"
        )

        st.download_button(
            label=f"📥 Download Consolidated Excel ({len(results)} invoice(s))",
            data=excel_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Reset sample flag after processing
        if st.session_state.use_sample:
            st.session_state.use_sample = False

else:
    # Empty state hint
    st.markdown(" ")
    st.info(
        "👆 Upload one or more invoice PDFs above, or click "
        "**Try with sample invoice** to see it in action."
    )
