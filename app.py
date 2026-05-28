from dotenv import load_dotenv
load_dotenv()

import os
import html
import logging
from pathlib import Path

import streamlit as st

from extractor import (
    extract_text_from_pdf,
    parse_invoice,
    generate_batch_excel,
)

logging.basicConfig(level=logging.INFO)

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Invoice Extractor",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Design system (CSS) ───────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    :root {
        --navy:    #1F3864;
        --blue:    #2E75B6;
        --ink:     #1A1F2B;
        --muted:   #6B7280;
        --line:    #E6E9EF;
        --bg-soft: #F8F9FA;
        --ok:      #2D6A3F;
        --warn:    #B7791F;
        --err:     #C0392B;
    }

    html, body, [class*="css"], .stApp, .stMarkdown, .stButton, .stTextInput {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    .block-container { padding-top: 1.5rem; max-width: 1120px; }
    #MainMenu, footer { visibility: hidden; }

    /* ── Hero ── */
    .hero {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        border-radius: 16px;
        padding: 2.2rem 2.4rem;
        color: #fff;
        box-shadow: 0 10px 30px rgba(31,56,100,0.18);
        margin-bottom: 1.6rem;
    }
    .hero h1 {
        font-size: 2.05rem; font-weight: 800; margin: 0 0 .4rem 0;
        letter-spacing: -0.02em; color: #fff;
    }
    .hero p { font-size: 1.02rem; color: #DCE6F5; margin: 0; max-width: 640px; line-height: 1.5; }
    .hero .pill {
        display: inline-block; margin-top: 1rem; padding: .28rem .8rem;
        background: rgba(255,255,255,0.14); border: 1px solid rgba(255,255,255,0.25);
        border-radius: 999px; font-size: .8rem; font-weight: 500; color: #EAF1FB;
    }

    /* ── KPI cards ── */
    .kpi {
        background: #fff; border: 1px solid var(--line); border-radius: 12px;
        padding: 1.1rem 1.2rem; box-shadow: 0 1px 2px rgba(16,24,40,0.04);
    }
    .kpi .val { font-size: 1.9rem; font-weight: 800; line-height: 1.1; color: var(--ink); }
    .kpi .lbl { font-size: .78rem; font-weight: 600; text-transform: uppercase;
                letter-spacing: .06em; color: var(--muted); margin-top: .25rem; }
    .kpi.ok   .val { color: var(--ok); }
    .kpi.err  .val { color: var(--err); }
    .kpi.accent .val { color: var(--blue); }

    /* ── Buttons ── */
    div[data-testid="stDownloadButton"] button,
    .stButton button[kind="primary"] {
        background: var(--navy); color: #fff; border: none; border-radius: 9px;
        padding: .6rem 1.4rem; font-weight: 600; font-size: 15px; width: 100%;
        transition: background .15s ease;
    }
    div[data-testid="stDownloadButton"] button:hover,
    .stButton button[kind="primary"]:hover { background: var(--blue); color:#fff; }

    .stButton button[kind="secondary"] {
        border-radius: 9px; border: 1px solid var(--blue); color: var(--navy);
        font-weight: 600; background: #fff;
    }
    .stButton button[kind="secondary"]:hover { background: var(--bg-soft); color: var(--navy); }

    /* ── Misc ── */
    .stAlert { border-radius: 10px; }
    section[data-testid="stSidebar"] { background: var(--bg-soft); border-right: 1px solid var(--line); }
    hr { margin: 1.1rem 0; border-color: var(--line); }
    .section-title { font-size: 1.15rem; font-weight: 700; color: var(--ink); margin: .2rem 0 .6rem 0; }
    .empty-card {
        border: 1.5px dashed var(--line); border-radius: 14px; padding: 2.4rem;
        text-align: center; color: var(--muted); background: #fff;
    }
    .empty-card .big { font-size: 2rem; }
    .footnote { color: var(--muted); font-size: .82rem; text-align: center; margin-top: 2rem; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────
def kpi_card(value, label, tone="default"):
    return (
        f'<div class="kpi {tone}">'
        f'<div class="val">{html.escape(str(value))}</div>'
        f'<div class="lbl">{html.escape(str(label))}</div></div>'
    )


# ── Hero ──────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>🧾 Invoice Data Extractor</h1>
    <p>Drop in your PDF invoices and get back clean, structured data plus a
    consolidated Excel — summary and every line item — in seconds. No manual entry, no typos.</p>
    <span class="pill">Powered by GPT-4o-mini · ~$0.0005 / invoice</span>
</div>
""", unsafe_allow_html=True)

# ── API key (secrets → env → input) ───────────────────────────
api_key = ""
secrets_path = Path(__file__).parent / ".streamlit" / "secrets.toml"
if secrets_path.exists():
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
            help="Never stored. Kept in memory for this session only.",
        )
        st.caption("Get a key at [platform.openai.com](https://platform.openai.com)")

    st.divider()
    st.markdown("**How it works**")
    st.markdown(
        "1. Upload one or more PDF invoices\n"
        "2. AI extracts every key field\n"
        "3. Review the consolidated table\n"
        "4. Download one clean Excel file"
    )
    st.divider()
    st.markdown("**Cost per invoice**")
    st.caption("~$0.0005 USD using GPT-4o-mini")
    st.divider()
    st.caption("Built by Alex · Finance Automation Specialist")

# ── Session state ─────────────────────────────────────────────
SAMPLE_PATH = Path(__file__).parent / "samples" / "sample_invoice.pdf"
st.session_state.setdefault("use_sample", False)

# ── Sample button ─────────────────────────────────────────────
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
    help="Text-based PDFs only. Scanned/image documents are not supported in this version.",
)

# ── Resolve which files to process ────────────────────────────
files_to_process = []
if st.session_state.use_sample:
    if SAMPLE_PATH.exists():
        files_to_process = [("sample_invoice.pdf", str(SAMPLE_PATH), "path")]
    else:
        st.session_state.use_sample = False
        st.error(
            "Sample invoice not found at `samples/sample_invoice.pdf`. "
            "Please upload your own PDF instead."
        )
elif uploaded_files:
    files_to_process = [(f.name, f, "file") for f in uploaded_files]

# ── Guard: no key ─────────────────────────────────────────────
if files_to_process and not api_key:
    st.warning("⚠️ Enter your OpenAI API key in the sidebar to continue.")
    st.stop()

# ── Process ───────────────────────────────────────────────────
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
        progress = st.progress(0.0)
        status_text = st.empty()
        results = []

        for i, (filename, file_obj, file_type) in enumerate(files_to_process):
            status_text.markdown(f"**Processing {i+1}/{len(files_to_process)}** — `{filename}`")
            try:
                if file_type == "file":
                    file_obj.seek(0)
                raw_text = extract_text_from_pdf(file_obj)
                data = parse_invoice(raw_text, api_key)
                results.append({"filename": filename, "status": "ok", "data": data})
            except (ValueError, RuntimeError) as e:
                logging.error(f"Failed on {filename}: {e}")
                results.append({"filename": filename, "status": "error", "error": str(e)})
            progress.progress((i + 1) / len(files_to_process))

        status_text.empty()
        progress.empty()

        ok_count = sum(1 for r in results if r["status"] == "ok")
        err_count = sum(1 for r in results if r["status"] == "error")

        # ── KPI cards ─────────────────────────────────────────
        st.markdown("<hr/>", unsafe_allow_html=True)
        k1, k2, k3 = st.columns(3)
        k1.markdown(kpi_card(len(results), "Total processed", "accent"), unsafe_allow_html=True)
        k2.markdown(kpi_card(ok_count, "Successful", "ok"), unsafe_allow_html=True)
        k3.markdown(kpi_card(err_count, "Failed", "err" if err_count else "default"),
                    unsafe_allow_html=True)

        if err_count:
            with st.expander(f"⚠️ {err_count} file(s) had errors", expanded=True):
                for r in results:
                    if r["status"] == "error":
                        st.markdown(f"**{r['filename']}** — {r['error']}")

        # ── Results table ─────────────────────────────────────
        if ok_count:
            st.markdown('<div class="section-title">📊 Extracted data</div>', unsafe_allow_html=True)
            table_data = []
            for r in results:
                if r["status"] == "ok":
                    d = r["data"]
                    table_data.append({
                        "File":      r["filename"],
                        "Vendor":    d.get("vendor_name") or "—",
                        "Invoice #": d.get("invoice_number") or "—",
                        "Date":      d.get("invoice_date") or "—",
                        "Due":       d.get("due_date") or "—",
                        "Subtotal":  d.get("subtotal") or "—",
                        "Tax":       d.get("tax") or "—",
                        "Total":     d.get("total") or "—",
                    })
            st.dataframe(table_data, use_container_width=True, hide_index=True)

            warnings = [
                (r["filename"], r["data"].get("validation_warning"))
                for r in results
                if r["status"] == "ok" and r["data"].get("validation_warning")
            ]
            if warnings:
                st.markdown("**⚠️ Validation warnings**")
                for fname, w in warnings:
                    st.warning(f"`{fname}` — {w}")

        # ── Download ──────────────────────────────────────────
        st.markdown("<hr/>", unsafe_allow_html=True)
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
            label=f"📥 Download consolidated Excel ({len(results)} invoice(s))",
            data=excel_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        if st.session_state.use_sample:
            st.session_state.use_sample = False

else:
    st.markdown("""
    <div class="empty-card">
        <div class="big">📥</div>
        <p style="margin:.6rem 0 0 0;">Upload one or more invoice PDFs above,<br/>
        or click <b>Try with sample invoice</b> to see it in action.</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown(
    '<div class="footnote">Invoice Data Extractor · Built by Alex, Finance Automation Specialist</div>',
    unsafe_allow_html=True,
)
