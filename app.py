from dotenv import load_dotenv
load_dotenv()

import os
import html
import logging
import time
from pathlib import Path

import streamlit as st

from extractor import (
    extract_text_from_pdf,
    parse_invoice,
    generate_batch_excel,
    redact,
    safe_filename,
    file_id,
    md_safe,
)

logging.basicConfig(level=logging.INFO)
# Note: a root-handler logging filter to scrub third-party leaks lives in
# extractor/security.py (install_log_redaction). It is intentionally NOT
# wired here — it touched Streamlit's own root handlers and could interfere
# with Tornado's websocket plumbing. Our own modules already redact at every
# raise/log site, so direct leakage is covered. Revisit only if we move
# behind a reverse proxy that handles logs separately.

# ── Demo guardrails ───────────────────────────────────────────
# Every OpenAI call costs real money. These caps bound the worst-case
# spend per anonymous visitor. Tune in one place; UI text reads from here.
MAX_INVOICES_PER_BATCH      = 3
MAX_FILE_SIZE_MB            = 5
MAX_BATCH_SIZE_MB           = 10
MAX_EXTRACTIONS_PER_SESSION = 9     # 3 full batches before quota locks

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
        --navy:    #1F3A5F;
        --blue:    #2E75B6;
        --ink:     #1A1A1A;
        --muted:   #666666;
        --line:    #C9D6E6;
        --bg-soft: #EAF3FB;
        --alt-row: #D9EAF7;
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
        background: linear-gradient(135deg, #1F3A5F 0%, #2E75B6 100%);
        border-radius: 16px;
        padding: 2.2rem 2.4rem;
        color: #fff;
        box-shadow: 0 10px 30px rgba(31,58,95,0.18);
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

    /* ── Buttons (unified system) ──
       Base: every button gets the same height, radius, weight, transition.
       Variants: primary (navy fill), secondary (white w/ border), download (gradient). */
    .stButton button,
    div[data-testid="stDownloadButton"] button {
        border-radius: 10px;
        padding: .65rem 1.25rem;
        font-weight: 600;
        font-size: 14px;
        letter-spacing: .005em;
        width: 100%;
        min-height: 44px;
        line-height: 1.2;
        transition: all .18s ease;
        box-shadow: 0 1px 2px rgba(16,24,40,0.05);
    }

    /* Default (sample-style) — clean ghost button */
    .stButton button {
        background: #fff;
        color: var(--navy);
        border: 1px solid var(--line);
    }
    .stButton button:hover {
        background: var(--bg-soft);
        border-color: var(--blue);
        color: var(--navy);
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(31,58,95,0.10);
    }
    .stButton button:active {
        transform: translateY(0);
        box-shadow: 0 1px 2px rgba(16,24,40,0.05);
    }

    /* Primary CTA — Extract Data */
    .stButton button[kind="primary"] {
        background: linear-gradient(135deg, #1F3A5F 0%, #2E75B6 100%);
        color: #fff;
        border: none;
        box-shadow: 0 4px 14px rgba(31,58,95,0.22);
    }
    .stButton button[kind="primary"]:hover {
        background: linear-gradient(135deg, #1A305A 0%, #286AA8 100%);
        color: #fff;
        transform: translateY(-1px);
        box-shadow: 0 6px 18px rgba(31,58,95,0.30);
    }
    .stButton button[kind="primary"]:active {
        transform: translateY(0);
        box-shadow: 0 2px 6px rgba(31,58,95,0.22);
    }

    /* Secondary kind — outline */
    .stButton button[kind="secondary"] {
        background: #fff;
        color: var(--navy);
        border: 1px solid var(--blue);
    }
    .stButton button[kind="secondary"]:hover {
        background: var(--bg-soft);
        color: var(--navy);
        border-color: var(--navy);
    }

    /* Download — solid navy, distinct from primary CTA */
    div[data-testid="stDownloadButton"] button {
        background: var(--navy);
        color: #fff;
        border: none;
        box-shadow: 0 4px 14px rgba(31,58,95,0.22);
    }
    div[data-testid="stDownloadButton"] button:hover {
        background: var(--blue);
        color: #fff;
        transform: translateY(-1px);
        box-shadow: 0 6px 18px rgba(46,117,182,0.28);
    }
    div[data-testid="stDownloadButton"] button:active {
        transform: translateY(0);
        box-shadow: 0 2px 6px rgba(31,58,95,0.22);
    }

    /* Focus ring — keyboard accessibility */
    .stButton button:focus-visible,
    div[data-testid="stDownloadButton"] button:focus-visible {
        outline: 2px solid var(--blue);
        outline-offset: 2px;
    }

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
    <span class="pill">AI-powered · From PDF to structured Excel in seconds</span>
</div>
""", unsafe_allow_html=True)

# ── API key (secrets → env) ───────────────────────────────────
api_key = ""
secrets_path = Path(__file__).parent / ".streamlit" / "secrets.toml"
if secrets_path.exists():
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        pass
if not api_key:
    api_key = os.getenv("OPENAI_API_KEY", "")

# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### How it works")
    st.markdown(
        "1. Upload one or more PDF invoices\n"
        "2. AI extracts every key field\n"
        "3. Review the consolidated table\n"
        "4. Download one clean Excel file"
    )
    st.divider()
    st.markdown("### Demo limits")
    st.session_state.setdefault("extractions_used", 0)
    remaining_quota = max(0, MAX_EXTRACTIONS_PER_SESSION - st.session_state.extractions_used)
    st.markdown(
        f"- **{MAX_INVOICES_PER_BATCH}** invoices max per batch\n"
        f"- **{MAX_FILE_SIZE_MB} MB** max per file\n"
        f"- **{remaining_quota} / {MAX_EXTRACTIONS_PER_SESSION}** extractions left this session"
    )
    st.caption("Need higher volume? Book a free audit below the results.")
    st.divider()
    st.markdown("Built by **Alex** · Finance Automation Specialist")

# ── Session state ─────────────────────────────────────────────
SAMPLE_PATH = Path(__file__).parent / "samples" / "sample_invoice.pdf"
st.session_state.setdefault("use_sample", False)
st.session_state.setdefault("use_sample_batch", False)

# ── Sample buttons ────────────────────────────────────────────
if SAMPLE_PATH.exists():
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("🎯 Try sample invoice", use_container_width=True):
            st.session_state.use_sample = True
            st.session_state.use_sample_batch = False
    with col2:
        if st.button(f"📦 Try {MAX_INVOICES_PER_BATCH} invoices →", use_container_width=True):
            st.session_state.use_sample_batch = True
            st.session_state.use_sample = False
    with col3:
        st.caption("No invoice handy? Try the demo with sample data.")
else:
    # TODO: add sample_invoice.pdf to samples/ folder
    pass

# ── Upload ────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    f"📁 Upload your invoices (PDF — up to {MAX_INVOICES_PER_BATCH} per batch)",
    type=["pdf"],
    accept_multiple_files=True,
    help=(
        f"Text-based PDFs only. Max {MAX_INVOICES_PER_BATCH} files per batch, "
        f"{MAX_FILE_SIZE_MB} MB per file, {MAX_BATCH_SIZE_MB} MB total. "
        f"Demo capped at {MAX_EXTRACTIONS_PER_SESSION} extractions per session."
    ),
)

# ── Upload validation (cap count, sizes — runs BEFORE any API call) ──
if uploaded_files:
    if len(uploaded_files) > MAX_INVOICES_PER_BATCH:
        st.error(
            f"⚠️ This demo accepts up to **{MAX_INVOICES_PER_BATCH} invoices per batch**. "
            f"You uploaded **{len(uploaded_files)}**. "
            f"Remove {len(uploaded_files) - MAX_INVOICES_PER_BATCH} file(s) and try again, "
            "or book a free audit (link below) for higher-volume processing."
        )
        st.stop()

    oversized = [
        f"`{f.name}` ({f.size / (1024*1024):.1f} MB)"
        for f in uploaded_files
        if f.size > MAX_FILE_SIZE_MB * 1024 * 1024
    ]
    if oversized:
        st.error(
            f"⚠️ The following file(s) exceed the **{MAX_FILE_SIZE_MB} MB per-file limit**: "
            + ", ".join(oversized)
            + ". Please trim or split before uploading."
        )
        st.stop()

    total_mb = sum(f.size for f in uploaded_files) / (1024 * 1024)
    if total_mb > MAX_BATCH_SIZE_MB:
        st.error(
            f"⚠️ Total batch size is **{total_mb:.1f} MB**, "
            f"exceeding the **{MAX_BATCH_SIZE_MB} MB** batch limit. "
            "Please remove some files."
        )
        st.stop()

# ── Resolve which files to process ────────────────────────────
files_to_process = []
if st.session_state.use_sample_batch:
    if SAMPLE_PATH.exists():
        files_to_process = [
            (f"sample_invoice_{i}.pdf", str(SAMPLE_PATH), "path")
            for i in range(1, MAX_INVOICES_PER_BATCH + 1)
        ]
    else:
        st.session_state.use_sample_batch = False
elif st.session_state.use_sample:
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

# ── Guard: no key configured ──────────────────────────────────
if files_to_process and not api_key:
    st.error("⚠️ This demo is temporarily unavailable. Please check back shortly.")
    st.stop()

# ── Per-session quota gate ────────────────────────────────────
# Hard stop the moment a request would exceed the per-session budget.
# Streamlit session_state survives page interactions but resets on full
# refresh — fine as a deterrent; for true rate-limiting move this to IP+Redis.
if files_to_process:
    used      = st.session_state.get("extractions_used", 0)
    remaining = MAX_EXTRACTIONS_PER_SESSION - used

    if remaining <= 0:
        st.error(
            f"⚠️ You've used all **{MAX_EXTRACTIONS_PER_SESSION}** demo extractions for "
            "this session. For unlimited processing, book a free audit (link below)."
        )
        st.stop()

    if len(files_to_process) > remaining:
        st.error(
            f"⚠️ This batch needs **{len(files_to_process)}** extractions but you only "
            f"have **{remaining}** left this session. Reduce the batch or book an audit."
        )
        st.stop()

# ── Process ───────────────────────────────────────────────────
if files_to_process:
    st.info(f"📄 **{len(files_to_process)} file(s)** ready to process.")
    with st.expander("Files queued", expanded=False):
        for name, _, _ in files_to_process:
            st.markdown(f"- {md_safe(name)}")

    process_clicked = st.button(
        f"⚡ Extract Data from {len(files_to_process)} Invoice(s)",
        use_container_width=True,
        type="primary",
    )

    if process_clicked:
        _start = time.time()          # for Sheet 3 processing-time metric
        progress    = st.progress(0.0)
        status_text = st.empty()
        results     = []

        for i, (filename, file_obj, file_type) in enumerate(files_to_process):
            status_text.markdown(
                f"**Processing {i+1}/{len(files_to_process)}** — `{md_safe(filename)}`"
            )
            try:
                if file_type == "file":
                    file_obj.seek(0)
                raw_text = extract_text_from_pdf(file_obj)
                data     = parse_invoice(raw_text, api_key)
                results.append({"filename": filename, "status": "ok", "data": data})
            except (ValueError, RuntimeError) as e:
                # Never log the raw filename or raw exception — both can be PII.
                logging.error(f"Extraction failed [{file_id(filename)}]: {redact(e)}")
                results.append({
                    "filename": filename,
                    "status": "error",
                    # User sees a scrubbed message; raw exception stays in logs.
                    "error": redact(e),
                })
            progress.progress((i + 1) / len(files_to_process))

        status_text.empty()
        progress.empty()

        ok_count  = sum(1 for r in results if r["status"] == "ok")
        err_count = sum(1 for r in results if r["status"] == "error")

        # ── KPI cards ─────────────────────────────────────────
        st.markdown("<hr/>", unsafe_allow_html=True)
        k1, k2, k3 = st.columns(3)
        k1.markdown(kpi_card(len(results), "Total processed", "accent"), unsafe_allow_html=True)
        k2.markdown(kpi_card(ok_count,  "Successful", "ok"),             unsafe_allow_html=True)
        k3.markdown(
            kpi_card(err_count, "Failed", "err" if err_count else "default"),
            unsafe_allow_html=True,
        )

        if err_count:
            with st.expander(f"⚠️ {err_count} file(s) had errors", expanded=True):
                for r in results:
                    if r["status"] == "error":
                        # Both filename and error are user-derived; escape both.
                        st.markdown(
                            f"**{md_safe(r['filename'])}** — {md_safe(r['error'])}"
                        )

        # ── Results table ─────────────────────────────────────
        if ok_count:
            st.markdown('<div class="section-title">📊 Extracted data</div>', unsafe_allow_html=True)
            table_data = []
            for r in results:
                if r["status"] == "ok":
                    d = r["data"]
                    table_data.append({
                        "File":      r["filename"],
                        "Vendor":    d.get("vendor_name")   or "—",
                        "Invoice #": d.get("invoice_number") or "—",
                        "Date":      d.get("invoice_date")   or "—",
                        "Due":       d.get("due_date")       or "—",
                        "Subtotal":  d.get("subtotal")       or "—",
                        "Tax":       d.get("tax")            or "—",
                        "Total":     d.get("total")          or "—",
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
                    st.warning(f"`{md_safe(fname)}` — {md_safe(w)}")

        # ── Download ──────────────────────────────────────────
        st.markdown("<hr/>", unsafe_allow_html=True)
        with st.spinner("Generating Excel file..."):
            try:
                excel_bytes = generate_batch_excel(results, start_time=_start)
            except Exception as e:
                # Log full (redacted) for debugging; show a generic message.
                logging.error(f"Excel generation failed: {redact(e)}")
                st.error("Could not generate the Excel file. Please try again.")
                st.stop()

        if len(results) > 1:
            download_name = "invoices_extracted.xlsx"
        else:
            stem = results[0]["filename"].rsplit(".pdf", 1)[0]
            download_name = safe_filename(
                f"{stem}_extracted.xlsx",
                default="invoice_extracted.xlsx",
            )
        st.download_button(
            label=f"📥 Download consolidated Excel ({len(results)} invoice(s))",
            data=excel_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # ── Charge the session budget (every attempt counts, success or fail) ──
        st.session_state.extractions_used = (
            st.session_state.get("extractions_used", 0) + len(results)
        )

        # ── CTA (Fix 3) ───────────────────────────────────────
        st.markdown("---")
        st.markdown(
            "💼 **Want this running automatically for your business?**  \n"
            "[Book a free 30-min audit →](https://yourwebsite.com)",
            unsafe_allow_html=False,
        )

        # Reset sample flags
        if st.session_state.use_sample:
            st.session_state.use_sample = False
        if st.session_state.use_sample_batch:
            st.session_state.use_sample_batch = False

else:
    st.markdown("""
    <div class="empty-card">
        <div class="big">📥</div>
        <p style="margin:.6rem 0 0 0;">Upload one or more invoice PDFs above,<br/>
        or click <b>Try sample invoice</b> to see it in action.</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown(
    '<div class="footnote">Invoice Data Extractor · Built by Alex, Finance Automation Specialist</div>',
    unsafe_allow_html=True,
)
