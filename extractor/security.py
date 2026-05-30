"""
Single source of truth for redaction and filename hardening.

Why this exists: error strings and log messages get aggregated by Streamlit
Cloud, OpenAI, and any external log sink. They can carry API keys, absolute
paths from the host, client emails embedded in PDFs, and user-uploaded
filenames containing PII. We scrub them at every boundary before they leave
the process or surface in the UI.
"""
import hashlib
import logging
import re
from pathlib import PurePath

# ── Redaction patterns ────────────────────────────────────────
# Order matters: more specific patterns first so they don't get clobbered.
_PATTERNS = [
    # OpenAI / Anthropic keys
    (re.compile(r"sk-[A-Za-z0-9_\-]{20,}"),       "[REDACTED-KEY]"),
    (re.compile(r"sk-ant-[A-Za-z0-9_\-]{20,}"),   "[REDACTED-KEY]"),
    # Bearer tokens
    (re.compile(r"Bearer\s+[A-Za-z0-9._\-]{20,}", re.IGNORECASE), "Bearer [REDACTED]"),
    # Email addresses
    (re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"), "[REDACTED-EMAIL]"),
    # Windows absolute paths (C:\Users\X\...)
    (re.compile(r"[A-Za-z]:\\[\w\\.\-\s]+"),     "[REDACTED-PATH]"),
    # POSIX absolute paths (3+ segments to avoid catching every '/foo')
    (re.compile(r"/(?:home|Users|var|tmp|opt|root)/[\w/.\-]+"), "[REDACTED-PATH]"),
]


def redact(message) -> str:
    """
    Return a copy of `message` with secrets, paths, and emails replaced by
    `[REDACTED-*]` markers. Safe to call on `Exception` instances directly —
    they are stringified first.
    """
    if message is None:
        return ""
    text = str(message)
    for pattern, replacement in _PATTERNS:
        text = pattern.sub(replacement, text)
    return text


# ── Filename hardening ────────────────────────────────────────
_FILENAME_BAD_CHARS = re.compile(r"[\r\n\t\x00-\x1f\"\\/:*?<>|;]")
_FILENAME_MAX_LEN   = 80


def safe_filename(name: str, default: str = "extracted.xlsx") -> str:
    """
    Sanitize a filename for use in HTTP Content-Disposition and on disk.

    Strips control chars, path separators, header-injection characters, and
    parent-directory references. Returns `default` if nothing meaningful is
    left. Length-capped to keep filenames bounded.
    """
    if not name or not isinstance(name, str):
        return default

    # Strip any directory prefix
    name = PurePath(name).name

    # Remove dangerous chars
    name = _FILENAME_BAD_CHARS.sub("", name).strip(". ")

    # Defend against traversal artifacts left after char stripping
    if not name or name in {".", ".."} or name.startswith("."):
        return default

    if len(name) > _FILENAME_MAX_LEN:
        stem = PurePath(name).stem[: _FILENAME_MAX_LEN - 10]
        suffix = PurePath(name).suffix[:10]
        name = f"{stem}{suffix}"

    return name or default


# ── File identifier for logs (never log raw filenames) ────────

def file_id(name) -> str:
    """
    Short, stable, non-reversible identifier for log lines.

    Lets you correlate log entries for the same file across the request
    lifecycle without writing the client's filename — which can itself be
    PII (e.g. 'Invoice-ACME-Corp-2026-05.pdf') — into the log sink.
    """
    if not name:
        return "file:unknown"
    digest = hashlib.sha256(str(name).encode("utf-8", "replace")).hexdigest()[:10]
    return f"file:{digest}"


# ── Markdown-safe display (defangs user-controlled text for st.markdown) ──
# Streamlit's markdown renderer is XSS-safe but still honors *, _, [], `, etc.
# A PDF named "**ACME**.pdf" would render in bold. We backslash-escape every
# CommonMark special so user-controlled strings render literally.
_MD_SPECIALS = r"\\`*_{}[]()#+-.!|<>~"
_MD_TRANS = str.maketrans({ch: f"\\{ch}" for ch in _MD_SPECIALS})


def md_safe(text) -> str:
    """
    Backslash-escape every CommonMark special so `text` renders verbatim
    when passed to `st.markdown(...)`. Use this on any user-controlled
    string (filenames, error messages from external sources) before
    interpolating into markdown.
    """
    if text is None:
        return ""
    return str(text).translate(_MD_TRANS)


# ── Root logging filter (catches third-party logs too) ────────

class _RedactingFilter(logging.Filter):
    """Pre-format a record's message and replace it with the redacted version.

    Setting `args = None` prevents the standard handler from re-formatting
    with the original args, which would re-introduce the secret we just
    stripped. Robust against records whose `msg` is not a string.
    """
    def filter(self, record):
        try:
            record.msg  = redact(record.getMessage())
            record.args = None
        except Exception:
            # Never let a logging filter raise — it would silently kill all logs.
            pass
        return True


def install_log_redaction():
    """
    Attach a redacting filter to every handler on the root logger.

    Why on handlers and not the logger itself: in stdlib `logging`, filters
    on a *logger* only see records emitted directly to that logger. Records
    propagated up from child loggers (`pdfplumber`, `openai`, `httpx`) bypass
    them and are only filtered by handler-level filters. Adding the filter
    to handlers catches everything that actually leaves the process.

    Call AFTER `logging.basicConfig(...)` so the default StreamHandler
    exists. Idempotent — safe to call repeatedly.
    """
    root = logging.getLogger()
    for handler in root.handlers:
        if not any(isinstance(f, _RedactingFilter) for f in handler.filters):
            handler.addFilter(_RedactingFilter())
