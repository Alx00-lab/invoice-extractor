"""
Pure-Python invoice validation and money parsing.

Why this exists: the AI must NEVER do arithmetic. GPT-4o-mini makes silent
summation errors, and the old prompt also used the wrong rule
(sum(line_items) == total), which falsely flags every invoice that has tax.

The correct, client-meaningful checks are done here in Python:
  1. subtotal + tax  ≈ total          (the number a client actually cares about)
  2. sum(line_items) ≈ subtotal       (catches OCR / extraction drift)
"""
import re
import logging

logger = logging.getLogger(__name__)

# Currency rounding slack. 0.05 absorbs legitimate 1-2 cent rounding without
# hiding a real error.
DEFAULT_TOLERANCE = 0.05


def parse_amount(value):
    """
    Parse a monetary value into a float, robust to currency symbols and to both
    US ('1,234.56') and European ('1.234,56') separators.

    Returns None when the value is missing or unparseable — callers must handle None.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    if not s:
        return None

    # Strip everything except digits, separators and a leading minus.
    cleaned = re.sub(r"[^\d.,\-]", "", s)
    if not cleaned or cleaned in ("-", ".", ","):
        return None

    if "," in cleaned and "." in cleaned:
        # Whichever separator comes last is the decimal separator.
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")  # 1.234,56 -> 1234.56
        else:
            cleaned = cleaned.replace(",", "")                    # 1,234.56 -> 1234.56
    elif "," in cleaned:
        parts = cleaned.split(",")
        # A single comma with 1-2 trailing digits is a decimal comma (e.g. '7,50').
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            cleaned = cleaned.replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")                    # thousands grouping

    try:
        return float(cleaned)
    except ValueError:
        return None


def detect_currency_symbol(values):
    """
    Best-effort: return the leading currency symbol/prefix from the first
    parseable string in `values` (e.g. '$', 'RD$', '€', 'US$'). Empty if none.
    """
    for v in values:
        if not v:
            continue
        m = re.match(r"^\s*([^\d\s.,\-]+)", str(v))
        if m:
            return m.group(1).strip()
    return ""


def validate_invoice(data, tolerance=DEFAULT_TOLERANCE):
    """
    Validate the invoice math. Returns a human-readable warning string, or None
    when everything reconciles (or there isn't enough data to check).
    """
    warnings = []

    subtotal = parse_amount(data.get("subtotal"))
    tax = parse_amount(data.get("tax"))
    total = parse_amount(data.get("total"))

    # ── Check 1: subtotal + tax ≈ total ───────────────────────────────────
    if subtotal is not None and total is not None:
        tax_val = tax if tax is not None else 0.0
        expected = subtotal + tax_val
        if abs(expected - total) > tolerance:
            warnings.append(
                f"Total mismatch: subtotal ({subtotal:,.2f}) + tax ({tax_val:,.2f}) "
                f"= {expected:,.2f}, but total shows {total:,.2f}"
            )

    # ── Check 2: sum(line_items) ≈ subtotal (fallback: total) ─────────────
    line_items = data.get("line_items") or []
    item_sum = 0.0
    counted = 0
    for item in line_items:
        amt = parse_amount(item.get("amount"))
        if amt is not None:
            item_sum += amt
            counted += 1

    if counted > 0:
        compare_to = subtotal if subtotal is not None else total
        label = "subtotal" if subtotal is not None else "total"
        if compare_to is not None and abs(item_sum - compare_to) > tolerance:
            warnings.append(
                f"Line items sum to {item_sum:,.2f} but {label} shows {compare_to:,.2f}"
            )

    return " | ".join(warnings) if warnings else None
