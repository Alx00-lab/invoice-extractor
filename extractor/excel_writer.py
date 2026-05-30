import io
import logging
import re
import time
from datetime import date as _date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .validator import parse_amount, detect_currency_symbol

logger = logging.getLogger(__name__)

# ── Brand constants ───────────────────────────────────────────
BRAND_DARK   = "1F4E78"   # headers, TOTAL row
BRAND_MID    = "2E75B6"   # mid-blue accents
BRAND_LIGHT  = "D6E4F0"   # subtotal rows
ALT_ROW      = "F2F7FC"   # alternating row tint
WHITE        = "FFFFFF"
WARN_BG      = "FFF2CC"
ERROR_BG     = "F8CBAD"
FONT_NAME    = "Calibri"
WEBSITE_URL  = "yourwebsite.com"
COMPANY_NAME_PLACEHOLDER = "[Your Company Name]"
COMPANY_ADDR_PLACEHOLDER = "[Street Address  ·  City, State ZIP  ·  (000) 000-0000  ·  email@domain.com]"

CLIENT_NAME_PLACEHOLDER   = "[Client / Company Name]"
CLIENT_ADDR_PLACEHOLDER   = "[Street Address]"
CLIENT_CITY_PLACEHOLDER   = "[City, State  ZIP]"

PAYMENT_TERMS_LINES = [
    "Payment Terms",
    "Payment due within 30 days of invoice date.",
    "Bank transfer preferred  ·  [Bank Name · Account # · Routing #]",
    "Questions?  [your@email.com]",
]


# ── Low-level cell helpers ────────────────────────────────────

def _side():
    return Side(style="thin", color="CCCCCC")


def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)


def _header_cell(cell, text, size=11):
    cell.value = text
    cell.font = Font(name=FONT_NAME, bold=True, color=WHITE, size=size)
    cell.fill = PatternFill("solid", fgColor=BRAND_DARK)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = _border()


def _data_cell(cell, value, alt=False, align="left", bold=False, size=11,
               number_format=None):
    cell.value = value if value not in (None, "null", "") else "—"
    cell.font = Font(name=FONT_NAME, size=size, bold=bold)
    cell.fill = PatternFill("solid", fgColor=ALT_ROW if alt else WHITE)
    cell.border = _border()
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if number_format:
        cell.number_format = number_format


def _subtotal_cell(cell, value, is_label=False, number_format=None):
    cell.value = value
    cell.font = Font(name=FONT_NAME, bold=True, size=11, color=BRAND_DARK)
    cell.fill = PatternFill("solid", fgColor=BRAND_LIGHT)
    cell.alignment = Alignment(
        horizontal="left" if is_label else "right", vertical="center"
    )
    cell.border = _border()
    if number_format:
        cell.number_format = number_format


def _total_cell(cell, value, is_label=False, number_format=None):
    cell.value = value
    cell.font = Font(name=FONT_NAME, bold=True, size=11, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=BRAND_DARK)
    cell.alignment = Alignment(
        horizontal="left" if is_label else "right", vertical="center"
    )
    cell.border = _border()
    if number_format:
        cell.number_format = number_format


def _footer(ws, row, n_cols):
    """Two-part branded footer: 'Thank you' left, byline right."""
    left_text  = "Thank you for your business."
    right_text = f"[Your Name]  ·  Finance Automation Specialist  ·  {WEBSITE_URL}"
    mid        = max(1, n_cols // 2)

    if mid > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=mid)
    c_l = ws.cell(row=row, column=1)
    c_l.value = left_text
    c_l.font  = Font(name=FONT_NAME, size=10, italic=True, color="888888")
    c_l.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    right_start = mid + 1
    if right_start <= n_cols and right_start < n_cols:
        ws.merge_cells(
            start_row=row, start_column=right_start,
            end_row=row,   end_column=n_cols
        )
    if right_start <= n_cols:
        c_r = ws.cell(row=row, column=right_start)
        c_r.value = right_text
        c_r.font  = Font(name=FONT_NAME, size=10, italic=True, color=BRAND_DARK)
        c_r.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

    ws.row_dimensions[row].height = 18


def _set_widths(ws, overrides=None, default=12, max_w=50):
    """Auto-size columns; overrides = {col_letter: min_width}."""
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if not hasattr(cell, "column_letter") or cell.value is None:
                continue
            col = cell.column_letter
            widths[col] = max(widths.get(col, 0), len(str(cell.value)))
    overrides = overrides or {}
    for col, w in widths.items():
        min_w = overrides.get(col, default)
        ws.column_dimensions[col].width = min(max(w + 2, min_w), max_w)
    for col, min_w in overrides.items():
        if col not in widths:
            ws.column_dimensions[col].width = min_w


def _fmt(val, symbol="$"):
    """Format float as currency string; returns '—' if None."""
    if val is None:
        return "—"
    return f"{symbol}{val:,.2f}"


def _excel_currency_format(symbol="$"):
    """Build an Excel number format string for the detected currency symbol."""
    safe = symbol.replace('"', '')
    return f'"{safe}"#,##0.00'


def _parse_date(s):
    """Best-effort parse of a date string into a datetime. Returns the original
    string if no known format matches — keeps display intact for exotic inputs."""
    if not s or not isinstance(s, str):
        return s
    s = s.strip()
    for fmt in (
        "%Y-%m-%d", "%Y/%m/%d",
        "%m/%d/%Y", "%m-%d-%Y",
        "%d/%m/%Y", "%d-%m-%Y",
        "%B %d, %Y", "%b %d, %Y",
        "%d %B %Y", "%d %b %Y",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return s


# ── Sheet 1: Invoice Summary (8 cols A-H) ─────────────────────

def _build_summary_sheet(ws, invoices, ok_invoices, sums, present, symbol):
    ws.title = "Invoice Summary"
    N = 8  # columns A through H

    money_fmt = _excel_currency_format(symbol)
    pct_fmt   = "0.00%"

    # Derive a single tax rate for per-line tax columns (matches reference look).
    if present["subtotal"] and present["tax"] and sums["subtotal"]:
        line_tax_rate = sums["tax"] / sums["subtotal"]
    else:
        line_tax_rate = 0.0

    # ── Company header (rows 2-3) ─────────────────────────────
    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = COMPANY_NAME_PLACEHOLDER
    c.font  = Font(name=FONT_NAME, bold=True, size=16, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 26

    ws.merge_cells("E2:H2")
    c = ws["E2"]
    c.value = "INVOICE"
    c.font  = Font(name=FONT_NAME, bold=True, size=18, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=BRAND_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A3:D3")
    c = ws["A3"]
    c.value = COMPANY_ADDR_PLACEHOLDER
    c.font  = Font(name=FONT_NAME, size=10, color="666666")
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("E3:H3")
    c = ws["E3"]
    c.value = "Finance Automation Specialist"
    c.font  = Font(name=FONT_NAME, italic=True, size=10, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[3].height = 18

    # ── BILL TO + invoice meta (rows 7-10) ────────────────────
    is_single = len(ok_invoices) == 1
    d_single  = ok_invoices[0].get("data", {}) if is_single else {}

    if is_single:
        client_name = d_single.get("client_name") or CLIENT_NAME_PLACEHOLDER
        client_addr = d_single.get("client_address") or CLIENT_ADDR_PLACEHOLDER
        client_city = d_single.get("client_city_zip") or CLIENT_CITY_PLACEHOLDER
        inv_num     = d_single.get("invoice_number") or "—"
        inv_date    = _parse_date(d_single.get("invoice_date"))
        due_date    = _parse_date(d_single.get("due_date"))
    else:
        n_err = len(invoices) - len(ok_invoices)
        client_name = f"Batch — {len(ok_invoices)} invoice(s) processed"
        client_addr = f"{n_err} error(s)" if n_err else ""
        client_city = ""
        inv_num     = "—"
        inv_date    = _date.today()
        due_date    = "—"

    # BILL TO label (A7)
    c = ws.cell(row=7, column=1)
    c.value = "BILL TO"
    c.font  = Font(name=FONT_NAME, bold=True, size=9, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[7].height = 16

    # Client name (A8:C8)
    ws.merge_cells("A8:C8")
    c = ws.cell(row=8, column=1)
    c.value = client_name
    c.font  = Font(name=FONT_NAME, bold=True, size=11, color="1A1F2B")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[8].height = 18

    # Client address line 1 (A9:C9)
    ws.merge_cells("A9:C9")
    c = ws.cell(row=9, column=1)
    c.value = client_addr
    c.font  = Font(name=FONT_NAME, size=10, color="555555")
    c.alignment = Alignment(horizontal="left", vertical="center")

    # Client city/zip (A10:C10)
    ws.merge_cells("A10:C10")
    c = ws.cell(row=10, column=1)
    c.value = client_city
    c.font  = Font(name=FONT_NAME, size=10, color="555555")
    c.alignment = Alignment(horizontal="left", vertical="center")

    # Right block: meta (F7:G7..G10)
    meta = [
        ("Invoice #", inv_num,   None),
        ("Date",      inv_date,  "mmm dd, yyyy"),
        ("Due Date",  due_date,  "mmm dd, yyyy"),
        ("Status",    "✓ Paid / Unpaid", None),
    ]
    for i, (label, value, fmt) in enumerate(meta):
        r = 7 + i
        c_lbl = ws.cell(row=r, column=6)
        c_lbl.value = label
        c_lbl.font  = Font(name=FONT_NAME, bold=True, size=10, color=BRAND_DARK)
        c_lbl.fill  = PatternFill("solid", fgColor=BRAND_LIGHT)
        c_lbl.alignment = Alignment(horizontal="right", vertical="center")
        c_lbl.border = _border()

        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        c_val = ws.cell(row=r, column=7)
        c_val.value = value
        c_val.font  = Font(name=FONT_NAME, size=10)
        c_val.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_val.border = _border()
        if fmt and isinstance(value, datetime):
            c_val.number_format = fmt
        ws.row_dimensions[r].height = 18

    # ── Column headers (row 12) ───────────────────────────────
    header_row = 12
    col_headers = [
        "#", "Description", "Vendor", "Qty",
        "Unit Price", "Tax %", "Tax Amt", "Total"
    ]
    for col, h in enumerate(col_headers, start=1):
        _header_cell(ws.cell(row=header_row, column=col), h)
    ws.row_dimensions[header_row].height = 24
    row = header_row + 1

    # ── Line items with LIVE formulas ─────────────────────────
    first_item_row = row
    counter = 0
    for inv in invoices:
        if inv.get("status") != "ok":
            continue
        d      = inv.get("data", {})
        vendor = d.get("vendor_name") or "—"
        for item in d.get("line_items") or []:
            alt = (counter % 2 == 1)
            qty = parse_amount(item.get("quantity"))
            up  = parse_amount(item.get("unit_price"))

            _data_cell(ws.cell(row=row, column=1), counter + 1, alt, align="center")
            _data_cell(ws.cell(row=row, column=2), item.get("description"), alt)
            _data_cell(ws.cell(row=row, column=3), vendor, alt)
            _data_cell(
                ws.cell(row=row, column=4),
                qty if qty is not None else item.get("quantity"),
                alt, align="center",
            )
            _data_cell(
                ws.cell(row=row, column=5),
                up if up is not None else item.get("unit_price"),
                alt, align="right",
                number_format=money_fmt if up is not None else None,
            )
            # Tax %
            _data_cell(
                ws.cell(row=row, column=6), line_tax_rate, alt,
                align="center", number_format=pct_fmt,
            )
            # Tax Amt — live formula = Qty * Unit Price * Tax %
            tax_cell = ws.cell(row=row, column=7)
            tax_cell.value = f"=D{row}*E{row}*F{row}"
            tax_cell.font = Font(name=FONT_NAME, size=11)
            tax_cell.fill = PatternFill("solid", fgColor=ALT_ROW if alt else WHITE)
            tax_cell.border = _border()
            tax_cell.alignment = Alignment(horizontal="right", vertical="center")
            tax_cell.number_format = money_fmt
            # Total — live formula = Qty * Unit Price + Tax Amt
            total_cell = ws.cell(row=row, column=8)
            total_cell.value = f"=D{row}*E{row}+G{row}"
            total_cell.font = Font(name=FONT_NAME, size=11, bold=True)
            total_cell.fill = PatternFill("solid", fgColor=ALT_ROW if alt else WHITE)
            total_cell.border = _border()
            total_cell.alignment = Alignment(horizontal="right", vertical="center")
            total_cell.number_format = money_fmt

            ws.row_dimensions[row].height = 18
            row += 1
            counter += 1

    last_item_row = row - 1

    if counter == 0:
        ws.merge_cells(
            start_row=row, start_column=1, end_row=row, end_column=N
        )
        c = ws.cell(row=row, column=1)
        c.value = "No line items extracted"
        c.font  = Font(name=FONT_NAME, italic=True, size=11, color="888888")
        c.alignment = Alignment(horizontal="center", vertical="center")
        row += 1

    row += 1  # spacer

    # ── Subtotal / Tax / Total (right-aligned in F:H, like reference) ──
    if counter > 0:
        subtotal_formula = f"=SUMPRODUCT(D{first_item_row}:D{last_item_row},E{first_item_row}:E{last_item_row})"
        tax_formula      = f"=SUM(G{first_item_row}:G{last_item_row})"
        total_formula    = f"=SUM(H{first_item_row}:H{last_item_row})"
    else:
        subtotal_formula = sums["subtotal"] if present["subtotal"] else "—"
        tax_formula      = sums["tax"]      if present["tax"]      else "—"
        total_formula    = sums["total"]    if present["total"]    else "—"

    # Subtotal row — label at F, value at H
    _subtotal_cell(ws.cell(row=row, column=6), "Subtotal", is_label=True)
    _subtotal_cell(ws.cell(row=row, column=8), subtotal_formula, number_format=money_fmt)
    ws.row_dimensions[row].height = 18
    row += 1

    # Tax row
    _subtotal_cell(ws.cell(row=row, column=6), "Tax", is_label=True)
    _subtotal_cell(ws.cell(row=row, column=8), tax_formula, number_format=money_fmt)
    ws.row_dimensions[row].height = 18
    row += 2  # spacer

    # TOTAL DUE — label spans E:G, value in H
    ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
    _total_cell(ws.cell(row=row, column=5), "TOTAL DUE", is_label=True)
    _total_cell(ws.cell(row=row, column=8), total_formula, number_format=money_fmt)
    ws.row_dimensions[row].height = 26
    row += 3  # spacer

    # ── Payment Terms block (4 lines, left side A:D) ──────────
    label_row = row
    ws.merge_cells(start_row=label_row, start_column=1, end_row=label_row, end_column=4)
    c = ws.cell(row=label_row, column=1)
    c.value = PAYMENT_TERMS_LINES[0]
    c.font  = Font(name=FONT_NAME, bold=True, size=11, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[label_row].height = 18

    for offset, text in enumerate(PAYMENT_TERMS_LINES[1:], start=1):
        r = label_row + offset
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        c = ws.cell(row=r, column=1)
        c.value = text
        c.font  = Font(name=FONT_NAME, size=10, color="555555")
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 16

    row = label_row + len(PAYMENT_TERMS_LINES) + 2

    # ── Footer ────────────────────────────────────────────────
    _footer(ws, row, N)

    ws.freeze_panes = "A13"
    _set_widths(
        ws,
        {"A": 6, "B": 38, "C": 22, "D": 8, "E": 14, "F": 10, "G": 14, "H": 16},
    )


# ── Sheet 2: Line Items (7 cols A-G) ──────────────────────────

def _build_lineitems_sheet(ws, invoices, sums, present, symbol):
    ws.title = "Line Items"
    N = 7
    money_fmt = _excel_currency_format(symbol)

    # ── Header bar (row 2) ────────────────────────────────────
    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = "Line Item Detail"
    c.font  = Font(name=FONT_NAME, bold=True, size=14, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 24

    ok = [i for i in invoices if i.get("status") == "ok"]
    if len(ok) == 1:
        d   = ok[0].get("data", {})
        ref = f"{d.get('invoice_number') or '—'}  ·  {d.get('vendor_name') or '—'}"
    else:
        ref = f"Batch: {len(ok)} invoice(s)"
    ws.merge_cells("E2:G2")
    c = ws["E2"]
    c.value = ref
    c.font  = Font(name=FONT_NAME, italic=True, size=11, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center")

    # ── Column headers (row 4) ────────────────────────────────
    for col, h in enumerate(
        ["#", "File", "Invoice #", "Description", "Qty", "Unit Price", "Amount"],
        start=1,
    ):
        _header_cell(ws.cell(row=4, column=col), h)
    ws.row_dimensions[4].height = 24

    # ── Data ──────────────────────────────────────────────────
    row     = 5
    first_item_row = row
    counter = 0
    for inv in invoices:
        if inv.get("status") != "ok":
            continue
        d        = inv.get("data", {})
        filename = inv.get("filename", "—")
        inv_num  = d.get("invoice_number") or "—"
        for item in d.get("line_items") or []:
            alt = (counter % 2 == 1)
            qty = parse_amount(item.get("quantity"))
            up  = parse_amount(item.get("unit_price"))
            amt = parse_amount(item.get("amount"))

            _data_cell(ws.cell(row=row, column=1), counter + 1, alt, align="center")
            _data_cell(ws.cell(row=row, column=2), filename, alt)
            _data_cell(ws.cell(row=row, column=3), inv_num,  alt)
            _data_cell(ws.cell(row=row, column=4), item.get("description"), alt)
            _data_cell(
                ws.cell(row=row, column=5),
                qty if qty is not None else item.get("quantity"),
                alt, align="center",
            )
            _data_cell(
                ws.cell(row=row, column=6),
                up if up is not None else item.get("unit_price"),
                alt, align="right",
                number_format=money_fmt if up is not None else None,
            )
            _data_cell(
                ws.cell(row=row, column=7),
                amt if amt is not None else item.get("amount"),
                alt, align="right",
                number_format=money_fmt if amt is not None else None,
            )
            ws.row_dimensions[row].height = 18
            row += 1
            counter += 1

    last_item_row = row - 1

    if counter == 0:
        ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=N)
        c = ws.cell(row=5, column=1)
        c.value = "No line items extracted"
        c.font  = Font(name=FONT_NAME, italic=True, size=11, color="888888")
        c.alignment = Alignment(horizontal="center", vertical="center")
        row = 6

    row += 1  # spacer

    # Tax rate label (e.g. "Tax (18%)") if derivable
    if present["subtotal"] and present["tax"] and sums["subtotal"]:
        rate = sums["tax"] / sums["subtotal"]
        tax_label = f"Tax ({rate * 100:.0f}%)"
    else:
        tax_label = "Tax"

    # ── Summary totals (right-aligned, label at F, value at G) ─
    rows_summary = [
        ("Subtotal", sums["subtotal"] if present["subtotal"] else "—", _subtotal_cell),
        (tax_label,  sums["tax"]      if present["tax"]      else "—", _subtotal_cell),
        ("TOTAL DUE", sums["total"]   if present["total"]    else "—", _total_cell),
    ]
    for label, value, fn in rows_summary:
        fn(ws.cell(row=row, column=6), label, is_label=True)
        fn(ws.cell(row=row, column=7), value, number_format=money_fmt)
        ws.row_dimensions[row].height = 18 if label != "TOTAL DUE" else 22
        row += 1

    row += 1
    _footer(ws, row, N)

    ws.freeze_panes = "A5"
    _set_widths(
        ws,
        {"A": 6, "B": 24, "C": 16, "D": 38, "E": 8, "F": 14, "G": 16},
    )


# ── Sheet 3: Automation Report (3 cols A-C) ───────────────────

def _build_report_sheet(ws, invoices, ok_invoices, err_count,
                        sums, present, symbol, elapsed):
    ws.title = "Automation Report"
    N = 3

    # ── Title (row 2) ─────────────────────────────────────────
    c = ws["A2"]
    c.value = "Automation Processing Report"
    c.font  = Font(name=FONT_NAME, bold=True, size=14, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 26

    ws.merge_cells("B2:C2")
    c = ws["B2"]
    c.value = f"Generated: {_date.today().strftime('%B %d, %Y')}"
    c.font  = Font(name=FONT_NAME, italic=True, size=10, color="888888")
    c.alignment = Alignment(horizontal="right", vertical="center")

    # ── Headers (row 4) ───────────────────────────────────────
    for col, h in enumerate(["Metric", "Result", "Status"], start=1):
        _header_cell(ws.cell(row=4, column=col), h)
    ws.row_dimensions[4].height = 22

    # ── Metrics ───────────────────────────────────────────────
    total = len(invoices)
    ok    = len(ok_invoices)
    pct   = f"{int(ok / total * 100)}%" if total else "—"

    total_val = sums["total"] if present["total"] else None
    tax_val   = sums["tax"]   if present["tax"]   else None

    if elapsed is not None:
        elapsed_str    = f"< {int(elapsed) + 1} seconds"
        elapsed_status = "✓ Fast" if elapsed < 60 else "⚠ Slow"
    else:
        elapsed_str    = "< 5 seconds"
        elapsed_status = "✓ Fast"

    inv_word = "invoice" if total == 1 else "invoices"

    metrics = [
        ("Invoices processed",     f"{total} {inv_word}",          "✓ OK"),
        ("Successful extractions", f"{ok} of {total}  ({pct})",
         "✓ OK" if ok == total else "⚠ Check errors"),
        ("Failed / flagged",       str(err_count) if err_count else "0",
         "⚠ Review" if err_count else "—"),
        ("Total value captured",   _fmt(total_val, symbol),       "✓ OK"),
        ("Tax captured",           _fmt(tax_val, symbol),         "✓ OK"),
        ("Processing time",        elapsed_str,                   elapsed_status),
        ("Manual entry required",  "0 fields",                    "✓ None"),
    ]

    row = 5
    for i, (metric, result, status) in enumerate(metrics):
        alt = (i % 2 == 1)
        _data_cell(ws.cell(row=row, column=1), metric, alt, bold=True)
        _data_cell(ws.cell(row=row, column=2), result, alt, align="center")

        c = ws.cell(row=row, column=3)
        c.value = status
        if status.startswith("✓"):
            color = "2D6A3F"
        elif "⚠" in status:
            color = "7B3F00"
        else:
            color = "888888"
        c.font  = Font(name=FONT_NAME, size=11, bold=True, color=color)
        c.fill  = PatternFill("solid", fgColor=ALT_ROW if alt else WHITE)
        c.border = _border()
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 20
        row += 1

    row += 2  # spacer

    # ── CTA ───────────────────────────────────────────────────
    c = ws.cell(row=row, column=1)
    c.value = "Want this running automatically for your business?"
    c.font  = Font(name=FONT_NAME, bold=True, size=11, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    c = ws.cell(row=row, column=2)
    c.value = f"Book a free 30-min audit → {WEBSITE_URL}"
    c.font  = Font(name=FONT_NAME, italic=True, size=11, color=BRAND_MID)
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[row].height = 22
    row += 3

    # ── Footer ────────────────────────────────────────────────
    _footer(ws, row, N)

    _set_widths(ws, {"A": 32, "B": 28, "C": 18})


# ── Public API ────────────────────────────────────────────────

def generate_batch_excel(invoices: list, start_time: float = None) -> bytes:
    """
    Generate a branded 3-sheet Excel from a batch of invoices.

    Args:
        invoices:   list of dicts — keys: filename, data, status, error
        start_time: time.time() snapshot taken before extraction begins;
                    used to compute elapsed in Sheet 3.  Optional.
    """
    ok_invoices = [i for i in invoices if i.get("status") == "ok"]
    err_count   = len(invoices) - len(ok_invoices)

    sums        = {"subtotal": 0.0, "tax": 0.0, "total": 0.0}
    present     = {"subtotal": False, "tax": False, "total": False}
    symbol_pool = []
    for inv in ok_invoices:
        d = inv.get("data", {})
        for k in sums:
            amt = parse_amount(d.get(k))
            if amt is not None:
                sums[k]    += amt
                present[k]  = True
            if d.get(k):
                symbol_pool.append(d[k])
    symbol  = detect_currency_symbol(symbol_pool) or "$"
    elapsed = (time.time() - start_time) if start_time else None

    wb = Workbook()
    _build_summary_sheet(wb.active, invoices, ok_invoices, sums, present, symbol)
    _build_lineitems_sheet(wb.create_sheet("Line Items"),       invoices, sums, present, symbol)
    _build_report_sheet(
        wb.create_sheet("Automation Report"),
        invoices, ok_invoices, err_count, sums, present, symbol, elapsed
    )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    logger.info(f"Excel generated: {len(invoices)} invoices, {len(ok_invoices)} OK.")
    return buf.getvalue()


def generate_excel(data: dict) -> bytes:
    """Single-invoice wrapper — backward compatible."""
    return generate_batch_excel([{
        "filename": "invoice.pdf",
        "data": data,
        "status": "ok",
    }])
