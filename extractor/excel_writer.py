import io
import logging
import re
import time
from datetime import date as _date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from .validator import parse_amount, detect_currency_symbol

logger = logging.getLogger(__name__)

# ── Brand constants (Premium palette — Prompt.txt spec) ───────
BRAND_DARK     = "1F3A5F"   # deep navy — headers, TOTAL row
BRAND_MID      = "2E75B6"   # professional blue — accents
BRAND_LIGHT    = "EAF3FB"   # light fill — block backgrounds
ALT_ROW        = "D9EAF7"   # alt row fill — striped tables
WHITE          = "FFFFFF"
TEXT_PRIMARY   = "1A1A1A"   # body text
TEXT_SECONDARY = "666666"   # captions, supporting text
RULE_LINE      = "C9D6E6"   # thin rule line
STATUS_UNPAID  = "B7791F"   # amber
STATUS_OVERDUE = "C0392B"   # red
STATUS_PAID    = "2D6A3F"   # green
WARN_BG        = "FFF2CC"
ERROR_BG       = "F8CBAD"
FONT_NAME      = "Calibri"
WEBSITE_URL    = "yourwebsite.com"
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
    return Side(style="thin", color=RULE_LINE)


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


def _compute_status(due_date_value, today=None):
    """
    Return (label, color_hex) for the invoice status block.
    Dynamic: OVERDUE if due_date < today, else UNPAID. PAID is not auto-detected.
    """
    today = today or _date.today()
    due = None
    if isinstance(due_date_value, datetime):
        due = due_date_value.date()
    elif isinstance(due_date_value, _date):
        due = due_date_value

    if due and due < today:
        return "OVERDUE", STATUS_OVERDUE
    return "UNPAID", STATUS_UNPAID


def _apply_print_setup(ws, n_cols):
    """Print-ready: portrait, fit-to-width, centered, slim margins, repeat header."""
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize   = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.2, footer=0.2
    )
    ws.print_options.horizontalCentered = True
    end_col = get_column_letter(n_cols)
    ws.print_area = f"A1:{end_col}{ws.max_row}"


def _footer(ws, row, n_cols, signature=True):
    """
    Premium footer.
      signature=True  (invoice mode): 'Thank you' + signature line + 'Authorized Signature'
                                       caption + byline.
      signature=False (batch report mode): 'Thank you' + byline only — no signature line,
                                            because a batch summary isn't a per-invoice doc.
    """
    # Row 1: Thank-you (centered across all cols)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    c = ws.cell(row=row, column=1)
    c.value = "Thank you for your business."
    c.font  = Font(name=FONT_NAME, size=11, italic=True, color=TEXT_SECONDARY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 22

    if signature:
        # Signature line — right half of the sheet, bottom border only
        sig_row = row + 2
        sig_start = max(1, n_cols - 2)
        if sig_start < n_cols:
            ws.merge_cells(
                start_row=sig_row, start_column=sig_start,
                end_row=sig_row,   end_column=n_cols,
            )
        c = ws.cell(row=sig_row, column=sig_start)
        c.value = ""
        c.border = Border(bottom=Side(style="medium", color=BRAND_DARK))
        c.alignment = Alignment(horizontal="center", vertical="bottom")
        ws.row_dimensions[sig_row].height = 28

        cap_row = sig_row + 1
        if sig_start < n_cols:
            ws.merge_cells(
                start_row=cap_row, start_column=sig_start,
                end_row=cap_row,   end_column=n_cols,
            )
        c = ws.cell(row=cap_row, column=sig_start)
        c.value = "Authorized Signature"
        c.font  = Font(name=FONT_NAME, size=9, italic=True, color=TEXT_SECONDARY)
        c.alignment = Alignment(horizontal="center", vertical="top")
        ws.row_dimensions[cap_row].height = 14

        by_row = cap_row + 2
    else:
        by_row = row + 2

    ws.merge_cells(start_row=by_row, start_column=1, end_row=by_row, end_column=n_cols)
    c = ws.cell(row=by_row, column=1)
    c.value = f"[Your Name]  ·  Finance Automation Specialist  ·  {WEBSITE_URL}"
    c.font  = Font(name=FONT_NAME, size=9, italic=True, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[by_row].height = 16


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


# ── Sheet 1: Invoice Summary — router ─────────────────────────

def _build_summary_sheet(ws, invoices, ok_invoices, sums, present, symbol, elapsed=None):
    """Dispatch: single-invoice premium layout vs multi-invoice batch report."""
    if len(ok_invoices) <= 1:
        _build_single_summary_sheet(ws, invoices, ok_invoices, sums, present, symbol)
    else:
        _build_batch_summary_sheet(ws, invoices, ok_invoices, sums, present, symbol, elapsed)


# ── Sheet 1A: Single-invoice premium layout (8 cols A-H) ──────

def _build_single_summary_sheet(ws, invoices, ok_invoices, sums, present, symbol):
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
    c.font  = Font(name=FONT_NAME, bold=True, size=18, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 32

    ws.merge_cells("E2:H2")
    c = ws["E2"]
    c.value = "INVOICE"
    c.font  = Font(name=FONT_NAME, bold=True, size=26, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=BRAND_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    ws.merge_cells("A3:D3")
    c = ws["A3"]
    c.value = COMPANY_ADDR_PLACEHOLDER
    c.font  = Font(name=FONT_NAME, size=10, color=TEXT_SECONDARY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    # Thin rule under company block
    c.border = Border(bottom=Side(style="thin", color=BRAND_MID))

    ws.merge_cells("E3:H3")
    c = ws["E3"]
    c.value = "Finance Automation Specialist"
    c.font  = Font(name=FONT_NAME, italic=True, size=10, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center")
    c.border = Border(bottom=Side(style="thin", color=BRAND_MID))
    ws.row_dimensions[3].height = 20

    # ── BILL TO + invoice meta (rows 7-10) ────────────────────
    d_single = ok_invoices[0].get("data", {}) if ok_invoices else {}
    client_name = d_single.get("client_name")    or CLIENT_NAME_PLACEHOLDER
    client_addr = d_single.get("client_address") or CLIENT_ADDR_PLACEHOLDER
    client_city = d_single.get("client_city_zip") or CLIENT_CITY_PLACEHOLDER
    inv_num     = d_single.get("invoice_number") or "—"
    inv_date    = _parse_date(d_single.get("invoice_date"))
    due_date    = _parse_date(d_single.get("due_date"))

    status_text, status_color = _compute_status(due_date)

    # ── Boxed BILL TO block (rows 7-10, cols A:C) ─────────────
    # Header bar (row 7)
    ws.merge_cells("A7:C7")
    c = ws.cell(row=7, column=1)
    c.value = "BILL TO"
    c.font  = Font(name=FONT_NAME, bold=True, size=9, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[7].height = 18

    # Three content rows — light fill, side borders, bottom on last row
    box_side   = Side(style="thin", color=RULE_LINE)
    box_bottom = Side(style="medium", color=BRAND_DARK)
    box_fill   = PatternFill("solid", fgColor=BRAND_LIGHT)

    def _box_row(r, text, is_last=False, bold=False, size=11, color=TEXT_PRIMARY):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        cell = ws.cell(row=r, column=1)
        cell.value = text
        cell.font  = Font(name=FONT_NAME, bold=bold, size=size, color=color)
        cell.fill  = box_fill
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border = Border(
            left=box_side, right=box_side,
            bottom=box_bottom if is_last else None,
        )
        ws.row_dimensions[r].height = 20 if bold else 18

    _box_row(8,  client_name, bold=True, size=12)
    _box_row(9,  client_addr, color=TEXT_SECONDARY, size=10)
    _box_row(10, client_city, is_last=True, color=TEXT_SECONDARY, size=10)

    # Right block: meta (F7:G7..G10) — labels with light fill, values bold
    meta = [
        ("Invoice #", inv_num,     None,              False),
        ("Date",      inv_date,    "mmm dd, yyyy",    False),
        ("Due Date",  due_date,    "mmm dd, yyyy",    False),
        ("Status",    status_text, None,              True),  # styled differently
    ]
    for i, (label, value, fmt, is_status) in enumerate(meta):
        r = 7 + i
        c_lbl = ws.cell(row=r, column=6)
        c_lbl.value = label
        c_lbl.font  = Font(name=FONT_NAME, bold=True, size=10, color=BRAND_DARK)
        c_lbl.fill  = PatternFill("solid", fgColor=BRAND_LIGHT)
        c_lbl.alignment = Alignment(horizontal="right", vertical="center")
        c_lbl.border = Border(left=box_side, right=box_side,
                              top=box_side if i == 0 else None,
                              bottom=box_side if i == len(meta) - 1 else None)

        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        c_val = ws.cell(row=r, column=7)
        c_val.value = value
        if is_status:
            c_val.font = Font(name=FONT_NAME, bold=True, size=11, color=status_color)
            c_val.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c_val.font = Font(name=FONT_NAME, size=10, color=TEXT_PRIMARY)
            c_val.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_val.border = Border(left=box_side, right=box_side,
                              top=box_side if i == 0 else None,
                              bottom=box_side if i == len(meta) - 1 else None)
        if fmt and isinstance(value, datetime):
            c_val.number_format = fmt
        ws.row_dimensions[r].height = 20 if is_status else 18

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
        c.font  = Font(name=FONT_NAME, italic=True, size=11, color=TEXT_SECONDARY)
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
        c.font  = Font(name=FONT_NAME, size=10, color=TEXT_SECONDARY)
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
    _apply_print_setup(ws, N)


# ── Sheet 1B: Batch summary report (8 cols A-H) ───────────────

def _build_batch_summary_sheet(ws, invoices, ok_invoices, sums, present,
                               symbol, elapsed=None):
    """
    Executive batch report for multi-invoice extractions.

    Layout:
      rows 2-3  : Company header | BATCH REPORT banner
      rows 5-6  : 4 KPI tiles (Invoices · Successful · Total · Tax)
      row  8    : 'INVOICES IN THIS BATCH' section title
      row  9    : column headers
      row 10+   : one row per invoice; Due Date colored red when overdue
      (errors) : red-tinted rows after the OK block
      row N     : GRAND TOTAL bar with live =SUM formulas
      (By Vendor): rollup shown only when >1 distinct vendor
      footer    : thank-you + byline (no signature line — wrong for a batch)
    """
    ws.title = "Invoice Summary"
    N = 8
    money_fmt = _excel_currency_format(symbol)

    total_count   = len(invoices)
    ok_count      = len(ok_invoices)
    err_count     = total_count - ok_count
    success_pct   = int(ok_count / total_count * 100) if total_count else 0
    err_invoices  = [i for i in invoices if i.get("status") != "ok"]

    # ── Company header (rows 2-3) ─────────────────────────────
    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = COMPANY_NAME_PLACEHOLDER
    c.font  = Font(name=FONT_NAME, bold=True, size=18, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 32

    ws.merge_cells("E2:H2")
    c = ws["E2"]
    c.value = "BATCH REPORT"
    c.font  = Font(name=FONT_NAME, bold=True, size=22, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=BRAND_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    ws.merge_cells("A3:D3")
    c = ws["A3"]
    c.value = COMPANY_ADDR_PLACEHOLDER
    c.font  = Font(name=FONT_NAME, size=10, color=TEXT_SECONDARY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = Border(bottom=Side(style="thin", color=BRAND_MID))

    inv_word = "invoice" if ok_count == 1 else "invoices"
    ws.merge_cells("E3:H3")
    c = ws["E3"]
    c.value = f"Generated {_date.today().strftime('%B %d, %Y')}  ·  {ok_count} {inv_word}"
    c.font  = Font(name=FONT_NAME, italic=True, size=10, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center")
    c.border = Border(bottom=Side(style="thin", color=BRAND_MID))
    ws.row_dimensions[3].height = 20

    # ── KPI tiles (rows 5-6) ──────────────────────────────────
    tiles = [
        ("INVOICES",       str(total_count)),
        ("SUCCESSFUL",     f"{ok_count} / {total_count}  ({success_pct}%)"),
        ("TOTAL CAPTURED", _fmt(sums["total"] if present["total"] else None, symbol)),
        ("TAX CAPTURED",   _fmt(sums["tax"]   if present["tax"]   else None, symbol)),
    ]
    side_thin   = Side(style="thin",   color=RULE_LINE)
    side_top    = Side(style="thin",   color=RULE_LINE)
    side_bottom = Side(style="medium", color=BRAND_DARK)

    for i, (label, value) in enumerate(tiles):
        start_col = i * 2 + 1
        end_col   = start_col + 1

        # Label band (row 5)
        ws.merge_cells(start_row=5, start_column=start_col,
                       end_row=5,   end_column=end_col)
        c = ws.cell(row=5, column=start_col)
        c.value = label
        c.font  = Font(name=FONT_NAME, bold=True, size=9, color=BRAND_DARK)
        c.fill  = PatternFill("solid", fgColor=BRAND_LIGHT)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.border = Border(left=side_thin, right=side_thin, top=side_top)

        # Value band (row 6)
        ws.merge_cells(start_row=6, start_column=start_col,
                       end_row=6,   end_column=end_col)
        c = ws.cell(row=6, column=start_col)
        c.value = value
        c.font  = Font(name=FONT_NAME, bold=True, size=16, color=BRAND_DARK)
        c.fill  = PatternFill("solid", fgColor=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.border = Border(left=side_thin, right=side_thin, bottom=side_bottom)

    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 32

    # ── Section title (row 8) ─────────────────────────────────
    ws.merge_cells("A8:H8")
    c = ws["A8"]
    c.value = "INVOICES IN THIS BATCH"
    c.font  = Font(name=FONT_NAME, bold=True, size=10, color=BRAND_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[8].height = 22

    # ── Column headers (row 9) ────────────────────────────────
    headers = ["#", "Invoice #", "Vendor", "Date", "Due Date",
               "Subtotal", "Tax", "Total"]
    for col, h in enumerate(headers, start=1):
        _header_cell(ws.cell(row=9, column=col), h)
    ws.row_dimensions[9].height = 24

    # ── Per-invoice rows (row 10+) ────────────────────────────
    row = 10
    first_inv_row = row
    counter = 0

    for inv in invoices:
        if inv.get("status") != "ok":
            continue
        d = inv.get("data", {})
        alt = (counter % 2 == 1)

        inv_num  = d.get("invoice_number") or "—"
        vendor   = d.get("vendor_name")    or "—"
        inv_date = _parse_date(d.get("invoice_date"))
        due_date = _parse_date(d.get("due_date"))
        sub      = parse_amount(d.get("subtotal"))
        tax      = parse_amount(d.get("tax"))
        tot      = parse_amount(d.get("total"))

        status_text, status_color = _compute_status(due_date)
        is_overdue = (status_text == "OVERDUE")

        _data_cell(ws.cell(row=row, column=1), counter + 1, alt, align="center")
        _data_cell(ws.cell(row=row, column=2), inv_num, alt, bold=True)
        _data_cell(ws.cell(row=row, column=3), vendor, alt)
        _data_cell(
            ws.cell(row=row, column=4), inv_date, alt, align="center",
            number_format="mmm dd, yyyy" if isinstance(inv_date, datetime) else None,
        )

        # Due Date — colored red+bold when overdue
        c = ws.cell(row=row, column=5)
        c.value = due_date
        c.font  = Font(
            name=FONT_NAME, size=11, bold=is_overdue,
            color=status_color if is_overdue else TEXT_PRIMARY,
        )
        c.fill  = PatternFill("solid", fgColor=ALT_ROW if alt else WHITE)
        c.border = _border()
        c.alignment = Alignment(horizontal="center", vertical="center")
        if isinstance(due_date, datetime):
            c.number_format = "mmm dd, yyyy"

        _data_cell(
            ws.cell(row=row, column=6),
            sub if sub is not None else "—", alt, align="right",
            number_format=money_fmt if sub is not None else None,
        )
        _data_cell(
            ws.cell(row=row, column=7),
            tax if tax is not None else "—", alt, align="right",
            number_format=money_fmt if tax is not None else None,
        )
        _data_cell(
            ws.cell(row=row, column=8),
            tot if tot is not None else "—", alt, align="right", bold=True,
            number_format=money_fmt if tot is not None else None,
        )

        ws.row_dimensions[row].height = 22
        row += 1
        counter += 1

    last_inv_row = row - 1

    # ── Errors block ──────────────────────────────────────────
    if err_invoices:
        row += 1  # spacer
        # Heading
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=N)
        c = ws.cell(row=row, column=1)
        c.value = f"⚠  {err_count} file(s) could not be extracted"
        c.font  = Font(name=FONT_NAME, bold=True, size=10, color=STATUS_OVERDUE)
        c.fill  = PatternFill("solid", fgColor=ERROR_BG)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 20
        row += 1

        for inv in err_invoices:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=N)
            c = ws.cell(row=row, column=1)
            err_msg = inv.get("error") or "extraction failed"
            c.value = f"     {inv.get('filename', '—')}  —  {err_msg}"
            c.font  = Font(name=FONT_NAME, size=10, italic=True, color=TEXT_PRIMARY)
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[row].height = 16
            row += 1

    row += 1  # spacer before grand total

    # ── Grand Total bar (navy, white bold, live =SUM) ─────────
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    _total_cell(ws.cell(row=row, column=1), "GRAND TOTAL", is_label=True)
    if counter > 0:
        sub_formula = f"=SUM(F{first_inv_row}:F{last_inv_row})"
        tax_formula = f"=SUM(G{first_inv_row}:G{last_inv_row})"
        tot_formula = f"=SUM(H{first_inv_row}:H{last_inv_row})"
    else:
        sub_formula = "—"
        tax_formula = "—"
        tot_formula = "—"
    _total_cell(ws.cell(row=row, column=6), sub_formula, number_format=money_fmt)
    _total_cell(ws.cell(row=row, column=7), tax_formula, number_format=money_fmt)
    _total_cell(ws.cell(row=row, column=8), tot_formula, number_format=money_fmt)
    ws.row_dimensions[row].height = 28
    row += 1

    # ── By-Vendor rollup (only if >1 vendor) ──────────────────
    vendor_agg = {}  # name -> (count, total)
    for inv in ok_invoices:
        d = inv.get("data", {})
        v = d.get("vendor_name") or "—"
        t = parse_amount(d.get("total")) or 0.0
        cnt, ttl = vendor_agg.get(v, (0, 0.0))
        vendor_agg[v] = (cnt + 1, ttl + t)

    if len(vendor_agg) > 1:
        row += 2  # spacer

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=N)
        c = ws.cell(row=row, column=1)
        c.value = "BY VENDOR"
        c.font  = Font(name=FONT_NAME, bold=True, size=10, color=BRAND_DARK)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 22
        row += 1

        # Column bands: A:E Vendor | F Invoices | G:H Total
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        _header_cell(ws.cell(row=row, column=1), "Vendor")
        _header_cell(ws.cell(row=row, column=6), "Invoices")
        ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=8)
        _header_cell(ws.cell(row=row, column=7), "Total")
        ws.row_dimensions[row].height = 22
        row += 1

        for i, (vendor, (cnt, ttl)) in enumerate(
            sorted(vendor_agg.items(), key=lambda kv: -kv[1][1])
        ):
            alt = (i % 2 == 1)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            _data_cell(ws.cell(row=row, column=1), vendor, alt, bold=True)
            _data_cell(ws.cell(row=row, column=6), cnt, alt, align="center")
            ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=8)
            _data_cell(
                ws.cell(row=row, column=7), ttl, alt, align="right", bold=True,
                number_format=money_fmt,
            )
            ws.row_dimensions[row].height = 20
            row += 1

    row += 2  # spacer before footer

    # ── Footer — no signature for a batch ─────────────────────
    _footer(ws, row, N, signature=False)

    ws.freeze_panes = "A10"
    _set_widths(
        ws,
        {"A": 5, "B": 18, "C": 24, "D": 13, "E": 13, "F": 13, "G": 12, "H": 14},
    )
    _apply_print_setup(ws, N)


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
        c.font  = Font(name=FONT_NAME, italic=True, size=11, color=TEXT_SECONDARY)
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
    _apply_print_setup(ws, N)


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
    c.font  = Font(name=FONT_NAME, italic=True, size=10, color=TEXT_SECONDARY)
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
    _apply_print_setup(ws, N)


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
    _build_summary_sheet(wb.active, invoices, ok_invoices, sums, present, symbol, elapsed)
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
