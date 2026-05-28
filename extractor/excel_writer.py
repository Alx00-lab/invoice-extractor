import io
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Colors
HEADER_COLOR  = "1F3864"
ACCENT_COLOR  = "2E75B6"
LIGHT_BLUE    = "D6E4F0"
WARNING_COLOR = "FFF2CC"
ERROR_COLOR   = "F8CBAD"
WHITE         = "FFFFFF"
GRAY          = "F2F2F2"


def _thin_border():
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)


def _style_header_cell(cell, text):
    cell.value = text
    cell.font = Font(bold=True, color=WHITE, size=11)
    cell.fill = PatternFill("solid", fgColor=HEADER_COLOR)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = _thin_border()


def _style_value_cell(cell, value, alternate=False):
    cell.value = value if value not in (None, "null") else "—"
    cell.font = Font(size=10)
    cell.fill = PatternFill("solid", fgColor=GRAY if alternate else WHITE)
    cell.border = _thin_border()
    cell.alignment = Alignment(wrap_text=True, vertical="center")


def _autosize_columns(ws, max_width=50):
    """Compute column widths, safely skipping merged cells."""
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if not hasattr(cell, "column_letter"):
                continue  # MergedCell — skip
            if cell.value is None:
                continue
            col = cell.column_letter
            length = len(str(cell.value))
            if length > widths.get(col, 0):
                widths[col] = length
    for col, length in widths.items():
        ws.column_dimensions[col].width = min(max(length + 2, 12), max_width)


def generate_batch_excel(invoices: list) -> bytes:
    """
    Generate a consolidated Excel from multiple invoices.

    Args:
        invoices: list of dicts, each with keys:
            - filename (str)
            - data (dict)  -> parsed invoice data
            - status ("ok" | "error")
            - error (str, optional)
    """
    wb = Workbook()

    # ── Sheet 1: Summary ──────────────────────────────────────
    summary = wb.active
    summary.title = "Summary"

    # Title row
    summary.merge_cells("A1:I1")
    title = summary["A1"]
    title.value = "Invoice Extraction Report"
    title.font = Font(bold=True, size=14, color=WHITE)
    title.fill = PatternFill("solid", fgColor=ACCENT_COLOR)
    title.alignment = Alignment(horizontal="center", vertical="center")
    summary.row_dimensions[1].height = 28

    # Column headers
    headers = [
        "File", "Vendor", "Invoice #", "Invoice Date",
        "Due Date", "Subtotal", "Tax", "Total", "Status / Warning"
    ]
    for col, h in enumerate(headers, start=1):
        _style_header_cell(summary.cell(row=2, column=col), h)
    summary.row_dimensions[2].height = 32

    # Data rows
    row = 3
    for i, inv in enumerate(invoices):
        alt = (i % 2 == 1)
        filename = inv.get("filename", "—")

        if inv.get("status") == "error":
            _style_value_cell(summary.cell(row=row, column=1), filename, alt)
            for c in range(2, 9):
                _style_value_cell(summary.cell(row=row, column=c), "—", alt)
            err_cell = summary.cell(row=row, column=9)
            err_cell.value = f"⚠ ERROR: {inv.get('error', 'Unknown error')}"
            err_cell.font = Font(bold=True, color="9C2D2D", size=10)
            err_cell.fill = PatternFill("solid", fgColor=ERROR_COLOR)
            err_cell.border = _thin_border()
            err_cell.alignment = Alignment(wrap_text=True, vertical="center")
        else:
            d = inv.get("data", {})
            values = [
                filename,
                d.get("vendor_name"),
                d.get("invoice_number"),
                d.get("invoice_date"),
                d.get("due_date"),
                d.get("subtotal"),
                d.get("tax"),
                d.get("total"),
            ]
            for col, val in enumerate(values, start=1):
                _style_value_cell(summary.cell(row=row, column=col), val, alt)

            warning = d.get("validation_warning")
            warn_cell = summary.cell(row=row, column=9)
            if warning:
                warn_cell.value = f"⚠ {warning}"
                warn_cell.font = Font(bold=True, color="7B3F00", size=10)
                warn_cell.fill = PatternFill("solid", fgColor=WARNING_COLOR)
            else:
                warn_cell.value = "✓ OK"
                warn_cell.font = Font(color="2D6A3F", size=10)
                warn_cell.fill = PatternFill("solid", fgColor=GRAY if alt else WHITE)
            warn_cell.border = _thin_border()
            warn_cell.alignment = Alignment(wrap_text=True, vertical="center")

        summary.row_dimensions[row].height = 22
        row += 1

    # Totals row (sum of successful invoices)
    ok_invoices = [i for i in invoices if i.get("status") == "ok"]
    if ok_invoices:
        row += 1
        totals_label = summary.cell(row=row, column=1)
        totals_label.value = f"TOTALS ({len(ok_invoices)} invoices)"
        totals_label.font = Font(bold=True, size=11, color=WHITE)
        totals_label.fill = PatternFill("solid", fgColor=HEADER_COLOR)
        totals_label.alignment = Alignment(horizontal="center", vertical="center")
        summary.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)

        # We display totals as strings (preserving original formatting)
        for col_idx, key in [(6, "subtotal"), (7, "tax"), (8, "total")]:
            cell = summary.cell(row=row, column=col_idx)
            cell.value = "see individual rows"
            cell.font = Font(italic=True, size=9, color="666666")
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = _thin_border()
        summary.row_dimensions[row].height = 22

    # Freeze top header
    summary.freeze_panes = "A3"
    _autosize_columns(summary)

    # ── Sheet 2: Line Items ───────────────────────────────────
    items_ws = wb.create_sheet("Line Items")

    items_ws.merge_cells("A1:F1")
    li_title = items_ws["A1"]
    li_title.value = "All Line Items"
    li_title.font = Font(bold=True, size=14, color=WHITE)
    li_title.fill = PatternFill("solid", fgColor=ACCENT_COLOR)
    li_title.alignment = Alignment(horizontal="center", vertical="center")
    items_ws.row_dimensions[1].height = 28

    li_headers = ["File", "Invoice #", "Description", "Quantity", "Unit Price", "Amount"]
    for col, h in enumerate(li_headers, start=1):
        _style_header_cell(items_ws.cell(row=2, column=col), h)
    items_ws.row_dimensions[2].height = 28

    row = 3
    counter = 0
    for inv in invoices:
        if inv.get("status") != "ok":
            continue
        d = inv.get("data", {})
        line_items = d.get("line_items") or []
        for item in line_items:
            alt = (counter % 2 == 1)
            _style_value_cell(items_ws.cell(row=row, column=1), inv.get("filename"), alt)
            _style_value_cell(items_ws.cell(row=row, column=2), d.get("invoice_number"), alt)
            _style_value_cell(items_ws.cell(row=row, column=3), item.get("description"), alt)
            _style_value_cell(items_ws.cell(row=row, column=4), item.get("quantity"), alt)
            _style_value_cell(items_ws.cell(row=row, column=5), item.get("unit_price"), alt)
            _style_value_cell(items_ws.cell(row=row, column=6), item.get("amount"), alt)
            items_ws.row_dimensions[row].height = 18
            row += 1
            counter += 1

    if counter == 0:
        items_ws.merge_cells(f"A3:F3")
        empty_cell = items_ws["A3"]
        empty_cell.value = "No line items extracted"
        empty_cell.font = Font(italic=True, color="888888")
        empty_cell.alignment = Alignment(horizontal="center", vertical="center")

    items_ws.freeze_panes = "A3"
    _autosize_columns(items_ws)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    logger.info(f"Batch Excel generated with {len(invoices)} invoices.")
    return buffer.getvalue()


# ── Backward-compatible single-invoice generator ──────────────
def generate_excel(data: dict) -> bytes:
    """Single invoice — wraps batch generator for compatibility."""
    return generate_batch_excel([{
        "filename": "invoice.pdf",
        "data": data,
        "status": "ok"
    }])
