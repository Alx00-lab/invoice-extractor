import io
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Colors
HEADER_COLOR = "1F3864"   # dark navy
ACCENT_COLOR = "2E75B6"   # blue
LIGHT_BLUE   = "D6E4F0"
WARNING_COLOR = "FFF2CC"
WHITE = "FFFFFF"
GRAY = "F2F2F2"


def _thin_border():
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)


def _header_style(cell, text):
    cell.value = text
    cell.font = Font(bold=True, color=WHITE, size=11)
    cell.fill = PatternFill("solid", fgColor=HEADER_COLOR)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _thin_border()


def _label_style(cell, text):
    cell.value = text
    cell.font = Font(bold=True, color="1F3864", size=10)
    cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    cell.border = _thin_border()


def _value_style(cell, text):
    cell.value = text if text not in (None, "null") else "—"
    cell.font = Font(size=10)
    cell.fill = PatternFill("solid", fgColor=WHITE)
    cell.border = _thin_border()
    cell.alignment = Alignment(wrap_text=True)


def generate_excel(data: dict) -> bytes:
    """
    Receives parsed invoice dict, returns Excel file as bytes.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Data"

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 35

    # ── Title ────────────────────────────────────────────────
    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = "Invoice Extraction Report"
    title_cell.font = Font(bold=True, size=14, color=WHITE)
    title_cell.fill = PatternFill("solid", fgColor=ACCENT_COLOR)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # ── Invoice Header Fields ─────────────────────────────────
    ws.merge_cells("A2:D2")
    ws["A2"].value = "INVOICE INFORMATION"
    ws["A2"].font = Font(bold=True, size=10, color=HEADER_COLOR)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18

    fields = [
        ("Vendor / Supplier",  data.get("vendor_name")),
        ("Invoice Number",     data.get("invoice_number")),
        ("Invoice Date",       data.get("invoice_date")),
        ("Due Date",           data.get("due_date")),
        ("Payment Terms",      data.get("payment_terms")),
    ]

    row = 3
    for i in range(0, len(fields), 2):
        left_label, left_value = fields[i]
        _label_style(ws.cell(row=row, column=1), left_label)
        _value_style(ws.cell(row=row, column=2), left_value)

        if i + 1 < len(fields):
            right_label, right_value = fields[i + 1]
            _label_style(ws.cell(row=row, column=3), right_label)
            _value_style(ws.cell(row=row, column=4), right_value)

        ws.row_dimensions[row].height = 18
        row += 1

    # ── Line Items ────────────────────────────────────────────
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"].value = "LINE ITEMS"
    ws[f"A{row}"].font = Font(bold=True, size=10, color=HEADER_COLOR)
    ws[f"A{row}"].fill = PatternFill("solid", fgColor=GRAY)
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 18
    row += 1

    li_headers = ["Description", "Quantity", "Unit Price", "Amount"]
    for col, h in enumerate(li_headers, start=1):
        _header_style(ws.cell(row=row, column=col), h)
    ws.row_dimensions[row].height = 20
    row += 1

    line_items = data.get("line_items") or []
    if line_items:
        for item in line_items:
            fill = PatternFill("solid", fgColor=WHITE if row % 2 == 0 else GRAY)
            values = [
                item.get("description"),
                item.get("quantity"),
                item.get("unit_price"),
                item.get("amount"),
            ]
            for col, val in enumerate(values, start=1):
                c = ws.cell(row=row, column=col)
                _value_style(c, val)
                c.fill = fill
            ws.row_dimensions[row].height = 16
            row += 1
    else:
        ws.merge_cells(f"A{row}:D{row}")
        ws[f"A{row}"].value = "No line items found"
        ws[f"A{row}"].font = Font(italic=True, color="888888")
        row += 1

    # ── Totals ────────────────────────────────────────────────
    row += 1
    totals = [
        ("Subtotal",  data.get("subtotal")),
        ("Tax",       data.get("tax")),
        ("TOTAL",     data.get("total")),
    ]
    for label, value in totals:
        ws.merge_cells(f"A{row}:B{row}")
        lc = ws[f"A{row}"]
        lc.value = label
        lc.font = Font(bold=True, size=10, color="1F3864")
        lc.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        lc.border = _thin_border()

        ws.merge_cells(f"C{row}:D{row}")
        vc = ws[f"C{row}"]
        _value_style(vc, value)
        if label == "TOTAL":
            vc.font = Font(bold=True, size=11, color="1F3864")
        ws.row_dimensions[row].height = 18
        row += 1

    # ── Validation Warning ────────────────────────────────────
    warning = data.get("validation_warning")
    if warning:
        row += 1
        ws.merge_cells(f"A{row}:D{row}")
        wc = ws[f"A{row}"]
        wc.value = f"⚠ {warning}"
        wc.font = Font(bold=True, color="7B3F00", size=10)
        wc.fill = PatternFill("solid", fgColor=WARNING_COLOR)
        wc.alignment = Alignment(wrap_text=True, horizontal="left")
        wc.border = _thin_border()
        ws.row_dimensions[row].height = 30

    # ── Freeze top row ────────────────────────────────────────
    ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    logger.info("Excel file generated successfully.")
    return buffer.getvalue()
