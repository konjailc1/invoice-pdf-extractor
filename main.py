"""Turn a folder of invoice PDFs into one Excel workbook.

Usage:

    python main.py sample_input
    python main.py sample_input -o output/invoices.xlsx

The workbook has three sheets:
- Invoices: one row per file, with status OK or ERROR
- Line Items: one row per item
- Issues: one row per problem found by the validator

ERROR rows on the Invoices sheet are filled red. Money columns use the
Excel format #,##0.
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from extractor import extract_invoice
from validator import validate_invoice

INVOICE_COLUMNS = [
    "File",
    "Invoice number",
    "Issue date",
    "Seller",
    "Tax code",
    "Buyer",
    "Subtotal (VND)",
    "VAT rate",
    "VAT amount (VND)",
    "Total (VND)",
    "Status",
]
LINE_COLUMNS = [
    "File",
    "Invoice number",
    "Line",
    "Description",
    "Qty",
    "Unit price (VND)",
    "Amount (VND)",
]
ISSUE_COLUMNS = ["File", "Invoice number", "Issue"]

MONEY_COLUMNS = {
    "Subtotal (VND)",
    "VAT amount (VND)",
    "Total (VND)",
    "Unit price (VND)",
    "Amount (VND)",
}

HEADER_FILL = PatternFill(start_color="1B3A4B", end_color="1B3A4B", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
ERROR_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
ERROR_FONT = Font(name="Calibri", color="9C0006")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
TAB_COLORS = {"Invoices": "1B3A4B", "Line Items": "0F6E6B", "Issues": "9C0006"}


def build_parser():
    """Command-line arguments for the extractor."""
    parser = argparse.ArgumentParser(
        description="Extract supplier invoices from PDFs and write an Excel workbook."
    )
    parser.add_argument("input_dir", help="Folder that contains invoice PDF files")
    parser.add_argument(
        "-o",
        "--output",
        default=str(Path("output") / "invoices.xlsx"),
        help="Where to write the workbook (default: output/invoices.xlsx)",
    )
    return parser


def main(argv=None):
    """Read every PDF in the folder, validate it, and write the workbook."""
    args = build_parser().parse_args(argv)
    input_dir = Path(args.input_dir)
    output_path = Path(args.output)

    if not input_dir.is_dir():
        print(f"Input folder not found: {input_dir}")
        return 1

    pdf_paths = sorted(input_dir.glob("*.pdf"))
    if not pdf_paths:
        print(f"No PDF files in {input_dir}")
        return 1

    invoice_rows = []
    line_rows = []
    issue_rows = []

    for pdf_path in pdf_paths:
        invoice, issues = read_one(pdf_path)
        status = "OK" if not issues else "ERROR"
        invoice_rows.append(invoice_row(invoice, status))
        line_rows.extend(line_item_rows(invoice))
        issue_rows.extend(issue_row(invoice, issue) for issue in issues)
        print_result(pdf_path.name, status, issues)

    try:
        write_workbook(invoice_rows, line_rows, issue_rows, output_path)
    except PermissionError:
        print(f"Could not write {output_path}. Close the file if it is open in Excel.")
        return 1

    ok_count = sum(1 for row in invoice_rows if row["Status"] == "OK")
    error_count = len(invoice_rows) - ok_count
    print(f"Processed {len(invoice_rows)} invoices: {ok_count} OK, {error_count} ERROR")
    print(f"Wrote {output_path}")
    return 0


def read_one(pdf_path):
    """Extract and validate one PDF. A broken file becomes an ERROR row."""
    try:
        invoice = extract_invoice(pdf_path)
    except Exception as exc:  # a bad PDF should not stop the rest of the folder
        invoice = {
            "file_name": pdf_path.name,
            "invoice_number": None,
            "issue_date": None,
            "seller_name": None,
            "tax_code": None,
            "buyer_name": None,
            "subtotal": None,
            "vat_rate": None,
            "vat_amount": None,
            "total": None,
            "line_items": [],
        }
        return invoice, [f"Could not read PDF: {exc}"]

    return invoice, validate_invoice(invoice)


def invoice_row(invoice, status):
    """One Invoices-sheet row."""
    return {
        "File": invoice.get("file_name"),
        "Invoice number": invoice.get("invoice_number"),
        "Issue date": parse_issue_date(invoice.get("issue_date")),
        "Seller": invoice.get("seller_name"),
        "Tax code": empty_as_none(invoice.get("tax_code")),
        "Buyer": invoice.get("buyer_name"),
        "Subtotal (VND)": invoice.get("subtotal"),
        "VAT rate": invoice.get("vat_rate"),
        "VAT amount (VND)": invoice.get("vat_amount"),
        "Total (VND)": invoice.get("total"),
        "Status": status,
    }


def line_item_rows(invoice):
    """One Line Items row per extracted item."""
    rows = []
    for index, item in enumerate(invoice.get("line_items") or [], start=1):
        rows.append(
            {
                "File": invoice.get("file_name"),
                "Invoice number": invoice.get("invoice_number"),
                "Line": index,
                "Description": item.get("description"),
                "Qty": item.get("qty"),
                "Unit price (VND)": item.get("unit_price"),
                "Amount (VND)": item.get("amount"),
            }
        )
    return rows


def issue_row(invoice, issue):
    """One Issues-sheet row."""
    return {
        "File": invoice.get("file_name"),
        "Invoice number": invoice.get("invoice_number"),
        "Issue": issue,
    }


def parse_issue_date(value):
    """Turn 12/01/2026 or 18.02.2026 into a real date. Unknown text is kept as-is."""
    if not value:
        return None
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(value), fmt).date()
        except ValueError:
            continue
    return value


def empty_as_none(value):
    """Store a blank tax code as an empty Excel cell rather than the word None."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def write_workbook(invoice_rows, line_rows, issue_rows, output_path):
    """Write the three sheets, then apply number formats, widths, and ERROR highlighting."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    frames = {
        "Invoices": pd.DataFrame(invoice_rows, columns=INVOICE_COLUMNS),
        "Line Items": pd.DataFrame(line_rows, columns=LINE_COLUMNS),
        "Issues": pd.DataFrame(issue_rows, columns=ISSUE_COLUMNS),
    }
    # Tax codes are text. A numeric column would drop the leading zero.
    frames["Invoices"]["Tax code"] = frames["Invoices"]["Tax code"].astype("string")
    frames["Invoices"]["Invoice number"] = frames["Invoices"]["Invoice number"].astype("string")
    frames["Line Items"]["Invoice number"] = frames["Line Items"]["Invoice number"].astype("string")
    frames["Issues"]["Invoice number"] = frames["Issues"]["Invoice number"].astype("string")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, frame in frames.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)

    style_workbook(output_path)


def style_workbook(output_path):
    """Format the file that pandas just wrote."""
    from openpyxl import load_workbook

    workbook = load_workbook(output_path)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet.sheet_properties.tabColor = TAB_COLORS.get(sheet_name, "1B3A4B")
        style_sheet(sheet)
        if sheet_name == "Invoices":
            highlight_error_rows(sheet)
    workbook.save(output_path)


def style_sheet(sheet):
    """Bold header, borders, number formats, and column widths."""
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 22

    headers = [cell.value for cell in sheet[1]]
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell, header in zip(row, headers):
            cell.border = THIN_BORDER
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = Alignment(vertical="center", horizontal=column_alignment(header))
            apply_number_format(cell, header)

    autosize_columns(sheet, headers)


def column_alignment(header):
    """Right-align numbers, center short codes, left-align text."""
    if header in MONEY_COLUMNS or header in ("Qty", "Line", "VAT rate"):
        return "right"
    if header in ("Status", "Issue date"):
        return "center"
    return "left"


def apply_number_format(cell, header):
    """VND amounts use #,##0. Dates and the VAT rate get their own formats."""
    if header == "Tax code":
        cell.number_format = "@"
    if cell.value is None:
        return
    if header in MONEY_COLUMNS or header == "Qty":
        cell.number_format = "#,##0"
    elif header == "VAT rate":
        cell.number_format = '0"%"'
    elif header == "Issue date":
        cell.number_format = "DD/MM/YYYY"


def highlight_error_rows(sheet):
    """Fill every cell in an ERROR invoice row."""
    headers = [cell.value for cell in sheet[1]]
    try:
        status_index = headers.index("Status")
    except ValueError:
        return

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=sheet.max_column):
        if row[status_index].value != "ERROR":
            continue
        for cell in row:
            cell.fill = ERROR_FILL
            cell.font = ERROR_FONT


def autosize_columns(sheet, headers):
    """Set each column just wide enough for its contents, with a sensible cap."""
    caps = {"Issue": 78, "Description": 42, "Seller": 46, "Buyer": 46}
    for index, header in enumerate(headers, start=1):
        letter = get_column_letter(index)
        longest = len(str(header))
        for cell in sheet[letter]:
            longest = max(longest, display_length(cell.value, header))
        cap = caps.get(header, 38)
        # Extra room so the filter arrow does not cover the header text.
        width = min(max(longest + 5, 14), cap)
        if header in MONEY_COLUMNS:
            width = max(width, 18)
        sheet.column_dimensions[letter].width = width


def display_length(value, header):
    """Length used for column width. Dates are measured as DD/MM/YYYY, not as timestamps."""
    if value is None:
        return 0
    if header == "Issue date" and hasattr(value, "strftime"):
        return len(value.strftime("%d/%m/%Y"))
    return len(str(value))


def print_result(file_name, status, issues):
    """Print one line per invoice, plus the problems under ERROR rows."""
    print(f"{file_name:<24} {status}")
    for issue in issues:
        print(f"  - {issue}")


if __name__ == "__main__":
    sys.exit(main())
