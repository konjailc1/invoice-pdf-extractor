"""Read a supplier invoice PDF and return its header fields and line items.

The sample PDFs use three layouts with different labels (see make_samples.py).
Header fields are found with regular expressions. The line-item grid is found
with pdfplumber's table reader.

Returned dictionary:

    {
        "file_name": "01_minh_phat.pdf",
        "invoice_number": "HD-2026-0142",
        "issue_date": "12/01/2026",
        "seller_name": "Công ty TNHH Thương mại Minh Phát",
        "tax_code": "0314567890",      # None when the PDF has no tax code
        "buyer_name": "Công ty TNHH Giải pháp Số Horizon",
        "subtotal": 6645000,
        "vat_rate": 10,
        "vat_amount": 664500,
        "total": 7309500,
        "line_items": [
            {"description": "Giấy in A4 Double A 70gsm", "qty": 10,
             "unit_price": 285000, "amount": 2850000},
        ],
    }
"""

import re
from pathlib import Path

import pdfplumber

# Each pattern accepts the wording used by the classic, modern, and compact layouts.
INVOICE_NUMBER_PATTERN = re.compile(
    r"(?:Số\s+hóa\s+đơn|Invoice\s+No\.?|Mã\s+HĐ|Mã\s+hóa\s+đơn)\s*:\s*(\S+)",
    re.IGNORECASE,
)
ISSUE_DATE_PATTERN = re.compile(
    r"(?:Ngày\s+lập|Ngày|Date)\s*:\s*"
    r"(\d{1,2}[./-]\d{1,2}[./-]\d{4}|\d{4}[./-]\d{1,2}[./-]\d{1,2})",
    re.IGNORECASE,
)
SELLER_PATTERN = re.compile(
    r"(?:Người\s+bán|Bên\s+bán|Đơn\s+vị\s+bán)\s*:\s*(.+)",
    re.IGNORECASE,
)
BUYER_PATTERN = re.compile(
    r"(?:Người\s+mua|Bên\s+mua|Khách\s+hàng)\s*:\s*(.+)",
    re.IGNORECASE,
)
TAX_CODE_PATTERN = re.compile(
    r"(?:Mã\s+số\s+thuế|MST)\s*:\s*([0-9][0-9\-]*)",
    re.IGNORECASE,
)
SUBTOTAL_PATTERN = re.compile(
    r"(?:Cộng\s+tiền\s+hàng|Tạm\s+tính|Tiền\s+hàng)\s*:\s*([\d\.,]+)\s*(?:VND)?",
    re.IGNORECASE,
)
VAT_PATTERN = re.compile(
    r"(?:Thuế\s+GTGT|VAT|GTGT)\s*\(?\s*(\d+)\s*%\)?\s*:\s*([\d\.,]+)\s*(?:VND)?",
    re.IGNORECASE,
)
TOTAL_PATTERN = re.compile(
    r"(?:Tổng\s+cộng\s+thanh\s+toán|Tổng\s+tiền|Số\s+tiền\s+phải\s+trả)"
    r"\s*:\s*([\d\.,]+)\s*(?:VND)?",
    re.IGNORECASE,
)

# Header cells we know how to map onto description / qty / unit price / amount.
DESCRIPTION_KEYS = ("mô tả", "hàng hóa", "mặt hàng", "diễn giải")
QTY_KEYS = ("số lượng", "qty", "sl")
UNIT_PRICE_KEYS = ("đơn giá",)
AMOUNT_KEYS = ("thành tiền",)


def extract_invoice(pdf_path):
    """Extract one invoice. ``pdf_path`` may be a string or a Path.

    Raises the underlying pdfplumber error if the file cannot be opened.
    Missing fields are returned as None, and missing tables as an empty list,
    so the validator can report them instead of the program crashing.
    """
    path = Path(pdf_path)
    with pdfplumber.open(path) as pdf:
        text_parts = []
        line_items = []
        for page in pdf.pages:
            text_parts.append(page.extract_text() or "")
            line_items.extend(extract_line_items(page))

    header = extract_header("\n".join(text_parts))
    header["file_name"] = path.name
    header["line_items"] = line_items
    return header


def extract_header(text):
    """Pull invoice number, parties, tax code, and totals out of page text."""
    cleaned = prepare_text(text)
    vat_match = VAT_PATTERN.search(cleaned)
    return {
        "invoice_number": match_group(INVOICE_NUMBER_PATTERN, cleaned),
        "issue_date": match_group(ISSUE_DATE_PATTERN, cleaned),
        "seller_name": match_group(SELLER_PATTERN, cleaned),
        "tax_code": match_group(TAX_CODE_PATTERN, cleaned),
        "buyer_name": match_group(BUYER_PATTERN, cleaned),
        "subtotal": parse_number(match_group(SUBTOTAL_PATTERN, cleaned)),
        "vat_rate": int(vat_match.group(1)) if vat_match else None,
        "vat_amount": parse_number(vat_match.group(2)) if vat_match else None,
        "total": parse_number(match_group(TOTAL_PATTERN, cleaned)),
    }


def extract_line_items(page):
    """Return the line items from the first real items table on this page.

    A title banner or a totals bar is ignored. A table counts only when its
    header row has a description column and an amount column.
    """
    table = find_items_table(page)
    if table is None:
        return []
    return table_to_items(table)


def prepare_text(text):
    """Make small punctuation differences so the patterns can stay simple."""
    text = text.replace("\u00a0", " ").replace("：", ":")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


def match_group(pattern, text, group=1):
    """Return a stripped capture, or None when the pattern is not in the text."""
    match = pattern.search(text)
    if not match:
        return None
    value = match.group(group).strip()
    return value or None


def parse_number(value):
    """Turn '1.250.000' or '5.000' into an int. Return None if it is not a number.

    Dots are thousands separators, which is the Vietnamese way of writing VND.
    A plain integer such as '20' is left as 20.
    """
    if value is None:
        return None
    text = str(value).strip()
    text = text.replace("\u00a0", "").replace(" ", "")
    text = text.replace("VND", "").replace("₫", "").replace("đ", "")
    text = text.strip(".,")
    if not text:
        return None
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
        return int(text.replace(".", ""))
    if re.fullmatch(r"\d{1,3}(,\d{3})+", text):
        return int(text.replace(",", ""))
    if re.fullmatch(r"\d+", text):
        return int(text)
    return None


def find_items_table(page):
    """Try line detection first, then a text-based fallback."""
    tables = page.extract_tables() or []
    found = first_items_table(tables)
    if found is not None:
        return found

    # Some PDFs draw rules that sit a fraction off the cell border.
    loose_lines = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 5,
        "join_tolerance": 5,
        "intersection_tolerance": 8,
    }
    tables = page.extract_tables(loose_lines) or []
    found = first_items_table(tables)
    if found is not None:
        return found

    by_text = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 4,
    }
    return first_items_table(page.extract_tables(by_text) or [])


def first_items_table(tables):
    """Return the first table that looks like a line-item grid."""
    for table in tables:
        if locate_header(table) is not None and column_map(table) is not None:
            return table
    return None


def table_to_items(table):
    """Convert a pdfplumber table into a list of item dictionaries."""
    header_index = locate_header(table)
    columns = column_map(table)
    if header_index is None or columns is None:
        return []

    items = []
    for row in table[header_index + 1 :]:
        description = clean_cell(row, columns["description"])
        if not description or is_header_text(description):
            continue
        if description.lower().startswith(("cộng", "tổng", "tạm tính", "total")):
            continue
        items.append(
            {
                "description": description,
                "qty": parse_number(clean_cell(row, columns["qty"])),
                "unit_price": parse_number(clean_cell(row, columns["unit_price"])),
                "amount": parse_number(clean_cell(row, columns["amount"])),
            }
        )
    return items


def column_map(table):
    """Map the columns we need. Return None if any of them is missing."""
    header_index = locate_header(table)
    if header_index is None:
        return None
    header = [_normalize(cell).lower() for cell in table[header_index]]
    mapping = {
        "description": find_column(header, DESCRIPTION_KEYS),
        "qty": find_column(header, QTY_KEYS),
        "unit_price": find_column(header, UNIT_PRICE_KEYS),
        "amount": find_column(header, AMOUNT_KEYS),
    }
    if any(index is None for index in mapping.values()):
        return None
    return mapping


def locate_header(table):
    """Find the header row. It is almost always row 0, but we check the first three."""
    if not table:
        return None
    for index, row in enumerate(table[:3]):
        names = [_normalize(cell).lower() for cell in row]
        if find_column(names, DESCRIPTION_KEYS) is None:
            continue
        if find_column(names, AMOUNT_KEYS) is None:
            continue
        return index
    return None


def find_column(names, keys):
    """Return the index of the header cell that matches one of ``keys``.

    Short keys such as 'sl' must match the whole cell (or the start of it).
    Longer keys may match inside a header like 'Đơn giá (VND)'.
    """
    for key in keys:
        for index, name in enumerate(names):
            if name == key or name.startswith(key + " ") or name.startswith(key + "("):
                return index
            if len(key) >= 4 and key in name:
                return index
    return None


def clean_cell(row, index):
    """Cell text with internal line breaks turned into single spaces."""
    if index is None or index >= len(row):
        return ""
    return _normalize(row[index])


def is_header_text(description):
    """True when a body row is really a repeated header."""
    lowered = description.lower()
    return any(key in lowered for key in DESCRIPTION_KEYS) and "thành tiền" in lowered


def _normalize(value):
    """Trim a cell and collapse whitespace. None becomes an empty string."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()
