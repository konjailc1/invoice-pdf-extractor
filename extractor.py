"""Read one Vietnamese VAT invoice PDF.

The sample layout (see make_samples.py) follows the usual VAT invoice order:
seller block, red title, buyer block, items table, totals, amount in words.

Header fields are found with regular expressions. The items grid is read with
pdfplumber. Money uses a dot as the thousands separator (1.000.000).

Returned dictionary (missing values are None):

    {
        "file_name": "01_sen_vang.pdf",
        "form_number": "01GTKT0/001",
        "serial": "AA/26E",
        "invoice_number": "0000101",
        "issue_date": "12/01/2026",
        "seller_name": "CÔNG TY TNHH THƯƠNG MẠI SEN VÀNG",
        "seller_tax_code": "0112345601",
        "seller_address": "45 Nguyễn Văn Cừ, Phường Bồ Đề, TP. Hà Nội",
        "buyer_person": "Phạm Thu Trang",
        "buyer_unit": "Công ty TNHH Giải pháp Số Bình Minh",
        "buyer_tax_code": "0101112201",
        "buyer_address": "...",
        "payment_method": "CK",
        "subtotal": 3060000,
        "vat_rate": 10,
        "vat_amount": 306000,
        "total": 3366000,
        "amount_in_words": "Ba triệu ...",
        "line_items": [
            {"line_no": 1, "description": "Giấy in A4 70gsm", "unit": "ram",
             "qty": 20, "unit_price": 65000, "amount": 1300000},
        ],
    }
"""

import re
from pathlib import Path

import pdfplumber

TITLE = "HÓA ĐƠN GIÁ TRỊ GIA TĂNG"

FORM_PATTERN = re.compile(r"Mẫu số\s*:\s*(\S+)", re.IGNORECASE)
SERIAL_PATTERN = re.compile(r"Ký hiệu\s*:\s*(\S+)", re.IGNORECASE)
# "Mẫu số" also contains "số", so require the value to be a long digit string.
NUMBER_PATTERN = re.compile(r"(?<!Mẫu )Số\s*:\s*(\d{6,8})")
DATE_PATTERN = re.compile(
    r"Ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})",
    re.IGNORECASE,
)
# The monogram can sit on the same line as the company name, so this is not anchored.
SELLER_PATTERN = re.compile(r"(CÔNG TY[^\n]+|DOANH NGHIỆP[^\n]+|HỘ KINH DOANH[^\n]+)")
TAX_PATTERN = re.compile(r"Mã số thuế\s*:\s*([0-9][0-9 \-]*)", re.IGNORECASE)
ADDRESS_PATTERN = re.compile(r"Địa chỉ\s*:\s*(.+)", re.IGNORECASE)
BUYER_PERSON_PATTERN = re.compile(r"Họ và tên người mua hàng\s*:\s*(.+)", re.IGNORECASE)
BUYER_UNIT_PATTERN = re.compile(r"Tên đơn vị\s*:\s*(.+)", re.IGNORECASE)
PAYMENT_PATTERN = re.compile(
    r"Hình thức thanh toán\s*:\s*(TM/CK|TM|CK)",
    re.IGNORECASE,
)
SUBTOTAL_PATTERN = re.compile(
    r"Cộng tiền hàng hóa,\s*dịch vụ\s*:\s*([\d.]+)",
    re.IGNORECASE,
)
VAT_RATE_PATTERN = re.compile(r"Thuế suất GTGT\s*:\s*(\d+)\s*%", re.IGNORECASE)
VAT_AMOUNT_PATTERN = re.compile(r"Tiền thuế GTGT\s*:\s*([\d.]+)", re.IGNORECASE)
TOTAL_PATTERN = re.compile(r"Tổng tiền thanh toán\s*:\s*([\d.]+)", re.IGNORECASE)
WORDS_PATTERN = re.compile(r"Số tiền bằng chữ\s*:\s*(.+)", re.IGNORECASE)

DESCRIPTION_KEYS = ("tên hàng", "hàng hóa")
UNIT_KEYS = ("đvt", "đơn vị tính")
QTY_KEYS = ("số lượng",)
UNIT_PRICE_KEYS = ("đơn giá",)
AMOUNT_KEYS = ("thành tiền",)


def extract_invoice(pdf_path):
    """Extract one invoice. ``pdf_path`` may be a string or a Path."""
    path = Path(pdf_path)
    with pdfplumber.open(path) as pdf:
        pages = [page.extract_text() or "" for page in pdf.pages]
        items = []
        for page in pdf.pages:
            items.extend(extract_line_items(page))

    header = extract_header("\n".join(pages))
    header["file_name"] = path.name
    header["line_items"] = items
    return header


def extract_header(text):
    """Pull the form fields, parties, totals, and amount in words from page text."""
    cleaned = prepare_text(text)
    head, _, tail = cleaned.partition(TITLE)
    # Fields drawn beside the title can land on either side of the split.
    whole = cleaned

    date_match = DATE_PATTERN.search(whole)
    issue_date = None
    if date_match:
        day, month, year = (int(part) for part in date_match.groups())
        issue_date = f"{day:02d}/{month:02d}/{year}"

    return {
        "form_number": match_group(FORM_PATTERN, whole),
        "serial": match_group(SERIAL_PATTERN, whole),
        "invoice_number": match_group(NUMBER_PATTERN, whole),
        "issue_date": issue_date,
        "seller_name": match_group(SELLER_PATTERN, head),
        "seller_tax_code": clean_tax_code(match_group(TAX_PATTERN, head)),
        "seller_address": match_group(ADDRESS_PATTERN, head),
        "buyer_person": match_group(BUYER_PERSON_PATTERN, tail),
        "buyer_unit": match_group(BUYER_UNIT_PATTERN, tail),
        "buyer_tax_code": buyer_tax_code(tail),
        "buyer_address": match_group(ADDRESS_PATTERN, tail),
        "payment_method": match_group(PAYMENT_PATTERN, tail),
        "subtotal": parse_number(match_group(SUBTOTAL_PATTERN, whole)),
        "vat_rate": _int_or_none(match_group(VAT_RATE_PATTERN, whole)),
        "vat_amount": parse_number(match_group(VAT_AMOUNT_PATTERN, whole)),
        "total": parse_number(match_group(TOTAL_PATTERN, whole)),
        "amount_in_words": clean_words(match_group(WORDS_PATTERN, whole)),
    }


def extract_line_items(page):
    """Return line items from the goods table. The 1..6 index row is skipped."""
    table = find_items_table(page)
    if table is None:
        return []
    return table_to_items(table)


def prepare_text(text):
    """Normalize spaces and the colon character so the patterns stay simple."""
    text = text.replace("\u00a0", " ").replace("：", ":")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


def match_group(pattern, text):
    """Return a stripped capture, or None when the pattern is absent."""
    match = pattern.search(text)
    if not match:
        return None
    value = match.group(1).strip()
    return value or None


def buyer_tax_code(tail):
    """The buyer's MST sits on the payment-method line, not in the seller block."""
    match = re.search(
        r"Hình thức thanh toán\s*:\s*(?:TM/CK|TM|CK)\s+Mã số thuế\s*:\s*([0-9][0-9 \-]*)",
        tail,
        re.IGNORECASE,
    )
    if match:
        return clean_tax_code(match.group(1))
    # Fall back to a MST that appears after the buyer heading.
    return clean_tax_code(match_group(TAX_PATTERN, tail))


def clean_tax_code(value):
    """Turn a letter-spaced code like '0 1 1 2 3 4 5 6 0 1' into 0112345601.

    A branch code keeps its hyphen: 0112345610-001.
    """
    if not value:
        return None
    compact = re.sub(r"\s+", "", value)
    compact = compact.strip(".:-")
    if not compact or not re.search(r"\d", compact):
        return None
    return compact


def clean_words(value):
    """Collapse a wrapped amount-in-words line into one sentence."""
    if not value:
        return None
    # Stop if the next section header was pulled onto this capture.
    value = re.split(r"\bNgười mua hàng\b|\bNgười bán hàng\b|\bHóa đơn mẫu\b", value)[0]
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value or None


def parse_number(value):
    """Turn '1.250.000' or '2.000' into an int. Return None if it is not a number.

    Dots are thousands separators. A plain integer such as '20' stays 20.
    """
    if value is None:
        return None
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    text = text.replace("VND", "").replace("₫", "").replace("đ", "").strip(".,")
    if not text:
        return None
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
        return int(text.replace(".", ""))
    if re.fullmatch(r"\d{1,3}(,\d{3})+", text):
        return int(text.replace(",", ""))
    if re.fullmatch(r"\d+", text):
        return int(text)
    return None


def _int_or_none(value):
    if value is None:
        return None
    return int(value)


def find_items_table(page):
    """Try the default line finder, then a slightly looser one."""
    found = first_items_table(page.extract_tables() or [])
    if found is not None:
        return found
    loose = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 5,
        "join_tolerance": 5,
        "intersection_tolerance": 8,
    }
    return first_items_table(page.extract_tables(loose) or [])


def first_items_table(tables):
    """Return the first table whose header names the goods column and the amount."""
    for table in tables:
        if locate_header(table) is not None and column_map(table) is not None:
            return table
    return None


def table_to_items(table):
    """Convert the goods table into item dicts. Skip the header and the 1..6 row."""
    header_index = locate_header(table)
    columns = column_map(table)
    if header_index is None or columns is None:
        return []

    items = []
    for row in table[header_index + 1 :]:
        if is_index_row(row) or is_header_row(row):
            continue
        description = clean_cell(row, columns["description"])
        if not description:
            continue
        line_no = parse_number(clean_cell(row, columns.get("stt")))
        items.append(
            {
                "line_no": line_no if line_no is not None else len(items) + 1,
                "description": description,
                "unit": clean_cell(row, columns.get("unit")),
                "qty": parse_number(clean_cell(row, columns["qty"])),
                "unit_price": parse_number(clean_cell(row, columns["unit_price"])),
                "amount": parse_number(clean_cell(row, columns["amount"])),
            }
        )
    return items


def column_map(table):
    """Map description, unit, qty, unit price, and amount. Unit may be absent."""
    header_index = locate_header(table)
    if header_index is None:
        return None
    header = [_normalize(cell).lower() for cell in table[header_index]]
    mapping = {
        "stt": find_column(header, ("stt",)),
        "description": find_column(header, DESCRIPTION_KEYS),
        "unit": find_column(header, UNIT_KEYS),
        "qty": find_column(header, QTY_KEYS),
        "unit_price": find_column(header, UNIT_PRICE_KEYS),
        "amount": find_column(header, AMOUNT_KEYS),
    }
    required = ("description", "qty", "unit_price", "amount")
    if any(mapping[name] is None for name in required):
        return None
    return mapping


def locate_header(table):
    """Find the header row. It is almost always row 0."""
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
    """Return the header index that matches one of ``keys``."""
    for key in keys:
        for index, name in enumerate(names):
            if name == key or name.startswith(key + " ") or name.startswith(key + "("):
                return index
            if len(key) >= 4 and key in name:
                return index
    return None


def is_index_row(row):
    """True for the column-index row '1 2 3 4 5 6=4x5'."""
    cells = [_normalize(cell) for cell in row]
    if "6=4x5" in cells or "6=4×5" in cells:
        return True
    return cells[:3] == ["1", "2", "3"]


def is_header_row(row):
    """True when a later row repeats the column titles."""
    names = [_normalize(cell).lower() for cell in row]
    return find_column(names, DESCRIPTION_KEYS) is not None and find_column(names, AMOUNT_KEYS) is not None


def clean_cell(row, index):
    """Cell text with line breaks turned into spaces."""
    if index is None or index >= len(row):
        return ""
    return _normalize(row[index])


def _normalize(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()
