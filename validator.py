"""Check one extracted invoice and return a list of problems.

An empty list means the invoice passed. Each problem has a Vietnamese
message and an English message:

    {"vi": "Thiếu mã số thuế người bán", "en": "Missing seller tax code"}

Checks:
- required fields are present
- seller and buyer tax codes are 10 digits, or 13 digits for a branch
- the line amounts add up to the subtotal
- subtotal x VAT rate equals the VAT amount
- subtotal + VAT equals the total
- the amount in words matches the total
"""

import re

from vnd_words import amount_in_words, normalize_amount_words

# (field name, Vietnamese message, English message)
REQUIRED_FIELDS = (
    ("form_number", "Thiếu mẫu số", "Missing form number"),
    ("serial", "Thiếu ký hiệu", "Missing serial"),
    ("invoice_number", "Thiếu số hóa đơn", "Missing invoice number"),
    ("issue_date", "Thiếu ngày lập", "Missing issue date"),
    ("seller_name", "Thiếu tên người bán", "Missing seller name"),
    ("seller_address", "Thiếu địa chỉ người bán", "Missing seller address"),
    ("buyer_unit", "Thiếu tên đơn vị mua", "Missing buyer name"),
    ("buyer_address", "Thiếu địa chỉ người mua", "Missing buyer address"),
    ("payment_method", "Thiếu hình thức thanh toán", "Missing payment method"),
    ("subtotal", "Thiếu tiền hàng", "Missing subtotal"),
    ("vat_rate", "Thiếu thuế suất GTGT", "Missing VAT rate"),
    ("vat_amount", "Thiếu tiền thuế GTGT", "Missing VAT amount"),
    ("total", "Thiếu tổng tiền thanh toán", "Missing total"),
    ("amount_in_words", "Thiếu số tiền bằng chữ", "Missing amount in words"),
)


def issue(vi, en):
    """One problem, written in Vietnamese and in English."""
    return {"vi": vi, "en": en}


def validate_invoice(invoice):
    """Return a list of issue dicts. No issues means the invoice is OK."""
    issues = []

    for field_name, vi, en in REQUIRED_FIELDS:
        if is_blank(invoice.get(field_name)):
            issues.append(issue(vi, en))

    issues.extend(tax_code_issues(invoice.get("seller_tax_code"), "người bán", "Seller"))
    issues.extend(tax_code_issues(invoice.get("buyer_tax_code"), "người mua", "Buyer"))
    issues.extend(line_amount_issues(invoice))
    issues.extend(vat_amount_issues(invoice))
    issues.extend(total_issues(invoice))
    issues.extend(words_issues(invoice))
    return issues


def tax_code_issues(tax_code, who_vi, who_en):
    """A Vietnamese MST is 10 digits, or 13 digits written as 0123456789-001."""
    if is_blank(tax_code):
        return [
            issue(
                f"Thiếu mã số thuế {who_vi}",
                f"Missing {who_en.lower()} tax code",
            )
        ]

    digits = re.sub(r"\D", "", str(tax_code))
    if len(digits) not in (10, 13):
        return [
            issue(
                f"Mã số thuế {who_vi} '{tax_code}' không phải 10 hoặc 13 chữ số",
                f"{who_en} tax code '{tax_code}' is not 10 or 13 digits",
            )
        ]
    return []


def line_amount_issues(invoice):
    """The printed subtotal should equal the sum of the line amounts."""
    items = invoice.get("line_items") or []
    if not items:
        return [issue("Không tìm thấy dòng hàng", "No line items found")]

    missing_rows = [
        str(item.get("line_no") or index)
        for index, item in enumerate(items, start=1)
        if item.get("amount") is None
    ]
    if missing_rows:
        rows = ", ".join(missing_rows)
        return [issue(f"Dòng {rows} thiếu thành tiền", f"Line {rows} is missing an amount")]

    if invoice.get("subtotal") is None:
        return []

    line_sum = sum(item["amount"] for item in items)
    subtotal = invoice["subtotal"]
    if line_sum != subtotal:
        return [
            issue(
                f"Tổng thành tiền các dòng ({vi_num(line_sum)}) không khớp tiền hàng ({vi_num(subtotal)})",
                f"Sum of line amounts ({en_num(line_sum)}) does not equal subtotal ({en_num(subtotal)})",
            )
        ]
    return []


def vat_amount_issues(invoice):
    """VAT amount should equal subtotal times the printed VAT rate."""
    subtotal = invoice.get("subtotal")
    rate = invoice.get("vat_rate")
    vat_amount = invoice.get("vat_amount")
    if None in (subtotal, rate, vat_amount):
        return []

    if subtotal * rate != vat_amount * 100:
        expected = subtotal * rate // 100
        return [
            issue(
                f"Tiền hàng × thuế suất ({vi_num(expected)}) không khớp tiền thuế GTGT ({vi_num(vat_amount)})",
                f"Subtotal × VAT rate ({en_num(expected)}) does not equal VAT amount ({en_num(vat_amount)})",
            )
        ]
    return []


def total_issues(invoice):
    """The printed total should equal subtotal + VAT."""
    subtotal = invoice.get("subtotal")
    vat_amount = invoice.get("vat_amount")
    total = invoice.get("total")
    if None in (subtotal, vat_amount, total):
        return []

    expected = subtotal + vat_amount
    if expected != total:
        return [
            issue(
                f"Tiền hàng + VAT ({vi_num(expected)}) không khớp tổng tiền ({vi_num(total)})",
                f"Subtotal + VAT ({en_num(expected)}) does not equal total ({en_num(total)})",
            )
        ]
    return []


def words_issues(invoice):
    """The sentence 'Số tiền bằng chữ' should read the same total that was printed."""
    total = invoice.get("total")
    words = invoice.get("amount_in_words")
    if total is None or is_blank(words):
        return []

    expected = normalize_amount_words(amount_in_words(total))
    actual = normalize_amount_words(words)
    if actual != expected:
        return [
            issue(
                f"Số tiền bằng chữ không khớp tổng tiền ({vi_num(total)})",
                f"Amount in words does not match the total ({en_num(total)})",
            )
        ]
    return []


def vi_num(amount):
    """1.000.000, the grouping used on the invoice itself."""
    return f"{amount:,}".replace(",", ".")


def en_num(amount):
    """1,000,000, the grouping used in the English message."""
    return f"{amount:,}"


def is_blank(value):
    """True for None and empty strings. The number 0 is a real value, not a blank."""
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False
