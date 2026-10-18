"""Check one extracted invoice and return a list of problems.

An empty list means the invoice passed. Each problem is a short English
sentence that can go straight into the Issues sheet.

Checks:
- required fields are present (number, date, seller, buyer, tax code, totals, lines)
- the tax code has 10 digits, or 13 digits when it is a branch code
- the line amounts add up to the subtotal
- subtotal + VAT equals the printed total
"""

import re

REQUIRED_FIELDS = (
    ("invoice_number", "Missing invoice number"),
    ("issue_date", "Missing issue date"),
    ("seller_name", "Missing seller name"),
    ("buyer_name", "Missing buyer name"),
    ("subtotal", "Missing subtotal"),
    ("vat_rate", "Missing VAT rate"),
    ("vat_amount", "Missing VAT amount"),
    ("total", "Missing total"),
)


def validate_invoice(invoice):
    """Return a list of issue strings. No issues means the invoice is OK."""
    issues = []

    for field_name, message in REQUIRED_FIELDS:
        if is_blank(invoice.get(field_name)):
            issues.append(message)

    issues.extend(tax_code_issues(invoice.get("tax_code")))
    issues.extend(line_amount_issues(invoice))
    issues.extend(total_issues(invoice))
    return issues


def tax_code_issues(tax_code):
    """A Vietnamese MST is 10 digits, or 13 digits for a branch (often written with a hyphen)."""
    if is_blank(tax_code):
        return ["Missing tax code"]

    digits = re.sub(r"\D", "", str(tax_code))
    if len(digits) not in (10, 13):
        return [f"Tax code '{tax_code}' is not 10 or 13 digits"]
    return []


def line_amount_issues(invoice):
    """The printed subtotal should equal the sum of the line amounts."""
    items = invoice.get("line_items") or []
    if not items:
        return ["No line items found"]

    missing_rows = [
        str(index)
        for index, item in enumerate(items, start=1)
        if item.get("amount") is None
    ]
    if missing_rows:
        rows = ", ".join(missing_rows)
        return [f"Line {rows} is missing an amount"]

    if invoice.get("subtotal") is None:
        return []

    line_sum = sum(item["amount"] for item in items)
    subtotal = invoice["subtotal"]
    if line_sum != subtotal:
        return [
            f"Sum of line amounts ({line_sum:,}) does not equal subtotal ({subtotal:,})"
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
        return [f"Subtotal + VAT ({expected:,}) does not equal total ({total:,})"]
    return []


def is_blank(value):
    """True for None and empty strings. The number 0 is a real value, not a blank."""
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False
