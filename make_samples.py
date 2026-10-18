"""Generate 10 fictional Vietnamese VAT invoices as A4 PDFs.

The layout follows a standard VAT invoice (hóa đơn GTGT): seller block, red
title, buyer block, items table, totals, amount in words, and signatures.

Every company, person, tax code, phone, email, and bank account is made up.
Addresses use real suburban Hanoi streets and the wards in force after the
1 July 2025 reform (no district level). House numbers are fictional.

Two invoices are wrong on purpose:
- 07_thiet_bi_so_lam.pdf prints a total that is 1,000,000 VND too high
- 09_co_khi_bac_nam.pdf leaves the seller tax code off the page

Run from the project folder:

    python make_samples.py
"""

from datetime import date
from pathlib import Path

from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle

from vnd_words import _self_check, amount_in_words

ROOT = Path(__file__).resolve().parent
FONT_DIR = ROOT / "fonts"
OUTPUT_DIR = ROOT / "sample_input"

PAGE_W, PAGE_H = A4
LEFT = 32
RIGHT = PAGE_W - 32
CONTENT_WIDTH = RIGHT - LEFT

RED = HexColor("#C00000")
NAVY = HexColor("#1B3A4B")
GRAY = HexColor("#E5E7EB")
GREEN = HexColor("#1B7A3A")
GREEN_BG = HexColor("#E5F6E8")
MUTED = HexColor("#52606D")

FORM_NUMBER = "01GTKT0/001"

# Administrative units follow the 1 July 2025 Hanoi reform: there is no
# district level, only phường/xã plus "TP. Hà Nội".
# Every street below is a real road in that suburban area. House numbers
# are fictional. Xã Dương Hòa and Xã Thuận An are omitted because no street
# there was certain enough to include.
ADDRESSES = [
    "45 Nguyễn Văn Cừ, Phường Bồ Đề, TP. Hà Nội",
    "128 Ngô Gia Tự, Phường Bồ Đề, TP. Hà Nội",
    "17 Cổ Linh, Phường Việt Hưng, TP. Hà Nội",
    "256 Nguyễn Văn Linh, Phường Phúc Lợi, TP. Hà Nội",
    "88 Đại lộ Thăng Long, Xã An Khánh, TP. Hà Nội",
    "9 Lại Yên, Xã Hoài Đức, TP. Hà Nội",
    "34 đường Tây Mỗ, Phường Tây Mỗ, TP. Hà Nội",
    "150 Quang Trung, Phường Hà Đông, TP. Hà Nội",
    "62 Tố Hữu, Phường Dương Nội, TP. Hà Nội",
    "21 Ỷ Lan, Xã Gia Lâm, TP. Hà Nội",
    "73 Nguyễn Văn Cừ, Phường Long Biên, TP. Hà Nội",
    "210 Nguyễn Văn Linh, Phường Việt Hưng, TP. Hà Nội",
    "14 An Khánh, Xã An Khánh, TP. Hà Nội",
    "55 Đại lộ Thăng Long, Xã Hoài Đức, TP. Hà Nội",
    "8 đường Đại Mỗ, Phường Đại Mỗ, TP. Hà Nội",
    "120 Hữu Hưng, Phường Đại Mỗ, TP. Hà Nội",
    "41 đường Xuân Phương, Phường Xuân Phương, TP. Hà Nội",
    "19 Lê Trọng Tấn, Phường Dương Nội, TP. Hà Nội",
    "6 Cổ Bi, Xã Gia Lâm, TP. Hà Nội",
    "230 Quang Trung, Phường Hà Đông, TP. Hà Nội",
]

# (description, unit, qty, unit price). Amounts are filled in later.
RAW_INVOICES = [
    {
        "file_name": "01_sen_vang.pdf",
        "monogram": "SV",
        "seller_name": "CÔNG TY TNHH THƯƠNG MẠI SEN VÀNG",
        "seller_tax_code": "0112345601",
        "seller_address": ADDRESSES[0],
        "seller_phone": "024 3555 0101",
        "seller_email": "lienhe@senvang.example",
        "seller_account": "190300010001 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AA/26E",
        "number": "0000101",
        "issue_date": date(2026, 1, 12),
        "buyer_person": "Phạm Thu Trang",
        "buyer_unit": "Công ty TNHH Giải pháp Số Bình Minh",
        "buyer_address": ADDRESSES[10],
        "buyer_account": "190300110001 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112201",
        "payment": "CK",
        "vat_rate": 10,
        "items": [
            ("Giấy in A4 70gsm", "ram", 20, 65_000),
            ("Mực in đen", "hộp", 6, 180_000),
            ("Bìa hồ sơ A4", "xấp", 10, 35_000),
            ("Bút bi xanh", "hộp", 15, 22_000),
        ],
    },
    {
        "file_name": "02_vat_lieu_bo_de.pdf",
        "monogram": "BĐ",
        "seller_name": "CÔNG TY CỔ PHẦN VẬT LIỆU BỒ ĐỀ",
        "seller_tax_code": "0112345602",
        "seller_address": ADDRESSES[1],
        "seller_phone": "024 3555 0102",
        "seller_email": "banhang@vatlieubode.example",
        "seller_account": "190300010002 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AB/26E",
        "number": "0000102",
        "issue_date": date(2026, 2, 3),
        "buyer_person": "Lê Quốc Bảo",
        "buyer_unit": "Công ty cổ phần Đầu tư Phố Xuân",
        "buyer_address": ADDRESSES[11],
        "buyer_account": "190300110002 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112202",
        "payment": "CK",
        "vat_rate": 8,
        "items": [
            ("Xi măng PCB40", "bao", 40, 95_000),
            ("Cát xây", "m3", 6, 250_000),
            ("Gạch ống", "viên", 2_000, 1_200),
        ],
    },
    {
        "file_name": "03_nang_mai.pdf",
        "monogram": "NM",
        "seller_name": "DOANH NGHIỆP TƯ NHÂN NẮNG MAI",
        "seller_tax_code": "0112345603",
        "seller_address": ADDRESSES[2],
        "seller_phone": "024 3555 0103",
        "seller_email": "bep@nangmai.example",
        "seller_account": "190300010003 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AC/26E",
        "number": "0000103",
        "issue_date": date(2026, 2, 18),
        "buyer_person": "Nguyễn Hà My",
        "buyer_unit": "Công ty TNHH Nhà hàng Biển Xanh Mẫu",
        "buyer_address": ADDRESSES[12],
        "buyer_account": "190300110003 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112203",
        "payment": "TM",
        "vat_rate": 8,
        "items": [
            ("Gạo túi 5kg", "bao", 20, 160_000),
            ("Dầu ăn 5 lít", "can", 12, 185_000),
            ("Nước mắm 500ml", "chai", 40, 35_000),
            ("Trứng gà vỉ 10", "vỉ", 25, 42_000),
            ("Rau củ theo mùa", "kg", 30, 28_000),
        ],
    },
    {
        "file_name": "04_la_xanh.pdf",
        "monogram": "LX",
        "seller_name": "CÔNG TY TNHH BAO BÌ LÁ XANH",
        "seller_tax_code": "0112345604",
        "seller_address": ADDRESSES[3],
        "seller_phone": "024 3555 0104",
        "seller_email": "xuat@laxanh.example",
        "seller_account": "190300010004 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AD/26E",
        "number": "0000104",
        "issue_date": date(2026, 3, 5),
        "buyer_person": "Trần Đức Long",
        "buyer_unit": "Công ty TNHH Bán lẻ Rau Sạch",
        "buyer_address": ADDRESSES[13],
        "buyer_account": "190300110004 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112204",
        "payment": "CK",
        "vat_rate": 10,
        "items": [
            ("Túi PE đóng gói", "kg", 80, 45_000),
            ("Màng co PVC", "cuộn", 10, 210_000),
        ],
    },
    {
        "file_name": "05_gio_moi.pdf",
        "monogram": "GM",
        "seller_name": "CÔNG TY TNHH VẬN TẢI GIÓ MỚI",
        "seller_tax_code": "0112345605",
        "seller_address": ADDRESSES[4],
        "seller_phone": "024 3555 0105",
        "seller_email": "dieuvan@giomoi.example",
        "seller_account": "190300010005 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AE/26E",
        "number": "0000105",
        "issue_date": date(2026, 3, 21),
        "buyer_person": "Vũ Mai Chi",
        "buyer_unit": "Công ty TNHH Thương mại Nam Dương",
        "buyer_address": ADDRESSES[14],
        "buyer_account": "190300110005 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112205",
        "payment": "CK",
        "vat_rate": 10,
        "items": [
            ("Cước vận chuyển nội thành", "chuyến", 4, 850_000),
            ("Phí bốc xếp", "công", 4, 150_000),
            ("Phí cầu đường", "chuyến", 4, 75_000),
            ("Phí lưu kho", "ngày", 3, 200_000),
            ("Dây đai niêm phong", "cuộn", 10, 25_000),
            ("Pallet gỗ", "cái", 6, 180_000),
        ],
    },
    {
        "file_name": "06_bep_nha_moc.pdf",
        "monogram": "BN",
        "seller_name": "HỘ KINH DOANH BẾP NHÀ MỘC",
        "seller_tax_code": "0112345606",
        "seller_address": ADDRESSES[5],
        "seller_phone": "024 3555 0106",
        "seller_email": "bep@nhamoc.example",
        "seller_account": "190300010006 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AF/26E",
        "number": "0000106",
        "issue_date": date(2026, 4, 2),
        "buyer_person": "Đỗ Anh Tú",
        "buyer_unit": "Trường mầm non Hoa Giấy",
        "buyer_address": ADDRESSES[15],
        "buyer_account": "190300110006 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112206",
        "payment": "TM",
        "vat_rate": 8,
        "items": [
            ("Suất cơm văn phòng", "suất", 60, 45_000),
            ("Nước suối 500ml", "thùng", 8, 55_000),
            ("Trái cây tráng miệng", "mâm", 6, 150_000),
        ],
    },
    {
        "file_name": "07_thiet_bi_so_lam.pdf",
        "monogram": "SL",
        "seller_name": "CÔNG TY CỔ PHẦN THIẾT BỊ SỐ LAM",
        "seller_tax_code": "0112345607",
        "seller_address": ADDRESSES[6],
        "seller_phone": "024 3555 0107",
        "seller_email": "sales@solam.example",
        "seller_account": "190300010007 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AG/26E",
        "number": "0000107",
        "issue_date": date(2026, 4, 16),
        "buyer_person": "Hoàng Khánh Linh",
        "buyer_unit": "Công ty TNHH Kiến trúc Nội thất Lá Xanh",
        "buyer_address": ADDRESSES[16],
        "buyer_account": "190300110007 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112207",
        "payment": "CK",
        "vat_rate": 10,
        "total_delta": 1_000_000,  # printed total is deliberately too high
        "items": [
            ("Máy in laser đơn sắc", "cái", 2, 4_500_000),
            ("Máy hủy tài liệu", "cái", 1, 2_200_000),
            ("Giấy A4", "ram", 20, 62_000),
            ("Mực in", "hộp", 4, 190_000),
        ],
    },
    {
        "file_name": "08_vuon_nho.pdf",
        "monogram": "VN",
        "seller_name": "CÔNG TY TNHH NÔNG SẢN VƯỜN NHỎ",
        "seller_tax_code": "0112345608",
        "seller_address": ADDRESSES[7],
        "seller_phone": "024 3555 0108",
        "seller_email": "vuon@vuonnho.example",
        "seller_account": "190300010008 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AH/26E",
        "number": "0000108",
        "issue_date": date(2026, 5, 7),
        "buyer_person": "Bùi Nhật Nam",
        "buyer_unit": "Công ty TNHH Bếp ăn Mai Vàng",
        "buyer_address": ADDRESSES[17],
        "buyer_account": "190300110008 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112208",
        "payment": "TM",
        "vat_rate": 8,
        "items": [
            ("Rau củ hữu cơ", "kg", 25, 32_000),
            ("Thịt heo sạch", "kg", 18, 150_000),
        ],
    },
    {
        "file_name": "09_co_khi_bac_nam.pdf",
        "monogram": "CK",
        "seller_name": "CÔNG TY TNHH CƠ KHÍ CHÍNH XÁC BẮC NAM",
        "seller_tax_code": None,  # deliberately missing
        "seller_address": ADDRESSES[8],
        "seller_phone": "024 3555 0109",
        "seller_email": "xuong@bacnam.example",
        "seller_account": "190300010009 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AI/26E",
        "number": "0000109",
        "issue_date": date(2026, 5, 19),
        "buyer_person": "Phan Gia Hân",
        "buyer_unit": "Công ty cổ phần Xây dựng Nền Mới",
        "buyer_address": ADDRESSES[18],
        "buyer_account": "190300110009 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112209",
        "payment": "CK",
        "vat_rate": 10,
        "items": [
            ("Bulong M12", "hộp", 15, 48_000),
            ("Vòng bi 6204", "cái", 12, 85_000),
            ("Dầu nhớt 18 lít", "can", 3, 1_150_000),
            ("Dây điện đơn", "cuộn", 8, 75_000),
            ("Ốc vít các loại", "hộp", 10, 32_000),
        ],
    },
    {
        "file_name": "10_muc_tuoi.pdf",
        "monogram": "MT",
        "seller_name": "CÔNG TY TNHH IN ẤN MỰC TƯƠI",
        "seller_tax_code": "0112345610-001",  # 13-digit branch code
        "seller_address": ADDRESSES[9],
        "seller_phone": "024 3555 0110",
        "seller_email": "in@muctuoi.example",
        "seller_account": "190300010010 tại Ngân hàng TMCP Thử Nghiệm",
        "serial": "AJ/26E",
        "number": "0000110",
        "issue_date": date(2026, 6, 1),
        "buyer_person": "Ngô Thanh Sơn",
        "buyer_unit": "Công ty TNHH Truyền thông Ánh Sao",
        "buyer_address": ADDRESSES[19],
        "buyer_account": "190300110010 tại Ngân hàng số Mẫu",
        "buyer_tax_code": "0101112210",
        "payment": "TM/CK",
        "vat_rate": 10,
        "items": [
            ("Catalogue A4 16 trang", "cuốn", 150, 22_000),
            ("Túi giấy có quai", "cái", 400, 8_500),
            ("Tem nhãn A6", "tờ", 800, 1_500),
        ],
    },
]


def register_fonts():
    """Register DejaVu Sans, which can draw Vietnamese letters. Bold is included."""
    regular = FONT_DIR / "DejaVuSans.ttf"
    bold = FONT_DIR / "DejaVuSans-Bold.ttf"
    if not regular.is_file() or not bold.is_file():
        raise SystemExit("Missing fonts/DejaVuSans.ttf or fonts/DejaVuSans-Bold.ttf")
    pdfmetrics.registerFont(TTFont("DejaVu", str(regular)))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(bold)))


def format_vnd(amount):
    """Format 1250000 as 1.250.000."""
    return f"{amount:,}".replace(",", ".")


def format_long_date(value):
    """12 January 2026 as 'Ngày 12 tháng 01 năm 2026'."""
    return f"Ngày {value.day:02d} tháng {value.month:02d} năm {value.year}"


def format_short_date(value):
    """12 January 2026 as 12/01/2026."""
    return value.strftime("%d/%m/%Y")


def prepare_invoice(raw):
    """Fill in line amounts, VAT, the printed total, and the amount in words.

    ``total_delta`` is added only to the printed total. The amount in words
    follows that printed total, so the words check still passes and the
    arithmetic check is the one that fails.
    """
    items = []
    for description, unit, qty, unit_price in raw["items"]:
        items.append(
            {
                "description": description,
                "unit": unit,
                "qty": qty,
                "unit_price": unit_price,
                "amount": qty * unit_price,
            }
        )
    subtotal = sum(item["amount"] for item in items)
    rate = raw["vat_rate"]
    if subtotal * rate % 100 != 0:
        raise SystemExit(f"{raw['file_name']}: VAT is not a whole number of VND")
    vat_amount = subtotal * rate // 100
    total = subtotal + vat_amount + raw.get("total_delta", 0)
    prepared = dict(raw)
    prepared["items"] = items
    prepared["subtotal"] = subtotal
    prepared["vat_amount"] = vat_amount
    prepared["total"] = total
    prepared["amount_in_words"] = amount_in_words(total)
    return prepared


def draw_border(c):
    """A double frame so the page reads as a form."""
    c.saveState()
    c.setStrokeColor(RED)
    c.setLineWidth(1.4)
    c.rect(14, 14, PAGE_W - 28, PAGE_H - 28)
    c.setStrokeColor(NAVY)
    c.setLineWidth(0.4)
    c.rect(18, 18, PAGE_W - 36, PAGE_H - 36)
    c.restoreState()


def draw_logo(c, monogram, x, y):
    """A plain monogram. It is not a real company logo."""
    c.saveState()
    c.setStrokeColor(NAVY)
    c.setFillColor(HexColor("#F4F7F8"))
    c.setLineWidth(1)
    c.circle(x + 22, y - 22, 20, fill=1, stroke=1)
    c.setFillColor(NAVY)
    c.setFont("DejaVu-Bold", 11)
    c.drawCentredString(x + 22, y - 26, monogram)
    c.restoreState()


def draw_seller(c, invoice, y):
    """Logo on the left, company details on the right. Returns the next y."""
    draw_logo(c, invoice["monogram"], LEFT, y)
    text_x = LEFT + 56
    c.setFillColor(RED)
    c.setFont("DejaVu-Bold", 11)
    c.drawString(text_x, y - 14, invoice["seller_name"])

    y_line = y - 32
    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    if invoice.get("seller_tax_code"):
        draw_spaced_tax_code(c, text_x, y_line, invoice["seller_tax_code"])
        y_line -= 14
    c.drawString(text_x, y_line, f"Địa chỉ: {invoice['seller_address']}")
    y_line -= 14
    c.drawString(
        text_x,
        y_line,
        f"Điện thoại: {invoice['seller_phone']}    Email: {invoice['seller_email']}",
    )
    y_line -= 14
    c.drawString(text_x, y_line, f"Số tài khoản: {invoice['seller_account']}")
    return y_line - 10


def draw_spaced_tax_code(c, x, y, tax_code):
    """Draw 'Mã số thuế:' then the digits with a gap between each one.

    Each digit is its own string. A text-state character spacing would leak
    into the table and the amount-in-words line drawn later.
    """
    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    label = "Mã số thuế: "
    c.drawString(x, y, label)
    cursor = x + c.stringWidth(label, "DejaVu", 9)
    for character in tax_code:
        c.drawString(cursor, y, character)
        cursor += c.stringWidth(character, "DejaVu", 9) + 2.4


def draw_title_block(c, invoice, y):
    """Centered red title, with mẫu số / ký hiệu / số on the right."""
    meta_x = RIGHT - 150
    center = (LEFT + meta_x - 8) / 2
    c.setFillColor(RED)
    c.setFont("DejaVu-Bold", 13)
    c.drawCentredString(center, y, "HÓA ĐƠN GIÁ TRỊ GIA TĂNG")
    draw_italic(c, "(Bản thể hiện của hóa đơn điện tử)", center, y - 14, 8, MUTED)
    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    c.drawCentredString(center, y - 28, format_long_date(invoice["issue_date"]))

    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    c.drawString(meta_x, y, f"Mẫu số: {FORM_NUMBER}")
    c.drawString(meta_x, y - 14, f"Ký hiệu: {invoice['serial']}")
    c.setFillColor(RED)
    c.setFont("DejaVu-Bold", 10)
    c.drawString(meta_x, y - 28, f"Số: {invoice['number']}")
    return y - 44


def draw_italic(c, text, center_x, y, size, color):
    """Shear DejaVu so the subtitle looks italic. There is no oblique font file."""
    c.saveState()
    c.setFillColor(color)
    c.setFont("DejaVu", size)
    width = c.stringWidth(text, "DejaVu", size)
    c.translate(center_x, y)
    c.transform(1, 0, 0.22, 1, 0, 0)
    c.drawString(-width / 2, 0, text)
    c.restoreState()


def draw_dotted_field(c, x, y, label, value, x_end):
    """'Label: value' followed by a dotted line out to ``x_end``."""
    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    written = f"{label}: {value}".rstrip()
    if value:
        written = f"{label}: {value}"
    else:
        written = f"{label}:"
    c.drawString(x, y, written)
    text_end = x + c.stringWidth(written, "DejaVu", 9) + 4
    if text_end < x_end - 8:
        c.saveState()
        c.setStrokeColor(HexColor("#9AA5B1"))
        c.setDash(1, 2)
        c.setLineWidth(0.6)
        c.line(text_end, y - 1, x_end, y - 1)
        c.restoreState()


def draw_buyer(c, invoice, y):
    """Buyer block. Payment method and buyer MST share the last line."""
    draw_dotted_field(c, LEFT, y, "Họ và tên người mua hàng", invoice["buyer_person"], RIGHT)
    y -= 16
    draw_dotted_field(c, LEFT, y, "Tên đơn vị", invoice["buyer_unit"], RIGHT)
    y -= 16
    draw_dotted_field(c, LEFT, y, "Địa chỉ", invoice["buyer_address"], RIGHT)
    y -= 16
    draw_dotted_field(c, LEFT, y, "Số tài khoản", invoice["buyer_account"], RIGHT)
    y -= 16
    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    c.drawString(LEFT, y, f"Hình thức thanh toán: {invoice['payment']}")
    tax = invoice.get("buyer_tax_code") or ""
    c.drawString(LEFT + 230, y, f"Mã số thuế: {tax}".rstrip())
    return y - 16


def items_table(invoice):
    """The goods grid, including the small 1 2 3 4 5 6=4x5 index row."""
    header = ["STT", "Tên hàng hóa, dịch vụ", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"]
    index = ["1", "2", "3", "4", "5", "6=4x5"]
    data = [header, index]
    for number, item in enumerate(invoice["items"], start=1):
        data.append(
            [
                str(number),
                item["description"],
                item["unit"],
                format_vnd(item["qty"]),
                format_vnd(item["unit_price"]),
                format_vnd(item["amount"]),
            ]
        )
    widths = scale_widths([32, 214, 46, 62, 86, 91], CONTENT_WIDTH)
    table = Table(data, colWidths=widths)
    style = [
        ("FONTNAME", (0, 0), (-1, 0), "DejaVu-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "DejaVu"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, 1), 7.5),
        ("FONTSIZE", (0, 2), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0), GRAY),
        ("ALIGN", (0, 0), (-1, 1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.7, HexColor("#334155")),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ALIGN", (1, 2), (1, -1), "LEFT"),
        ("ALIGN", (3, 2), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 2), (0, -1), "CENTER"),
        ("ALIGN", (2, 2), (2, -1), "CENTER"),
    ]
    table.setStyle(TableStyle(style))
    return table


def scale_widths(widths, total):
    """Scale column widths so they fill the frame inside the border."""
    factor = total / sum(widths)
    return [width * factor for width in widths]


def draw_totals(c, invoice, y):
    """Subtotal, VAT rate, VAT amount, and total. Returns the next y."""
    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    rows = [
        f"Cộng tiền hàng hóa, dịch vụ: {format_vnd(invoice['subtotal'])}",
        (
            f"Thuế suất GTGT: {invoice['vat_rate']}%"
            f"      Tiền thuế GTGT: {format_vnd(invoice['vat_amount'])}"
        ),
    ]
    for row in rows:
        c.drawRightString(RIGHT, y, row)
        y -= 14
    c.setFillColor(RED)
    c.setFont("DejaVu-Bold", 10)
    c.drawRightString(RIGHT, y, f"Tổng tiền thanh toán: {format_vnd(invoice['total'])}")
    return y - 18


def draw_words(c, invoice, y):
    """The total written out in Vietnamese, as one text run so it stays in order."""
    c.setFillColor(black)
    c.setFont("DejaVu", 9)
    c.drawString(LEFT, y, f"Số tiền bằng chữ: {invoice['amount_in_words']}")
    return y - 16


def draw_signatures(c, invoice, y):
    """Buyer and seller signature captions, plus a green 'signed by' box."""
    c.setFillColor(black)
    c.setFont("DejaVu-Bold", 9)
    c.drawCentredString(LEFT + 110, y, "Người mua hàng")
    c.drawCentredString(RIGHT - 110, y, "Người bán hàng")
    c.setFont("DejaVu", 8)
    c.setFillColor(MUTED)
    c.drawCentredString(LEFT + 110, y - 13, "(Ký, ghi rõ họ, tên)")
    c.drawCentredString(RIGHT - 110, y - 13, "(Ký, ghi rõ họ, tên)")

    box_x = RIGHT - 230
    box_y = y - 62
    c.setStrokeColor(GREEN)
    c.setFillColor(GREEN_BG)
    c.setLineWidth(1)
    c.roundRect(box_x, box_y, 220, 36, 4, fill=1, stroke=1)
    draw_check(c, box_x + 10, box_y + 14)
    c.setFillColor(GREEN)
    c.setFont("DejaVu", 7.5)
    c.drawString(box_x + 26, box_y + 20, f"Được ký bởi {invoice['seller_name']}")
    c.drawString(box_x + 26, box_y + 8, f"Ngày ký: {format_short_date(invoice['issue_date'])}")


def draw_check(c, x, y):
    """A small check mark drawn with two strokes."""
    c.saveState()
    c.setStrokeColor(GREEN)
    c.setLineWidth(1.6)
    c.setLineCap(1)
    c.line(x, y + 4, x + 4, y)
    c.line(x + 4, y, x + 12, y + 10)
    c.restoreState()


def draw_footer(c):
    """A short reminder that this page is not a real tax invoice."""
    c.setFillColor(MUTED)
    c.setFont("DejaVu", 7.5)
    c.drawCentredString(
        PAGE_W / 2,
        26,
        "Hóa đơn mẫu — dữ liệu giả lập, không có giá trị thanh toán.",
    )


def render_invoice(invoice, destination):
    """Build one single-page PDF."""
    c = canvas.Canvas(str(destination), pagesize=A4)
    c.setTitle(f"{invoice['number']} (demo)")
    c.setAuthor("invoice-pdf-extractor")
    draw_border(c)

    y = PAGE_H - 42
    y = draw_seller(c, invoice, y)
    c.setStrokeColor(HexColor("#CBD2D9"))
    c.setLineWidth(0.8)
    c.line(LEFT, y, RIGHT, y)
    y -= 22
    y = draw_title_block(c, invoice, y)
    y -= 6
    y = draw_buyer(c, invoice, y)
    y -= 6

    table = items_table(invoice)
    _, height = table.wrap(CONTENT_WIDTH, 600)
    table.drawOn(c, LEFT, y - height)
    y = y - height - 16
    y = draw_totals(c, invoice, y)
    y = draw_words(c, invoice, y)
    draw_signatures(c, invoice, y - 8)
    draw_footer(c)
    c.showPage()
    c.save()


def main():
    """Delete the old sample PDFs and write the ten VAT invoices."""
    _self_check()
    register_fonts()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for old in OUTPUT_DIR.glob("*.pdf"):
        old.unlink()

    seen = [invoice["seller_address"] for invoice in RAW_INVOICES]
    seen += [invoice["buyer_address"] for invoice in RAW_INVOICES]
    if len(seen) != len(set(seen)):
        raise SystemExit("Each seller and buyer needs a different address")

    print(f"Writing {len(RAW_INVOICES)} VAT invoices to {OUTPUT_DIR}")
    for raw in RAW_INVOICES:
        invoice = prepare_invoice(raw)
        render_invoice(invoice, OUTPUT_DIR / invoice["file_name"])
        flags = []
        if raw.get("total_delta"):
            flags.append("WRONG_TOTAL")
        if not raw.get("seller_tax_code"):
            flags.append("MISSING_SELLER_TAX_CODE")
        flag_text = f"  [{', '.join(flags)}]" if flags else ""
        print(
            f"  {invoice['file_name']}  items={len(invoice['items'])}"
            f"  vat={invoice['vat_rate']}  subtotal={invoice['subtotal']}"
            f"  tax={invoice['vat_amount']}  total={invoice['total']}{flag_text}"
        )
    print("Done.")


if __name__ == "__main__":
    main()
