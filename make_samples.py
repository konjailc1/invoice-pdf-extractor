"""Generate fictional Vietnamese supplier invoices as PDF files.

The PDFs are the "before" picture for this project: realistic enough to
practice on, but every company, tax code, and amount is made up.

Two of the ten invoices are wrong on purpose:
- 07_an_khang.pdf prints a total that is 1,000,000 VND too high
- 09_dong_a.pdf leaves the seller tax code (MST) off the page

Run from the project folder:

    python make_samples.py
"""

import tarfile
import urllib.request
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).resolve().parent
FONT_DIR = ROOT / "fonts"
OUTPUT_DIR = ROOT / "sample_input"
CONTENT_WIDTH = A4[0] - 72  # page width minus 36pt margins on each side

# Official DejaVu 2.37 release. Raw GitHub file links are often blocked,
# so we pull the two faces we need out of the published archive.
FONT_ARCHIVE_URL = (
    "https://github.com/dejavu-fonts/dejavu-fonts/releases/download/"
    "version_2_37/dejavu-fonts-ttf-2.37.tar.bz2"
)
FONT_FILES = ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf")

# Visual recipe for each layout. Labels differ on purpose so the extractor
# has to recognise more than one wording.
LAYOUTS = {
    "classic": {
        "title": "HÓA ĐƠN GIÁ TRỊ GIA TĂNG",
        "accent": "#1B3A4B",
        "soft": "#E6EEF2",
        "number": "Số hóa đơn",
        "date": "Ngày lập",
        "seller": "Người bán",
        "buyer": "Người mua",
        "tax": "Mã số thuế",
        "subtotal": "Cộng tiền hàng",
        "vat": "Thuế GTGT ({rate}%)",
        "total": "Tổng cộng thanh toán",
        "date_style": "slash",
        "field_order": "classic",
        "title_size": 15,
        "columns": [
            ("stt", "STT", 30),
            ("description", "Mô tả hàng hóa", 213),
            ("unit", "ĐVT", 48),
            ("qty", "Số lượng", 58),
            ("unit_price", "Đơn giá", 82),
            ("amount", "Thành tiền", 92),
        ],
    },
    "modern": {
        "title": "HÓA ĐƠN BÁN HÀNG",
        "accent": "#0F6E6B",
        "soft": "#E5F4F3",
        "number": "Invoice No",
        "date": "Date",
        "seller": "Bên bán",
        "buyer": "Bên mua",
        "tax": "MST",
        "subtotal": "Tạm tính",
        "vat": "VAT {rate}%",
        "total": "Tổng tiền",
        "date_style": "slash",
        "field_order": "seller_first",
        "title_size": 15,
        "columns": [
            ("stt", "#", 28),
            ("description", "Hàng hóa / Dịch vụ", 237),
            ("qty", "SL", 50),
            ("unit_price", "Đơn giá (VND)", 100),
            ("amount", "Thành tiền (VND)", 108),
        ],
    },
    "compact": {
        "title": "PHIẾU HÓA ĐƠN",
        "accent": "#6B3E26",
        "soft": "#F6EFEA",
        "number": "Mã HĐ",
        "date": "Ngày",
        "seller": "Đơn vị bán",
        "buyer": "Khách hàng",
        "tax": "MST",
        "subtotal": "Tiền hàng",
        "vat": "GTGT {rate}%",
        "total": "Số tiền phải trả",
        "date_style": "dot",
        "field_order": "buyer_first",
        "title_size": 13,
        "columns": [
            ("description", "Mặt hàng", 255),
            ("qty", "Số lượng", 60),
            ("unit_price", "Đơn giá", 100),
            ("amount", "Thành tiền", 108),
        ],
    },
}


# All names, tax codes, and prices below are fictional.
RAW_INVOICES = [
    {
        "file_name": "01_minh_phat.pdf",
        "layout": "classic",
        "invoice_number": "HD-2026-0142",
        "issue_date": date(2026, 1, 12),
        "seller_name": "Công ty TNHH Thương mại Minh Phát",
        "seller_address": "45 Nguyễn Huệ, Quận 1, TP. Hồ Chí Minh",
        "seller_phone": "028 7300 1122",
        "tax_code": "0314567890",
        "buyer_name": "Công ty TNHH Giải pháp Số Horizon",
        "vat_rate": 10,
        "items": [
            ("Giấy in A4 Double A 70gsm", "thùng", 10, 285_000),
            ("Mực in HP 680 đen", "hộp", 8, 320_000),
            ("Bìa hồ sơ A4", "xấp", 15, 45_000),
            ("Bút bi Thiên Long", "hộp", 20, 28_000),
        ],
    },
    {
        "file_name": "02_hoa_binh.pdf",
        "layout": "modern",
        "invoice_number": "INV-2026-0203",
        "issue_date": date(2026, 2, 3),
        "seller_name": "Công ty CP Vật tư Xây dựng Hòa Bình",
        "seller_address": "18 Trần Phú, Ba Đình, Hà Nội",
        "seller_phone": "024 3733 2201",
        "tax_code": "0102345678",
        "buyer_name": "Công ty CP Đầu tư Sông Hồng",
        "vat_rate": 8,
        "items": [
            ("Xi măng Hà Tiên PCB40", "bao", 80, 95_000),
            ("Cát xây tô", "m3", 12, 280_000),
            ("Gạch ống 4 lỗ", "viên", 5_000, 1_200),
        ],
    },
    {
        "file_name": "03_anh_duong.pdf",
        "layout": "compact",
        "invoice_number": "PX-2026-0088",
        "issue_date": date(2026, 2, 18),
        "seller_name": "Doanh nghiệp tư nhân Ánh Dương",
        "seller_address": "9 Lê Hồng Phong, Quy Nhơn",
        "seller_phone": "0256 394 1188",
        "tax_code": "3701234567",
        "buyer_name": "Công ty TNHH Nhà hàng Biển Xanh",
        "vat_rate": 8,
        "items": [
            ("Gạo ST25 túi 5kg", "bao", 40, 165_000),
            ("Dầu ăn Tường An 5 lít", "can", 24, 195_000),
            ("Nước mắm Phú Quốc 500ml", "chai", 60, 42_000),
        ],
    },
    {
        "file_name": "04_dai_viet.pdf",
        "layout": "classic",
        "invoice_number": "HD-2026-0305",
        "issue_date": date(2026, 3, 5),
        "seller_name": "Công ty TNHH Sản xuất Nhựa Đại Việt",
        "seller_address": "KCN Tân Bình, TP. Hồ Chí Minh",
        "seller_phone": "028 3811 9090",
        "tax_code": "0309876543-001",  # 13-digit branch code
        "buyer_name": "Công ty TNHH Bán lẻ GreenMart",
        "vat_rate": 10,
        "items": [
            ("Túi PE 20x30", "kg", 50, 48_000),
            ("Màng co PVC khổ 40cm", "cuộn", 12, 220_000),
            ("Hộp nhựa 500ml", "thùng", 8, 310_000),
        ],
    },
    {
        "file_name": "05_sao_mai.pdf",
        "layout": "modern",
        "invoice_number": "INV-2026-0321",
        "issue_date": date(2026, 3, 21),
        "seller_name": "Công ty TNHH Vận tải Sao Mai",
        "seller_address": "210 Xa lộ Hà Nội, TP. Thủ Đức",
        "seller_phone": "028 3722 6644",
        "tax_code": "0311122233",
        "buyer_name": "Văn phòng đại diện Pacific Trade",
        "vat_rate": 10,
        "items": [
            ("Cước vận chuyển nội thành", "chuyến", 6, 850_000),
            ("Phí bốc xếp", "công", 6, 180_000),
            ("Phí cầu đường", "chuyến", 6, 75_000),
        ],
    },
    {
        "file_name": "06_ba_nam.pdf",
        "layout": "compact",
        "invoice_number": "PX-2026-0402",
        "issue_date": date(2026, 4, 2),
        "seller_name": "Hộ kinh doanh Cơm tấm Bà Năm",
        "seller_address": "27 Nguyễn Trãi, Quận 5, TP. Hồ Chí Minh",
        "seller_phone": "028 3833 4411",
        "tax_code": "8012345678",
        "buyer_name": "Trường Mầm non Hoa Sen",
        "vat_rate": 8,
        "items": [
            ("Suất cơm văn phòng", "suất", 80, 45_000),
            ("Nước suối 500ml", "thùng", 10, 55_000),
            ("Trái cây theo mùa", "mâm", 8, 180_000),
        ],
    },
    {
        "file_name": "07_an_khang.pdf",
        "layout": "classic",
        "invoice_number": "HD-2026-0416",
        "issue_date": date(2026, 4, 16),
        "seller_name": "Công ty CP Thiết bị Văn phòng An Khang",
        "seller_address": "88 Láng Hạ, Đống Đa, Hà Nội",
        "seller_phone": "024 3514 7788",
        "tax_code": "0105566778",
        "buyer_name": "Công ty TNHH Kiến trúc Nội thất Lá",
        "vat_rate": 10,
        "total_delta": 1_000_000,  # printed total is deliberately too high
        "items": [
            ("Máy in laser đơn sắc", "cái", 2, 4_500_000),
            ("Máy hủy tài liệu", "cái", 1, 2_200_000),
            ("Giấy A4", "ram", 30, 65_000),
        ],
    },
    {
        "file_name": "08_greenleaf.pdf",
        "layout": "modern",
        "invoice_number": "INV-2026-0507",
        "issue_date": date(2026, 5, 7),
        "seller_name": "Công ty TNHH Thực phẩm Sạch GreenLeaf",
        "seller_address": "15 Võ Văn Tần, Quận 3, TP. Hồ Chí Minh",
        "seller_phone": "028 3930 5566",
        "tax_code": "0319988776",
        "buyer_name": "Công ty TNHH Nhà hàng Biển Xanh",
        "vat_rate": 8,
        "items": [
            ("Rau củ hữu cơ", "kg", 40, 35_000),
            ("Trứng gà ta vỉ 10", "vỉ", 30, 42_000),
            ("Thịt heo sạch", "kg", 20, 145_000),
        ],
    },
    {
        "file_name": "09_dong_a.pdf",
        "layout": "compact",
        "invoice_number": "PX-2026-0519",
        "issue_date": date(2026, 5, 19),
        "seller_name": "Công ty TNHH Cơ khí Chính Xác Đông Á",
        "seller_address": "KCN Hòa Khánh, Đà Nẵng",
        "seller_phone": "0236 373 9099",
        "tax_code": None,  # deliberately missing from the PDF
        "buyer_name": "Công ty CP Đầu tư Sông Hồng",
        "vat_rate": 10,
        "items": [
            ("Bulong M12", "hộp", 20, 45_000),
            ("Vòng bi 6204", "cái", 15, 85_000),
            ("Dầu nhớt công nghiệp 18L", "can", 4, 1_150_000),
        ],
    },
    {
        "file_name": "10_phuong_nam.pdf",
        "layout": "classic",
        "invoice_number": "HD-2026-0601",
        "issue_date": date(2026, 6, 1),
        "seller_name": "Công ty TNHH In ấn Phương Nam",
        "seller_address": "64 Cách Mạng Tháng 8, Quận 3, TP. Hồ Chí Minh",
        "seller_phone": "028 3848 1212",
        "tax_code": "0312233445",
        "buyer_name": "Công ty TNHH Truyền thông Ánh Sao",
        "vat_rate": 10,
        "items": [
            ("Catalogue A4 16 trang", "cuốn", 200, 25_000),
            ("Túi giấy kraft có quai", "cái", 500, 8_500),
            ("Tem nhãn decal A6", "tờ", 1_000, 1_200),
        ],
    },
]


def ensure_fonts():
    """Download DejaVu Sans into fonts/ when the files are not already there.

    DejaVu covers Vietnamese letters such as ă, ê, ơ, ư, and đ. ReportLab's
    built-in fonts do not.
    """
    FONT_DIR.mkdir(parents=True, exist_ok=True)
    missing = [name for name in FONT_FILES if not (FONT_DIR / name).is_file()]
    if not missing:
        return

    archive_path = FONT_DIR / "dejavu-ttf.tar.bz2"
    print("Downloading DejaVu Sans (Vietnamese Unicode font)...")
    request = urllib.request.Request(FONT_ARCHIVE_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=120) as response:
        archive_path.write_bytes(response.read())

    with tarfile.open(archive_path, "r:bz2") as archive:
        for member in archive.getmembers():
            if Path(member.name).name in FONT_FILES and member.isfile():
                target = FONT_DIR / Path(member.name).name
                source = archive.extractfile(member)
                if source is None:
                    continue
                target.write_bytes(source.read())
    archive_path.unlink(missing_ok=True)

    still_missing = [name for name in FONT_FILES if not (FONT_DIR / name).is_file()]
    if still_missing:
        raise SystemExit(f"Could not download font files: {', '.join(still_missing)}")


def register_fonts():
    """Register DejaVu under names the PDF layout can use, including bold."""
    ensure_fonts()
    pdfmetrics.registerFont(TTFont("DejaVu", str(FONT_DIR / "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(FONT_DIR / "DejaVuSans-Bold.ttf")))
    pdfmetrics.registerFontFamily(
        "DejaVu",
        normal="DejaVu",
        bold="DejaVu-Bold",
        italic="DejaVu",
        boldItalic="DejaVu-Bold",
    )


def format_vnd(amount):
    """Format 1250000 as 1.250.000, the usual grouping on Vietnamese invoices."""
    return f"{amount:,}".replace(",", ".")


def format_issue_date(value, style):
    """Print a date as 12/01/2026 or 12.01.2026 depending on the layout."""
    if style == "dot":
        return value.strftime("%d.%m.%Y")
    return value.strftime("%d/%m/%Y")


def prepare_invoice(raw):
    """Fill in line amounts, VAT, and the total that will be printed.

    ``total_delta`` is added only to the printed total. Line amounts, the
    subtotal, and VAT stay correct so the validator can spot the mistake.
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
    return prepared


def make_styles(layout):
    """Paragraph styles for one layout. Colors follow that layout's accent."""
    accent = HexColor(layout["accent"])
    body = dict(fontName="DejaVu", textColor=HexColor("#1F2933"))
    return {
        "title": ParagraphStyle(
            "title",
            fontName="DejaVu-Bold",
            fontSize=layout["title_size"],
            leading=layout["title_size"] + 4,
            textColor=white,
            alignment=TA_CENTER,
        ),
        "banner_note": ParagraphStyle(
            "banner_note",
            fontName="DejaVu",
            fontSize=8,
            leading=11,
            textColor=white,
            alignment=TA_CENTER,
        ),
        "label": ParagraphStyle("label", fontSize=10, leading=14, **body),
        "summary": ParagraphStyle("summary", fontSize=10, leading=14, alignment=TA_RIGHT, **body),
        "total": ParagraphStyle(
            "total",
            fontName="DejaVu-Bold",
            fontSize=11,
            leading=14,
            textColor=white,
            alignment=TA_RIGHT,
        ),
        "header": ParagraphStyle(
            "header",
            fontName="DejaVu-Bold",
            fontSize=8,
            leading=11,
            textColor=white,
            alignment=TA_CENTER,
        ),
        "desc": ParagraphStyle("desc", fontSize=8.5, leading=11, alignment=TA_LEFT, **body),
        "note": ParagraphStyle(
            "note",
            fontName="DejaVu",
            fontSize=8,
            leading=11,
            textColor=HexColor("#52606D"),
            alignment=TA_LEFT,
        ),
        "accent": accent,
        "soft": HexColor(layout["soft"]),
    }


def labeled_line(label, value, style):
    """One 'Label: value' line. The extractor looks for these labels."""
    return Paragraph(f"<b>{escape(label)}:</b> {escape(str(value))}", style)


def party_flowables(invoice, layout, styles):
    """Seller name, address, phone, and tax code (when the invoice has one)."""
    rows = [
        (layout["seller"], invoice["seller_name"]),
        ("Địa chỉ", invoice["seller_address"]),
        ("Điện thoại", invoice["seller_phone"]),
    ]
    if invoice.get("tax_code"):
        rows.append((layout["tax"], invoice["tax_code"]))
    return [labeled_line(label, value, styles["label"]) for label, value in rows]


def buyer_flowables(invoice, layout, styles):
    """The customer name line."""
    return [labeled_line(layout["buyer"], invoice["buyer_name"], styles["label"])]


def meta_flowables(invoice, layout, styles):
    """Invoice number and issue date."""
    printed_date = format_issue_date(invoice["issue_date"], layout["date_style"])
    rows = [
        (layout["number"], invoice["invoice_number"]),
        (layout["date"], printed_date),
    ]
    return [labeled_line(label, value, styles["label"]) for label, value in rows]


def header_flowables(invoice, layout, styles):
    """Title bar plus the header fields, in the order this layout uses."""
    story = [
        banner(layout["title"], "Bản demo — dữ liệu giả lập, không dùng để kê khai thuế", styles),
        Spacer(1, 10),
    ]
    blocks = {
        "meta": meta_flowables(invoice, layout, styles),
        "seller": party_flowables(invoice, layout, styles),
        "buyer": buyer_flowables(invoice, layout, styles),
    }
    if layout["field_order"] == "seller_first":
        sequence = ("seller", "buyer", "meta")
    elif layout["field_order"] == "buyer_first":
        sequence = ("meta", "buyer", "seller")
    else:
        sequence = ("meta", "seller", "buyer")

    for name in sequence:
        story.extend(blocks[name])
        story.append(Spacer(1, 6))
    return story


def banner(title, note, styles):
    """A full-width colored title. No grid, so it is not read as a line-item table."""
    block = Table(
        [
            [Paragraph(escape(title), styles["title"])],
            [Paragraph(escape(note), styles["banner_note"])],
        ],
        colWidths=[CONTENT_WIDTH],
    )
    block.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), styles["accent"]),
                ("TOPPADDING", (0, 0), (-1, 0), 10),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 8),
                ("TOPPADDING", (0, 1), (-1, 1), 1),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return block


def items_table(invoice, layout, styles):
    """The line-item grid. Visible rules let pdfplumber find the table."""
    columns = layout["columns"]
    widths = scale_widths([width for _, _, width in columns], CONTENT_WIDTH)
    header = [Paragraph(escape(title), styles["header"]) for _, title, _ in columns]
    data = [header]
    for index, item in enumerate(invoice["items"], start=1):
        data.append([item_cell(item, key, index, styles) for key, _, _ in columns])

    table = Table(data, colWidths=widths, repeatRows=1)
    numeric_keys = {"qty", "unit_price", "amount"}
    style_commands = [
        ("FONTNAME", (0, 0), (-1, -1), "DejaVu"),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("BACKGROUND", (0, 0), (-1, 0), styles["accent"]),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.7, HexColor("#334155")),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, styles["soft"]]),
    ]
    for column_index, (key, _, _) in enumerate(columns):
        alignment = "RIGHT" if key in numeric_keys else "CENTER"
        if key == "description":
            alignment = "LEFT"
        style_commands.append(("ALIGN", (column_index, 1), (column_index, -1), alignment))
    table.setStyle(TableStyle(style_commands))
    return table


def item_cell(item, key, index, styles):
    """One table cell. Descriptions are paragraphs so long names can wrap."""
    if key == "stt":
        return str(index)
    if key == "description":
        return Paragraph(escape(item["description"]), styles["desc"])
    if key == "unit":
        return item["unit"]
    if key in ("qty", "unit_price", "amount"):
        return format_vnd(item[key])
    raise KeyError(key)


def scale_widths(widths, total):
    """Scale hand-tuned column widths so they exactly fill the text frame."""
    factor = total / sum(widths)
    return [width * factor for width in widths]


def summary_flowables(invoice, layout, styles):
    """Subtotal, VAT, and total. The total sits on a colored bar."""
    vat_label = layout["vat"].format(rate=invoice["vat_rate"])
    lines = [
        Spacer(1, 10),
        Paragraph(
            f"{escape(layout['subtotal'])}: {format_vnd(invoice['subtotal'])} VND",
            styles["summary"],
        ),
        Paragraph(
            f"{escape(vat_label)}: {format_vnd(invoice['vat_amount'])} VND",
            styles["summary"],
        ),
        Spacer(1, 4),
        total_bar(
            f"{layout['total']}: {format_vnd(invoice['total'])} VND",
            styles,
        ),
        Spacer(1, 8),
        Paragraph(
            "Hóa đơn mẫu cho dự án trích xuất PDF. Không có giá trị thanh toán.",
            styles["note"],
        ),
    ]
    return lines


def total_bar(text, styles):
    """Highlighted total line, drawn without a grid so it stays out of the item table."""
    block = Table(
        [[Paragraph(escape(text), styles["total"])]],
        colWidths=[CONTENT_WIDTH],
    )
    block.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), styles["accent"]),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return block


def draw_footer(canvas, doc):
    """Small footer on every page. Kept clear of the totals regex."""
    canvas.saveState()
    canvas.setFillColor(HexColor("#7B8794"))
    canvas.setFont("DejaVu", 8)
    canvas.drawString(36, 24, "invoice-pdf-extractor  ·  dữ liệu giả lập")
    canvas.restoreState()


def render_invoice(invoice, destination):
    """Build one PDF invoice using the layout named on the invoice."""
    layout = LAYOUTS[invoice["layout"]]
    styles = make_styles(layout)
    story = []
    story.extend(header_flowables(invoice, layout, styles))
    story.append(items_table(invoice, layout, styles))
    story.extend(summary_flowables(invoice, layout, styles))

    document = SimpleDocTemplate(
        str(destination),
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=42,
        title=f"{invoice['invoice_number']} (demo)",
        author="invoice-pdf-extractor",
    )
    document.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)


def main():
    """Write all sample invoices into sample_input/."""
    register_fonts()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Writing {len(RAW_INVOICES)} sample invoices to {OUTPUT_DIR}")
    for raw in RAW_INVOICES:
        invoice = prepare_invoice(raw)
        render_invoice(invoice, OUTPUT_DIR / invoice["file_name"])
        flags = []
        if raw.get("total_delta"):
            flags.append("WRONG_TOTAL")
        if not raw.get("tax_code"):
            flags.append("MISSING_TAX_CODE")
        flag_text = f"  [{', '.join(flags)}]" if flags else ""
        print(
            f"  {invoice['file_name']}  {invoice['layout']}"
            f"  subtotal={invoice['subtotal']}  vat={invoice['vat_amount']}"
            f"  total={invoice['total']}{flag_text}"
        )
    print("Done.")


if __name__ == "__main__":
    main()
