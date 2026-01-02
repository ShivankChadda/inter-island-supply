"""
Simple Flask app to generate slips (order / purchase / invoice) from Master Roll.xlsx.

Run:
    export FLASK_APP=app.py
    flask run

Then open http://127.0.0.1:5000 and use the form.
"""
import base64  # for embedding clipart inline
import io  # in-memory buffers for PDF/ZIP creation
import re  # simple string sanitizing for filenames
from datetime import date, datetime  # for stamping current date/time
import os  # file mtime for status bar
from zoneinfo import ZoneInfo  # local timezone handling

import pandas as pd  # Excel parsing and data shaping
from flask import Flask, render_template_string, request, send_file  # web server + templating
from PIL import Image as PILImage  # label image generation
from PIL import ImageDraw, ImageFont  # drawing text/shapes on labels
from reportlab.lib import colors  # table styling
from reportlab.lib.pagesizes import A4  # PDF page size for slips
from reportlab.lib.styles import ParagraphStyle  # text styling in PDFs
from reportlab.lib.units import cm, inch  # convenient unit conversions
from reportlab.lib.utils import ImageReader  # (kept for table PDF path)
from reportlab.pdfgen import canvas  # (kept for compatibility)
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak  # PDF layout

app = Flask(__name__)

# Path to default workbook bundled with the app
DEFAULT_XLSX = "Master Roll.xlsx"
# In-memory pointer to the most recent uploaded workbook; falls back to DEFAULT_XLSX
CURRENT_XLSX = DEFAULT_XLSX
LAST_PROCESSED_AT = None  # track last processed timestamp for UI (localized)
LOCAL_TZ = ZoneInfo("Asia/Kolkata")  # adjust to your local timezone
DELIVERY_CACHE = {"pdf": None, "filename": None}  # simple cache for delivery slip PDF
DELIVERY_XLSX = None  # uploaded delivery master roll (per session)

# Branding asset and label sizing in inches / pixels
LOGO_PATH = "Farmers_Wordmark_Badge_Transparent_1_3000px.png"
CLIPART_PATH = "Excel clipart.png"
MARKER_DIR = "Marker box"
TEMPLATE_PATH = "Master Roll Template.xlsx"
DELIVERY_TEMPLATE = "Delivery Slip Master Roll.xlsx"  # optional default path if bundled
LABEL_WIDTH_IN = 5  # label width in inches (landscape 5x3)
LABEL_HEIGHT_IN = 3  # label height in inches
LABEL_DPI = 300  # DPI for high-quality PNG output
LABEL_WIDTH = LABEL_WIDTH_IN * inch  # also keep reportlab units for consistency
LABEL_HEIGHT = LABEL_HEIGHT_IN * inch


# Normalization helpers -----------------------------------------------------
# Known typo map for item names; extend as needed
ITEM_NAME_MAP = {
    "green chilli": "Green Chilli",
    "green chillie": "Green Chilli",
    "brocoli": "Broccoli",
}


def normalize_item_name(name: str) -> str:
    """Collapse whitespace, lowercase for lookup, and apply known typo fixes."""
    if not isinstance(name, str):
        return name
    base = " ".join(name.strip().split())  # collapse extra spaces (multiple -> single)
    key = base.lower()
    if key in ITEM_NAME_MAP:
        return ITEM_NAME_MAP[key]
    return base


def format_qty(x):
    """Format quantity to max 1 decimal, trim trailing zeros/dots."""
    try:
        val = float(x)
    except (TypeError, ValueError):
        return x
    return f"{val:.1f}".rstrip("0").rstrip(".")


def first_non_empty(series, default=""):
    """Return the first non-null value from a Series or a default."""
    if series is None:
        return default
    ser = series.dropna()
    return ser.iloc[0] if not ser.empty else default


def load_clipart_data_uri() -> str:
    """Embed clipart as a data URI so it can be used in the template without a static folder."""
    try:
        with open(CLIPART_PATH, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("ascii")
            return f"data:image/png;base64,{encoded}"
    except Exception:
        return ""


CLIPART_DATA_URI = load_clipart_data_uri()


def load_logo_data_uri() -> str:
    """Embed main logo as data URI for the top bar."""
    try:
        with open(LOGO_PATH, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("ascii")
            return f"data:image/png;base64,{encoded}"
    except Exception:
        return ""


LOGO_DATA_URI = load_logo_data_uri()


def load_marker_images():
    """Load marker images from the marker folder keyed by lowercase basename (without extension)."""
    mapping = {}
    if not os.path.isdir(MARKER_DIR):
        return mapping
    for fname in os.listdir(MARKER_DIR):
        if not fname.lower().endswith(".png"):
            continue
        key = os.path.splitext(fname)[0].lower().strip()
        path = os.path.join(MARKER_DIR, fname)
        try:
            mapping[key] = PILImage.open(path).convert("RGBA")
        except Exception:
            continue
    return mapping


def normalize_source_name(source: str) -> str:
    """Normalize Source to match marker filenames."""
    if not isinstance(source, str):
        return ""
    s = source.strip().lower()
    aliases = {
        "fgn": "fgn",
        "f.g.n": "fgn",
        "rsn": "rsn",
        "r.s.n": "rsn",
        "prashanto": "prashanto",
        "prashanto.": "prashanto",
    }
    return aliases.get(s, s)


MARKER_IMAGES = load_marker_images()

# Default delivery master roll if present on disk
if os.path.exists(DELIVERY_TEMPLATE):
    DELIVERY_XLSX = DELIVERY_TEMPLATE


def load_delivery_row(vendor: str, path: str) -> dict:
    """Load delivery master and return row dict for the vendor."""
    required_cols = [
        "Sailing_Date",
        "Ship_Name",
        "Vendor_Code",
        "Vendor_Name",
        "Vendor_Location",
        "Number of packages",
        "Number of boxes",
        "Number of trays",
    ]
    df = pd.read_excel(path)
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Delivery master missing columns: {', '.join(missing)}")
    mask = df["Vendor_Name"].fillna("").str.strip().str.lower() == vendor.strip().lower()
    if not mask.any():
        raise ValueError(f"Vendor '{vendor}' not found in Delivery Slip Master Roll.")
    row = df[mask].iloc[0]
    def val_int(x):
        try:
            n = int(x)
            return n if n >= 0 else 0
        except Exception:
            return 0
    meta = {
        "vendor_name": str(row["Vendor_Name"]).strip(),
        "vendor_marker": str(row["Vendor_Code"]).strip(),
        "vendor_location": str(row.get("Vendor_Location", "")).strip(),
        "sailing_date": row["Sailing_Date"],
        "ship_name": str(row["Ship_Name"]).strip(),
        "packages": val_int(row["Number of packages"]),
        "boxes": val_int(row["Number of boxes"]),
        "trays": val_int(row["Number of trays"]),
    }
    return meta

COLS = ["Serial Number", "Item", "Quantity", "Unit"]


def load_matches(df: pd.DataFrame, slip_type: str, identifier: str) -> tuple[pd.DataFrame, dict]:
    """Filter rows based on slip type and identifier; return rows and meta."""
    ident_l = identifier.lower()
    if slip_type in {"order", "invoice", "label"}:
        # Vendor-based slips: match by vendor name or code (case-insensitive)
        name_series = df["Vendor_Name"].fillna("").str.lower()
        code_series = df["Vendor_Code"].fillna("").str.lower()
        matches = df[(name_series == ident_l) | (code_series == ident_l)].copy()
        if matches.empty:
            raise ValueError(f"No rows found for vendor '{identifier}' (by name or code)")
        vendor_code = first_non_empty(matches["Vendor_Code"], identifier)
        customer_name = vendor_code  # per requirement: use code in meta
        quantity_column = "Packed_Quantity" if slip_type == "invoice" else "Quantity"
    else:  # purchase
        # Purchase slips: match by Source column
        source_series = df["Source"].fillna("").str.lower()
        matches = df[source_series == ident_l].copy()
        if matches.empty:
            raise ValueError(f"No rows found for source '{identifier}'")
        customer_name = identifier
        quantity_column = "Quantity"

    # Try to pull ship name and sailing date, with safe fallbacks
    ship_name = first_non_empty(matches["Ship_Name"], "")
    sail_val = first_non_empty(matches["Sailing_Date"], "")
    if hasattr(sail_val, "strftime"):
        sailing_date = sail_val.strftime("%d %B")
    else:
        sailing_date = str(sail_val) if sail_val else ""

    # Meta is reused across slips and labels
    meta = {
        "customer_name": customer_name,
        "vendor_name": first_non_empty(matches["Vendor_Name"], identifier),
        "location": first_non_empty(matches["Vendor_Location"]) if "Vendor_Location" in matches else "",
        "source": first_non_empty(matches["Source"]) if "Source" in matches else "",
        "ship_name": ship_name,
        "sailing_date": sailing_date,
        "quantity_column": quantity_column,
    }
    return matches, meta


def build_table(matches: pd.DataFrame, meta: dict, slip_type: str) -> pd.DataFrame:
    """Construct the output table with sorting, serials, totals, and proper quantity column."""
    # Start from the minimal columns we need and rename quantity consistently
    items_df = matches[["Item_Name", meta["quantity_column"], "Unit"]].copy()
    items_df.rename(columns={meta["quantity_column"]: "Quantity"}, inplace=True)
    items_df["Quantity"] = pd.to_numeric(items_df["Quantity"], errors="coerce")
    items_df = items_df.dropna(subset=["Item_Name"])
    # Normalize item names to fix common typos and collapse whitespace
    items_df["Item_Name"] = items_df["Item_Name"].apply(normalize_item_name)

    if slip_type == "invoice":
        items_df = items_df[items_df["Quantity"] > 0]  # skip zero/NaN

    # Aggregate quantities per item/unit to collapse duplicates (including typos)
    items_df = items_df.groupby(["Item_Name", "Unit"], as_index=False)["Quantity"].sum()

    items_df = items_df.sort_values("Item_Name")
    if items_df.empty:
        raise ValueError(f"No items found for {slip_type}")

    items_df = items_df.assign(**{"Serial Number": range(1, len(items_df) + 1)})
    items_df = items_df[["Serial Number", "Item_Name", "Quantity", "Unit"]]
    items_df = items_df.rename(columns={"Item_Name": "Item"}).reindex(columns=COLS)

    header_row = pd.DataFrame([{c: c for c in COLS}])
    total_row = pd.DataFrame(
        [{"Serial Number": "", "Item": "Total", "Quantity": items_df["Quantity"].sum(), "Unit": "Kg"}]
    )
    return pd.concat([header_row, items_df, total_row], ignore_index=True)


def build_label_items(matches: pd.DataFrame, meta: dict) -> list[dict]:
    """Prepare aggregated item rows for labels."""
    # Similar to build_table but returns a simple list of dicts for PNG labels
    items_df = matches[["Item_Name", meta["quantity_column"], "Unit", "Source"]].copy()
    items_df.rename(columns={meta["quantity_column"]: "Quantity"}, inplace=True)
    items_df["Quantity"] = pd.to_numeric(items_df["Quantity"], errors="coerce")
    items_df = items_df.dropna(subset=["Item_Name"])
    items_df["Item_Name"] = items_df["Item_Name"].apply(normalize_item_name)
    items_df["Source"] = items_df["Source"].apply(normalize_source_name)

    # aggregate to unique items (sum quantities per item/unit)
    items_df = (
        items_df.groupby(["Item_Name", "Unit", "Source"], as_index=False)["Quantity"]
        .sum()
    )
    items_df = items_df.sort_values("Item_Name")
    if items_df.empty:
        raise ValueError("No items found for labels")

    items = []
    for _, row in items_df.iterrows():
        items.append(
            {
                "item": row["Item_Name"],
                "quantity": row["Quantity"],
                "unit": row["Unit"],
                "source_key": row["Source"],
            }
        )
    return items


def make_pdf(table_df: pd.DataFrame, meta: dict, slip_type: str, identifier: str) -> tuple[bytes, str]:
    """Generate PDF bytes and filename for the given table/meta."""
    header_idx = 0
    total_idx = len(table_df) - 1

    data = table_df.copy()
    data["Quantity"] = data["Quantity"].apply(format_qty)
    table_data = data.values.tolist()

    # Set up a basic reportlab document
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []

    # Meta block at top of PDF
    meta_style = ParagraphStyle(name="Meta", fontName="Times-Roman", fontSize=12, leading=15, spaceAfter=4)
    for label, value in [
        ("Date", date.today().strftime("%d %B %Y")),
        ("Customer Name", meta["customer_name"]),
        ("Sailing Date", meta["sailing_date"]),
        ("Ship", meta["ship_name"]),
    ]:
        story.append(Paragraph(f"<b>{label}</b>: {value}", meta_style))
    story.append(Spacer(1, 12))

    # Table body with styling
    table = Table(table_data, colWidths=[2.5 * cm, 6 * cm, 3 * cm, 2.5 * cm])
    style = TableStyle(
        [
            ("FONTNAME", (0, 0), (-1, 0), "Times-Bold"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, header_idx + 1), (-1, total_idx), "Times-Roman"),
            ("ALIGN", (0, header_idx + 1), (0, total_idx), "CENTER"),
            ("ALIGN", (1, header_idx + 1), (2, total_idx), "LEFT"),
            ("ALIGN", (3, header_idx + 1), (3, total_idx), "CENTER"),
            ("ITALIC", (3, header_idx + 1), (3, total_idx), True),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ]
    )
    style.add("BACKGROUND", (0, total_idx), (-1, total_idx), colors.HexColor("#e6f4e6"))
    style.add("FONTNAME", (0, total_idx), (-1, total_idx), "Times-Bold")
    table.setStyle(style)
    story.append(table)

    # Render the PDF
    doc.build(story)
    buf.seek(0)

    # Safe filename
    safe_id = re.sub(r"[^a-z0-9]+", "_", identifier.lower()).strip("_") or "id"
    safe_sail = re.sub(r"[^a-z0-9]+", "_", meta["sailing_date"].lower()).strip("_") or "sailing_date"
    filename = f"{slip_type}_{safe_id}_{safe_sail}_slip.pdf"
    return buf.read(), filename


def make_bulk_pdf(sections: list[tuple[str, pd.DataFrame, dict]], slip_type: str) -> tuple[bytes, str]:
    """Generate a multi-section PDF (one per identifier)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    for idx, (identifier, table_df, meta) in enumerate(sections):
        data = table_df.copy()
        data["Quantity"] = data["Quantity"].apply(format_qty)
        table_data = data.values.tolist()

        meta_style = ParagraphStyle(name="Meta", fontName="Times-Roman", fontSize=12, leading=15, spaceAfter=4)
        story.append(Paragraph(f"<b>{slip_type.title()}</b>: {identifier}", meta_style))
        for label, value in [
            ("Date", date.today().strftime("%d %B %Y")),
            ("Customer Name", meta.get("customer_name", "")),
            ("Sailing Date", meta.get("sailing_date", "")),
            ("Ship", meta.get("ship_name", "")),
        ]:
            story.append(Paragraph(f"<b>{label}</b>: {value}", meta_style))
        story.append(Spacer(1, 12))

        header_idx = 0
        total_idx = len(data) - 1
        table = Table(table_data, colWidths=[2.5 * cm, 6 * cm, 3 * cm, 2.5 * cm])
        style = TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Times-Bold"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTNAME", (0, header_idx + 1), (-1, total_idx), "Times-Roman"),
                ("ALIGN", (0, header_idx + 1), (0, total_idx), "CENTER"),
                ("ALIGN", (1, header_idx + 1), (2, total_idx), "LEFT"),
                ("ALIGN", (3, header_idx + 1), (3, total_idx), "CENTER"),
                ("ITALIC", (3, header_idx + 1), (3, total_idx), True),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ]
        )
        style.add("BACKGROUND", (0, total_idx), (-1, total_idx), colors.HexColor("#e6f4e6"))
        style.add("FONTNAME", (0, total_idx), (-1, total_idx), "Times-Bold")
        table.setStyle(style)
        story.append(table)
        if idx < len(sections) - 1:
            story.append(PageBreak())

    doc.build(story)
    buf.seek(0)
    filename = f"{slip_type}_all.pdf"
    return buf.read(), filename


def make_label_zip(items: list[dict], meta: dict, identifier: str) -> tuple[bytes, str]:
    """Generate label images (PNG) and return a ZIP as bytes."""
    # Compute pixel dimensions from inches and DPI
    width_px = int(LABEL_WIDTH_IN * LABEL_DPI)
    height_px = int(LABEL_HEIGHT_IN * LABEL_DPI)

    # load fonts (fallback to default if missing)
    def load_font(size, bold=False):
        candidates = [
            ("Times New Roman Bold.ttf", "Times New Roman.ttf") if bold else ("Times New Roman.ttf",),
            ("DejaVuSerif-Bold.ttf", "DejaVuSerif.ttf") if bold else ("DejaVuSerif.ttf",),
        ]
        for family in candidates:
            for path in family:
                try:
                    return ImageFont.truetype(path, size)
                except Exception:
                    continue
        return ImageFont.load_default()

    font_label = load_font(40, bold=True)
    font_value = load_font(40, bold=False)

    # pre-load logo
    logo_img = None
    try:
        logo_img = PILImage.open(LOGO_PATH).convert("RGBA")
    except Exception:
        logo_img = None

    images = []
    # Fixed composition in px for 5x3" @300dpi (1500x900)
    border_margin = 60  # 0.20" from edges
    inner_border_offset = 2  # gap between double borders
    logo_top_space = 66  # ~0.22" below inner border
    logo_bottom_space = 60  # gap below logo before text
    text_left_margin = 105  # ~0.35" inset from inner border
    line_min_gap = 46  # approx 1.15x 40px
    marker_offset = 54  # inset from inner border
    marker_size = 156  # ~0.52"

    for item in items:
        # Blank white canvas per label
        img = PILImage.new("RGB", (width_px, height_px), "white")
        draw = ImageDraw.Draw(img)

        # Double border inset from edge
        outer_rect = [border_margin, border_margin, width_px - border_margin, height_px - border_margin]
        inner_rect = [
            outer_rect[0] + inner_border_offset,
            outer_rect[1] + inner_border_offset,
            outer_rect[2] - inner_border_offset,
            outer_rect[3] - inner_border_offset,
        ]
        draw.rectangle(outer_rect, outline="black", width=2)
        draw.rectangle(inner_rect, outline="black", width=2)

        # logo placement (centered, scaled to 55-60% width, <=20% height)
        y_cursor = border_margin + logo_top_space
        if logo_img:
            max_w = int(width_px * 0.5) if width_px * 0.5 < 750 else 750
            max_w = max(max_w, 660)
            max_h = int(height_px * 0.2)
            lw, lh = logo_img.size
            scale = min(max_w / lw, max_h / lh)
            new_size = (int(lw * scale), int(lh * scale))
            logo_resized = logo_img.resize(new_size, PILImage.LANCZOS)
            x_logo = (width_px - new_size[0]) // 2
            y_logo = y_cursor
            img.paste(logo_resized, (x_logo, y_logo), logo_resized)
            y_cursor = y_logo + new_size[1] + logo_bottom_space
        # Ensure text starts below logo with defined gap
        text_y = y_cursor

        lines = [
            ("Vendor Name", meta.get("vendor_name", "")),
            ("Location", meta.get("location", "")),
            ("Marker", meta.get("customer_name", "")),
            ("Item", item.get("item", "")),
            ("Weight", f"{format_qty(item.get('quantity'))} {item.get('unit', '')}".strip()),
        ]

        # measure text using getbbox (compatible with newer PIL)
        def text_width(txt, font):
            bbox = font.getbbox(txt)
            return bbox[2] - bbox[0] if bbox else 0

        # Draw each Label: Value line with spacing (1.45x line height, min 5mm)
        line_spacing = max(int(font_value.size * 1.45), line_min_gap)
        x_text = text_left_margin + border_margin
        for label, val in lines:
            label_txt = f"{label}:"
            draw.text((x_text, text_y), label_txt, font=font_label, fill="black")
            label_w = text_width(label_txt, font_label)
            draw.text((x_text + label_w + 10, text_y), str(val), font=font_value, fill="black")
            text_y += line_spacing

        # marker box bottom right with icon based on Source
        box_size = marker_size
        box_x = width_px - border_margin - marker_offset - box_size
        box_y = height_px - border_margin - marker_offset - box_size
        source_key = item.get("source_key") or normalize_source_name(meta.get("source", ""))
        marker_img = MARKER_IMAGES.get(source_key)
        if not marker_img:
            # fallback: substring match
            for key in MARKER_IMAGES.keys():
                if key in source_key:
                    marker_img = MARKER_IMAGES.get(key)
                    break
        if marker_img:
            # Scale marker image to occupy the marker area, preserve aspect, no container
            miw, mih = marker_img.size
            scale = min(box_size / miw, box_size / mih)
            new_size = (int(miw * scale), int(mih * scale))
            marker_resized = marker_img.resize(new_size, PILImage.LANCZOS)
            mx = box_x + (box_size - new_size[0]) // 2
            my = box_y + (box_size - new_size[1]) // 2
            img.paste(marker_resized, (mx, my), marker_resized)
        else:
            # default simple dot if no marker available
            draw.ellipse([box_x, box_y, box_x + box_size, box_y + box_size], fill="black", outline=None)

        # Save this label to in-memory PNG
        bio = io.BytesIO()
        img.save(bio, format="PNG")
        bio.seek(0)
        safe_item = re.sub(r"[^a-z0-9]+", "_", str(item.get("item", "")).lower()).strip("_") or "item"
        images.append((f"{safe_item}.png", bio.read()))

    # build zip of all labels
    zip_buf = io.BytesIO()
    import zipfile

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in images:
            zf.writestr(fname, data)
    zip_buf.seek(0)
    safe_id = re.sub(r"[^a-z0-9]+", "_", identifier.lower()).strip("_") or "id"
    filename = f"label_{safe_id}.zip"
    return zip_buf.read(), filename


def make_delivery_pdf(meta: dict, images: list[tuple[str, bytes]]) -> tuple[bytes, str]:
    """Generate a delivery slip PDF with structured header, summary, and 3x3 photo grid."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []

    vendor = meta.get("vendor_name", "")
    ship_name = meta.get("ship_name", "")
    sailing_raw = meta.get("sailing_date", "")
    if hasattr(sailing_raw, "strftime"):
        sailing_date = sailing_raw.strftime("%d %B %Y")
    else:
        sailing_date = str(sailing_raw)
    packages = meta.get("packages") or 0
    trays = meta.get("trays") or 0
    boxes = meta.get("boxes") or 0

    # Header band
    header_bg = colors.HexColor("#f6f3ed")
    title_style = ParagraphStyle(name="Title", fontName="Times-Roman", fontSize=12, leading=14, textColor=colors.HexColor("#4a433b"))
    vendor_style = ParagraphStyle(name="Vendor", fontName="Times-Bold", fontSize=24, leading=28, textColor=colors.HexColor("#2f2a24"))
    meta_style = ParagraphStyle(name="MetaSmall", fontName="Times-Roman", fontSize=11, leading=13, textColor=colors.HexColor("#5a544d"))

    header_rows = []
    # logo
    logo_flow = ""
    try:
        logo = Image(LOGO_PATH, width=150, height=60)
        logo_flow = logo
    except Exception:
        pass
    header_rows.append([logo_flow])
    header_rows.append([Paragraph("Delivery Slip", title_style)])
    header_rows.append([Paragraph(vendor, vendor_style)])
    header_rows.append([Paragraph(f"Location: {meta.get('vendor_location','')}", meta_style)])
    header_rows.append([Paragraph(f"Ship: {ship_name}", meta_style)])
    header_rows.append([Paragraph(f"Sailing Date: {sailing_date}", meta_style)])
    header = Table(header_rows, colWidths=[doc.width])
    header.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), header_bg),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
        ("RIGHTPADDING", (0,0), (-1,-1), 10),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("ALIGN", (0,0), (-1,-1), "LEFT"),
    ]))
    story.append(header)
    story.append(Spacer(1, 14))

    # Summary boxes
    summary_items = []
    if packages > 0:
        summary_items.append(("Packages", packages))
    if trays > 0:
        summary_items.append(("Trays", trays))
    if boxes > 0:
        summary_items.append(("Boxes", boxes))
    if summary_items:
        cols = len(summary_items)
        col_widths = [doc.width / cols] * cols
        data = []
        row = []
        for label, val in summary_items:
            row.append([
                Paragraph(str(val), ParagraphStyle(name="Num", fontName="Times-Bold", fontSize=16, alignment=1)),
                Paragraph(label, ParagraphStyle(name="Lbl", fontName="Times-Roman", fontSize=11, textColor=colors.HexColor("#5a544d"), alignment=1))
            ])
        data.append(row)
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("INNERGRID", (0,0), (-1,-1), 0.5, colors.HexColor("#d1c6b8")),
            ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#d1c6b8")),
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#fbf9f6")),
            ("TOPPADDING", (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ]))
        story.append(table)
        story.append(Spacer(1, 14))

    # Divider
    story.append(Paragraph("Package Photos", ParagraphStyle(name="Div", fontName="Times-Bold", fontSize=12, leading=14, textColor=colors.HexColor("#4a433b"))))
    story.append(Spacer(1, 4))
    story.append(Table([[""]], colWidths=[doc.width], style=TableStyle([("LINEBELOW",(0,0),(-1,0),0.5,colors.HexColor("#b0a495"))])))
    story.append(Spacer(1, 10))

    # Images grid 3x3 per page with card styling
    cell_w = (doc.width - 12) / 3
    cell_h = (A4[1] - doc.topMargin - doc.bottomMargin - 200) / 3

    def img_card(name, data, idx):
        img = PILImage.open(io.BytesIO(data))
        iw, ih = img.size
        # maintain aspect but fit within card image area
        max_w = cell_w - 12
        max_h = cell_h - 30
        scale = min(max_w / iw, max_h / ih)
        new_size = (int(iw * scale), int(ih * scale))
        bio = io.BytesIO()
        img.resize(new_size, PILImage.LANCZOS).save(bio, format="PNG")
        bio.seek(0)
        cap_style = ParagraphStyle(name="Cap", fontName="Times-Roman", fontSize=10, leading=12, alignment=1, textColor=colors.HexColor("#4a433b"))
        label = f"Package {idx+1}"
        table = Table(
            [[Image(bio, width=new_size[0], height=new_size[1])],
             [Paragraph(label, cap_style)]],
            colWidths=[cell_w - 8]
        )
        table.setStyle(TableStyle([
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#d1c6b8")),
            ("BACKGROUND",(0,0),(-1,-1),colors.whitesmoke),
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]))
        return table

    cards = [img_card(name, data, idx) for idx, (name, data) in enumerate(images)]
    rows = []
    for i in range(0, len(cards), 3):
        row = cards[i:i+3]
        while len(row) < 3:
            row.append("")
        rows.append(row)
        if len(rows) == 3:
            story.append(Table(rows, colWidths=[cell_w]*3, style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("ALIGN",(0,0),(-1,-1),"CENTER")])))
            story.append(PageBreak())
            rows = []
    if rows:
        story.append(Table(rows, colWidths=[cell_w]*3, style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("ALIGN",(0,0),(-1,-1),"CENTER")])))

    doc.build(story)
    buf.seek(0)
    safe_vendor = re.sub(r"[^a-z0-9]+", "_", vendor.lower()).strip("_") or "vendor"
    filename = f"delivery_{safe_vendor}.pdf"
    return buf.read(), filename

def render_html(table_df: pd.DataFrame, meta: dict, slip_type: str, identifier: str) -> str:
    """Build a simple HTML view of the slip."""
    df = table_df.copy()
    df["Quantity"] = df["Quantity"].apply(format_qty)

    # Build header and rows manually to keep styling consistent
    rows_html = []
    for i, row in df.iterrows():
        is_header = i == 0
        is_total = i == len(df) - 1
        cells = []
        for col in COLS:
            val = row[col]
            if pd.isna(val):
                val = ""
            align = "center" if col in ("Serial Number", "Unit") else "left"
            style = "padding:6px; border:1px solid #000;"
            if is_header:
                style += "font-weight:bold; text-align:center;"
            elif is_total:
                style += "font-weight:bold; background:#e6f4e6;"
            else:
                if col == "Unit":
                    style += "font-style:italic; text-align:center;"
                elif col == "Serial Number":
                    style += "text-align:center;"
                else:
                    style += "text-align:left;"
            cells.append(f'<td style="{style}">{val}</td>')
        rows_html.append("<tr>" + "".join(cells) + "</tr>")

    meta_block = (
        "<div style=\"font-family:'Bell MT','CMU Serif','Computer Modern',serif; "
        "font-size:14px; color:black; background:white; line-height:1.5; margin-bottom:12px;\">"
        f"<div><strong>Date</strong>: {date.today().strftime('%d %B %Y')}</div>"
        f"<div><strong>Customer Name</strong>: {meta['customer_name']}</div>"
        f"<div><strong>Sailing Date</strong>: {meta['sailing_date']}</div>"
        f"<div><strong>Ship</strong>: {meta['ship_name']}</div>"
        "</div>"
    )

    table_html = (
        "<table style='border-collapse:collapse; background:white; color:black; "
        "font-family:'Bell MT','CMU Serif','Computer Modern',serif; font-size:14px;'>"
        f"{''.join(rows_html)}</table>"
    )

    return meta_block + table_html


def render_label_preview(items: list[dict], meta: dict) -> str:
    """Simple HTML preview list for labels."""
    rows = []
    for item in items:
        rows.append(
            f"<li>{item.get('item','')} — {format_qty(item.get('quantity'))} {item.get('unit','')}</li>"
        )
    meta_block = (
        "<div style=\"font-family:'Bell MT','CMU Serif','Computer Modern',serif; "
        "font-size:14px; color:black; background:white; line-height:1.5; margin-bottom:12px;\">"
        f"<div><strong>Vendor Name</strong>: {meta.get('vendor_name','')}</div>"
        f"<div><strong>Location</strong>: {meta.get('location','')}</div>"
        f"<div><strong>Marker</strong>: {meta.get('customer_name','')}</div>"
        "</div>"
    )
    list_block = "<ul style='padding-left:16px;'>" + "".join(rows) + "</ul>"
    return meta_block + list_block


@app.route("/", methods=["GET", "POST"])
def home():
    global CURRENT_XLSX
    global LAST_PROCESSED_AT
    global DELIVERY_CACHE
    global DELIVERY_XLSX
    error = None
    html_snippet = None
    delivery_snippet = None
    slip_type = ""
    identifier = ""
    uploaded_name = None
    upload_only = False
    file_rows = None
    file_cols = []
    mode = request.args.get("mode") or request.form.get("mode") or "dispatch"
    delivery_uploaded_name = None

    if request.method == "POST":
        if mode == "delivery":
            vendor_name = request.form.get("delivery_vendor", "").strip()
            delivery_upload = request.files.get("delivery_workbook")
            photos = request.files.getlist("photos")
            try:
                if not vendor_name:
                    raise ValueError("Vendor Name is required.")
                if delivery_upload and delivery_upload.filename:
                    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", delivery_upload.filename) or "delivery_master.xlsx"
                    tmp_path = f"/tmp/{safe_name}"
                    delivery_upload.save(tmp_path)
                    DELIVERY_XLSX = tmp_path
                    delivery_uploaded_name = delivery_upload.filename
                if not DELIVERY_XLSX:
                    raise ValueError("Please upload the Delivery Slip Master Roll (.xlsx).")
                meta = load_delivery_row(vendor_name, DELIVERY_XLSX)
                images = []
                for f in photos:
                    if not f or not f.filename:
                        continue
                    name = f.filename
                    ext = os.path.splitext(name)[1].lower()
                    if ext not in {".png", ".jpg", ".jpeg"}:
                        raise ValueError(f"Unsupported file type: {ext}")
                    data = f.read()
                    images.append((name, data))
                if not images:
                    raise ValueError("Please upload photos for the delivery slip.")
                images.sort(key=lambda t: t[0].lower())
                pdf_bytes, filename = make_delivery_pdf(meta, images)
                DELIVERY_CACHE = {"pdf": pdf_bytes, "filename": filename}
                delivery_snippet = f"Delivery slip ready for {meta.get('vendor_name','')}. Photos: {len(images)}"
            except Exception as exc:
                error = str(exc)
        else:
            # Allow resetting the file without needing other inputs
            if request.form.get("reset_file") == "1":
                CURRENT_XLSX = DEFAULT_XLSX
                uploaded_name = None
                slip_type = ""
                identifier = ""
                LAST_PROCESSED_AT = None
            else:
                # Read form inputs
                slip_type = request.form.get("slip_type", "").strip().lower()
                identifier = request.form.get("identifier", "").strip()
                upload = request.files.get("workbook")
                upload_only = request.form.get("upload_only") == "1"
                try:
                    # Process upload first (always)
                    if upload and upload.filename:
                        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", upload.filename) or "uploaded.xlsx"
                        tmp_path = f"/tmp/{safe_name}"
                        upload.save(tmp_path)
                        CURRENT_XLSX = tmp_path
                        uploaded_name = upload.filename
                        LAST_PROCESSED_AT = datetime.now(LOCAL_TZ).strftime("%d %b %Y, %I:%M %p")
                    if CURRENT_XLSX == DEFAULT_XLSX and not upload:
                        raise ValueError("Please upload the Master Roll (.xlsx) to continue.")

                    # If upload-only, stop after updating file state
                    if upload_only:
                        raw_df = pd.read_excel(CURRENT_XLSX, nrows=5000)
                        file_rows = len(raw_df)
                        key_cols = [
                            "Vendor_Name",
                            "Vendor_Code",
                            "Source",
                            "Item_Name",
                            "Quantity",
                            "Packed_Quantity",
                            "Ship_Name",
                            "Sailing_Date",
                        ]
                        file_cols = [c for c in key_cols if c in raw_df.columns]
                    else:
                        if slip_type not in {"order", "purchase", "invoice", "label"}:
                            raise ValueError("Slip type must be order, purchase, invoice, or label.")

                        all_mode = request.form.get("all_mode") == "1"
                        workbook_path = CURRENT_XLSX
                        raw_df = pd.read_excel(workbook_path)

                        if all_mode:
                            sections = []
                            if slip_type in {"order", "invoice"}:
                                # unique vendors by code/name
                                vendors = raw_df[["Vendor_Code", "Vendor_Name"]].dropna(how="all").drop_duplicates()
                                for _, row in vendors.iterrows():
                                    ident = row["Vendor_Code"] if pd.notna(row["Vendor_Code"]) else row["Vendor_Name"]
                                    if pd.isna(ident):
                                        continue
                                    matches, meta = load_matches(raw_df, slip_type, str(ident))
                                    table_df = build_table(matches, meta, slip_type)
                                    sections.append((str(ident), table_df, meta))
                            elif slip_type == "purchase":
                                sources = raw_df["Source"].dropna().unique()
                                for src in sources:
                                    matches, meta = load_matches(raw_df, slip_type, str(src))
                                    table_df = build_table(matches, meta, slip_type)
                                    sections.append((str(src), table_df, meta))
                            LAST_PROCESSED_AT = datetime.now(LOCAL_TZ).strftime("%d %b %Y, %I:%M %p")
                            html_snippet = f"Ready to download {len(sections)} slips."
                            identifier = ""  # clear
                        else:
                            if not identifier:
                                raise ValueError("Identifier is required.")
                            matches, meta = load_matches(raw_df, slip_type, identifier)
                            LAST_PROCESSED_AT = datetime.now(LOCAL_TZ).strftime("%d %b %Y, %I:%M %p")
                            if slip_type == "label":
                                # Labels: show a simple preview list
                                label_items = build_label_items(matches, meta)
                                html_snippet = render_label_preview(label_items, meta)
                            else:
                                # Slips: render HTML table preview
                                table_df = build_table(matches, meta, slip_type)
                                html_snippet = render_html(table_df, meta, slip_type, identifier)
                except Exception as exc:  # broad to show message to user
                    error = str(exc)

    # status info for header
    file_display = uploaded_name or os.path.basename(CURRENT_XLSX)
    try:
        file_modified = datetime.fromtimestamp(os.path.getmtime(CURRENT_XLSX), tz=LOCAL_TZ).strftime("%d %b %Y, %I:%M %p")
    except Exception:
        file_modified = "unknown"
    today_str = datetime.now(LOCAL_TZ).strftime("%d %b %Y")
    status_text = "Master Roll loaded" if CURRENT_XLSX != DEFAULT_XLSX else "Default Master Roll"
    last_processed = LAST_PROCESSED_AT or "Not processed"
    delivery_file_name = os.path.basename(DELIVERY_XLSX) if DELIVERY_XLSX else ""

    # file summary for preview card (if not already from upload-only branch)
    if file_rows is None:
        try:
            tmp_df = pd.read_excel(CURRENT_XLSX, nrows=5000)  # limit for speed
            file_rows = len(tmp_df)
            key_cols = [
                "Vendor_Name",
                "Vendor_Code",
                "Source",
                "Item_Name",
                "Quantity",
                "Packed_Quantity",
                "Ship_Name",
                "Sailing_Date",
            ]
            file_cols = [c for c in key_cols if c in tmp_df.columns]
        except Exception:
            file_rows = None

    template = """
    <!doctype html>
    <html>
    <head>
      <title>Slip Generator</title>
      <style>
        :root {
          --primary: #B5906D;
          --bg: #F1E7D5;
          --border: #15322A;
        }
        body { font-family: 'Bell MT','CMU Serif','Computer Modern',serif; background:var(--bg); padding:16px 16px 12px 16px; font-size:18px; }
        .page { max-width: 1050px; margin: 0 auto; }
        .topbar { display:flex; align-items:center; justify-content:space-between; padding:10px 14px; border:1px solid var(--border); border-radius:10px; background:white; box-shadow:0 6px 16px rgba(0,0,0,0.08); margin-bottom:14px; }
        .top-left { display:flex; align-items:center; gap:10px; font-size:22px; font-weight:bold; color: var(--border); }
        .top-right { text-align:right; font-family: 'Courier New', monospace; font-size:14px; line-height:1.4; color:#333; }
        .grid { display: grid; grid-template-columns: 1fr; gap: 12px; }
        @media (min-width: 900px) { .grid { grid-template-columns: 1fr 1.2fr; } }
        .card { background:white; padding:20px 24px; border-radius:12px; box-shadow: 0 8px 24px rgba(0,0,0,0.08); border:1px solid #d6c9b5; }
        .card-primary { border:2px solid var(--border); box-shadow: 0 10px 26px rgba(0,0,0,0.1); }
        .card-secondary { border:1px solid #e8dcc7; }
        .card-secondary.active { border:2px solid var(--primary); box-shadow: 0 10px 26px rgba(181,144,109,0.25); }
        .card h2 { margin-top:0; text-align:left; }
        label { display:block; margin-top:18px; margin-bottom:8px; font-weight:bold; text-align:left; }
        input[type="text"], select { padding:10px; width: 100%; font-size:18px; font-family: 'Bell MT','CMU Serif','Computer Modern',serif; text-align:left; border:1px solid var(--border); border-radius:8px; box-sizing:border-box; background:white; margin-bottom:16px; }
        input[type="checkbox"] { width:auto; height:auto; margin:0; accent-color: var(--border); }
        .btn { margin-top:16px; padding:14px 18px; background:#9a7754; color:white; border:1px solid var(--border); cursor:pointer; font-size:18px; font-family: 'Bell MT','CMU Serif','Computer Modern',serif; border-radius:8px; width:100%; transition: all 0.15s ease; }
        .btn:hover { background:#826340; box-shadow: 0 6px 12px rgba(0,0,0,0.12); }
        .microcopy { font-size:13px; color:#444; margin-top:6px; text-align:left; }
        .error { color:#b00020; margin-top:12px; }
        .alert { background:#fdecea; color:#611a15; border:1px solid #f5c6cb; padding:10px 12px; border-radius:8px; margin-top:12px; }
        .result { margin-top:10px; text-align:left; }
        .preview-card { min-height: 300px; }
        .header-title { display:none; }
        .dropzone { border:1.5px dashed #a6937f; border-radius:12px; background:rgba(181,144,109,0.05); padding:16px; text-align:center; cursor:pointer; transition: all 0.2s ease; box-shadow: none; }
        .dropzone.hover { background: rgba(181,144,109,0.08); border:1.5px solid #8c7458; box-shadow: 0 8px 18px rgba(0,0,0,0.08); }
        .dropzone img { width: 52px; height: auto; display:block; margin:0 auto 10px; }
        .dropzone .dz-title { font-weight:600; margin-bottom:8px; font-size:17px; color:#3a342c; }
        .dropzone .dz-sub { color:#555; font-size:13px; margin-bottom:6px; }
        .dropzone .dz-hint { color:#777; font-size:12px; }
        .dropzone input { display:none; }
        .file-preview { border:1px solid var(--border); border-radius:10px; padding:14px; background:#fff; display:flex; align-items:center; gap:12px; box-shadow:0 6px 16px rgba(0,0,0,0.08); margin-bottom:16px; }
        .file-preview img { width:52px; height:auto; }
        .file-meta { flex:1; text-align:left; }
        .file-meta .name { font-weight:600; color:var(--border); font-size:15px; }
        .file-meta .status-pill { display:inline-block; padding:4px 8px; border-radius:10px; background:#e1f3e1; color:#2e7d32; font-size:12px; margin-top:4px; }
        .file-meta .meta-line { font-size:13px; color:#555; margin-top:4px; }
        .file-actions { display:flex; gap:10px; }
        .file-actions button { padding:8px 12px; font-size:14px; border-radius:10px; border:1px solid var(--border); background:#f8f1e4; cursor:pointer; height:34px; }
        .file-actions button.primary { background:#fff; }
        .file-actions button:hover { background:#e6d7c0; }
        footer { margin-top:18px; text-align:center; font-size:14px; color:#444; }
        .nav-row { display:flex; gap:8px; margin:8px 0 12px 0; }
        .nav-btn { padding:10px 14px; border:1px solid var(--border); border-radius:10px; background:#fff; cursor:pointer; font-size:16px; text-decoration:none; color:var(--border); }
        .nav-btn.active { background:var(--primary); color:#fff; }
      </style>
    </head>
    <body>
      <div class="page">
        <div class="topbar">
          <div class="top-left">
            <span class="top-title">FGN Dispatch Desk</span>
          </div>
          <div class="top-right">
            <div>Status: {{ status_text }}</div>
            <div>File: {{ file_display }}</div>
            <div>Last processed: {{ last_processed }}</div>
            <div>File modified: {{ file_modified }}</div>
            <div>Today: {{ today_str }}</div>
          </div>
        </div>
        <div class="nav-row">
          <a class="nav-btn {% if mode=='dispatch' %}active{% endif %}" href="/?mode=dispatch">Dispatch Desk</a>
          <a class="nav-btn {% if mode=='delivery' %}active{% endif %}" href="/?mode=delivery">Delivery Slip</a>
        </div>
        <div style="margin: 8px 0 12px 0; display:flex; justify-content:flex-end;">
          <a href="/download-template" style="text-decoration:none;">
            <button class="btn" type="button" style="width:auto; padding:10px 14px;">Download Master Roll Template</button>
          </a>
        </div>
        <div class="grid">
          {% if mode == 'delivery' %}
          <div class="card card-primary">
            <h2>Delivery Slip Generator</h2>
            <form method="post" enctype="multipart/form-data">
              <input type="hidden" name="mode" value="delivery" />
              <label for="delivery_vendor">Vendor Name *</label>
              <input type="text" id="delivery_vendor" name="delivery_vendor" required placeholder="e.g., Anand" />
              <label for="delivery_workbook">Delivery Slip Master Roll (.xlsx)</label>
              <div id="delivery-workbook-drop" class="dropzone">
                <img src="{{ clipart_data }}" alt="Excel file" />
                <div class="dz-title">Drop Delivery Slip Master Roll</div>
                <div class="dz-sub">or click to browse (.xlsx)</div>
                <div class="dz-hint">Excel files only (.xlsx)</div>
                <input type="file" id="delivery_workbook" name="delivery_workbook" accept=".xlsx" {% if not delivery_file_name %}required{% endif %} />
              </div>
              {% if delivery_file_name %}
                <div class="microcopy">Using: {{ delivery_file_name }}</div>
              {% endif %}
              <label for="photos">Upload photos (folder or select multiple)</label>
              <div id="delivery-photos-drop" class="dropzone">
                <div class="dz-title">Drop photos here</div>
                <div class="dz-sub">or click to browse (PNG, JPG, JPEG)</div>
                <div class="dz-hint">You can select multiple files</div>
                <input type="file" id="photos" name="photos" accept="image/png,image/jpeg" multiple required />
              </div>
              <button class="btn" type="submit">Generate</button>
              <div class="microcopy">Generates preview before download.</div>
            </form>
              {% if error %}<div class="alert">Error: {{error}}</div>{% endif %}
          </div>
          <div class="card preview-card {% if delivery_snippet %}card-secondary active{% else %}card-secondary{% endif %}">
            <h2>Preview & Download</h2>
            {% if delivery_snippet %}
              <div class="result">
                {{ delivery_snippet }}
                <form action="/delivery-pdf" method="get" style="margin-top:12px;">
                  <button class="btn" type="submit">Download PDF</button>
                </form>
              </div>
            {% else %}
              <p style="color:#666;">Generate a delivery slip to see the preview here.</p>
            {% endif %}
          </div>
          {% else %}
          <div class="card card-primary">
            <h2>Slip Generator</h2>
            <form method="post" enctype="multipart/form-data">
              <input type="hidden" name="mode" value="dispatch" />
              <input type="hidden" name="upload_only" id="upload_only" value="">
              <label for="slip_type">Slip type</label>
              <select name="slip_type" id="slip_type" required>
                <option value="" {% if not slip_type %}selected{% endif %}></option>
                <option value="order" {% if slip_type=='order' %}selected{% endif %}>Order slip (by Vendor)</option>
                <option value="invoice" {% if slip_type=='invoice' %}selected{% endif %}>Final invoice (Packed Quantity, by Vendor)</option>
                <option value="purchase" {% if slip_type=='purchase' %}selected{% endif %}>Purchase slip (by Source)</option>
                <option value="label" {% if slip_type=='label' %}selected{% endif %}>Labels (by Vendor)</option>
              </select>
              <div style="display:flex; align-items:center; gap:8px; margin:10px 0 6px 0;">
                <input type="checkbox" id="all_mode" name="all_mode" value="1" {% if request.form.get('all_mode') %}checked{% endif %} />
                <label for="all_mode" style="margin:0; font-weight:normal; display:inline-block;">Generate for all (no identifier needed)</label>
              </div>
              <label for="identifier">Vendor name/code or Source</label>
              <input type="text" id="identifier" name="identifier" value="{{identifier}}" required placeholder="e.g., Prabhu, SKT C/BAY, or RSN" />
              <label for="workbook">Upload your Master Roll</label>
              {% if has_custom_file %}
                <div class="file-preview">
                  <img src="{{ clipart_data }}" alt="Excel file" />
                  <div class="file-meta">
                    <div class="name">{{ file_display }}</div>
                    <div class="status-pill">Loaded</div>
                    {% if file_rows is not none %}<div class="meta-line">Rows: {{ file_rows }} | Columns: {{ file_cols|join(', ') }}</div>{% endif %}
                    <div class="meta-line">Last processed: {{ last_processed }}</div>
                    <div class="meta-line">File modified: {{ file_modified }}</div>
                  </div>
                  <div class="file-actions">
                    <button type="button" class="primary" id="replace-btn">Replace</button>
                    <button type="submit" name="reset_file" value="1">Remove</button>
                  </div>
                </div>
                <input type="file" id="workbook" name="workbook" accept=".xlsx" style="display:none;" />
              {% else %}
                <div id="dispatch-drop" class="dropzone">
                  <img src="{{ clipart_data }}" alt="Excel file" />
                  <div class="dz-title">Drop Master Roll here</div>
                  <div class="dz-sub">or click to browse (.xlsx)</div>
                  <div class="dz-hint">Excel files only (.xlsx)</div>
                  <input type="file" id="workbook" name="workbook" accept=".xlsx" required />
                </div>
              {% endif %}
              <button class="btn" type="submit">Generate</button>
              <div class="microcopy">Generates preview before download.</div>
              {% if uploaded_name %}<div style="margin-top:6px; color:#444; text-align:left;">Using uploaded file: {{uploaded_name}}</div>{% endif %}
            </form>
              {% if error %}<div class="alert">Error: {{error}}<br/>Check spelling or confirm column values in Master Roll.</div>{% endif %}
          </div>
          <div class="card preview-card {% if html_snippet %}card-secondary active{% else %}card-secondary{% endif %}">
            <h2>Preview & Download</h2>
            {% if html_snippet %}
              <div class="result">
                {{ html_snippet | safe }}
                <form action="/pdf" method="get" style="margin-top:12px;">
                  <input type="hidden" name="slip_type" value="{{slip_type}}">
                  <input type="hidden" name="identifier" value="{{identifier}}">
                  {% if request.form.get('all_mode') %}<input type="hidden" name="all_mode" value="1">{% endif %}
                  <button class="btn" type="submit">
                    {% if slip_type=='label' %}Download Labels (ZIP){% else %}Download PDF{% endif %}
                  </button>
                </form>
              </div>
            {% else %}
              <p style="color:#666;">Generate a slip to see the preview here.</p>
            {% endif %}
          </div>
          {% endif %}
        </div>
        <footer>designed by Shivank Chadda • Internal tool • FGN Operations • v1.0</footer>
      </div>
      <script>
        const uploadOnly = document.getElementById('upload_only');
        const slipSelect = document.getElementById('slip_type');
        const identifierInput = document.getElementById('identifier');
        const allMode = document.getElementById('all_mode');

        function toggleIdentifierRequired() {
          const isAll = allMode && allMode.checked;
          if (isAll) {
            identifierInput.removeAttribute('required');
          } else {
            identifierInput.setAttribute('required', 'required');
          }
        }
        if (allMode) {
          allMode.addEventListener('change', toggleIdentifierRequired);
          toggleIdentifierRequired();
        }

        function attachDropzone(zoneId, inputId, autoSubmit) {
          const dz = document.getElementById(zoneId);
          const input = document.getElementById(inputId);
          if (!dz || !input) return;
          dz.addEventListener('click', () => input.click());
          dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('hover'); });
          dz.addEventListener('dragleave', () => dz.classList.remove('hover'));
          dz.addEventListener('drop', (e) => {
            e.preventDefault();
            dz.classList.remove('hover');
            if (e.dataTransfer.files && e.dataTransfer.files.length) {
              input.files = e.dataTransfer.files;
              if (autoSubmit && uploadOnly) {
                if (slipSelect) slipSelect.removeAttribute('required');
                if (identifierInput) identifierInput.removeAttribute('required');
                uploadOnly.value = "1";
                input.form.submit();
              }
            }
          });
          input.addEventListener('change', (e) => {
            if (e.target.files && e.target.files.length) {
              if (autoSubmit && uploadOnly) {
                if (slipSelect) slipSelect.removeAttribute('required');
                if (identifierInput) identifierInput.removeAttribute('required');
                uploadOnly.value = "1";
                input.form.submit();
              }
            }
          });
        }

        attachDropzone('dispatch-drop', 'workbook', true);
        attachDropzone('delivery-workbook-drop', 'delivery_workbook', false);
        attachDropzone('delivery-photos-drop', 'photos', false);

        const replaceBtn = document.getElementById('replace-btn');
        const replaceInput = document.getElementById('workbook');
        if (replaceBtn && replaceInput) {
          replaceBtn.addEventListener('click', () => replaceInput.click());
          replaceInput.addEventListener('change', (e) => {
            if (e.target.files && e.target.files.length) {
              if (uploadOnly) {
                if (slipSelect) slipSelect.removeAttribute('required');
                if (identifierInput) identifierInput.removeAttribute('required');
                uploadOnly.value = "1";
              }
              replaceInput.form.submit();
            }
          });
        }
      </script>
    </body>
    </html>
    """
    return render_template_string(
        template,
        error=error,
        html_snippet=html_snippet,
        delivery_snippet=delivery_snippet,
        slip_type=slip_type,
        identifier=identifier,
        clipart_data=CLIPART_DATA_URI,
        logo_data=LOGO_DATA_URI,
        today_str=today_str,
        file_display=file_display,
        status_text=status_text,
        file_rows=file_rows,
        file_cols=file_cols,
        has_custom_file=(CURRENT_XLSX != DEFAULT_XLSX),
        file_modified=file_modified,
        last_processed=last_processed,
        mode=mode,
        delivery_file_name=delivery_file_name,
    )


@app.route("/pdf")
def pdf():
    slip_type = request.args.get("slip_type", "").strip().lower()
    identifier = request.args.get("identifier", "").strip()
    all_mode = request.args.get("all_mode") == "1"
    if slip_type not in {"order", "purchase", "invoice", "label"} or (not identifier and not all_mode):
        return "Missing or invalid parameters", 400
    try:
        # Use current workbook (uploaded or default)
        workbook_path = CURRENT_XLSX
        raw_df = pd.read_excel(workbook_path)
        if all_mode:
            sections = []
            if slip_type in {"order", "invoice"}:
                vendors = raw_df[["Vendor_Code", "Vendor_Name"]].dropna(how="all").drop_duplicates()
                for _, row in vendors.iterrows():
                    ident = row["Vendor_Code"] if pd.notna(row["Vendor_Code"]) else row["Vendor_Name"]
                    if pd.isna(ident):
                        continue
                    matches, meta = load_matches(raw_df, slip_type, str(ident))
                    table_df = build_table(matches, meta, slip_type)
                    sections.append((str(ident), table_df, meta))
            elif slip_type == "purchase":
                sources = raw_df["Source"].dropna().unique()
                for src in sources:
                    matches, meta = load_matches(raw_df, slip_type, str(src))
                    table_df = build_table(matches, meta, slip_type)
                    sections.append((str(src), table_df, meta))
            else:
                return "Labels do not support all-mode", 400
            if not sections:
                return "No slips found", 400
            pdf_bytes, filename = make_bulk_pdf(sections, slip_type)
        else:
            matches, meta = load_matches(raw_df, slip_type, identifier)
            if slip_type == "label":
                # Labels -> ZIP of PNGs
                label_items = build_label_items(matches, meta)
                pdf_bytes, filename = make_label_zip(label_items, meta, identifier)
            else:
                # All other slips -> PDF
                table_df = build_table(matches, meta, slip_type)
                pdf_bytes, filename = make_pdf(table_df, meta, slip_type, identifier)
    except Exception as exc:
        return f"Error: {exc}", 400
    mimetype = "application/zip" if slip_type == "label" else "application/pdf"
    return send_file(io.BytesIO(pdf_bytes), as_attachment=True, download_name=filename, mimetype=mimetype)


@app.route("/debug-markers")
def debug_markers():
    """Simple debug endpoint to inspect marker availability on the server."""
    out = []
    out.append(f"Marker files on disk: {list(MARKER_IMAGES.keys())}")
    out.append(f"Current workbook: {CURRENT_XLSX}")
    try:
        df = pd.read_excel(CURRENT_XLSX, nrows=50)
        sources = df["Source"].dropna().unique()
        for src in sources:
            key = normalize_source_name(src)
            has = key in MARKER_IMAGES
            out.append(f"Source='{src}' -> key='{key}' -> marker_exists={has}")
    except Exception as e:
        out.append(f"Error reading workbook: {e}")
    return "<br>".join(out)


@app.route("/download-template")
def download_template():
    """Download the Master Roll template with a dated filename."""
    if not os.path.exists(TEMPLATE_PATH):
        return "Template not found on server", 404
    dated = datetime.now(LOCAL_TZ).strftime("%Y-%m-%d")
    filename = f"Master_Roll_{dated}.xlsx"
    return send_file(TEMPLATE_PATH, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/delivery-pdf")
def delivery_pdf():
    if not DELIVERY_CACHE.get("pdf"):
        return "No delivery slip generated yet", 400
    return send_file(
        io.BytesIO(DELIVERY_CACHE["pdf"]),
        as_attachment=True,
        download_name=DELIVERY_CACHE.get("filename", "delivery.pdf"),
        mimetype="application/pdf",
    )


if __name__ == "__main__":
    app.run(debug=True)
