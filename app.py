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
from datetime import date  # for stamping current date
import os  # file mtime for status bar

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
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle  # PDF layout

app = Flask(__name__)

# Path to default workbook bundled with the app
DEFAULT_XLSX = "Master Roll.xlsx"
# In-memory pointer to the most recent uploaded workbook; falls back to DEFAULT_XLSX
CURRENT_XLSX = DEFAULT_XLSX

# Branding asset and label sizing in inches / pixels
LOGO_PATH = "Farmers_Wordmark_Badge_Transparent_1_3000px.png"
CLIPART_PATH = "Excel clipart.png"
MARKER_DIR = "Marker box"
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


def mm_to_px(mm: float) -> int:
    """Convert millimeters to pixels at LABEL_DPI."""
    return int(mm / 25.4 * LABEL_DPI)


def load_marker_images():
    """Load marker images from the marker folder keyed by lowercase basename (without extension)."""
    mapping = {}
    for key in ("fgn", "rsn", "prashanto"):
        path = os.path.join(MARKER_DIR, f"{key.upper()}.png")
        if not os.path.exists(path):
            path = os.path.join(MARKER_DIR, f"{key}.png")
        if os.path.exists(path):
            try:
                mapping[key] = PILImage.open(path).convert("RGBA")
            except Exception:
                continue
    return mapping


MARKER_IMAGES = load_marker_images()

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
    items_df = matches[["Item_Name", meta["quantity_column"], "Unit"]].copy()
    items_df.rename(columns={meta["quantity_column"]: "Quantity"}, inplace=True)
    items_df["Quantity"] = pd.to_numeric(items_df["Quantity"], errors="coerce")
    items_df = items_df.dropna(subset=["Item_Name"])
    items_df["Item_Name"] = items_df["Item_Name"].apply(normalize_item_name)

    # aggregate to unique items (sum quantities per item/unit)
    items_df = items_df.groupby(["Item_Name", "Unit"], as_index=False)["Quantity"].sum()
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

    font_label = load_font(46, bold=True)
    font_value = load_font(46, bold=False)

    # pre-load logo
    logo_img = None
    try:
        logo_img = PILImage.open(LOGO_PATH).convert("RGBA")
    except Exception:
        logo_img = None

    images = []
    border_margin = mm_to_px(11)  # outer border inset
    inner_border_offset = mm_to_px(2)  # gap between double borders
    logo_top_space = mm_to_px(14)
    logo_bottom_space = mm_to_px(11)
    text_left_margin = mm_to_px(19)
    text_top_min = mm_to_px(21)
    line_min_gap = mm_to_px(5)
    marker_offset = mm_to_px(9)
    marker_size = mm_to_px(19)

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
        y_cursor = logo_top_space + border_margin
        if logo_img:
            max_w = int(width_px * 0.58)
            max_h = int(height_px * 0.2)
            lw, lh = logo_img.size
            scale = min(max_w / lw, max_h / lh)
            new_size = (int(lw * scale), int(lh * scale))
            logo_resized = logo_img.resize(new_size, PILImage.LANCZOS)
            x_logo = (width_px - new_size[0]) // 2
            y_logo = y_cursor
            img.paste(logo_resized, (x_logo, y_logo), logo_resized)
            y_cursor = y_logo + new_size[1] + logo_bottom_space
        y_cursor = max(y_cursor, text_top_min + border_margin)

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
            draw.text((x_text, y_cursor), label_txt, font=font_label, fill="black")
            label_w = text_width(label_txt, font_label)
            draw.text((x_text + label_w + 10, y_cursor), str(val), font=font_value, fill="black")
            y_cursor += line_spacing

        # marker box bottom right with icon based on Source
        box_size = marker_size
        box_x = width_px - border_margin - marker_offset - box_size
        box_y = height_px - border_margin - marker_offset - box_size
        draw.rectangle([box_x, box_y, box_x + box_size, box_y + box_size], outline="black", width=2)
        source_key = str(meta.get("source", "") or "").lower()
        marker_img = None
        for key in ("fgn", "rsn", "prashanto"):
            if key in source_key:
                marker_img = MARKER_IMAGES.get(key)
                break
        if marker_img:
            target = int(box_size * 0.6)
            miw, mih = marker_img.size
            scale = min(target / miw, target / mih)
            new_size = (int(miw * scale), int(mih * scale))
            marker_resized = marker_img.resize(new_size, PILImage.LANCZOS)
            mx = box_x + (box_size - new_size[0]) // 2
            my = box_y + (box_size - new_size[1]) // 2
            img.paste(marker_resized, (mx, my), marker_resized)
        else:
            # default circle if no specific marker
            draw.ellipse(
                [box_x + box_size * 0.3, box_y + box_size * 0.3, box_x + box_size * 0.7, box_y + box_size * 0.7],
                fill="black",
                outline=None,
            )

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
    error = None
    html_snippet = None
    slip_type = ""
    identifier = ""
    uploaded_name = None

    if request.method == "POST":
        # Allow resetting the file without needing other inputs
        if request.form.get("reset_file") == "1":
            CURRENT_XLSX = DEFAULT_XLSX
            uploaded_name = None
            slip_type = ""
            identifier = ""
        else:
        # Read form inputs
            slip_type = request.form.get("slip_type", "").strip().lower()
            identifier = request.form.get("identifier", "").strip()
            upload = request.files.get("workbook")
            try:
                if slip_type not in {"order", "purchase", "invoice", "label"}:
                    raise ValueError("Slip type must be order, purchase, invoice, or label.")
                if not identifier:
                    raise ValueError("Identifier is required.")

                # Decide which workbook to read: uploaded or default/current
                workbook_path = CURRENT_XLSX
                if upload and upload.filename:
                    # Save uploaded workbook to a temp path and switch CURRENT_XLSX
                    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", upload.filename) or "uploaded.xlsx"
                    tmp_path = f"/tmp/{safe_name}"
                    upload.save(tmp_path)
                    CURRENT_XLSX = tmp_path
                    workbook_path = CURRENT_XLSX
                    uploaded_name = upload.filename
                elif CURRENT_XLSX == DEFAULT_XLSX:
                    # If nothing has been uploaded yet, require an upload
                    raise ValueError("Please upload the Master Roll (.xlsx) to continue.")

                raw_df = pd.read_excel(workbook_path)
                matches, meta = load_matches(raw_df, slip_type, identifier)
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
        last_updated = date.fromtimestamp(os.path.getmtime(CURRENT_XLSX)).strftime("%d %B %Y, %H:%M")
    except Exception:
        last_updated = "unknown"
    today_str = date.today().strftime("%d %b %Y")
    status_text = "Master Roll loaded" if CURRENT_XLSX != DEFAULT_XLSX else "Default Master Roll"

    # file summary for preview card
    file_rows = None
    file_cols = []
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
        input, select { padding:10px; width: 100%; font-size:18px; font-family: 'Bell MT','CMU Serif','Computer Modern',serif; text-align:left; border:1px solid var(--border); border-radius:8px; box-sizing:border-box; background:white; margin-bottom:16px; }
        .btn { margin-top:16px; padding:14px 18px; background:#9a7754; color:white; border:1px solid var(--border); cursor:pointer; font-size:18px; font-family: 'Bell MT','CMU Serif','Computer Modern',serif; border-radius:8px; width:100%; transition: all 0.15s ease; }
        .btn:hover { background:#826340; box-shadow: 0 6px 12px rgba(0,0,0,0.12); }
        .microcopy { font-size:13px; color:#444; margin-top:6px; text-align:left; }
        .error { color:#b00020; margin-top:12px; }
        .alert { background:#fdecea; color:#611a15; border:1px solid #f5c6cb; padding:10px 12px; border-radius:8px; margin-top:12px; }
        .result { margin-top:10px; text-align:left; }
        .preview-card { min-height: 300px; }
        .header-title { display:none; }
        .dropzone { border:1px dashed var(--border); border-radius:12px; background:#fff; padding:16px; text-align:center; cursor:pointer; transition: background 0.2s, border-color 0.2s; }
        .dropzone.hover { background: #f8f1e4; border-color: var(--primary); }
        .dropzone img { width: 48px; height: auto; display:block; margin:0 auto 10px; }
        .dropzone .dz-title { font-weight:bold; margin-bottom:4px; }
        .dropzone .dz-sub { color:#555; font-size:16px; }
        .dropzone input { display:none; }
        .file-preview { border:1px solid var(--border); border-radius:10px; padding:14px; background:#fff; display:flex; align-items:center; gap:12px; box-shadow:0 6px 16px rgba(0,0,0,0.08); margin-bottom:16px; }
        .file-preview img { width:48px; height:auto; }
        .file-meta { flex:1; text-align:left; }
        .file-meta .name { font-weight:bold; color:var(--border); }
        .file-meta .meta-line { font-size:14px; color:#555; }
        .file-actions { display:flex; gap:8px; }
        .file-actions button { padding:8px 10px; font-size:14px; border-radius:8px; border:1px solid var(--border); background:#f8f1e4; cursor:pointer; }
        .file-actions button:hover { background:#e6d7c0; }
        footer { margin-top:18px; text-align:center; font-size:14px; color:#444; }
      </style>
    </head>
    <body>
      <div class="page">
        <div class="topbar">
          <div class="top-left">
            <img src="{{ logo_data }}" alt="FGN">
            <span class="top-title">FGN Dispatch Desk</span>
          </div>
          <div class="top-right">
            <div>Status: {{ status_text }}</div>
            <div>File: {{ file_display }}</div>
            <div>Last Updated: {{ last_updated }}</div>
            <div>Today: {{ today_str }}</div>
          </div>
        </div>
        <div class="grid">
          <div class="card card-primary">
            <h2>Slip Generator</h2>
            <form method="post" enctype="multipart/form-data">
              <label for="slip_type">Slip type</label>
              <select name="slip_type" id="slip_type" required>
                <option value="" {% if not slip_type %}selected{% endif %}></option>
                <option value="order" {% if slip_type=='order' %}selected{% endif %}>Order slip (by Vendor)</option>
                <option value="invoice" {% if slip_type=='invoice' %}selected{% endif %}>Final invoice (Packed Quantity, by Vendor)</option>
                <option value="purchase" {% if slip_type=='purchase' %}selected{% endif %}>Purchase slip (by Source)</option>
                <option value="label" {% if slip_type=='label' %}selected{% endif %}>Labels (by Vendor)</option>
              </select>
              <label for="identifier">Vendor name/code or Source</label>
              <input type="text" id="identifier" name="identifier" value="{{identifier}}" required placeholder="e.g., Prabhu, SKT C/BAY, or RSN" />
              <label for="workbook">Upload your Master Roll</label>
              {% if has_custom_file %}
                <div class="file-preview">
                  <img src="{{ clipart_data }}" alt="Excel file" />
                  <div class="file-meta">
                    <div class="name">{{ file_display }}</div>
                    <div class="meta-line">Status: Master Roll loaded</div>
                    {% if file_rows is not none %}<div class="meta-line">Rows: {{ file_rows }} | Columns: {{ file_cols|join(', ') }}</div>{% endif %}
                  </div>
                  <div class="file-actions">
                    <button type="button" id="replace-btn">Replace</button>
                    <button type="submit" name="reset_file" value="1">Remove</button>
                  </div>
                </div>
                <input type="file" id="workbook" name="workbook" accept=".xlsx" style="display:none;" />
              {% else %}
                <div id="drop-zone" class="dropzone">
                  <img src="{{ clipart_data }}" alt="Excel file" />
                  <div class="dz-title">Drag & Drop</div>
                  <div class="dz-sub">or click to browse (.xlsx)</div>
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
                  <button class="btn" type="submit">
                    {% if slip_type=='label' %}Download Labels (ZIP){% else %}Download PDF{% endif %}
                  </button>
                </form>
              </div>
            {% else %}
              <p style="color:#666;">Generate a slip to see the preview here.</p>
            {% endif %}
          </div>
        </div>
        <footer>designed by Shivank Chadda • Internal tool • FGN Operations • v1.0</footer>
      </div>
      <script>
        const dropZone = document.getElementById('drop-zone');
        const fileInput = document.getElementById('workbook');
        const replaceBtn = document.getElementById('replace-btn');
        if (dropZone) {
          dropZone.addEventListener('click', () => fileInput.click());
          dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('hover');
          });
          dropZone.addEventListener('dragleave', () => dropZone.classList.remove('hover'));
          dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('hover');
            if (e.dataTransfer.files && e.dataTransfer.files.length) {
              fileInput.files = e.dataTransfer.files;
            }
          });
        }
        if (replaceBtn) {
          replaceBtn.addEventListener('click', () => fileInput.click());
        }
      </script>
    </body>
    </html>
    """
    return render_template_string(
        template,
        error=error,
        html_snippet=html_snippet,
        slip_type=slip_type,
        identifier=identifier,
        clipart_data=CLIPART_DATA_URI,
        logo_data=LOGO_DATA_URI,
        today_str=today_str,
        file_display=file_display,
        last_updated=last_updated,
        status_text=status_text,
        file_rows=file_rows,
        file_cols=file_cols,
        has_custom_file=(CURRENT_XLSX != DEFAULT_XLSX),
    )


@app.route("/pdf")
def pdf():
    slip_type = request.args.get("slip_type", "").strip().lower()
    identifier = request.args.get("identifier", "").strip()
    if slip_type not in {"order", "purchase", "invoice", "label"} or not identifier:
        return "Missing or invalid parameters", 400
    try:
        # Use current workbook (uploaded or default)
        workbook_path = CURRENT_XLSX
        raw_df = pd.read_excel(workbook_path)
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


if __name__ == "__main__":
    app.run(debug=True)
