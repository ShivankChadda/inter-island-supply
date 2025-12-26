"""
Simple Flask app to generate slips (order / purchase / invoice) from Master Roll.xlsx.

Run:
    export FLASK_APP=app.py
    flask run

Then open http://127.0.0.1:5000 and use the form.
"""
import io
import re
from datetime import date

import pandas as pd
from flask import Flask, render_template_string, request, send_file
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm, inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

app = Flask(__name__)

# Path to default workbook bundled with the app
DEFAULT_XLSX = "Master Roll.xlsx"
# In-memory pointer to the most recent uploaded workbook; falls back to DEFAULT_XLSX
CURRENT_XLSX = DEFAULT_XLSX

LOGO_PATH = "Farmers_Wordmark_Badge_Transparent_1_3000px.png"
LABEL_WIDTH = 3 * inch
LABEL_HEIGHT = 5 * inch


# Normalization helpers -----------------------------------------------------
ITEM_NAME_MAP = {
    "green chilli": "Green Chilli",
    "green chillie": "Green Chilli",
    "brocoli": "Broccoli",
}


def normalize_item_name(name: str) -> str:
    """Collapse whitespace, lowercase for lookup, and apply known typo fixes."""
    if not isinstance(name, str):
        return name
    base = " ".join(name.strip().split())  # collapse extra spaces
    key = base.lower()
    if key in ITEM_NAME_MAP:
        return ITEM_NAME_MAP[key]
    return base


def format_qty(x):
    try:
        val = float(x)
    except (TypeError, ValueError):
        return x
    return f"{val:.1f}".rstrip("0").rstrip(".")


def first_non_empty(series, default=""):
    if series is None:
        return default
    ser = series.dropna()
    return ser.iloc[0] if not ser.empty else default

COLS = ["Serial Number", "Item", "Quantity", "Unit"]


def load_matches(df: pd.DataFrame, slip_type: str, identifier: str) -> tuple[pd.DataFrame, dict]:
    """Filter rows based on slip type and identifier; return rows and meta."""
    ident_l = identifier.lower()
    if slip_type in {"order", "invoice", "label"}:
        name_series = df["Vendor_Name"].fillna("").str.lower()
        code_series = df["Vendor_Code"].fillna("").str.lower()
        matches = df[(name_series == ident_l) | (code_series == ident_l)].copy()
        if matches.empty:
            raise ValueError(f"No rows found for vendor '{identifier}' (by name or code)")
        vendor_code = first_non_empty(matches["Vendor_Code"], identifier)
        customer_name = vendor_code  # per requirement: use code in meta
        quantity_column = "Packed_Quantity" if slip_type == "invoice" else "Quantity"
    else:  # purchase
        source_series = df["Source"].fillna("").str.lower()
        matches = df[source_series == ident_l].copy()
        if matches.empty:
            raise ValueError(f"No rows found for source '{identifier}'")
        customer_name = identifier
        quantity_column = "Quantity"

    ship_name = first_non_empty(matches["Ship_Name"], "")
    sail_val = first_non_empty(matches["Sailing_Date"], "")
    if hasattr(sail_val, "strftime"):
        sailing_date = sail_val.strftime("%d %B")
    else:
        sailing_date = str(sail_val) if sail_val else ""

    meta = {
        "customer_name": customer_name,
        "vendor_name": first_non_empty(matches["Vendor_Name"], identifier),
        "location": first_non_empty(matches["Vendor_Location"]) if "Vendor_Location" in matches else "",
        "ship_name": ship_name,
        "sailing_date": sailing_date,
        "quantity_column": quantity_column,
    }
    return matches, meta


def build_table(matches: pd.DataFrame, meta: dict, slip_type: str) -> pd.DataFrame:
    """Construct the output table with sorting, serials, totals, and proper quantity column."""
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

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []

    meta_style = ParagraphStyle(name="Meta", fontName="Times-Roman", fontSize=12, leading=15, spaceAfter=4)
    for label, value in [
        ("Date", date.today().strftime("%d %B %Y")),
        ("Customer Name", meta["customer_name"]),
        ("Sailing Date", meta["sailing_date"]),
        ("Ship", meta["ship_name"]),
    ]:
        story.append(Paragraph(f"<b>{label}</b>: {value}", meta_style))
    story.append(Spacer(1, 12))

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

    doc.build(story)
    buf.seek(0)

    safe_id = re.sub(r"[^a-z0-9]+", "_", identifier.lower()).strip("_") or "id"
    safe_sail = re.sub(r"[^a-z0-9]+", "_", meta["sailing_date"].lower()).strip("_") or "sailing_date"
    filename = f"{slip_type}_{safe_id}_{safe_sail}_slip.pdf"
    return buf.read(), filename


def make_label_pdf(items: list[dict], meta: dict, identifier: str) -> tuple[bytes, str]:
    """Generate label PDF (one label per item) sized 3x5 inches, portrait."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(LABEL_WIDTH, LABEL_HEIGHT))

    logo_reader = None
    try:
        logo_reader = ImageReader(LOGO_PATH)
    except Exception:
        logo_reader = None

    for item in items:
        c.saveState()
        # Draw logo
        top_margin = 0.35 * inch
        available_height = LABEL_HEIGHT - 2 * top_margin
        logo_height_target = available_height * 0.18  # 15-18% of label height
        logo_width_target = LABEL_WIDTH * 0.8
        y_cursor = LABEL_HEIGHT - top_margin

        if logo_reader:
            iw, ih = logo_reader.getSize()
            scale = min(logo_width_target / iw, logo_height_target / ih)
            lw, lh = iw * scale, ih * scale
            x_logo = (LABEL_WIDTH - lw) / 2
            y_logo = y_cursor - lh
            c.drawImage(logo_reader, x_logo, y_logo, width=lw, height=lh, preserveAspectRatio=True, mask='auto')
            y_cursor = y_logo - 0.25 * inch
        else:
            y_cursor -= 0.25 * inch

        # Text block
        text_style_label = ("Times-Bold", 14)
        text_style_value = ("Times-Roman", 14)
        line_leading = 18
        lines = [
            ("Vendor Name", meta.get("vendor_name", "")),
            ("Location", meta.get("location", "")),
            ("Marker", meta.get("customer_name", "")),
            ("Item", item.get("item", "")),
            ("Weight", f"{format_qty(item.get('quantity'))} {item.get('unit', '')}".strip()),
        ]

        text = c.beginText()
        text.setTextOrigin(0.5 * inch, y_cursor)
        for label, val in lines:
            text.setFont(*text_style_label)
            text.textOut(f"{label}: ")
            text.setFont(*text_style_value)
            text.textLine(str(val))
            text.setFont(*text_style_label)
        c.drawText(text)

        # Marker box with filled circle at bottom right
        box_size = 0.45 * inch
        box_x = LABEL_WIDTH - box_size - 0.35 * inch
        box_y = 0.35 * inch
        c.rect(box_x, box_y, box_size, box_size, stroke=1, fill=0)
        c.circle(box_x + box_size / 2, box_y + box_size / 2, box_size / 5, stroke=0, fill=1)

        c.restoreState()
        c.showPage()

    c.save()
    buf.seek(0)
    safe_id = re.sub(r"[^a-z0-9]+", "_", identifier.lower()).strip("_") or "id"
    filename = f"label_{safe_id}.pdf"
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
    error = None
    html_snippet = None
    slip_type = ""
    identifier = ""
    uploaded_name = None

    if request.method == "POST":
        slip_type = request.form.get("slip_type", "").strip().lower()
        identifier = request.form.get("identifier", "").strip()
        upload = request.files.get("workbook")
        try:
            if slip_type not in {"order", "purchase", "invoice", "label"}:
                raise ValueError("Slip type must be order, purchase, invoice, or label.")
            if not identifier:
                raise ValueError("Identifier is required.")

            workbook_path = CURRENT_XLSX
            if upload and upload.filename:
                # Save uploaded workbook to a temp path and switch CURRENT_XLSX
                safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", upload.filename) or "uploaded.xlsx"
                tmp_path = f"/tmp/{safe_name}"
                upload.save(tmp_path)
                CURRENT_XLSX = tmp_path
                workbook_path = CURRENT_XLSX
                uploaded_name = upload.filename

            raw_df = pd.read_excel(workbook_path)
            matches, meta = load_matches(raw_df, slip_type, identifier)
            if slip_type == "label":
                label_items = build_label_items(matches, meta)
                html_snippet = render_label_preview(label_items, meta)
            else:
                table_df = build_table(matches, meta, slip_type)
                html_snippet = render_html(table_df, meta, slip_type, identifier)
        except Exception as exc:  # broad to show message to user
            error = str(exc)

    template = """
    <!doctype html>
    <html>
    <head>
      <title>Slip Generator</title>
      <style>
        body { font-family: 'Bell MT','CMU Serif','Computer Modern',serif; background:#f7f7f7; padding:20px; font-size:20px; text-align:center; }
        .card { background:white; padding:16px 20px; border:1px solid #ddd; border-radius:6px; max-width:900px; margin: 0 auto; font-size:20px; }
        label { display:block; margin-top:10px; font-weight:bold; text-align:center; }
        input, select { padding:8px; width: 300px; font-size:20px; font-family: 'Bell MT','CMU Serif','Computer Modern',serif; text-align:center; }
        .btn { margin-top:12px; padding:10px 16px; background:#000; color:white; border:none; cursor:pointer; font-size:20px; font-family: 'Bell MT','CMU Serif','Computer Modern',serif; }
        .error { color:#b00020; margin-top:12px; }
        .result { margin-top:20px; text-align:left; }
      </style>
    </head>
    <body>
      <div class="card">
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
          <label for="workbook">Upload Excel (optional, .xlsx)</label>
          <input type="file" id="workbook" name="workbook" accept=".xlsx" />
          <div><button class="btn" type="submit">Generate</button></div>
          {% if uploaded_name %}<div style="margin-top:6px; color:#444;">Using uploaded file: {{uploaded_name}}</div>{% endif %}
        </form>
        {% if error %}<div class="error">{{error}}</div>{% endif %}
        {% if html_snippet %}
          <div class="result">
            {{ html_snippet | safe }}
            <form action="/pdf" method="get" style="margin-top:12px;">
              <input type="hidden" name="slip_type" value="{{slip_type}}">
              <input type="hidden" name="identifier" value="{{identifier}}">
              <button class="btn" type="submit">Download PDF</button>
            </form>
          </div>
        {% endif %}
      </div>
    </body>
    </html>
    """
    return render_template_string(template, error=error, html_snippet=html_snippet, slip_type=slip_type, identifier=identifier)


@app.route("/pdf")
def pdf():
    slip_type = request.args.get("slip_type", "").strip().lower()
    identifier = request.args.get("identifier", "").strip()
    if slip_type not in {"order", "purchase", "invoice", "label"} or not identifier:
        return "Missing or invalid parameters", 400
    try:
        workbook_path = CURRENT_XLSX
        raw_df = pd.read_excel(workbook_path)
        matches, meta = load_matches(raw_df, slip_type, identifier)
        if slip_type == "label":
            label_items = build_label_items(matches, meta)
            pdf_bytes, filename = make_label_pdf(label_items, meta, identifier)
        else:
            table_df = build_table(matches, meta, slip_type)
            pdf_bytes, filename = make_pdf(table_df, meta, slip_type, identifier)
    except Exception as exc:
        return f"Error: {exc}", 400
    return send_file(io.BytesIO(pdf_bytes), as_attachment=True, download_name=filename, mimetype="application/pdf")


if __name__ == "__main__":
    app.run(debug=True)
