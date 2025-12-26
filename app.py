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
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

app = Flask(__name__)

# Path to default workbook bundled with the app
DEFAULT_XLSX = "Master Roll.xlsx"
# In-memory pointer to the most recent uploaded workbook; falls back to DEFAULT_XLSX
CURRENT_XLSX = DEFAULT_XLSX

COLS = ["Serial Number", "Item", "Quantity", "Unit"]


def load_matches(df: pd.DataFrame, slip_type: str, identifier: str) -> tuple[pd.DataFrame, dict]:
    """Filter rows based on slip type and identifier; return rows and meta."""
    ident_l = identifier.lower()
    if slip_type in {"order", "invoice"}:
        name_series = df["Vendor_Name"].fillna("").str.lower()
        code_series = df["Vendor_Code"].fillna("").str.lower()
        matches = df[(name_series == ident_l) | (code_series == ident_l)].copy()
        if matches.empty:
            raise ValueError(f"No rows found for vendor '{identifier}' (by name or code)")
        vendor_code = (
            matches["Vendor_Code"].dropna().iloc[0] if not matches["Vendor_Code"].dropna().empty else identifier
        )
        customer_name = vendor_code  # per requirement: use code in meta
        quantity_column = "Packed_Quantity" if slip_type == "invoice" else "Quantity"
    else:  # purchase
        source_series = df["Source"].fillna("").str.lower()
        matches = df[source_series == ident_l].copy()
        if matches.empty:
            raise ValueError(f"No rows found for source '{identifier}'")
        customer_name = identifier
        quantity_column = "Quantity"

    ship_name = matches["Ship_Name"].dropna().iloc[0] if not matches["Ship_Name"].dropna().empty else ""
    sail_val = matches["Sailing_Date"].dropna().iloc[0] if not matches["Sailing_Date"].dropna().empty else ""
    if hasattr(sail_val, "strftime"):
        sailing_date = sail_val.strftime("%d %B")
    else:
        sailing_date = str(sail_val) if sail_val else ""

    meta = {
        "customer_name": customer_name,
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

    if slip_type == "invoice":
        items_df = items_df[items_df["Quantity"] > 0]  # skip zero/NaN

    if slip_type == "purchase":
        # aggregate quantities per item/unit
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


def make_pdf(table_df: pd.DataFrame, meta: dict, slip_type: str, identifier: str) -> tuple[bytes, str]:
    """Generate PDF bytes and filename for the given table/meta."""
    header_idx = 0
    total_idx = len(table_df) - 1

    def fmt_qty(x):
        try:
            val = float(x)
        except (TypeError, ValueError):
            return x
        return f"{val:.1f}".rstrip("0").rstrip(".")

    data = table_df.copy()
    data["Quantity"] = data["Quantity"].apply(fmt_qty)
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


def render_html(table_df: pd.DataFrame, meta: dict, slip_type: str, identifier: str) -> str:
    """Build a simple HTML view of the slip."""
    def fmt_qty(x):
        try:
            val = float(x)
        except (TypeError, ValueError):
            return x
        return f"{val:.1f}".rstrip("0").rstrip(".")

    df = table_df.copy()
    df["Quantity"] = df["Quantity"].apply(fmt_qty)

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


@app.route("/", methods=["GET", "POST"])
def home():
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
            if slip_type not in {"order", "purchase", "invoice"}:
                raise ValueError("Slip type must be order, purchase, or invoice.")
            if not identifier:
                raise ValueError("Identifier is required.")

            workbook_path = CURRENT_XLSX
            if upload and upload.filename:
                # Save uploaded workbook to a temp path and switch CURRENT_XLSX
                safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", upload.filename) or "uploaded.xlsx"
                tmp_path = f"/tmp/{safe_name}"
                upload.save(tmp_path)
                global CURRENT_XLSX
                CURRENT_XLSX = tmp_path
                workbook_path = CURRENT_XLSX
                uploaded_name = upload.filename

            raw_df = pd.read_excel(workbook_path)
            matches, meta = load_matches(raw_df, slip_type, identifier)
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
        body { font-family: 'Bell MT','CMU Serif','Computer Modern',serif; background:#f7f7f7; padding:20px; }
        .card { background:white; padding:16px 20px; border:1px solid #ddd; border-radius:6px; max-width:900px; }
        label { display:block; margin-top:10px; font-weight:bold; }
        input, select { padding:6px; width: 280px; }
        .btn { margin-top:12px; padding:8px 14px; background:#000; color:white; border:none; cursor:pointer; }
        .error { color:#b00020; margin-top:12px; }
        .result { margin-top:20px; }
      </style>
    </head>
    <body>
      <div class="card">
        <h2>Slip Generator</h2>
        <form method="post">
          <label for="slip_type">Slip type</label>
          <select name="slip_type" id="slip_type" required>
            <option value="" {% if not slip_type %}selected{% endif %}></option>
            <option value="order" {% if slip_type=='order' %}selected{% endif %}>Order slip (by Vendor)</option>
            <option value="invoice" {% if slip_type=='invoice' %}selected{% endif %}>Final invoice (Packed Quantity, by Vendor)</option>
            <option value="purchase" {% if slip_type=='purchase' %}selected{% endif %}>Purchase slip (by Source)</option>
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
    if slip_type not in {"order", "purchase", "invoice"} or not identifier:
        return "Missing or invalid parameters", 400
    try:
        raw_df = pd.read_excel("Master Roll.xlsx")
        matches, meta = load_matches(raw_df, slip_type, identifier)
        table_df = build_table(matches, meta, slip_type)
        pdf_bytes, filename = make_pdf(table_df, meta, slip_type, identifier)
    except Exception as exc:
        return f"Error: {exc}", 400
    return send_file(io.BytesIO(pdf_bytes), as_attachment=True, download_name=filename, mimetype="application/pdf")


if __name__ == "__main__":
    app.run(debug=True)
