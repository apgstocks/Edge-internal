"""
Edge Metals Invoice Portal - With Separate Documents Support
Run:  python app.py
Open: http://localhost:5000
"""

from flask import Flask, render_template, request, jsonify, send_file
import subprocess
import os
import csv
import io
import glob
import tempfile
import shutil
import requests
import zipfile

app = Flask(__name__)

# Configuration
GOOGLE_SHEET_ID = "1QsCeuqeRKODuouzO2PfKbxG9qJpN8yAbIurSzhI--6s"
MAIN_SHEET_GID  = "571096144"

def read_google_sheet():
    """Read Google Sheet via CSV export."""
    url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv&gid={MAIN_SHEET_GID}"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        content = response.content.decode('utf-8')
        reader  = csv.reader(io.StringIO(content))
        all_rows = list(reader)
        if not all_rows:
            return [], []
        headers   = all_rows[0]
        data_rows = [row for row in all_rows[1:] if any(cell.strip() for cell in row)]
        print(f"✅ Loaded {len(data_rows)} invoices")
        return headers, data_rows
    except Exception as e:
        print(f"❌ Error: {e}")
        return [], []


def find_container_column(headers):
    for i, header in enumerate(headers):
        if "container" in header.lower():
            return i
    return 5


def find_all_container_rows(container_no):
    """Return ALL rows matching the container number, plus headers."""
    headers, rows = read_google_sheet()
    if not rows:
        return [], [], None

    container_col = find_container_column(headers)
    container_no  = container_no.upper().strip()
    matched = []

    for i, row in enumerate(rows):
        if len(row) > container_col:
            if str(row[container_col]).upper().strip() == container_no:
                matched.append((i + 2, row))   # (1-based sheet row, row data)

    return matched, headers, container_col


def safe_get(row, idx, default=""):
    try:
        return str(row[idx]).strip() if len(row) > idx and row[idx] else default
    except:
        return default


def is_blank_or_zero(val) -> bool:
    """True for '', whitespace, or a numeric-zero placeholder like '0' / '0.0'.

    `value or "TO BE ADVISED"` treats "0.0" as truthy (non-empty string), so a
    sheet formula that defaults to a literal zero instead of a true blank cell
    renders as a port name ("0.0") instead of falling through to the real
    default. Mirrors the same fix applied in invoice_gen.py — kept as a local
    copy here rather than importing across files, since app.py already avoids
    a top-level dependency on invoice_gen.py (it's only imported dynamically,
    for the proforma endpoints).
    """
    s = str(val).strip() if val is not None else ""
    if not s:
        return True
    try:
        return float(s) == 0.0
    except (ValueError, TypeError):
        return False


PACKING_SHEET_GID = "1340048377"

PORT_ALIASES = {
    "LA": "LOS ANGELES", "LAX": "LOS ANGELES", "LB": "LONG BEACH",
    "NY": "NEW YORK", "CHI": "CHICAGO", "HOU": "HOUSTON",
    "SAV": "SAVANNAH", "OAK": "OAKLAND",
}

def read_packing_data(container_no):
    """
    Read packing sheet for a container.
    Returns (port_loading, port_discharge, place_of_receipt, packing_rows)
    where packing_rows is a list of dicts with per-item weights.
    """
    url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv&gid={PACKING_SHEET_GID}"
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        reader = csv.reader(io.StringIO(resp.content.decode("utf-8")))
        rows = list(reader)
    except Exception:
        return "LOS ANGELES", "TO BE ADVISED", "", []

    container_no     = container_no.upper().strip()
    packing_rows     = []
    port_loading     = "LOS ANGELES"
    port_discharge   = "TO BE ADVISED"
    place_of_receipt = ""

    for row in rows[1:]:
        if len(row) > 8 and row[8].upper().strip() == container_no:
            # Parse ports from first matching row's carrier column
            if not packing_rows:
                carrier_raw = row[0].strip()
                if "/" in carrier_raw:
                    parts = [p.strip() for p in carrier_raw.split("/")]
                    if len(parts) >= 3:
                        place_of_receipt = parts[0]
                        port_loading     = parts[1]
                        port_discharge   = parts[2]
                    else:
                        place_of_receipt = ""
                        port_loading     = parts[0]
                        port_discharge   = parts[1]
                else:
                    port_loading = carrier_raw or "LOS ANGELES"

                port_loading     = PORT_ALIASES.get(port_loading.upper(),     port_loading)
                port_discharge   = PORT_ALIASES.get(port_discharge.upper(),   port_discharge)
                place_of_receipt = PORT_ALIASES.get(place_of_receipt.upper(), place_of_receipt)

            def sg(r, i): return r[i].strip() if i < len(r) else ""

            packing_rows.append({
                "item_desc":        sg(row, 10),
                "gross_weight_lbs": sg(row, 11),
                "truck_lbs":        sg(row, 12),
                "container_tare":   sg(row, 13),
                "chassis_lbs":      sg(row, 14),
                "boxes_lbs":        sg(row, 15),
                "net_weight_lbs":   sg(row, 17),
                "net_weight_mt":    sg(row, 18),
            })

    return (
        port_loading or "LOS ANGELES",
        port_discharge or "TO BE ADVISED",
        place_of_receipt,
        packing_rows,
    )


def read_packing_ports(container_no):
    """Backwards-compat wrapper — used by PDF generator path."""
    loading, discharge, receipt, _ = read_packing_data(container_no)
    return loading, discharge, receipt


@app.route("/")
def landing():
    return render_template("index.html")

@app.route("/generate")
def generate_page():
    return render_template("generate.html")

@app.route("/verify")
def verify_page():
    return render_template("verify.html")

@app.route("/proforma")
def proforma_page():
    return render_template("proforma.html")


@app.route("/api/preview", methods=["POST"])
def api_preview():
    """Preview invoice data — returns all line items across one or more containers.

    Each line item keeps its OWN container/booking/seal/packing-weight data —
    combining containers must not silently stamp every row with only the
    first container's identifiers (that was the actual multi-container bug:
    the UI to add containers was missing, but even the merge logic beneath
    it collapsed per-row identity down to a single container).
    """
    body = request.json or {}
    container_nos = body.get("container_nos") or []
    if not container_nos:
        single = body.get("container_no", "").strip()
        if single:
            container_nos = [single]
    container_nos = [c.strip().upper() for c in container_nos if c.strip()]

    if not container_nos:
        return jsonify({"error": "Please enter a container number."}), 400

    # Gather matched rows per container, tagging each row with its own container.
    # (row_num, row, container_no) — preserves per-row identity through the merge.
    combined = []
    packing_cache = {}   # container_no -> packing_rows, fetched once per container
    for c in container_nos:
        matched, headers, container_col = find_all_container_rows(c)
        if not matched:
            return jsonify({"error": f"Container '{c}' not found."}), 404
        for row_num, row in matched:
            combined.append((row_num, row, c))
        if c not in packing_cache:
            _, port_discharge_pack, _, packing_rows = read_packing_data(c)
            packing_cache[c] = packing_rows

    # Header fields (invoice-level, not container-level) come from the FIRST
    # container's first row — same convention as invoice_gen.py's rows_to_invoice_data.
    first_container = container_nos[0]
    first_row_num, first_row, _ = combined[0]

    inv_no    = safe_get(first_row,  1, "N/A")
    inv_date  = safe_get(first_row,  2, "N/A")
    consignee = safe_get(first_row,  0, "N/A")
    terms     = safe_get(first_row,  8, "N/A")
    # Port fields are invoice-level too (typically the same port for every
    # container in one combined shipment) — sourced from the first container.
    port_loading, port_discharge_pack, place_of_receipt, _ = read_packing_data(first_container)
    port_dis = "TO BE ADVISED" if is_blank_or_zero(port_discharge_pack) else port_discharge_pack

    # Build one line-item entry per matched row, across ALL containers —
    # booking/seal/container/packing-weights are per-item, not invoice-level.
    line_items = []
    subtotal   = 0.0
    freight    = 0.0
    efs        = 0.0

    for row_num, row, row_container in combined:
        item_desc = safe_get(row, 12, "N/A")
        weight    = safe_get(row, 13, "0")
        price     = safe_get(row, 14, "0")
        freight_r = safe_get(row, 21, "0")
        efs_r     = safe_get(row, 25, "0")
        row_booking = safe_get(row, 4, "N/A")
        row_seal    = safe_get(row, 6, "N/A")

        try:
            w = float(weight.replace(",", ""))
            p = float(price.replace(",", "").replace("$", ""))
            amt = w * p
        except:
            w = p = amt = 0.0

        # Only take freight/efs from first non-zero row (container-level)
        if not freight:
            try:
                freight = float(freight_r.replace(",", "").replace("$", ""))
            except:
                freight = 0.0
        if not efs:
            try:
                efs = float(efs_r.replace(",", "").replace("$", ""))
            except:
                efs = 0.0

        subtotal += amt
        # Match packing row by description — within THIS row's own container
        row_packing_rows = packing_cache.get(row_container, [])
        pr = next((r for r in row_packing_rows
                   if r["item_desc"].upper() in item_desc.upper()
                   or item_desc.upper() in r["item_desc"].upper()), {})
        line_items.append({
            "row":              row_num,
            "item_desc":        item_desc,
            "weight":           f"{w:.3f} MT" if w <= 1000 else f"{w:,.2f} LBS",
            "price":            f"${p:,.2f}",
            "amount":           f"${amt:,.2f}",
            "gross_weight_lbs": pr.get("gross_weight_lbs", "—"),
            "truck_lbs":        pr.get("truck_lbs",        "—"),
            "container_tare":   pr.get("container_tare",   "—"),
            "chassis_lbs":      pr.get("chassis_lbs",      "—"),
            "boxes_lbs":        pr.get("boxes_lbs",        "—"),
            "net_weight_lbs":   pr.get("net_weight_lbs",   "—"),
            "net_weight_mt":    pr.get("net_weight_mt",    f"{w:.3f}"),
            "container_no":     row_container,
            "booking_no":       row_booking,
            "seal_no":          row_seal,
        })

    final_amt = subtotal - freight - efs

    return jsonify({
        "container_no":  container_nos[0],
        "container_nos": container_nos,
        "first_row":     first_row_num,
        "preview": {
            "invoice": {
                "Invoice No":   inv_no,
                "Invoice Date": inv_date,
                "Reference":    safe_get(first_row, 11),
                "Proforma Date": safe_get(first_row, 10),
            },
            "shipment": {
                "Container #":        ", ".join(container_nos),
                "Booking #":          line_items[0]["booking_no"] if line_items else "N/A",
                "Seal #":             line_items[0]["seal_no"] if line_items else "N/A",
                "Port of Loading":    port_loading,
                "Port of Discharge":  port_dis or "TO BE ADVISED",
                "Place of Receipt":   place_of_receipt,
            },
            "buyer": {
                "Company": consignee,
                "Terms":   terms,
            },
            "financials": {
                "Freight Deduction": f"${freight:,.2f}" if freight > 0 else "—",
                "EFS":               f"${efs:,.2f}"     if efs > 0     else "—",
                "Final Amount":      f"${final_amt:,.2f}",
            },
            "packing": {
                "Total Net Weight (MT)": f"{sum(float(item['weight'].replace(' MT','').replace(' LBS','').replace(',','')) for item in line_items):.3f}",
            }
        },
        "line_items": line_items,
    })


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """Generate PDF — supports single or multiple containers."""
    body             = request.json or {}
    separate         = body.get("separate",         False)
    vessel           = body.get("vessel",           "").strip()
    terms            = body.get("terms",            "").strip()
    payment_terms    = body.get("payment_terms",    "").strip()
    inv_no_override  = body.get("inv_no",           "").strip()
    inv_date_override= body.get("inv_date",         "").strip()
    reference_override   = body.get("reference",      "").strip()
    proforma_override    = body.get("proforma_date", "").strip()
    booking_no_override = body.get("booking_no", "").strip()
    seal_no_override    = body.get("seal_no",    "").strip()
    pol_override     = body.get("port_loading",     "").strip().upper()
    pod_override     = body.get("port_discharge",   "").strip().upper()
    receipt_override = body.get("place_of_receipt", "").strip()
    freight_override = body.get("freight",          "").strip()
    efs_override     = body.get("efs",              "").strip()
    line_items_override = body.get("line_items_override", [])
    note_rows           = body.get("note_rows", [])

    container_nos = body.get("container_nos") or []
    if not container_nos:
        single = body.get("container_no", "").strip()
        if single:
            container_nos = [single]
    container_nos = [c.strip().upper() for c in container_nos if c.strip()]

    if not container_nos:
        return jsonify({"error": "At least one container required."}), 400

    for c in container_nos:
        matched, _, _ = find_all_container_rows(c)
        if not matched:
            return jsonify({"error": f"Container '{c}' not found."}), 404

    first_matched, _, _ = find_all_container_rows(container_nos[0])
    first_row    = first_matched[0][1]
    inv_no_label = safe_get(first_row, 1, container_nos[0]).replace("/", "_").replace(" ", "_")
    container_label = "_".join(c.replace("/", "_").replace(" ", "_") for c in container_nos)
    if len(container_nos) > 1:
        inv_no_label += f"_+{len(container_nos)-1}more"

    temp_dir = tempfile.mkdtemp()

    try:
        print(f"\n{'='*60}")
        print(f"🔄 Generating {'SEPARATE' if separate else 'COMBINED'} documents")
        print(f"📦 Containers: {container_nos}  |  Vessel: '{vessel}'")
        print(f"{'='*60}\n")

        cmd = [
            "python", "/home/invoice007/invoice/invoice_gen.py",
            "--container", container_nos[0],
            "--output",    temp_dir,
        ]
        if len(container_nos) > 1:
            cmd += ["--extra-containers"] + container_nos[1:]
        if separate:          cmd.append("--separate")
        if vessel:            cmd += ["--vessel",           vessel]
        if terms:             cmd += ["--terms",            terms]
        if payment_terms:     cmd += ["--payment-terms",    payment_terms]
        cmd += ["--inv-no",    inv_no_override    or "__CLEAR__"]
        cmd += ["--inv-date",  inv_date_override  or "__CLEAR__"]
        cmd += ["--reference",     reference_override or "__CLEAR__"]
        cmd += ["--proforma-date", proforma_override    or "__CLEAR__"]
        if booking_no_override: cmd += ["--booking-no", booking_no_override]
        if seal_no_override:    cmd += ["--seal-no",    seal_no_override]
        if pol_override:      cmd += ["--port-loading",     pol_override]
        if pod_override:      cmd += ["--port-discharge",   pod_override]
        if receipt_override:  cmd += ["--place-of-receipt", receipt_override]
        if freight_override:  cmd += ["--freight",          freight_override]
        if efs_override:      cmd += ["--efs",              efs_override]

        if line_items_override:
            import json
            li_file = os.path.join(temp_dir, "line_items_override.json")
            with open(li_file, "w") as f:
                json.dump(line_items_override, f)
            cmd += ["--line-items-file", li_file]

        if note_rows:
            import json
            notes_file = os.path.join(temp_dir, "note_rows.json")
            with open(notes_file, "w") as f:
                json.dump(note_rows, f)
            cmd += ["--notes-file", notes_file]

        print(f"Running: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        print(f"Return code: {result.returncode}")
        if result.stdout: print(f"Output: {result.stdout}")
        if result.stderr: print(f"Errors: {result.stderr}")

        pdf_files = glob.glob(os.path.join(temp_dir, "*.pdf"))
        print(f"📄 Found {len(pdf_files)} PDF(s)")

        if not pdf_files:
            shutil.rmtree(temp_dir, ignore_errors=True)
            error_detail = result.stderr[:300] if result.stderr else "Unknown error"
            return jsonify({"error": f"PDF generation failed. {error_detail}"}), 500

        if separate:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for pdf_file in sorted(pdf_files):
                    with open(pdf_file, "rb") as pf:
                        zipf.writestr(os.path.basename(pdf_file), pf.read())
            zip_buffer.seek(0)
            shutil.rmtree(temp_dir, ignore_errors=True)
            return send_file(
                zip_buffer,
                mimetype="application/zip",
                as_attachment=True,
                download_name=f"{inv_no_label}_DOCUMENTS.zip"
            )
        else:
            pdf_path = pdf_files[0]
            with open(pdf_path, "rb") as f:
                pdf_buffer = io.BytesIO(f.read())
            shutil.rmtree(temp_dir, ignore_errors=True)
            pdf_buffer.seek(0)
            return send_file(
                pdf_buffer,
                mimetype="application/pdf",
                as_attachment=True,
                download_name=f"Invoice and Packing list_{container_label}.pdf"
            )

    except subprocess.TimeoutExpired:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": "Generation timed out."}), 500
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        import traceback
        print(f"\n❌ ERROR:\n{traceback.format_exc()}")
        return jsonify({"error": f"Error: {str(e)}"}), 500


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    return jsonify({"message": "Data refreshed from Google Sheets!"})


@app.route("/api/preview-pdf", methods=["POST"])
def api_preview_pdf():
    """Returns first PDF inline for browser preview (not as attachment)."""
    body             = request.json or {}
    vessel           = body.get("vessel",           "").strip()
    terms            = body.get("terms",            "").strip()
    payment_terms    = body.get("payment_terms",    "").strip()
    inv_no_override  = body.get("inv_no",           "").strip()
    inv_date_override= body.get("inv_date",         "").strip()
    reference_override   = body.get("reference",      "").strip()
    proforma_override    = body.get("proforma_date", "").strip()
    booking_no_override = body.get("booking_no", "").strip()
    seal_no_override    = body.get("seal_no",    "").strip()
    pol_override     = body.get("port_loading",     "").strip().upper()
    pod_override     = body.get("port_discharge",   "").strip().upper()
    receipt_override = body.get("place_of_receipt", "").strip()
    freight_override = body.get("freight",          "").strip()
    efs_override     = body.get("efs",              "").strip()
    line_items_override = body.get("line_items_override", [])
    note_rows           = body.get("note_rows", [])

    container_nos = body.get("container_nos") or []
    if not container_nos:
        single = body.get("container_no", "").strip()
        if single:
            container_nos = [single]
    container_nos = [c.strip().upper() for c in container_nos if c.strip()]

    if not container_nos:
        return jsonify({"error": "At least one container required."}), 400

    temp_dir = tempfile.mkdtemp()
    try:
        cmd = [
            "python", "/home/invoice007/invoice/invoice_gen.py",
            "--container", container_nos[0],
            "--output",    temp_dir,
        ]
        if len(container_nos) > 1:
            cmd += ["--extra-containers"] + container_nos[1:]
        if vessel:            cmd += ["--vessel",           vessel]
        if terms:             cmd += ["--terms",            terms]
        if payment_terms:     cmd += ["--payment-terms",    payment_terms]
        cmd += ["--inv-no",    inv_no_override    or "__CLEAR__"]
        cmd += ["--inv-date",  inv_date_override  or "__CLEAR__"]
        cmd += ["--reference",     reference_override or "__CLEAR__"]
        cmd += ["--proforma-date", proforma_override    or "__CLEAR__"]
        if booking_no_override: cmd += ["--booking-no", booking_no_override]
        if seal_no_override:    cmd += ["--seal-no",    seal_no_override]
        if pol_override:      cmd += ["--port-loading",     pol_override]
        if pod_override:      cmd += ["--port-discharge",   pod_override]
        if receipt_override:  cmd += ["--place-of-receipt", receipt_override]
        if freight_override:  cmd += ["--freight",          freight_override]
        if efs_override:      cmd += ["--efs",              efs_override]

        if line_items_override:
            import json
            li_file = os.path.join(temp_dir, "line_items_override.json")
            with open(li_file, "w") as f:
                json.dump(line_items_override, f)
            cmd += ["--line-items-file", li_file]

        if note_rows:
            import json
            notes_file = os.path.join(temp_dir, "note_rows.json")
            with open(notes_file, "w") as f:
                json.dump(note_rows, f)
            cmd += ["--notes-file", notes_file]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        pdf_files = sorted(glob.glob(os.path.join(temp_dir, "*.pdf")))

        if not pdf_files:
            shutil.rmtree(temp_dir, ignore_errors=True)
            error_detail = result.stderr[:300] if result.stderr else "Unknown error"
            return jsonify({"error": f"PDF generation failed. {error_detail}"}), 500

        with open(pdf_files[0], "rb") as f:
            pdf_buffer = io.BytesIO(f.read())
        shutil.rmtree(temp_dir, ignore_errors=True)
        pdf_buffer.seek(0)
        return send_file(pdf_buffer, mimetype="application/pdf", as_attachment=False)

    except subprocess.TimeoutExpired:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": "Generation timed out."}), 500
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": f"Error: {str(e)}"}), 500




@app.route("/api/debug-row/<container_no>")
def api_debug_row(container_no):
    """Debug: show ALL column values for a container row with indices."""
    matched, headers, _ = find_all_container_rows(container_no)
    if not matched:
        return jsonify({"error": f"Container {container_no} not found"})

    row_num, row = matched[0]
    result = {}
    for i, val in enumerate(row):
        result[f"col_{i:02d}"] = val

    # Highlight likely freight columns
    highlights = {
        "col_19": row[19] if len(row) > 19 else "OUT OF RANGE",
        "col_20": row[20] if len(row) > 20 else "OUT OF RANGE",
        "col_21": row[21] if len(row) > 21 else "OUT OF RANGE",
        "col_22": row[22] if len(row) > 22 else "OUT OF RANGE",
        "col_23": row[23] if len(row) > 23 else "OUT OF RANGE",
        "col_24": row[24] if len(row) > 24 else "OUT OF RANGE",
        "col_25": row[25] if len(row) > 25 else "OUT OF RANGE",
    }
    return jsonify({
        "row_number": row_num,
        "total_columns": len(row),
        "cols_19_to_25": highlights,
        "all_columns": result,
    })

# ── REPLACE the verify section in app.py with this ───────────────────────────
# Replaces: extract_statement_records, read_sheet_hbl_map, api_verify

import pdfplumber
import re as _re
from datetime import datetime
from collections import defaultdict

def extract_statement_records(pdf_bytes):
    """
    Extract HBL number, amount AND date from ZIMEX GLT statement PDF.
    Returns list of dicts: {hbl, amount, date}
    Multiple rows with same HBL are kept separate (will be summed later).
    """
    records = []
    seen_lines = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split('\n'):
                m = _re.search(
                    r'(\d{2}/\d{2}/\d{4})\s+GLTOER-\d+\s+(GLTOEH-\d+)\s+GLTINV-\d+\s+([\d,]+\.00)',
                    line
                )
                if m:
                    date_str = m.group(1)
                    hbl      = m.group(2).strip()
                    amount   = float(m.group(3).replace(',', ''))
                    key      = (date_str, hbl, amount)
                    if key not in seen_lines:
                        seen_lines.add(key)
                        try:
                            date = datetime.strptime(date_str, "%m/%d/%Y")
                        except:
                            date = None
                        records.append({"hbl": hbl, "amount": amount, "date": date})

    return records


def read_sheet_hbl_map():
    """
    Read Google Sheet and return HBL -> total freight amount.
    Sums multiple rows with the same HBL number.
    """
    headers, rows = read_google_sheet()
    if not rows:
        return {}

    hbl_totals = defaultdict(float)

    for row in rows:
        hbl         = safe_get(row, 3, "").strip()
        freight_raw = safe_get(row, 21, "0").replace(",", "").replace("$", "").strip()
        if not hbl:
            continue
        try:
            freight = float(freight_raw) if freight_raw else 0.0
        except:
            freight = 0.0
        if freight > 0:
            hbl_totals[hbl] += freight

    return dict(hbl_totals)


@app.route("/api/verify", methods=["POST"])
def api_verify():
    try:
        pdf_files    = request.files.getlist("pdf_files")
        year_str     = request.form.get("year",  "").strip()
        month_str    = request.form.get("month", "").strip()

        if not pdf_files:
            return jsonify({"error": "No PDF files uploaded."}), 400

        filter_year  = int(year_str)  if year_str.isdigit()  else None
        filter_month = int(month_str) if month_str.isdigit() else None

        print(f"[VERIFY] Filter: year={filter_year} month={filter_month}")

        # Extract all records from PDFs
        all_records = []
        for pdf_file in pdf_files:
            try:
                records = extract_statement_records(pdf_file.read())
                all_records.extend(records)
                print(f"[VERIFY] {pdf_file.filename}: {len(records)} records")
            except Exception as e:
                print(f"[VERIFY] Failed: {pdf_file.filename}: {e}")

        if not all_records:
            return jsonify({"error": "No records extracted from uploaded PDFs."}), 422

        # Apply month/year filter
        if filter_year or filter_month:
            filtered = []
            for r in all_records:
                d = r.get("date")
                if d is None:
                    filtered.append(r)
                    continue
                if (filter_year  is None or d.year  == filter_year) and \
                   (filter_month is None or d.month == filter_month):
                    filtered.append(r)
            print(f"[VERIFY] After filter: {len(filtered)}/{len(all_records)}")
            all_records = filtered

        if not all_records:
            return jsonify({
                "total_pdf": 0, "matched": 0, "missing": [], "mismatch": [],
                "info": f"No records found for {filter_month}/{filter_year} in the uploaded PDF."
            })

        # Sum PDF amounts by HBL
        pdf_hbl_totals = defaultdict(float)
        pdf_hbl_dates  = {}  # HBL -> earliest date string
        for r in all_records:
            hbl = r["hbl"]
            pdf_hbl_totals[hbl] += r["amount"]
            if hbl not in pdf_hbl_dates and r["date"]:
                pdf_hbl_dates[hbl] = r["date"].strftime("%m/%d/%Y")

        # Load sheet HBL totals (already summed)
        sheet_hbl_totals = read_sheet_hbl_map()

        # Compare
        matched  = []
        missing  = []
        mismatch = []

        for hbl, pdf_total in pdf_hbl_totals.items():
            date_s = pdf_hbl_dates.get(hbl, "—")

            if hbl not in sheet_hbl_totals:
                missing.append({
                    "hbl"       : hbl,
                    "date"      : date_s,
                    "pdf_amount": f"${pdf_total:,.2f}",
                })
            elif abs(sheet_hbl_totals[hbl] - pdf_total) > 0.01:
                mismatch.append({
                    "hbl"         : hbl,
                    "date"        : date_s,
                    "pdf_amount"  : f"${pdf_total:,.2f}",
                    "sheet_amount": f"${sheet_hbl_totals[hbl]:,.2f}",
                    "difference"  : f"${abs(sheet_hbl_totals[hbl] - pdf_total):,.2f}",
                })
            else:
                matched.append(hbl)

        print(f"[VERIFY] HBLs: total={len(pdf_hbl_totals)} matched={len(matched)} missing={len(missing)} mismatch={len(mismatch)}")

        return jsonify({
            "total_pdf": len(pdf_hbl_totals),
            "matched"  : len(matched),
            "missing"  : missing,
            "mismatch" : mismatch,
        })

    except Exception as e:
        import traceback
        print(f"[VERIFY] Error:\n{traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500

GEMINI_API_KEY_BACKEND = "AQ.Ab8RN6KzHJnCcTocHKXzx3DbYOTMnlil85ww3SCbeY55A1Hp1A"

@app.route("/api/proforma-smartfill", methods=["POST"])
def api_proforma_smartfill():
    """
    Parse natural-language shipment note and return dc-2 form JSON.
    Input:  { "text": "taewon confirmed 3 containers AL combo @1150 and regular combo @695" }
    Output: dc-2 fields + containers array (qty aggregated per commodity).
    """
    import re, json as _json
    body = request.get_json(force=True) or {}
    text = body.get("text", "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400

    prompt = f"""You are a trade-document parser for Edge Metals Inc., a US scrap metal exporter.

Parse the following natural-language shipment note and return ONLY a valid JSON object — no markdown, no explanation, no code fences.

Input: "{text}"

Rules:
- consignee: match to known buyers (use FULL legal name):
    "Taewon Automotive Co., Ltd." — keywords: taewon, tae won
    "ZIMEX Co., Ltd." — keywords: zimex
    "Daehan Smelting Co., Ltd." — keywords: daehan
  If unrecognised, use the name as given.
- items: array of {{"desc": string, "qty_per_container": number, "rate": number}}
  qty_per_container is MT per container. Default 21 if not stated.
  Map product names: "al combo" -> "AL Combo", "regular combo" -> "Regular Combo",
  "auto cast" -> "Auto Cast", "auto parts" -> "Scrap Auto Parts"
- num_containers: integer. Default 1.
- trade_terms: e.g. "TT · CIF Busan, South Korea". Infer from consignee country if not stated.
- port_discharge: infer from consignee country if not stated (Korea -> "Busan, South Korea")
- currency: always "US Dollar ($)"
- payment_term: default "T/T 100% Against Shipping Documents"
- shipment_allowance: default "+/- 10% on weights"
- packaging: default "Loose"
- country_of_origin: default "USA"
- memo: any special instructions (e.g. "MMR Required"), else ""
- inv_date: today as MM/DD/YYYY

Return exactly:
{{
  "consignee": "",
  "inv_date": "",
  "trade_terms": "",
  "currency": "US Dollar ($)",
  "port_discharge": "",
  "payment_term": "",
  "shipment_allowance": "",
  "packaging": "Loose",
  "country_of_origin": "USA",
  "memo": "",
  "num_containers": 1,
  "items": [{{"desc": "", "qty_per_container": 0, "rate": 0}}]
}}"""

    try:
        import requests as req
        resp = req.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY_BACKEND}",
            json={"contents": [{"parts": [{"text": prompt}]}]},
            timeout=25
        )
        resp.raise_for_status()
        raw = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
        raw = re.sub(r"```json|```", "", raw).strip()
        parsed = _json.loads(raw)
    except Exception as e:
        return jsonify({"error": f"Gemini parse failed: {e}"}), 500

    # Build containers array
    num_c = int(parsed.get("num_containers", 1))
    items = parsed.get("items", [])
    containers = []
    for i in range(num_c):
        containers.append({
            "container_no": f"(Container {i+1})",
            "items": [
                {
                    "desc":   it.get("desc", ""),
                    "qty":    float(it.get("qty_per_container", 21)),
                    "rate":   float(it.get("rate", 0)),
                    "amount": float(it.get("qty_per_container", 21)) * float(it.get("rate", 0))
                }
                for it in items
            ]
        })

    # Consignee address lookup
    consignee = parsed.get("consignee", "")
    consignee_address = []
    if consignee:
        try:
            invoice_gen_path = os.path.join(os.path.dirname(__file__), "invoice_gen.py")
            import importlib.util
            spec = importlib.util.spec_from_file_location("invoice_gen", invoice_gen_path)
            ig = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(ig)
            consignee_address = ig.get_buyer_address(consignee, ig.read_address_lookup(ig.ADDRESS_DOC_ID))
        except Exception:
            pass

    from datetime import date
    today = date.today()
    inv_no = f"EM-PI-{today.strftime('%Y%m%d')}"

    return jsonify({
        "inv_no":             inv_no,
        "inv_date":           parsed.get("inv_date", today.strftime("%m/%d/%Y")),
        "consignee":          consignee,
        "consignee_address":  consignee_address,
        "buyer_po":           "",
        "purchase_person":    "Marc Kang",
        "prepared_by":        "Marc Kang",
        "currency":           parsed.get("currency", "US Dollar ($)"),
        "trade_terms":        parsed.get("trade_terms", ""),
        "port_discharge":     parsed.get("port_discharge", ""),
        "payment_term":       parsed.get("payment_term", "T/T 100% Against Shipping Documents"),
        "shipment_allowance": parsed.get("shipment_allowance", "+/- 10% on weights"),
        "packaging":          parsed.get("packaging", "Loose"),
        "country_of_origin":  parsed.get("country_of_origin", "USA"),
        "freight_label":      "CIF (freight included)",
        "memo":               parsed.get("memo", ""),
        "bank_beneficiary":   "Edge Metals Inc.",
        "bank_name":          "",
        "bank_account_swift": "",
        "containers":         containers,
    })


@app.route("/api/generate-proforma-dc2", methods=["POST"])
def api_generate_proforma_dc2():
    """Generate dc-2 style proforma PDF (spec-list, aggregated commodities)."""
    body = request.get_json(force=True) or {}
    if not body:
        return jsonify({"error": "No data provided."}), 400

    consignee = body.get("consignee", "").strip()
    if consignee and not body.get("consignee_address"):
        try:
            invoice_gen_path = os.path.join(os.path.dirname(__file__), "invoice_gen.py")
            import importlib.util
            spec = importlib.util.spec_from_file_location("invoice_gen", invoice_gen_path)
            ig = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(ig)
            body["consignee_address"] = ig.get_buyer_address(consignee, ig.read_address_lookup(ig.ADDRESS_DOC_ID))
        except Exception:
            body["consignee_address"] = []

    inv_no   = body.get("inv_no", "PROFORMA").replace("/", "_").replace(" ", "_")
    temp_dir = tempfile.mkdtemp()
    out_path = os.path.join(temp_dir, f"{inv_no}_PROFORMA.pdf")

    try:
        invoice_gen_path = os.path.join(os.path.dirname(__file__), "invoice_gen.py")
        import importlib.util
        spec = importlib.util.spec_from_file_location("invoice_gen", invoice_gen_path)
        ig = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(ig)
        ig.generate_proforma_pdf_dc2(body, out_path)

        with open(out_path, "rb") as f:
            pdf_bytes = io.BytesIO(f.read())
        shutil.rmtree(temp_dir, ignore_errors=True)
        pdf_bytes.seek(0)
        return send_file(pdf_bytes, mimetype="application/pdf",
                         as_attachment=True,
                         download_name=f"{inv_no}_PROFORMA.pdf")
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        import traceback
        print(f"\n❌ DC2 ERROR:\n{traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("\n🚀 Edge Metals Invoice Portal")
    print("=" * 50)
    print("Server: http://localhost:5000")
    print("=" * 50)
    app.run(debug=True, port=5000, host='0.0.0.0')
