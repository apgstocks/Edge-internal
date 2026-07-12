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
import json
import tempfile
import shutil
import requests
import zipfile
from datetime import datetime

app = Flask(__name__)

# Configuration
GOOGLE_SHEET_ID = "1QsCeuqeRKODuouzO2PfKbxG9qJpN8yAbIurSzhI--6s"
MAIN_SHEET_GID  = "571096144"

# ─────────────────────────────────────────────────────────────────────
#  GENERATED-INVOICE PERSISTENCE
# ─────────────────────────────────────────────────────────────────────
# Every successful /api/generate saves the POST body + metadata as JSON
# on disk, keyed by the (sorted) container list. On subsequent previews,
# the UI is told whether saved overrides exist so it can offer to reload
# them. Rationale: users edit invoices days or weeks later; re-typing
# every override from scratch is error-prone. The saved JSON IS the
# source of truth for "what did the user last generate for this
# container?" — no PDF-parsing, no ambiguity.
#
# Format:
#   generated/{container_key}__{iso_timestamp}.json
# where container_key = "_".join(sorted(container_nos_upper)).
# Sorting means multi-container invoices resolve to the same file
# regardless of the order they were originally entered.
#
# No cleanup/rotation today — sub-100 invoices/month × ~5KB each is
# negligible even after 10 years. If this ever grows unexpectedly, add
# a nightly cron to prune files older than N months; but the on-request
# path stays simple.

GENERATED_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated")
os.makedirs(GENERATED_DIR, exist_ok=True)


def _container_key(container_nos):
    """Canonical filesystem key for a container list. Sorted so search
    order doesn't matter — searching 'B, A' finds saved edits for 'A, B'."""
    return "_".join(sorted(c.upper().strip() for c in container_nos if c and c.strip()))


def save_generation_state(container_nos, body):
    """Persist the exact POST body that produced a PDF, plus a timestamp.

    Called AFTER successful PDF generation. Non-fatal: if the save fails
    (disk full, permissions), the download still returns to the user —
    we just log and move on. Losing history is annoying; losing the
    download would be a regression.
    """
    try:
        key = _container_key(container_nos)
        if not key:
            return
        ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
        path = os.path.join(GENERATED_DIR, f"{key}__{ts}.json")
        record = {
            "container_key": key,
            "container_nos": sorted(c.upper().strip() for c in container_nos),
            "saved_at_utc":  ts,
            "saved_at_iso":  datetime.utcnow().isoformat() + "Z",
            "body":          body,
        }
        with open(path, "w") as f:
            json.dump(record, f, indent=2)
        print(f"💾 Saved generation state: {path}")
    except Exception as e:
        # Don't propagate — download must not fail because we couldn't
        # persist state.
        print(f"⚠️  Could not save generation state: {e}")


def load_latest_generation(container_nos):
    """Return the most-recent saved state for these containers, or None.

    Used by /api/preview to tell the UI whether to show the "Load previous
    edits" banner. Returns a small metadata dict; the full body is loaded
    on demand by /api/load-saved to keep the preview response small.
    """
    try:
        key = _container_key(container_nos)
        if not key:
            return None
        pattern = os.path.join(GENERATED_DIR, f"{key}__*.json")
        matches = sorted(glob.glob(pattern), reverse=True)   # newest first
        if not matches:
            return None
        latest_path = matches[0]
        with open(latest_path) as f:
            record = json.load(f)
        # Return metadata only — the actual body ships via /api/load-saved.
        # The path name (not the record contents) is the load token — that
        # way even if the file is corrupt, listing still works.
        return {
            "container_key":  record.get("container_key", key),
            "saved_at_iso":   record.get("saved_at_iso", ""),
            "total_versions": len(matches),
            "load_token":     os.path.basename(latest_path),   # filename-only, not full path
        }
    except Exception as e:
        print(f"⚠️  Could not load saved state: {e}")
        return None


def load_saved_by_token(load_token):
    """Fetch the full body for a specific saved generation.

    load_token is the raw filename returned by load_latest_generation —
    NOT a user-supplied path. We defend against directory traversal by
    rejecting anything with a path separator or leading dot.
    """
    if not load_token or "/" in load_token or "\\" in load_token or load_token.startswith("."):
        return None
    if not load_token.endswith(".json"):
        return None
    path = os.path.join(GENERATED_DIR, load_token)
    if not os.path.isfile(path):
        return None
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return None



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
                # Country of Origin has no source column in the packing sheet
                # (or the main sheet) — Edge Metals is a US exporter so the
                # PDF-side default is already "USA". Sending "USA" here lets
                # the UI prefill it as an editable field, so the user only
                # needs to touch it in the rare case a shipment originates
                # elsewhere. Backend still treats blank as "use the default"
                # — an empty POST value doesn't override anything.
                "Country of Origin":  "USA",
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
        # Tell the UI whether previously-generated edits exist for these
        # containers. Small metadata only — actual body loads via /api/load-saved
        # when the user clicks the banner. Absent key = no history (UI stays quiet).
        "saved_overrides": load_latest_generation(container_nos),
    })


@app.route("/api/load-saved/<load_token>")
def api_load_saved(load_token):
    """Return the full saved POST body for a specific past generation.

    load_token comes from the saved_overrides.load_token that /api/preview
    returned — an opaque-ish filename the client just echoes back. Path
    traversal is defended against in load_saved_by_token (no slashes, no
    leading dots, must end in .json, must exist inside GENERATED_DIR).
    """
    record = load_saved_by_token(load_token)
    if not record:
        return jsonify({"error": "Saved state not found."}), 404
    return jsonify(record)


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
    country_override = body.get("country_of_origin", "").strip()
    total_override   = body.get("total_override",    "").strip()
    consignee_override = body.get("consignee_override", "").strip()
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
        # Country of Origin: skip-when-empty (not __CLEAR__) — empty means
        # "use the PDF's default of USA", not "force blank". Matches the same
        # pattern as port_loading/discharge above rather than the inv_no
        # forced-clear pattern; a blank country row on the PDF would be a
        # customs red flag, so we specifically do NOT allow it to be blanked.
        if country_override:  cmd += ["--country-of-origin", country_override]
        # Total override: skip-when-empty too. When set, invoice_gen.py
        # prints the value silently (no annotation on the PDF). Empty here
        # = normal computed total.
        if total_override:    cmd += ["--total-override",    total_override]
        # Consignee override — only forwarded when the UI marked it dirty
        # (see dataset.original comparison in generate.html). Empty means
        # "use the Google Doc canonical name" (existing behavior — a blank
        # override MUST NOT print an empty buyer name, hence skip-when-empty).
        if consignee_override: cmd += ["--consignee-override", consignee_override]

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
            # Persist AFTER we know generation succeeded — no point storing
            # state for a failed run. Save happens before returning the file
            # so a network hiccup between save and download doesn't leave us
            # with a downloaded PDF and no matching saved state (which would
            # break the whole "reload previous edits" flow).
            save_generation_state(container_nos, body)
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
            # Same save point as the ZIP branch above — kept in sync so
            # combined and separate downloads produce identical history.
            save_generation_state(container_nos, body)
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
    country_override = body.get("country_of_origin", "").strip()
    total_override   = body.get("total_override",    "").strip()
    consignee_override = body.get("consignee_override", "").strip()
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
        # Kept in sync with api_generate — inline preview and final download
        # must produce identical PDFs; any override skipped here would show
        # the user one PDF and hand them a different one on Download.
        if country_override:  cmd += ["--country-of-origin", country_override]
        if total_override:    cmd += ["--total-override",    total_override]
        if consignee_override: cmd += ["--consignee-override", consignee_override]

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
