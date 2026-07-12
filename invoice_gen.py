"""
Edge Metals Invoice Generator - Simplified Version
Reads directly from Google Sheets using CSV export (no API needed)

Requirements:
1. Your Google Sheet must be set to "Anyone with the link can view"
2. Install: pip install reportlab requests

Usage:
    python invoice_gen.py --all
    python invoice_gen.py --row 5
"""

import argparse
import os
import sys
from datetime import datetime
from typing import List, Dict, Any
import io

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    import requests
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Install with: pip install reportlab requests")
    sys.exit(1)


# ═══════════════════════════════════════════════
#  CONFIGURATION - CHANGE THESE
# ═══════════════════════════════════════════════

GOOGLE_SHEET_ID = "1QsCeuqeRKODuouzO2PfKbxG9qJpN8yAbIurSzhI--6s"

MAIN_SHEET_GID    = "571096144"
PACKING_SHEET_GID = "1340048377"

SHEET_NAME = "Sheet1"

ADDRESS_DOC_ID = "1u-hKBqVvqS1GIpckUXWbT5AQtTGWjWK69rre5IlSHio"

# Signature file — place signature.png in the same folder as this script
SIGNATURE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "signature.png")

COLUMNS = {
    "consignee":        0,
    "inv_no":           1,
    "inv_date":         2,
    "hbl_no":           3,
    "booking_no":       4,
    "container_no":     5,
    "seal_no":          6,
    "supplier":         7,
    "terms":            8,
    "customer":         9,
    "proforma_date":   10,
    "reference":       11,
    "item_desc":       12,
    "weight":          13,
    "inv_price":       14,
    "commissions":     15,
    "invoice_amt":     16,
    "received_amt":    17,
    "received_date":   18,
    "freight_etd":     20,
    "freight_charge":  21,
    "eta":             22,
    # NOTE: port_discharge is intentionally NOT mapped to a column here.
    # It never had a real source column in the main sheet — column 23 held
    # unrelated/stray values (sometimes "0" or "0.0") that were incorrectly
    # taking priority over the correctly-parsed packing-sheet carrier split
    # ("HOUSTON/BUSAN" -> loading/discharge). port_loading and place_of_receipt
    # were never mapped to a main-sheet column either — port_discharge should
    # behave the same way: packing-sheet-derived by default, explicit
    # CLI/UI override only when the person actually types one in.
    "efs":             25,
}

PACKING_COLUMNS = {
    "carrier":              0,
    "trucking":             2,
    "date":                 3,
    "supplier":             4,
    "customer":             5,
    "inv_no":               6,
    "booking_no":           7,
    "container_no":         8,
    "seal_no":              9,
    "item_desc":            10,
    "gross_weight_lbs":    11,
    "truck_lbs":           12,
    "container_tare_lbs":  13,
    "chassis_lbs":         14,
    "boxes_weight_lbs":    15,
    "total_lbs":           16,
    "net_weight_lbs":      17,
    "net_weight_mt":       18,
    "price":               19,
    "suppl_invoice_amt":   20,
    "invoice_nbr":         21,
    "pier_pass":           22,
    "advance_trucking":    23,
    "balance":             24,
    "loan_deduc":          25,
    "balance2":            26,
    "paid_by_edge":        27,
    "fas":                 28,
    "total_invoice_price": 29,
    "buyer_selling_amt":   30,
    "freight":             31,
    "edge_net":            32,
    "wire_charge":         33,
    "commissions":         34,
    "edge_net2":           35,
    "photo_link":          36,
}

EXPORTER = {
    "name":    "EDGE METALS INC",
    "address": "14750 DEVONSHIRE LN, FRISCO, TX 75035",
    "tel":     "TEL: (310) 938-2525",
    "fax":     "FAX: (425) 940-9408",
}

PORT_OF_LOADING = "LOS ANGELES"
COUNTRY_OF_ORIGIN = "USA"

ITEM_CODE_MAP = {
    "AL": "ALUMINIUM COMBO",
    "AP": "SCRAP AUTO PARTS",
    "RC": "REGULAR COMBO",
    "BT": "BATTERY",
    "AW": "ALUMINIUM WHEELS",
    "CW": "CHROME WHEELS",
    "TT": "TAINT TABOUR",
    "AC": "AUTO CAST",
    "ML": "MIXED LOAD",
    "SU": "SEALED UNITS",
    "MM": "MIXED MOTORS",
    "RD": "ROTORS AND DRUMS",
    "MC": "MIXED COMBO",
}


def get_item_label(inv_no: str) -> str:
    """
    Extract the 2-letter code from the invoice number and return
    'CODE-DESCRIPTION', e.g. '260416_ML_26MK45' -> 'ML-MIXED LOAD'.
    Falls back to empty string if no match.
    """
    import re
    # Find all 2-letter uppercase segments in the invoice number
    parts = re.findall(r'[A-Z]{2}', inv_no.upper())
    for part in parts:
        if part in ITEM_CODE_MAP:
            return f"{ITEM_CODE_MAP[part]}"
    return ""


# ═══════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════

def safe_str(val) -> str:
    if val is None or val == "":
        return ""
    if isinstance(val, datetime):
        return val.strftime("%m/%d/%Y")
    return str(val).strip()


def safe_float(val) -> float:
    try:
        return float(str(val).replace(",", "").replace("$", "").strip())
    except Exception:
        return 0.0


def is_blank_or_zero(val) -> bool:
    """True for '', whitespace, or a numeric-zero placeholder like '0' / '0.0' / '0.00'.

    Every port_loading/port_discharge/place_of_receipt fallback chain in this
    file uses `value or fallback`. Python truthy-`or` treats "0.0" as a valid
    non-empty string, not as blank — so when a sheet formula defaults to a
    literal zero instead of a true empty cell, it silently wins the `or` and
    renders as a port name ("0.0") instead of falling through to the real
    default. This is the fix for that: treat numeric-zero strings as blank too.
    """
    s = safe_str(val).strip()
    if not s:
        return True
    try:
        return float(s) == 0.0
    except (ValueError, TypeError):
        return False


def first_meaningful(*vals) -> str:
    """Return the first value that isn't blank or a numeric-zero placeholder.

    Replaces chained `a or b or c` for port fields — same intent, but treats
    "0"/"0.0" as blank instead of as a valid answer. Last argument is the
    hard fallback and is always returned as-is if nothing earlier qualifies.
    """
    for v in vals[:-1]:
        s = safe_str(v).strip()
        if not is_blank_or_zero(s):
            return s
    return safe_str(vals[-1]) if vals else ""


def format_date(val) -> str:
    if isinstance(val, datetime):
        return val.strftime("%m/%d/%Y")
    if val:
        s = str(val).strip()
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]:
            try:
                dt = datetime.strptime(s, fmt)
                return dt.strftime("%m/%d/%Y")
            except:
                continue
        return s
    return ""


def format_rate(value) -> str:
    """Format a per-unit rate preserving exactly the precision that was entered.

    0.42 -> $0.42, 0.4256 -> $0.4256, 0.425 -> $0.425 — no fixed 2 or 3 decimal
    cap that would round off real digits the person typed. Only exception:
    a currency floor of 2 decimals, so whole numbers still read as $1.00, not $1.
    Uses Decimal(str(value)) rather than Decimal(value) directly, since the
    latter exposes binary-float noise (0.425 -> 0.42499999999999998...) that
    Python's own str() repr already avoids.
    """
    from decimal import Decimal
    d = Decimal(str(value)).normalize()
    if d.as_tuple().exponent > 0:
        d = d.quantize(Decimal(1))          # flatten exponent notation, e.g. 1E+2 -> 100
    if d.as_tuple().exponent > -2:
        d = d.quantize(Decimal("0.01"))     # enforce currency floor of 2 decimals
    return f"${d:,f}"


def wrap_text_to_width(text: str, max_width: float, font: str, font_size: float) -> list:
    """Word-wrap text to fit max_width. Returns list of lines (min 1).

    Used for the Description column in every line-items table (invoice, invoice-only,
    packing list). ReportLab's drawCentredString has no wrapping — a long
    description would silently overflow the cell and collide with neighboring
    columns. This pre-splits at word boundaries; the caller then grows row height
    to fit len(lines) * line_height. A single word longer than max_width is
    kept on its own line rather than character-split, since commodity descriptions
    are always real words and mid-word breaks would look broken.
    """
    from reportlab.pdfbase.pdfmetrics import stringWidth
    text = safe_str(text)
    if not text:
        return [""]
    words = text.split()
    if not words:
        return [""]
    lines, current = [], words[0]
    for word in words[1:]:
        candidate = current + " " + word
        if stringWidth(candidate, font, font_size) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def eval_freight(expr: str) -> float:
    import re
    expr = expr.replace(",", "").replace("$", "").strip()
    if re.fullmatch(r'[\d\.\+\-\*/\s]+', expr):
        try:
            return float(eval(expr))
        except Exception:
            pass
    try:
        return float(expr)
    except Exception:
        return 0.0


def detect_template(data: Dict[str, Any]) -> str:
    consignee = safe_str(data.get("consignee", "")).upper()
    terms = safe_str(data.get("terms", "")).upper()
    if "ESWARI" in consignee or "INDIA" in terms or "MANGALORE" in terms:
        return "eswari"
    else:
        return "mk_trading"


# ═══════════════════════════════════════════════
#  COMMODITY CODE → DESCRIPTION MAPPING
# ═══════════════════════════════════════════════

COMMODITY_CODES = {
    "AL": "ALUMINIUM COMBO",
    "AP": "SCRAP AUTO PARTS",
    "RC": "REGULAR COMBO",
    "BT": "BATTERY",
    "AW": "ALUMINIUM WHEELS",
    "CW": "CHROME WHEELS",
    "TT": "TAINT TABOUR",
    "AC": "AUTO CAST",
    "ML": "MIXED LOAD",
    "SU": "SEALED UNITS",
    "MM": "MIXED MOTORS",
    "RD": "ROTORS AND DRUMS",
    "MC": "MIXED COMBO",
}


def resolve_item_desc(item_desc: str, inv_no: str) -> str:
    """
    If item_desc is blank, extract commodity code from invoice number and map it.
    Invoice format: YYMMDD_XX_CONTAINERCODE  e.g. 260416_ML_26MK45
    The code is the middle segment (ML in this example).
    If item_desc is already populated, strip any leading XX- code prefix.
    """
    if item_desc and item_desc.strip():
        desc = item_desc.strip()
        # Strip leading code prefix like "AC-" or "ML-"
        import re
        desc = re.sub(r'^[A-Z]{2}-', '', desc).strip()
        return desc
    # Try to extract code from invoice number
    parts = safe_str(inv_no).split("_")
    for part in parts:
        code = part.strip().upper()
        if code in COMMODITY_CODES:
            return COMMODITY_CODES[code]
    return item_desc or "SCRAP METALS"


# ═══════════════════════════════════════════════
#  GOOGLE SHEETS CSV READERS
# ═══════════════════════════════════════════════

def read_google_sheet_csv(sheet_id: str, sheet_name: str = "Sheet1") -> tuple:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={MAIN_SHEET_GID}"
    print(f"Fetching data from Google Sheets...")
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"\n❌ ERROR: Could not read Google Sheet\nError: {e}")
        sys.exit(1)

    import csv
    content = response.content.decode('utf-8')
    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    if not rows:
        print("❌ ERROR: Sheet is empty")
        sys.exit(1)

    headers = rows[0]
    data_rows = [row for row in rows[1:] if any(cell.strip() for cell in row)]
    print(f"✓ Loaded {len(data_rows)} rows from main sheet")
    return headers, data_rows


def read_packing_lookup_sheet(sheet_id: str) -> Dict[str, Dict[str, Any]]:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={PACKING_SHEET_GID}"
    print(f"Fetching packing data from lookup sheet (GID: {PACKING_SHEET_GID})...")
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"⚠️  WARNING: Could not read packing lookup sheet: {e}")
        return {}

    import csv
    content = response.content.decode('utf-8')
    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    if not rows:
        return {}

    data_rows = rows[1:] if len(rows) > 1 else []
    lookup = {}
    for row in data_rows:
        if not any(cell.strip() for cell in row):
            continue
        container_col = PACKING_COLUMNS["container_no"]
        if container_col >= len(row):
            continue
        container_key = safe_str(row[container_col]).upper().strip()
        if not container_key:
            continue

        packing = {}
        for field, idx in PACKING_COLUMNS.items():
            packing[field] = row[idx] if idx < len(row) else ""

        carrier_raw = safe_str(packing.get("carrier", ""))
        if "/" in carrier_raw:
            parts = [p.strip() for p in carrier_raw.split("/")]
            if len(parts) >= 3:
                packing["place_of_receipt"] = parts[0]
                packing["port_loading"]     = parts[1]
                packing["port_discharge"]   = parts[2]
            else:
                packing["place_of_receipt"] = ""
                packing["port_loading"]     = parts[0]
                packing["port_discharge"]   = parts[1]
        else:
            packing["place_of_receipt"] = ""
            packing["port_loading"]     = carrier_raw
            packing["port_discharge"]   = ""

        # Store as list — multiple rows per container (one per item)
        if container_key not in lookup:
            lookup[container_key] = []
        lookup[container_key].append(packing)

    print(f"✓ Loaded packing data for {len(lookup)} containers")
    return lookup


def read_address_lookup(doc_id: str) -> Dict[str, list]:
    url = f"https://docs.google.com/document/d/{doc_id}/export?format=txt"
    print(f"Fetching buyer addresses from Google Doc...")
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"⚠️  WARNING: Could not read address Google Doc: {e}")
        return {}

    text = response.content.decode("utf-8")
    lookup = {}
    current_key = None
    current_lines = []

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if line.startswith("[") and line.endswith("]"):
            if current_key is not None:
                lookup[current_key] = [l for l in current_lines if l]
            current_key = line[1:-1].upper().strip()
            current_lines = []
        elif current_key is not None:
            current_lines.append(line)

    if current_key is not None:
        lookup[current_key] = [l for l in current_lines if l]

    print(f"✓ Loaded addresses for {len(lookup)} buyers: {list(lookup.keys())}")
    return lookup


def get_buyer_address(consignee: str, address_lookup: Dict[str, list]) -> list:
    key = consignee.upper().strip()
    if key in address_lookup:
        return address_lookup[key]
    for doc_key, lines in address_lookup.items():
        if doc_key in key or key in doc_key:
            return lines
    return []


# ═══════════════════════════════════════════════
#  MULTI-ROW GROUPING  ← NEW
# ═══════════════════════════════════════════════

def group_rows_by_container(rows: list) -> Dict[str, list]:
    """
    Group all sheet rows by container number.
    Returns { "CONTAINER_NO": [row1, row2, ...], ... }
    Preserves sheet order within each group.
    """
    groups: Dict[str, list] = {}
    container_col = COLUMNS["container_no"]
    for row in rows:
        key = safe_str(
            row[container_col] if container_col < len(row) else ""
        ).upper().strip()
        if not key:
            continue
        groups.setdefault(key, []).append(row)
    return groups


def rows_to_invoice_data(container_rows: list) -> Dict[str, Any]:
    """
    Merge multiple sheet rows into one invoice dict.

    - Header fields (consignee, inv_no, dates, terms …) come from the FIRST
      row — these are invoice-level fields, same for the whole document.
    - booking_no/container_no/seal_no are CONTAINER-level, not invoice-level:
      each line item carries its OWN row's booking/container/seal, so
      combining rows from multiple containers doesn't silently stamp every
      row with only the first container's identifiers. The top-level
      data["booking_no"/"container_no"/"seal_no"] (from the first row) is
      kept too, purely as a fallback for old code paths and for the
      single-container case where every row shares the same identifiers anyway.
    - line_items is a list of dicts, one entry per sheet row.
    """
    first = row_to_dict(container_rows[0])

    line_items = []
    for row in container_rows:
        d         = row_to_dict(row)
        weight    = safe_float(d.get("weight",    0))
        inv_price = safe_float(d.get("inv_price", 0))
        inv_amt   = (weight * inv_price) if (weight and inv_price) \
                    else safe_float(d.get("invoice_amt", 0))
        line_items.append({
            "item_desc":      resolve_item_desc(safe_str(d.get("item_desc", "")), safe_str(first.get("inv_no", ""))),
            "weight":         weight,
            "inv_price":      inv_price,
            "invoice_amt":    inv_amt,
            "freight_charge": safe_str(d.get("freight_charge", "")),
            "efs":            safe_str(d.get("efs",            "")),
            "container_no":   safe_str(d.get("container_no", "")),
            "booking_no":     safe_str(d.get("booking_no",   "")),
            "seal_no":        safe_str(d.get("seal_no",      "")),
        })

    # Collect freight/efs — scan ALL rows, take first non-zero value
    # Also check the first row dict directly (covers single-row containers)
    freight_val = 0.0
    efs_val     = 0.0
    for item in line_items:
        if not freight_val:
            fv = eval_freight(item["freight_charge"]) if item["freight_charge"] else 0.0
            if fv > 0:
                freight_val = fv
        if not efs_val:
            ev = eval_freight(item.get("efs", "")) if item.get("efs") else 0.0
            if ev > 0:
                efs_val = ev
    # Fallback: read directly from first row (handles formula results etc.)
    if not freight_val:
        freight_val = eval_freight(safe_str(first.get("freight_charge", "")))
    if not efs_val:
        efs_val = eval_freight(safe_str(first.get("efs", "")))
    print(f"  💰 freight collected: {freight_val}, efs: {efs_val}")
    print(f"  💰 line item freight values: {[item['freight_charge'] for item in line_items]}")
    first["freight_charge"] = str(freight_val) if freight_val else ""
    first["efs"]            = str(efs_val)     if efs_val     else ""
    first["line_items"] = line_items
    return first


def row_to_dict(row: list) -> Dict[str, Any]:
    d = {}
    for key, idx in COLUMNS.items():
        d[key] = row[idx] if idx < len(row) else ""
    return d


# ═══════════════════════════════════════════════
#  LINE-ITEMS TABLE HELPERS
# ═══════════════════════════════════════════════
#
# Constants for wrap-aware row rendering. Extracted so all three renderers
# (draw_mk_trading_invoice / draw_invoice_only / draw_packing_list_only)
# share the same visual rhythm — if you tune LINE_H or MIN_ROW_H in one
# place, packing list and invoice stay in sync.
DESC_LINE_H  = 3.2 * mm   # vertical gap between wrapped description lines
DESC_MIN_ROW = 8   * mm   # single-line row height (matches pre-wrap default)
DESC_PAD     = 1.5 * mm   # horizontal padding inside description cell


def draw_line_item_row(c_obj, row_vals, col_widths, desc_col_idx, margin, y, base_row_h=DESC_MIN_ROW):
    """Draw one line-item row with description-column word-wrap.

    Row height grows dynamically to fit however many wrapped description
    lines are needed; every other column centers vertically within the
    resulting box, so short columns don't look stranded at the top when
    the description takes 3+ lines. Returns the new `y` after the row.

    Deliberately NOT integrated with the packing list at present — packing
    descriptions are short by convention (mapped commodity names) so paying
    the wrap cost there gives near-zero visual benefit. Kept generic in case
    that changes later.
    """
    desc_text = row_vals[desc_col_idx]
    desc_w    = col_widths[desc_col_idx] - DESC_PAD * 2

    desc_lines = wrap_text_to_width(desc_text, desc_w, "Helvetica", 7)
    # Row height = padding + N text lines + padding; floor at MIN so single-line
    # rows stay identical to the pre-wrap layout (no visual regression).
    row_h = max(base_row_h, len(desc_lines) * DESC_LINE_H + 3 * mm)

    x = margin
    c_obj.setFillColor(colors.black)
    for i, (val, w) in enumerate(zip(row_vals, col_widths)):
        c_obj.rect(x, y - row_h, w, row_h, fill=0, stroke=1)
        c_obj.setFont("Helvetica", 7)
        if i == desc_col_idx:
            # Vertically center the multi-line block: compute total block height,
            # then start the FIRST line at (row_center + block_height/2 - one_line).
            total_text_h = len(desc_lines) * DESC_LINE_H
            first_y = y - (row_h - total_text_h) / 2 - DESC_LINE_H + 1 * mm
            for j, line in enumerate(desc_lines):
                c_obj.drawCentredString(x + w / 2, first_y - j * DESC_LINE_H, line)
        else:
            c_obj.drawCentredString(x + w / 2, y - row_h / 2 - 1 * mm, val)
        x += w
    return y - row_h


# ═══════════════════════════════════════════════
#  PDF GENERATION - MK TRADING TEMPLATE
# ═══════════════════════════════════════════════


def find_packing_row(packing_rows, item_desc: str) -> dict:
    """Find the packing row matching item_desc. Falls back to first row if no match."""
    if not packing_rows:
        return {}
    item_key = item_desc.strip().upper()
    for row in packing_rows:
        row_desc = safe_str(row.get("item_desc", "")).strip().upper()
        if row_desc and (row_desc in item_key or item_key in row_desc):
            return row
    return packing_rows[0]   # fallback to first row


def draw_note_rows(c_obj, data, margin, y, col_widths, row_h):
    """Render user-added note rows below the line items.

    Cells are merged (S.No→Rate) because note text is free-form and usually
    longer than any single column. A light tint fill signals that the merge
    is intentional — every other row in this table has full column dividers,
    so an unshaded merged row reads as a broken grid, not a styled one.
    Regular weight + centered to match every other data row; only the label
    text (not weight) differs, since notes are full sentences, not captions.
    Returns (new_y, notes_sum) — notes_sum folds into the invoice TOTAL.
    """
    note_rows = data.get("note_rows") or []
    notes_sum = 0.0
    if not note_rows:
        return y, notes_sum

    text_w = sum(col_widths[:-1])
    amt_w  = col_widths[-1]
    TINT   = colors.HexColor("#F4F6F8")   # neutral tint — distinct from header/total blue, packing orange
    from reportlab.pdfbase.pdfmetrics import stringWidth

    for note in note_rows:
        txt = safe_str(note.get("text", ""))
        amt = note.get("amount", None)

        c_obj.setFillColor(TINT)
        c_obj.rect(margin, y - row_h, text_w, row_h, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        fs = 7
        while fs > 5 and stringWidth(txt, "Helvetica", fs) > text_w - 6 * mm:
            fs -= 0.5
        c_obj.setFont("Helvetica", fs)
        c_obj.drawCentredString(margin + text_w / 2, y - row_h / 2 - 1 * mm, txt)

        c_obj.setFillColor(TINT)
        c_obj.rect(margin + text_w, y - row_h, amt_w, row_h, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        if amt is not None and amt != 0:
            notes_sum += amt
            amt_str = f"-${abs(amt):,.2f}" if amt < 0 else f"${amt:,.2f}"
            c_obj.setFont("Helvetica", 7)
            c_obj.drawCentredString(margin + text_w + amt_w / 2, y - row_h / 2 - 1 * mm, amt_str)
        y -= row_h

    return y, notes_sum


def draw_summary_rows(c_obj, margin, y, col_widths, row_h, subtotal, notes_sum, freight, efs, LIGHT_BLUE, total_override=None):
    """Draw FREIGHT DEDUCTION / EFS / TOTAL rows below the line items (and notes, if any).

    Merged into one label-cell + one amount-cell per row — same grid convention as
    draw_note_rows() — instead of 8 separate boxes with 5 left empty. The old
    per-column version left 4-5 blank bordered cells per row, which read as a
    broken/fragmented grid once a merged note row sat directly above it.

    total_override: when a non-None number, the printed TOTAL uses that value
    instead of the computed (subtotal + notes − freight − efs). The override
    is printed silently — no visible marker on the PDF — so the printed
    invoice matches exactly what the user typed in the UI. This is a
    deliberate product decision: the UI's amber tint on the Final Amount
    field is the sole audit signal, and it does not survive PDF generation.
    Passing None (default) preserves the pre-existing computed behavior
    exactly — no behavior change for callers that don't opt in.

    Returns (new_y, final_amt) where final_amt is the value actually printed.
    """
    computed = subtotal + notes_sum - freight - efs

    def _row(label, amount_str, tint=None, bold_amt=False):
        nonlocal y
        text_w = sum(col_widths[:-1])
        amt_w  = col_widths[-1]
        fill_color = tint if tint else colors.white
        c_obj.setFillColor(fill_color)
        c_obj.rect(margin, y - row_h, text_w, row_h, fill=1, stroke=1)
        c_obj.rect(margin + text_w, y - row_h, amt_w, row_h, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 8 if bold_amt else 7)
        c_obj.drawCentredString(margin + text_w / 2, y - row_h / 2 - 1 * mm, label)
        c_obj.drawCentredString(margin + text_w + amt_w / 2, y - row_h / 2 - 1 * mm, amount_str)
        y -= row_h

    if freight > 0:
        _row("FREIGHT DEDUCTION", f"-${freight:,.2f}")
    if efs > 0:
        _row("EFS", f"-${efs:,.2f}")

    final_amt = float(total_override) if total_override is not None else computed
    _row("TOTAL", f"${final_amt:,.2f}", tint=LIGHT_BLUE, bold_amt=True)

    return y, final_amt


def draw_mk_trading_invoice(c_obj, data, page_width, page_height, packing_lookup=None, address_lookup=None):
    """Draw MK Trading LLC invoice — supports multiple line items per container."""

    BLUE       = colors.HexColor("#2E6DA4")
    LIGHT_BLUE = colors.HexColor("#D9E8F5")

    margin = 15 * mm
    y      = page_height - margin
    width  = page_width - 2 * margin

    # ── HEADER BAR ────────────────────────────────────────────────────────────
    header_h = 10 * mm
    c_obj.setFillColor(BLUE)
    c_obj.rect(margin, y - header_h, width, header_h, fill=1, stroke=0)
    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.setFont("Helvetica-Bold", 14)
    c_obj.drawCentredString(page_width / 2, y - header_h + 3 * mm, "Commercial Invoice")
    y -= header_h

    # ── ROW 1: Exporter | Invoice No / Date / Reference ───────────────────────
    lbl_h     = 5 * mm
    top_val   = 10 * mm
    ref_val_h = 8 * mm
    right_total = lbl_h + top_val + lbl_h + ref_val_h   # 28 mm
    exp_val_h   = right_total - lbl_h                    # 23 mm

    left_x  = margin
    left_w  = width / 2
    right_x = margin + width / 2
    right_w = width / 2
    inv_w   = right_w * 0.65
    date_w  = right_w * 0.35
    date_x  = right_x + inv_w

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(left_x, y - lbl_h, left_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h + 1.5 * mm, "Exporter")
    c_obj.rect(left_x, y - lbl_h - exp_val_h, left_w, exp_val_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 5  * mm, EXPORTER["name"])
    c_obj.setFont("Helvetica", 7)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 10 * mm, EXPORTER["address"])
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 15 * mm, EXPORTER["tel"])
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 20 * mm, EXPORTER["fax"])

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, y - lbl_h, right_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - lbl_h + 1.5 * mm, "Invoice No & Items")
    c_obj.drawString(date_x  + 2 * mm, y - lbl_h + 1.5 * mm, "DATE")

    c_obj.rect(right_x, y - lbl_h - top_val, inv_w, top_val, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - lbl_h - 4 * mm, safe_str(data.get("inv_no", "")))
    _item_label = get_item_label(safe_str(data.get("inv_no", "")))
    if _item_label:
        c_obj.setFont("Helvetica", 6.5)
        c_obj.drawString(right_x + 2 * mm, y - lbl_h - 8 * mm, _item_label)

    c_obj.rect(date_x, y - lbl_h - top_val, date_w, top_val, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(date_x + 2 * mm,
                     y - lbl_h - top_val + (top_val - 2.5 * mm) / 2,
                     format_date(data.get("inv_date", "")))

    ref_top = y - lbl_h - top_val
    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ref_top - lbl_h, right_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h + 1.5 * mm, "Other Reference(s)")
    c_obj.rect(right_x, ref_top - lbl_h - ref_val_h, right_w, ref_val_h, fill=0, stroke=1)
    ref_str   = safe_str(data.get("reference", ""))
    proforma  = format_date(data.get("proforma_date", ""))
    other_ref = " | ".join(filter(None, [proforma, ref_str]))
    c_obj.setFont("Helvetica", 7)
    # Word-wrap other_ref if it exceeds cell width
    from reportlab.pdfbase.pdfmetrics import stringWidth
    max_w = right_w - 4 * mm
    font_size = 7
    if stringWidth(other_ref, "Helvetica", font_size) <= max_w:
        c_obj.drawString(right_x + 2 * mm,
                         ref_top - lbl_h - ref_val_h + (ref_val_h - 2.5 * mm) / 2,
                         other_ref)
    else:
        # Split into two lines at the midpoint word boundary
        words = other_ref.split()
        line1, line2 = "", ""
        for i, word in enumerate(words):
            test = " ".join(words[:i+1])
            if stringWidth(test, "Helvetica", font_size) <= max_w:
                line1 = test
            else:
                line2 = " ".join(words[i:])
                break
        c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h - ref_val_h + 5.5 * mm, line1)
        if line2:
            c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h - ref_val_h + 2.0 * mm, line2)

    y -= right_total

    # ── ROW 2: Buyer | Terms / Vessel / Comments ──────────────────────────────
    buyer_lbl_h = 5 * mm
    cell_lbl_h  = 5 * mm
    cell_val_h  = 6 * mm
    buyer_val_h = cell_val_h + (cell_lbl_h + cell_val_h) * 2   # 28 mm

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(left_x, y - buyer_lbl_h, left_w, buyer_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - buyer_lbl_h + 1.5 * mm, "Buyer")

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, y - buyer_lbl_h, right_w, buyer_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - buyer_lbl_h + 1.5 * mm, "Terms")

    y -= buyer_lbl_h

    c_obj.rect(left_x, y - buyer_val_h, left_w, buyer_val_h, fill=0, stroke=1)
    consignee_name = safe_str(data.get("consignee", ""))
    addr_lines = get_buyer_address(consignee_name, address_lookup or {})
    # Consignee override: when the user edits the Company field in the UI,
    # that value prints as the buyer name on the PDF. Address lines below
    # are STILL looked up from the Google Doc using consignee_name — the
    # override only replaces the display name on line 1, not the address
    # data source. When unset, we keep the historical behavior of using
    # the Doc's first line (which is the "canonical" formal name for that
    # buyer) — that way years of previously-generated invoices stay
    # visually consistent for any invoice generated without touching the
    # field. Same skip-when-empty pattern as country_of_origin.
    consignee_override = safe_str(data.get("consignee_override", "")).strip()
    if addr_lines:
        c_obj.setFont("Helvetica-Bold", 9)
        display_name = consignee_override if consignee_override else addr_lines[0].strip()
        c_obj.drawString(left_x + 2 * mm, y - 4 * mm, display_name)
        c_obj.setFont("Helvetica", 7)
        for i, line in enumerate(addr_lines[1:7]):
            c_obj.drawString(left_x + 2 * mm, y - (8 + i * 3.5) * mm, line.strip())

    ry = y
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica", 7)
    terms_val = safe_str(data.get("terms", ""))
    if terms_val:
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, terms_val)
    ry -= cell_val_h

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ry - cell_lbl_h, right_w, cell_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ry - cell_lbl_h + 1.5 * mm, "Vessel / Flight No")
    ry -= cell_lbl_h
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    vessel_val = safe_str(data.get("vessel", ""))
    if vessel_val:
        c_obj.setFont("Helvetica", 7)
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, vessel_val)
    ry -= cell_val_h

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ry - cell_lbl_h, right_w, cell_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ry - cell_lbl_h + 1.5 * mm, "Payment Terms")
    ry -= cell_lbl_h
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    payment_terms_val = safe_str(data.get("payment_terms", ""))
    if payment_terms_val:
        c_obj.setFont("Helvetica", 7)
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, payment_terms_val)

    y -= buyer_val_h

    # ── ROW 3: 4 port cells ───────────────────────────────────────────────────
    four_lbl_h = 5 * mm
    four_val_h = 7 * mm
    cw = width / 4

    container_key_ports = safe_str(data.get("container_no", "")).upper().strip()
    _packing_rows_ports = (packing_lookup or {}).get(container_key_ports, [])
    packing_ports = _packing_rows_ports[0] if _packing_rows_ports else {}

    PORT_ALIASES = {
        "LA": "LOS ANGELES", "LAX": "LOS ANGELES", "LB": "LONG BEACH",
        "NY": "NEW YORK", "CHI": "CHICAGO", "HOU": "HOUSTON", "SAV": "SAVANNAH",
    }
    def expand_port(val: str) -> str:
        v = val.strip().upper()
        return PORT_ALIASES.get(v, val.strip())

    place_of_receipt = first_meaningful(data.get("place_of_receipt", ""),
                                        expand_port(safe_str(packing_ports.get("place_of_receipt", ""))),
                                        "").upper()
    port_loading     = first_meaningful(data.get("port_loading", ""),
                                        expand_port(safe_str(packing_ports.get("port_loading", ""))),
                                        PORT_OF_LOADING).upper()
    port_discharge   = first_meaningful(data.get("port_discharge", ""),
                                        packing_ports.get("port_discharge", ""),
                                        "TO BE ADVISED").upper()

    four_cells = [
        ("Country of Origin",           safe_str(data.get("country_of_origin", "")).upper() or COUNTRY_OF_ORIGIN),
        ("Place of Receipt by Carrier", place_of_receipt),
        ("Port of Loading",             port_loading),
        ("Port of Discharge",           port_discharge),
    ]

    for i, (label, value) in enumerate(four_cells):
        cx = margin + i * cw
        c_obj.setFillColor(LIGHT_BLUE)
        c_obj.rect(cx, y - four_lbl_h, cw, four_lbl_h, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 6)
        c_obj.drawString(cx + 1.5 * mm, y - four_lbl_h + 1.2 * mm, label)
        c_obj.rect(cx, y - four_lbl_h - four_val_h, cw, four_val_h, fill=0, stroke=1)
        if value:
            c_obj.setFont("Helvetica", 7)
            c_obj.drawString(cx + 1.5 * mm,
                             y - four_lbl_h - four_val_h + (four_val_h - 2.5 * mm) / 2,
                             value)

    y -= (four_lbl_h + four_val_h)

    # Spacer
    spacer_h = 2 * mm
    c_obj.rect(margin, y - spacer_h, width, spacer_h, fill=0, stroke=1)
    y -= spacer_h

    # ═════════════════════════════════════════════════════════════════════════
    # LINE ITEMS TABLE — one row per commodity  ← UPDATED
    # ═════════════════════════════════════════════════════════════════════════
    col_widths = [
        width * 0.05, width * 0.12, width * 0.14, width * 0.10,
        width * 0.20, width * 0.10, width * 0.12, width * 0.17,
    ]
    DESC_COL_IDX = 4   # keep in sync with the "Description" position in tbl_headers

    # Pull line_items built by rows_to_invoice_data(); fall back for old callers
    line_items = data.get("line_items")
    if not line_items:
        weight    = safe_float(data.get("weight",    0))
        inv_price = safe_float(data.get("inv_price", 0))
        inv_amt   = weight * inv_price if (weight and inv_price) \
                    else safe_float(data.get("invoice_amt", 0))
        line_items = [{
            "item_desc":      resolve_item_desc(safe_str(data.get("item_desc", "")), safe_str(data.get("inv_no", ""))),
            "weight":         weight,
            "inv_price":      inv_price,
            "invoice_amt":    inv_amt,
            "freight_charge": safe_str(data.get("freight_charge", "")),
        }]

    # Determine unit from the first item's weight
    use_lbs  = line_items[0]["weight"] > 1000 if line_items else False
    qty_unit = "LBS" if use_lbs else "MT"

    # Table header row
    tbl_headers = [
        "S. No", "Booking #", "Container #", "Seal #", "Description",
        f"Quantity\n{qty_unit}", f"Rate\nUS$/{qty_unit}", "Amount\nUS$",
    ]

    c_obj.setFillColor(LIGHT_BLUE)
    x = margin
    for h, w in zip(tbl_headers, col_widths):
        c_obj.rect(x, y - 8 * mm, w, 8 * mm, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 7)
        lines = h.split("\n")
        if len(lines) == 2:
            c_obj.drawCentredString(x + w / 2, y - 3 * mm, lines[0])
            c_obj.drawCentredString(x + w / 2, y - 6 * mm, lines[1])
        else:
            c_obj.drawCentredString(x + w / 2, y - 4.5 * mm, h)
        x += w
        c_obj.setFillColor(LIGHT_BLUE)
    y -= 8 * mm

    # One data row per line item — row height grows with wrapped description
    subtotal = 0.0
    for s_no, item in enumerate(line_items, start=1):
        weight    = item["weight"]
        inv_price = item["inv_price"]
        inv_amt   = item["invoice_amt"]
        subtotal += inv_amt

        qty_display = f"{weight:,.2f}" if use_lbs else f"{weight:.3f}"

        # Per-item container/booking/seal (from its own sheet row) take
        # priority; falls back to the invoice-level value for old override
        # paths (e.g. manual line-item edits) that don't carry these per item.
        row_vals = [
            str(s_no),
            item.get("booking_no")   or safe_str(data.get("booking_no",   "")),
            item.get("container_no") or safe_str(data.get("container_no", "")),
            item.get("seal_no")      or safe_str(data.get("seal_no",      "")),
            item["item_desc"],
            qty_display,
            format_rate(inv_price),
            f"${inv_amt:,.2f}",
        ]

        y = draw_line_item_row(c_obj, row_vals, col_widths, DESC_COL_IDX, margin, y)

    # Note rows (user-added) — rendered after line items, before deductions
    y, notes_sum = draw_note_rows(c_obj, data, margin, y, col_widths, 7 * mm)

    # Freight — read from top-level data dict (set by rows_to_invoice_data)
    freight_raw = safe_str(data.get("freight_charge", ""))
    freight     = eval_freight(freight_raw) if freight_raw else 0.0
    efs_raw     = safe_str(data.get("efs", ""))
    efs         = eval_freight(efs_raw) if efs_raw else 0.0
    row_h       = 7 * mm

    # Total override — passed through from CLI/UI. Kept as None (not 0.0) when
    # unset so a legitimate $0 override could theoretically be distinguished
    # from "no override" — safe_float would collapse both to 0.0 and hide it.
    total_override_raw = safe_str(data.get("total_override", ""))
    total_override     = safe_float(total_override_raw) if total_override_raw else None
    y, final_amt = draw_summary_rows(c_obj, margin, y, col_widths, row_h,
                                      subtotal, notes_sum, freight, efs, LIGHT_BLUE,
                                      total_override=total_override)
    y -= 3 * mm

    # ═════════════════════════════════════════════════════════════════════════
    # PACKING LIST
    # ═════════════════════════════════════════════════════════════════════════
    PACK_ORANGE       = colors.HexColor("#B5531A")
    PACK_ORANGE_LIGHT = colors.HexColor("#FAE8D8")

    c_obj.setFillColor(PACK_ORANGE)
    c_obj.rect(margin, y - 8 * mm, width, 8 * mm, fill=1, stroke=0)
    c_obj.setFillColor(PACK_ORANGE_LIGHT)
    c_obj.setFont("Helvetica-Bold", 12)
    c_obj.drawCentredString(page_width / 2, y - 5 * mm, "PACKING LIST")
    y -= 8 * mm

    pack_cols = [
        ("Container",             width * 0.15),
        ("Gross Weight\n(lbs)",   width * 0.13),
        ("Truck\n(lbs)",          width * 0.11),
        ("Container\nTare (lbs)", width * 0.12),
        ("Chassis\n(lbs)",        width * 0.11),
        ("Boxes\n(lbs)",          width * 0.12),
        ("Net Weight\n(lbs)",     width * 0.13),
        ("Net Weight\n(MT)",      width * 0.13),
    ]

    c_obj.setFillColor(PACK_ORANGE_LIGHT)
    x = margin
    for lbl, w in pack_cols:
        c_obj.rect(x, y - 10 * mm, w, 10 * mm, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 6.5)
        lines = lbl.split("\n")
        if len(lines) == 2:
            c_obj.drawCentredString(x + w / 2, y - 4   * mm, lines[0])
            c_obj.drawCentredString(x + w / 2, y - 7.5 * mm, lines[1])
        else:
            c_obj.drawCentredString(x + w / 2, y - 6 * mm, lbl)
        x += w
        c_obj.setFillColor(PACK_ORANGE_LIGHT)
    y -= 10 * mm

    # NOTE: packing data is now looked up PER ITEM (each item's own container),
    # not once for the whole invoice — a merged multi-container invoice would
    # otherwise show every item's weights from only the first container's
    # packing data, silently wrong for every other container in the merge.
    default_container_key = safe_str(data.get("container_no", "")).upper().strip()

    # Total net for TOTAL row
    # Accumulate totals from displayed values so TOTAL row matches data rows exactly
    total_net_lbs_sum = 0
    total_net_mt_sum  = 0.0

    # One data row per line item — match packing row by description, within its OWN container
    row_h_pack = 10 * mm
    for s_no, item in enumerate(line_items, start=1):
        item_container_key = safe_str(item.get("container_no", "")).upper().strip() or default_container_key
        packing_rows = (packing_lookup or {}).get(item_container_key, [])
        if packing_rows:
            print(f"  📦 Found {len(packing_rows)} packing row(s) for {item_container_key}")
        else:
            print(f"  ⚠️  No packing data for '{item_container_key}', using calculated values")

        item_wt_mt  = safe_float(item["weight"])
        item_wt_lbs = round(item_wt_mt * 2204.62)

        pr = find_packing_row(packing_rows, item["item_desc"])
        if pr:
            p_gross = safe_str(pr.get("gross_weight_lbs",   ""))
            p_truck = safe_str(pr.get("truck_lbs",          ""))
            p_tare  = safe_str(pr.get("container_tare_lbs", ""))
            p_chas  = safe_str(pr.get("chassis_lbs",        ""))
            p_box   = safe_str(pr.get("boxes_weight_lbs",   ""))
            p_nlbs  = safe_str(pr.get("net_weight_lbs",     f"{item_wt_lbs:,}"))
            p_nmt   = safe_str(pr.get("net_weight_mt",      f"{item_wt_mt:.3f}"))
        else:
            p_gross = str(item_wt_lbs + 31200)
            p_truck = p_tare = p_chas = p_box = ""
            p_nlbs  = f"{item_wt_lbs:,}"
            p_nmt   = f"{item_wt_mt:.3f}"

        # Accumulate from displayed values
        try:    total_net_lbs_sum += int(p_nlbs.replace(",", ""))
        except: total_net_lbs_sum += item_wt_lbs
        try:    total_net_mt_sum  += float(p_nmt.replace(",", ""))
        except: total_net_mt_sum  += item_wt_mt

        pack_values = [
            item_container_key,
            p_gross, p_truck, p_tare, p_chas, p_box, p_nlbs, p_nmt,
        ]

        x = margin
        c_obj.setFillColor(colors.black)
        for val, (lbl, w) in zip(pack_values, pack_cols):
            c_obj.rect(x, y - row_h_pack, w, row_h_pack, fill=0, stroke=1)
            c_obj.setFont("Helvetica", 7)
            c_obj.drawCentredString(x + w / 2, y - row_h_pack / 2 - 1 * mm, val)
            x += w
        y -= row_h_pack

    # TOTAL row
    x = margin
    for i, (lbl, w) in enumerate(pack_cols):
        c_obj.setFillColor(PACK_ORANGE_LIGHT)
        c_obj.rect(x, y - row_h_pack, w, row_h_pack, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 7)
        if i == 0:
            c_obj.drawCentredString(x + w / 2, y - row_h_pack / 2 - 1 * mm, "TOTAL")
        elif i == 6:
            c_obj.drawCentredString(x + w / 2, y - row_h_pack / 2 - 1 * mm, f"{total_net_lbs_sum:,}")
        elif i == 7:
            c_obj.drawCentredString(x + w / 2, y - row_h_pack / 2 - 1 * mm, f"{total_net_mt_sum:.3f}")
        x += w
    y -= row_h_pack + 3 * mm

    # ═════════════════════════════════════════════════════════════════════════
    # FOOTER
    # ═════════════════════════════════════════════════════════════════════════
    footer_h = 15 * mm
    c_obj.rect(margin, y - footer_h, width * 0.5, footer_h, fill=0, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica", 7)
    c_obj.drawString(margin + 2 * mm, y - 4 * mm,
                     "We declare that this invoice shows the actual price of the goods")
    c_obj.drawString(margin + 2 * mm, y - 8 * mm,
                     "described and that all particulars are true and correct.")

    c_obj.rect(margin + width * 0.5, y - footer_h, width * 0.5, footer_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawCentredString(page_width / 2 + width * 0.25, y - 3 * mm, "Authorised Signatory")

    if SIGNATURE_FILE and os.path.exists(SIGNATURE_FILE):
        try:
            sig_width  = 40 * mm
            sig_height =  6 * mm
            sig_x = page_width / 2 + width * 0.25 - sig_width / 2
            sig_y = y - 9.5 * mm
            c_obj.drawImage(SIGNATURE_FILE, sig_x, sig_y,
                            width=sig_width, height=sig_height,
                            preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    c_obj.drawCentredString(page_width / 2 + width * 0.25, y - 13 * mm, "for Edge Metals Inc")

    # Outer border
    border_bottom = y - footer_h
    border_top    = page_height - margin
    c_obj.setStrokeColor(colors.black)
    c_obj.setLineWidth(0.8)
    c_obj.rect(margin, border_bottom, width, border_top - border_bottom, fill=0, stroke=1)
    c_obj.setLineWidth(0.5)


# ═══════════════════════════════════════════════
#  INVOICE ONLY (no packing list)
# ═══════════════════════════════════════════════

def draw_invoice_only(c_obj, data, page_width, page_height, packing_lookup=None, address_lookup=None):
    """Commercial Invoice page WITHOUT the packing list section."""
    BLUE       = colors.HexColor("#2E6DA4")
    LIGHT_BLUE = colors.HexColor("#D9E8F5")

    margin = 15 * mm
    y      = page_height - margin
    width  = page_width - 2 * margin

    # Header bar
    header_h = 10 * mm
    c_obj.setFillColor(BLUE)
    c_obj.rect(margin, y - header_h, width, header_h, fill=1, stroke=0)
    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.setFont("Helvetica-Bold", 14)
    c_obj.drawCentredString(page_width / 2, y - header_h + 3 * mm, "Commercial Invoice")
    y -= header_h

    # ── ROW 1: Exporter | Invoice No / Date / Reference ──────────────────────
    lbl_h     = 5 * mm
    top_val   = 10 * mm
    ref_val_h = 8 * mm
    right_total = lbl_h + top_val + lbl_h + ref_val_h
    exp_val_h   = right_total - lbl_h

    left_x  = margin
    left_w  = width / 2
    right_x = margin + width / 2
    right_w = width / 2
    inv_w   = right_w * 0.65
    date_w  = right_w * 0.35
    date_x  = right_x + inv_w

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(left_x, y - lbl_h, left_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h + 1.5 * mm, "Exporter")
    c_obj.rect(left_x, y - lbl_h - exp_val_h, left_w, exp_val_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 5  * mm, EXPORTER["name"])
    c_obj.setFont("Helvetica", 7)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 10 * mm, EXPORTER["address"])
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 15 * mm, EXPORTER["tel"])
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 20 * mm, EXPORTER["fax"])

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, y - lbl_h, right_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - lbl_h + 1.5 * mm, "Invoice No & Items")
    c_obj.drawString(date_x  + 2 * mm, y - lbl_h + 1.5 * mm, "DATE")

    c_obj.rect(right_x, y - lbl_h - top_val, inv_w, top_val, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - lbl_h - 4 * mm, safe_str(data.get("inv_no", "")))
    _item_label = get_item_label(safe_str(data.get("inv_no", "")))
    if _item_label:
        c_obj.setFont("Helvetica", 6.5)
        c_obj.drawString(right_x + 2 * mm, y - lbl_h - 8 * mm, _item_label)

    c_obj.rect(date_x, y - lbl_h - top_val, date_w, top_val, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(date_x + 2 * mm,
                     y - lbl_h - top_val + (top_val - 2.5 * mm) / 2,
                     format_date(data.get("inv_date", "")))

    ref_top = y - lbl_h - top_val
    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ref_top - lbl_h, right_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h + 1.5 * mm, "Other Reference(s)")
    c_obj.rect(right_x, ref_top - lbl_h - ref_val_h, right_w, ref_val_h, fill=0, stroke=1)
    ref_str   = safe_str(data.get("reference", ""))
    proforma  = format_date(data.get("proforma_date", ""))
    other_ref = " | ".join(filter(None, [proforma, ref_str]))
    c_obj.setFont("Helvetica", 7)
    # Word-wrap other_ref if it exceeds cell width
    from reportlab.pdfbase.pdfmetrics import stringWidth
    max_w = right_w - 4 * mm
    font_size = 7
    if stringWidth(other_ref, "Helvetica", font_size) <= max_w:
        c_obj.drawString(right_x + 2 * mm,
                         ref_top - lbl_h - ref_val_h + (ref_val_h - 2.5 * mm) / 2,
                         other_ref)
    else:
        # Split into two lines at the midpoint word boundary
        words = other_ref.split()
        line1, line2 = "", ""
        for i, word in enumerate(words):
            test = " ".join(words[:i+1])
            if stringWidth(test, "Helvetica", font_size) <= max_w:
                line1 = test
            else:
                line2 = " ".join(words[i:])
                break
        c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h - ref_val_h + 5.5 * mm, line1)
        if line2:
            c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h - ref_val_h + 2.0 * mm, line2)
    y -= right_total

    # ── ROW 2: Buyer | Terms / Vessel / Comments ─────────────────────────────
    buyer_lbl_h = 5 * mm
    cell_lbl_h  = 5 * mm
    cell_val_h  = 6 * mm
    buyer_val_h = cell_val_h + (cell_lbl_h + cell_val_h) * 2

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(left_x, y - buyer_lbl_h, left_w, buyer_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - buyer_lbl_h + 1.5 * mm, "Buyer")

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, y - buyer_lbl_h, right_w, buyer_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - buyer_lbl_h + 1.5 * mm, "Terms")
    y -= buyer_lbl_h

    c_obj.rect(left_x, y - buyer_val_h, left_w, buyer_val_h, fill=0, stroke=1)
    consignee_name = safe_str(data.get("consignee", ""))
    addr_lines = get_buyer_address(consignee_name, address_lookup or {})
    # Consignee override: when the user edits the Company field in the UI,
    # that value prints as the buyer name on the PDF. Address lines below
    # are STILL looked up from the Google Doc using consignee_name — the
    # override only replaces the display name on line 1, not the address
    # data source. When unset, we keep the historical behavior of using
    # the Doc's first line (which is the "canonical" formal name for that
    # buyer) — that way years of previously-generated invoices stay
    # visually consistent for any invoice generated without touching the
    # field. Same skip-when-empty pattern as country_of_origin.
    consignee_override = safe_str(data.get("consignee_override", "")).strip()
    if addr_lines:
        c_obj.setFont("Helvetica-Bold", 9)
        display_name = consignee_override if consignee_override else addr_lines[0].strip()
        c_obj.drawString(left_x + 2 * mm, y - 4 * mm, display_name)
        c_obj.setFont("Helvetica", 7)
        for i, line in enumerate(addr_lines[1:7]):
            c_obj.drawString(left_x + 2 * mm, y - (8 + i * 3.5) * mm, line.strip())

    ry = y
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica", 7)
    terms_val = safe_str(data.get("terms", ""))
    if terms_val:
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, terms_val)
    ry -= cell_val_h

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ry - cell_lbl_h, right_w, cell_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ry - cell_lbl_h + 1.5 * mm, "Vessel / Flight No")
    ry -= cell_lbl_h
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    vessel_val = safe_str(data.get("vessel", ""))
    if vessel_val:
        c_obj.setFont("Helvetica", 7)
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, vessel_val)
    ry -= cell_val_h

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ry - cell_lbl_h, right_w, cell_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ry - cell_lbl_h + 1.5 * mm, "Payment Terms")
    ry -= cell_lbl_h
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    payment_terms_val = safe_str(data.get("payment_terms", ""))
    if payment_terms_val:
        c_obj.setFont("Helvetica", 7)
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, payment_terms_val)
    y -= buyer_val_h

    # ── ROW 3: 4 port cells ───────────────────────────────────────────────────
    four_lbl_h = 5 * mm
    four_val_h = 7 * mm
    cw = width / 4

    container_key_ports = safe_str(data.get("container_no", "")).upper().strip()
    _packing_rows_ports = (packing_lookup or {}).get(container_key_ports, [])
    packing_ports = _packing_rows_ports[0] if _packing_rows_ports else {}

    PORT_ALIASES = {
        "LA": "LOS ANGELES", "LAX": "LOS ANGELES", "LB": "LONG BEACH",
        "NY": "NEW YORK", "CHI": "CHICAGO", "HOU": "HOUSTON", "SAV": "SAVANNAH",
    }
    def expand_port(val: str) -> str:
        v = val.strip().upper()
        return PORT_ALIASES.get(v, val.strip())

    place_of_receipt = first_meaningful(data.get("place_of_receipt", ""),
                                        expand_port(safe_str(packing_ports.get("place_of_receipt", ""))),
                                        "").upper()
    port_loading     = first_meaningful(data.get("port_loading", ""),
                                        expand_port(safe_str(packing_ports.get("port_loading", ""))),
                                        PORT_OF_LOADING).upper()
    port_discharge   = first_meaningful(data.get("port_discharge", ""),
                                        packing_ports.get("port_discharge", ""),
                                        "TO BE ADVISED").upper()

    four_cells = [
        ("Country of Origin",           safe_str(data.get("country_of_origin", "")).upper() or COUNTRY_OF_ORIGIN),
        ("Place of Receipt by Carrier", place_of_receipt),
        ("Port of Loading",             port_loading),
        ("Port of Discharge",           port_discharge),
    ]
    for i, (label, value) in enumerate(four_cells):
        cx = margin + i * cw
        c_obj.setFillColor(LIGHT_BLUE)
        c_obj.rect(cx, y - four_lbl_h, cw, four_lbl_h, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 6)
        c_obj.drawString(cx + 1.5 * mm, y - four_lbl_h + 1.2 * mm, label)
        c_obj.rect(cx, y - four_lbl_h - four_val_h, cw, four_val_h, fill=0, stroke=1)
        if value:
            c_obj.setFont("Helvetica", 7)
            c_obj.drawString(cx + 1.5 * mm,
                             y - four_lbl_h - four_val_h + (four_val_h - 2.5 * mm) / 2,
                             value)
    y -= (four_lbl_h + four_val_h)

    spacer_h = 2 * mm
    c_obj.rect(margin, y - spacer_h, width, spacer_h, fill=0, stroke=1)
    y -= spacer_h

    # ── Line items ────────────────────────────────────────────────────────────
    col_widths = [
        width * 0.05, width * 0.12, width * 0.14, width * 0.10,
        width * 0.20, width * 0.10, width * 0.12, width * 0.17,
    ]
    DESC_COL_IDX = 4   # keep in sync with the "Description" position in tbl_headers

    line_items = data.get("line_items")
    if not line_items:
        weight    = safe_float(data.get("weight",    0))
        inv_price = safe_float(data.get("inv_price", 0))
        inv_amt   = weight * inv_price if (weight and inv_price) \
                    else safe_float(data.get("invoice_amt", 0))
        line_items = [{
            "item_desc":      resolve_item_desc(safe_str(data.get("item_desc", "")), safe_str(data.get("inv_no", ""))),
            "weight":         weight,
            "inv_price":      inv_price,
            "invoice_amt":    inv_amt,
            "freight_charge": safe_str(data.get("freight_charge", "")),
        }]

    use_lbs  = line_items[0]["weight"] > 1000 if line_items else False
    qty_unit = "LBS" if use_lbs else "MT"

    tbl_headers = [
        "S. No", "Booking #", "Container #", "Seal #", "Description",
        f"Quantity\n{qty_unit}", f"Rate\nUS$/{qty_unit}", "Amount\nUS$",
    ]
    c_obj.setFillColor(LIGHT_BLUE)
    x = margin
    for h, w in zip(tbl_headers, col_widths):
        c_obj.rect(x, y - 8 * mm, w, 8 * mm, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 7)
        lines = h.split("\n")
        if len(lines) == 2:
            c_obj.drawCentredString(x + w / 2, y - 3 * mm, lines[0])
            c_obj.drawCentredString(x + w / 2, y - 6 * mm, lines[1])
        else:
            c_obj.drawCentredString(x + w / 2, y - 4.5 * mm, h)
        x += w
        c_obj.setFillColor(LIGHT_BLUE)
    y -= 8 * mm

    subtotal = 0.0
    for s_no, item in enumerate(line_items, start=1):
        weight    = item["weight"]
        inv_price = item["inv_price"]
        inv_amt   = item["invoice_amt"]
        subtotal += inv_amt
        qty_display = f"{weight:,.2f}" if use_lbs else f"{weight:.3f}"
        # Per-item container/booking/seal (from its own sheet row) take
        # priority; falls back to the invoice-level value for old override
        # paths (e.g. manual line-item edits) that don't carry these per item.
        row_vals = [
            str(s_no),
            item.get("booking_no")   or safe_str(data.get("booking_no",   "")),
            item.get("container_no") or safe_str(data.get("container_no", "")),
            item.get("seal_no")      or safe_str(data.get("seal_no",      "")),
            item["item_desc"],
            qty_display,
            format_rate(inv_price),
            f"${inv_amt:,.2f}",
        ]
        y = draw_line_item_row(c_obj, row_vals, col_widths, DESC_COL_IDX, margin, y)

    # Note rows (user-added) — rendered after line items, before deductions
    y, notes_sum = draw_note_rows(c_obj, data, margin, y, col_widths, 7 * mm)

    freight_raw = safe_str(data.get("freight_charge", ""))
    freight     = eval_freight(freight_raw) if freight_raw else 0.0
    efs_raw     = safe_str(data.get("efs", ""))
    efs         = eval_freight(efs_raw) if efs_raw else 0.0
    row_h       = 7 * mm

    # Same total-override wiring as the combined renderer — kept in sync so
    # `--separate` and combined output show identical TOTAL values.
    total_override_raw = safe_str(data.get("total_override", ""))
    total_override     = safe_float(total_override_raw) if total_override_raw else None
    y, final_amt = draw_summary_rows(c_obj, margin, y, col_widths, row_h,
                                      subtotal, notes_sum, freight, efs, LIGHT_BLUE,
                                      total_override=total_override)
    y -= 3 * mm

    # ── Footer (no packing list) ──────────────────────────────────────────────
    footer_h = 15 * mm
    c_obj.rect(margin, y - footer_h, width * 0.5, footer_h, fill=0, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica", 7)
    c_obj.drawString(margin + 2 * mm, y - 4 * mm,
                     "We declare that this invoice shows the actual price of the goods")
    c_obj.drawString(margin + 2 * mm, y - 8 * mm,
                     "described and that all particulars are true and correct.")

    c_obj.rect(margin + width * 0.5, y - footer_h, width * 0.5, footer_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawCentredString(page_width / 2 + width * 0.25, y - 3 * mm, "Authorised Signatory")

    if SIGNATURE_FILE and os.path.exists(SIGNATURE_FILE):
        try:
            sig_width  = 40 * mm
            sig_height =  6 * mm
            sig_x = page_width / 2 + width * 0.25 - sig_width / 2
            sig_y = y - 9.5 * mm
            c_obj.drawImage(SIGNATURE_FILE, sig_x, sig_y,
                            width=sig_width, height=sig_height,
                            preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    c_obj.drawCentredString(page_width / 2 + width * 0.25, y - 13 * mm, "for Edge Metals Inc")


# ═══════════════════════════════════════════════
#  PACKING LIST ONLY
# ═══════════════════════════════════════════════

def draw_packing_list_only(c_obj, data, page_width, page_height, packing_lookup=None, address_lookup=None):
    """Standalone Packing List — same header layout as the invoice."""
    BLUE       = colors.HexColor("#7B2D42")   # deep burgundy for packing list
    LIGHT_BLUE = colors.HexColor("#F2DCE2")   # light rose

    margin = 15 * mm
    y      = page_height - margin
    width  = page_width - 2 * margin

    line_items      = data.get("line_items", [])
    total_weight_mt = sum(item["weight"] for item in line_items) if line_items \
                      else safe_float(data.get("weight", 0))

    # ── HEADER BAR ────────────────────────────────────────────────────────────
    header_h = 10 * mm
    c_obj.setFillColor(BLUE)
    c_obj.rect(margin, y - header_h, width, header_h, fill=1, stroke=0)
    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.setFont("Helvetica-Bold", 14)
    c_obj.drawCentredString(page_width / 2, y - header_h + 3 * mm, "Packing List")  # green header below
    y -= header_h

    # ── ROW 1: Exporter | Invoice No / Date / Reference ──────────────────────
    lbl_h     = 5 * mm
    top_val   = 10 * mm
    ref_val_h = 8 * mm
    right_total = lbl_h + top_val + lbl_h + ref_val_h   # 28 mm
    exp_val_h   = right_total - lbl_h                    # 23 mm

    left_x  = margin
    left_w  = width / 2
    right_x = margin + width / 2
    right_w = width / 2
    inv_w   = right_w * 0.65
    date_w  = right_w * 0.35
    date_x  = right_x + inv_w

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(left_x, y - lbl_h, left_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h + 1.5 * mm, "Exporter")
    c_obj.rect(left_x, y - lbl_h - exp_val_h, left_w, exp_val_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 5  * mm, EXPORTER["name"])
    c_obj.setFont("Helvetica", 7)
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 10 * mm, EXPORTER["address"])
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 15 * mm, EXPORTER["tel"])
    c_obj.drawString(left_x + 2 * mm, y - lbl_h - 20 * mm, EXPORTER["fax"])

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, y - lbl_h, right_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - lbl_h + 1.5 * mm, "Invoice No & Items")
    c_obj.drawString(date_x  + 2 * mm, y - lbl_h + 1.5 * mm, "DATE")

    c_obj.rect(right_x, y - lbl_h - top_val, inv_w, top_val, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - lbl_h - 4 * mm, safe_str(data.get("inv_no", "")))
    _item_label = get_item_label(safe_str(data.get("inv_no", "")))
    if _item_label:
        c_obj.setFont("Helvetica", 6.5)
        c_obj.drawString(right_x + 2 * mm, y - lbl_h - 8 * mm, _item_label)

    c_obj.rect(date_x, y - lbl_h - top_val, date_w, top_val, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(date_x + 2 * mm,
                     y - lbl_h - top_val + (top_val - 2.5 * mm) / 2,
                     format_date(data.get("inv_date", "")))

    ref_top = y - lbl_h - top_val
    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ref_top - lbl_h, right_w, lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h + 1.5 * mm, "Other Reference(s)")
    c_obj.rect(right_x, ref_top - lbl_h - ref_val_h, right_w, ref_val_h, fill=0, stroke=1)
    ref_str   = safe_str(data.get("reference", ""))
    proforma  = format_date(data.get("proforma_date", ""))
    other_ref = " | ".join(filter(None, [proforma, ref_str]))
    c_obj.setFont("Helvetica", 7)
    # Word-wrap other_ref if it exceeds cell width
    from reportlab.pdfbase.pdfmetrics import stringWidth
    max_w = right_w - 4 * mm
    font_size = 7
    if stringWidth(other_ref, "Helvetica", font_size) <= max_w:
        c_obj.drawString(right_x + 2 * mm,
                         ref_top - lbl_h - ref_val_h + (ref_val_h - 2.5 * mm) / 2,
                         other_ref)
    else:
        # Split into two lines at the midpoint word boundary
        words = other_ref.split()
        line1, line2 = "", ""
        for i, word in enumerate(words):
            test = " ".join(words[:i+1])
            if stringWidth(test, "Helvetica", font_size) <= max_w:
                line1 = test
            else:
                line2 = " ".join(words[i:])
                break
        c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h - ref_val_h + 5.5 * mm, line1)
        if line2:
            c_obj.drawString(right_x + 2 * mm, ref_top - lbl_h - ref_val_h + 2.0 * mm, line2)
    y -= right_total

    # ── ROW 2: Buyer | Terms / Vessel / Comments ─────────────────────────────
    buyer_lbl_h = 5 * mm
    cell_lbl_h  = 5 * mm
    cell_val_h  = 6 * mm
    buyer_val_h = cell_val_h + (cell_lbl_h + cell_val_h) * 2   # 28 mm

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(left_x, y - buyer_lbl_h, left_w, buyer_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawString(left_x + 2 * mm, y - buyer_lbl_h + 1.5 * mm, "Buyer")

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, y - buyer_lbl_h, right_w, buyer_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, y - buyer_lbl_h + 1.5 * mm, "Terms")
    y -= buyer_lbl_h

    c_obj.rect(left_x, y - buyer_val_h, left_w, buyer_val_h, fill=0, stroke=1)
    consignee_name = safe_str(data.get("consignee", ""))
    addr_lines = get_buyer_address(consignee_name, address_lookup or {})
    # Consignee override: when the user edits the Company field in the UI,
    # that value prints as the buyer name on the PDF. Address lines below
    # are STILL looked up from the Google Doc using consignee_name — the
    # override only replaces the display name on line 1, not the address
    # data source. When unset, we keep the historical behavior of using
    # the Doc's first line (which is the "canonical" formal name for that
    # buyer) — that way years of previously-generated invoices stay
    # visually consistent for any invoice generated without touching the
    # field. Same skip-when-empty pattern as country_of_origin.
    consignee_override = safe_str(data.get("consignee_override", "")).strip()
    if addr_lines:
        c_obj.setFont("Helvetica-Bold", 9)
        display_name = consignee_override if consignee_override else addr_lines[0].strip()
        c_obj.drawString(left_x + 2 * mm, y - 4 * mm, display_name)
        c_obj.setFont("Helvetica", 7)
        for i, line in enumerate(addr_lines[1:7]):
            c_obj.drawString(left_x + 2 * mm, y - (8 + i * 3.5) * mm, line.strip())

    ry = y
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica", 7)
    terms_val = safe_str(data.get("terms", ""))
    if terms_val:
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, terms_val)
    ry -= cell_val_h

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ry - cell_lbl_h, right_w, cell_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ry - cell_lbl_h + 1.5 * mm, "Vessel / Flight No")
    ry -= cell_lbl_h
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    vessel_val = safe_str(data.get("vessel", ""))
    if vessel_val:
        c_obj.setFont("Helvetica", 7)
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, vessel_val)
    ry -= cell_val_h

    c_obj.setFillColor(LIGHT_BLUE)
    c_obj.rect(right_x, ry - cell_lbl_h, right_w, cell_lbl_h, fill=1, stroke=1)
    c_obj.setFillColor(colors.black)
    c_obj.setFont("Helvetica-Bold", 7)
    c_obj.drawString(right_x + 2 * mm, ry - cell_lbl_h + 1.5 * mm, "Payment Terms")
    ry -= cell_lbl_h
    c_obj.rect(right_x, ry - cell_val_h, right_w, cell_val_h, fill=0, stroke=1)
    payment_terms_val = safe_str(data.get("payment_terms", ""))
    if payment_terms_val:
        c_obj.setFont("Helvetica", 7)
        c_obj.drawString(right_x + 2 * mm, ry - cell_val_h + 1.5 * mm, payment_terms_val)
    y -= buyer_val_h

    # ── ROW 3: 4 port cells ───────────────────────────────────────────────────
    four_lbl_h = 5 * mm
    four_val_h = 7 * mm
    cw = width / 4

    container_key_ports = safe_str(data.get("container_no", "")).upper().strip()
    _packing_rows_ports = (packing_lookup or {}).get(container_key_ports, [])
    packing_ports = _packing_rows_ports[0] if _packing_rows_ports else {}

    PORT_ALIASES = {
        "LA": "LOS ANGELES", "LAX": "LOS ANGELES", "LB": "LONG BEACH",
        "NY": "NEW YORK", "CHI": "CHICAGO", "HOU": "HOUSTON", "SAV": "SAVANNAH",
    }
    def expand_port(val: str) -> str:
        v = val.strip().upper()
        return PORT_ALIASES.get(v, val.strip())

    place_of_receipt = first_meaningful(data.get("place_of_receipt", ""),
                                        expand_port(safe_str(packing_ports.get("place_of_receipt", ""))),
                                        "").upper()
    port_loading     = first_meaningful(data.get("port_loading", ""),
                                        expand_port(safe_str(packing_ports.get("port_loading", ""))),
                                        PORT_OF_LOADING).upper()
    port_discharge   = first_meaningful(data.get("port_discharge", ""),
                                        packing_ports.get("port_discharge", ""),
                                        "TO BE ADVISED").upper()

    four_cells = [
        ("Country of Origin",           safe_str(data.get("country_of_origin", "")).upper() or COUNTRY_OF_ORIGIN),
        ("Place of Receipt by Carrier", place_of_receipt),
        ("Port of Loading",             port_loading),
        ("Port of Discharge",           port_discharge),
    ]
    for i, (label, value) in enumerate(four_cells):
        cx = margin + i * cw
        c_obj.setFillColor(LIGHT_BLUE)
        c_obj.rect(cx, y - four_lbl_h, cw, four_lbl_h, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 6)
        c_obj.drawString(cx + 1.5 * mm, y - four_lbl_h + 1.2 * mm, label)
        c_obj.rect(cx, y - four_lbl_h - four_val_h, cw, four_val_h, fill=0, stroke=1)
        if value:
            c_obj.setFont("Helvetica", 7)
            c_obj.drawString(cx + 1.5 * mm,
                             y - four_lbl_h - four_val_h + (four_val_h - 2.5 * mm) / 2,
                             value)
    y -= (four_lbl_h + four_val_h)

    # Spacer
    spacer_h = 2 * mm
    c_obj.rect(margin, y - spacer_h, width, spacer_h, fill=0, stroke=1)
    y -= spacer_h

    # ── PACKING TABLE — one row per line item ────────────────────────────────
    # Description col intentionally sits at index 1 here (unlike the invoice
    # tables at index 4) — draw_line_item_row is column-agnostic, we just
    # pass DESC_COL_IDX = 1 below. The other 8 columns are all short numeric
    # weights that fit fine without wrap.
    pack_cols = [
        ("S.No",                  width * 0.05),
        ("Description",           width * 0.17),
        ("Container",             width * 0.11),
        ("Gross Weight\n(lbs)",   width * 0.10),
        ("Truck\n(lbs)",          width * 0.08),
        ("Container\nTare (lbs)", width * 0.10),
        ("Chassis\n(lbs)",        width * 0.08),
        ("Boxes\n(lbs)",          width * 0.08),
        ("Net Weight\n(lbs)",     width * 0.12),
        ("Net Weight\n(MT)",      width * 0.11),
    ]
    pack_col_widths = [w for _, w in pack_cols]
    PACK_DESC_COL_IDX = 1

    # Column headers
    c_obj.setFillColor(LIGHT_BLUE)
    x = margin
    for lbl, w in pack_cols:
        c_obj.rect(x, y - 10 * mm, w, 10 * mm, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 6.5)
        lines = lbl.split("\n")
        if len(lines) == 2:
            c_obj.drawCentredString(x + w / 2, y - 4   * mm, lines[0])
            c_obj.drawCentredString(x + w / 2, y - 7.5 * mm, lines[1])
        else:
            c_obj.drawCentredString(x + w / 2, y - 6 * mm, lbl)
        x += w
        c_obj.setFillColor(LIGHT_BLUE)
    y -= 10 * mm

    default_container_key = safe_str(data.get("container_no", "")).upper().strip()

    # Total net for TOTAL row
    # Accumulate totals from displayed values so TOTAL row matches data rows exactly
    total_net_lbs = 0
    total_net_mt  = 0.0

    # One data row per line item — match packing row by description, within its OWN container
    for s_no, item in enumerate(line_items, start=1):
        item_container_key = safe_str(item.get("container_no", "")).upper().strip() or default_container_key
        packing_rows = (packing_lookup or {}).get(item_container_key, [])

        item_wt_mt  = safe_float(item["weight"])
        item_wt_lbs = round(item_wt_mt * 2204.62)

        pr = find_packing_row(packing_rows, item["item_desc"])
        if pr:
            p_gross = safe_str(pr.get("gross_weight_lbs",   ""))
            p_truck = safe_str(pr.get("truck_lbs",          ""))
            p_tare  = safe_str(pr.get("container_tare_lbs", ""))
            p_chas  = safe_str(pr.get("chassis_lbs",        ""))
            p_box   = safe_str(pr.get("boxes_weight_lbs",   ""))
            p_nlbs  = safe_str(pr.get("net_weight_lbs",     f"{item_wt_lbs:,}"))
            p_nmt   = safe_str(pr.get("net_weight_mt",      f"{item_wt_mt:.3f}"))
        else:
            p_gross = str(item_wt_lbs + 31200)
            p_truck = p_tare = p_chas = p_box = ""
            p_nlbs  = f"{item_wt_lbs:,}"
            p_nmt   = f"{item_wt_mt:.3f}"

        # Accumulate from displayed values
        try:    total_net_lbs += int(p_nlbs.replace(",", ""))
        except: total_net_lbs += item_wt_lbs
        try:    total_net_mt  += float(p_nmt.replace(",", ""))
        except: total_net_mt  += item_wt_mt

        row_vals = [
            str(s_no),
            item["item_desc"],
            item_container_key,
            p_gross, p_truck, p_tare, p_chas, p_box, p_nlbs, p_nmt,
        ]

        # Wrap-aware row: pack list uses base_row_h=8mm (matches its previous
        # fixed row_h), grows if description wraps. The TOTAL row below still
        # uses the fixed 8mm — it doesn't have a wrapping description column,
        # so keeping it fixed keeps the summary tight.
        y = draw_line_item_row(c_obj, row_vals, pack_col_widths, PACK_DESC_COL_IDX,
                               margin, y, base_row_h=8 * mm)

    # Totals row — sum of all line item weights (fixed height, no wrap needed)
    row_h = 8 * mm
    x = margin
    for i, (lbl, w) in enumerate(pack_cols):
        c_obj.setFillColor(LIGHT_BLUE)
        c_obj.rect(x, y - row_h, w, row_h, fill=1, stroke=1)
        c_obj.setFillColor(colors.black)
        c_obj.setFont("Helvetica-Bold", 7)
        if i == 1:
            c_obj.drawCentredString(x + w / 2, y - row_h / 2 - 1 * mm, "TOTAL")
        elif i == 8:   # Net Weight (lbs)
            c_obj.drawCentredString(x + w / 2, y - row_h / 2 - 1 * mm, f"{total_net_lbs:,}")
        elif i == 9:   # Net Weight (MT)
            c_obj.drawCentredString(x + w / 2, y - row_h / 2 - 1 * mm, f"{total_net_mt:.3f}")
        x += w
    y -= row_h + 3 * mm

    # ── FOOTER ────────────────────────────────────────────────────────────────
    footer_h = 15 * mm
    c_obj.rect(margin, y - footer_h, width * 0.5, footer_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica", 7)
    c_obj.drawString(margin + 2 * mm, y - 4 * mm,
                     "We declare that this packing list is true and correct")
    c_obj.drawString(margin + 2 * mm, y - 8 * mm,
                     "and corresponds to the commercial invoice.")

    c_obj.rect(margin + width * 0.5, y - footer_h, width * 0.5, footer_h, fill=0, stroke=1)
    c_obj.setFont("Helvetica-Bold", 8)
    c_obj.drawCentredString(page_width / 2 + width * 0.25, y - 3 * mm, "Authorised Signatory")

    if SIGNATURE_FILE and os.path.exists(SIGNATURE_FILE):
        try:
            sig_width  = 40 * mm
            sig_height =  6 * mm
            sig_x = page_width / 2 + width * 0.25 - sig_width / 2
            sig_y = y - 9.5 * mm
            c_obj.drawImage(SIGNATURE_FILE, sig_x, sig_y,
                            width=sig_width, height=sig_height,
                            preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    c_obj.drawCentredString(page_width / 2 + width * 0.25, y - 13 * mm, "for Edge Metals Inc")

    # Outer border
    border_bottom = y - footer_h
    border_top    = page_height - margin
    c_obj.setStrokeColor(colors.black)
    c_obj.setLineWidth(0.8)
    c_obj.rect(margin, border_bottom, width, border_top - border_bottom, fill=0, stroke=1)
    c_obj.setLineWidth(0.5)


# ═══════════════════════════════════════════════
#  MAIN GENERATION
# ═══════════════════════════════════════════════

def generate_invoice(row_data: Dict[str, Any], output_path: str, packing_lookup=None, address_lookup=None):
    """Generate combined invoice + packing list PDF."""
    page_size   = A4
    page_width, page_height = page_size
    c_obj = canvas.Canvas(output_path, pagesize=page_size)
    draw_mk_trading_invoice(c_obj, row_data, page_width, page_height,
                            packing_lookup=packing_lookup,
                            address_lookup=address_lookup)
    c_obj.save()
    print(f"  ✓ {output_path}")


def generate_separate(row_data: Dict[str, Any], output_dir: str,
                      packing_lookup=None, address_lookup=None) -> list:
    """
    Generate two separate PDFs: _INVOICE.pdf and _PACKING_LIST.pdf
    Returns list of both file paths.
    """
    os.makedirs(output_dir, exist_ok=True)
    inv_no     = safe_str(row_data.get("inv_no", "INVOICE")).replace("/", "_").replace(" ", "_")
    page_size  = A4
    page_width, page_height = page_size

    invoice_path = os.path.join(output_dir, f"{inv_no}_INVOICE.pdf")
    c_inv = canvas.Canvas(invoice_path, pagesize=page_size)
    draw_invoice_only(c_inv, row_data, page_width, page_height,
                      packing_lookup=packing_lookup, address_lookup=address_lookup)
    c_inv.save()
    print(f"  ✓ Invoice: {invoice_path}")

    packing_path = os.path.join(output_dir, f"{inv_no}_PACKING_LIST.pdf")
    c_pack = canvas.Canvas(packing_path, pagesize=page_size)
    draw_packing_list_only(c_pack, row_data, page_width, page_height,
                           packing_lookup=packing_lookup, address_lookup=address_lookup)
    c_pack.save()
    print(f"  ✓ Packing List: {packing_path}")

    return [invoice_path, packing_path]


def main():
    parser = argparse.ArgumentParser(description="Edge Metals Invoice Generator (Simplified)")
    parser.add_argument("--row",       type=int,            help="Generate specific row (1-based)")
    parser.add_argument("--container", default="",          help="Generate by container number (preferred over --row)")
    parser.add_argument("--all",       action="store_true", help="Generate all invoices (one PDF per unique container)")
    parser.add_argument("--separate",  action="store_true", help="Generate invoice and packing list as separate PDFs")
    parser.add_argument("--vessel",           default="", help="Vessel / Flight No")
    parser.add_argument("--payment-terms",    default="", help="Payment Terms cell override")
    parser.add_argument("--terms",            default="", help="Payment/delivery terms override")
    parser.add_argument("--inv-no",           default="", help="Invoice number override")
    parser.add_argument("--inv-date",         default="", help="Invoice date override")
    parser.add_argument("--reference",        default="", help="Other reference override")
    parser.add_argument("--proforma-date",    default="", help="Proforma date override")
    parser.add_argument("--booking-no",       default="", help="Booking number override")
    parser.add_argument("--seal-no",          default="", help="Seal number override")
    parser.add_argument("--port-loading",     default="", help="Port of loading override")
    parser.add_argument("--port-discharge",   default="", help="Port of discharge override")
    parser.add_argument("--place-of-receipt", default="", help="Place of receipt override")
    parser.add_argument("--freight",          default="", help="Freight charge override")
    parser.add_argument("--efs",              default="", help="EFS charge override")
    parser.add_argument("--country-of-origin",default="", help="Country of Origin override (default: USA)")
    parser.add_argument("--total-override",   default="", help="Manually override the printed TOTAL; printed as-is with no annotation")
    parser.add_argument("--consignee-override",default="", help="Override the printed buyer/consignee name on the PDF; address lines still come from the Google Doc lookup")
    parser.add_argument("--output",           default="invoices", help="Output directory")
    parser.add_argument("--extra-containers", nargs="*", default=[], help="Additional container numbers")
    parser.add_argument("--line-items-file",  default="", help="Path to JSON file with line item overrides")
    parser.add_argument("--notes-file",       default="", help="Path to JSON file with note rows [{text, amount}]")
    args = parser.parse_args()

    headers, rows  = read_google_sheet_csv(GOOGLE_SHEET_ID, SHEET_NAME)
    packing_lookup = read_packing_lookup_sheet(GOOGLE_SHEET_ID)
    address_lookup = read_address_lookup(ADDRESS_DOC_ID)
    os.makedirs(args.output, exist_ok=True)

    # Group ALL rows by container number upfront
    groups = group_rows_by_container(rows)

    def _run(container_rows, label):
        import json as _json
        data = rows_to_invoice_data(container_rows)
        overrides = {
            "vessel":           args.vessel,
            "payment_terms":    getattr(args, "payment_terms",    ""),
            "terms":            args.terms,
            "inv_no":           getattr(args, "inv_no",           ""),
            "inv_date":         getattr(args, "inv_date",         ""),
            "reference":        getattr(args, "reference",        ""),
            "proforma_date":    getattr(args, "proforma_date",    ""),
            "booking_no":       getattr(args, "booking_no",       ""),
            "seal_no":          getattr(args, "seal_no",          ""),
            "port_loading":     getattr(args, "port_loading",     ""),
            "port_discharge":   getattr(args, "port_discharge",   ""),
            "place_of_receipt": getattr(args, "place_of_receipt", ""),
            "freight_charge":   getattr(args, "freight",          ""),
            "efs":              getattr(args, "efs",              ""),
            "country_of_origin":getattr(args, "country_of_origin",""),
            "total_override":   getattr(args, "total_override",   ""),
            "consignee_override": getattr(args, "consignee_override", ""),
        }
        for k, v in overrides.items():
            if v == "__CLEAR__":
                data[k] = ""
            elif v:
                data[k] = v

        # If booking_no/container_no/seal_no were EXPLICITLY overridden above,
        # that override must win for every row, not just the invoice-level
        # fallback — per-item values (from rows_to_invoice_data) are real
        # sheet data and are almost never empty, so they'd otherwise always
        # beat an explicit override in the row_vals fallback chain.
        for k in ("booking_no", "container_no", "seal_no"):
            v = overrides.get(k, "")
            if v == "__CLEAR__" or v:
                for item in data.get("line_items", []):
                    item[k] = data[k]


        # Apply line items override from JSON file if provided
        li_file = getattr(args, "line_items_file", "")
        if li_file and os.path.exists(li_file):
            try:
                with open(li_file) as f:
                    li_override = _json.load(f)
                if li_override:
                    converted = []
                    for item in li_override:
                        def parse_num(s):
                            try: return float(str(s).replace(",","").replace("$","").replace(" MT","").replace(" LBS","").strip())
                            except: return 0.0
                        w = parse_num(item.get("weight", 0))
                        p = parse_num(item.get("price",  0))
                        a = parse_num(item.get("amount", 0)) or (w * p)
                        converted.append({
                            "item_desc":        item.get("item_desc", ""),
                            "weight":           w,
                            "inv_price":        p,
                            "invoice_amt":      a,
                            "freight_charge":   "",
                            "efs":              "",
                            "gross_weight_lbs": item.get("gross_weight_lbs", ""),
                            "truck_lbs":        item.get("truck_lbs",        ""),
                            "container_tare":   item.get("container_tare",   ""),
                            "chassis_lbs":      item.get("chassis_lbs",      ""),
                            "boxes_lbs":        item.get("boxes_lbs",        ""),
                            "net_weight_lbs":   item.get("net_weight_lbs",   ""),
                            "net_weight_mt":    item.get("net_weight_mt",    ""),
                            # Preserve per-item container/booking/seal through the
                            # override round-trip — without this, editing ANY line
                            # item (even just fixing a typo in description) would
                            # silently collapse every row back to the invoice-level
                            # fallback, undoing the multi-container per-row fix.
                            "container_no":     item.get("container_no", "") or safe_str(data.get("container_no", "")),
                            "booking_no":        item.get("booking_no",   "") or safe_str(data.get("booking_no",   "")),
                            "seal_no":           item.get("seal_no",      "") or safe_str(data.get("seal_no",      "")),
                        })
                    if converted:
                        data["line_items"] = converted
                        print(f"  ✏️  Applied {len(converted)} line item override(s)")
            except Exception as e:
                print(f"  ⚠️  Could not load line items override: {e}")

        # Apply note rows from JSON file if provided
        notes_file = getattr(args, "notes_file", "")
        if notes_file and os.path.exists(notes_file):
            try:
                with open(notes_file) as f:
                    notes_raw = _json.load(f)
                note_rows = []
                for n in (notes_raw or []):
                    txt = safe_str(n.get("text", "")).strip()
                    amt_raw = safe_str(n.get("amount", "")).strip()
                    if not txt and not amt_raw:
                        continue
                    amt = None
                    if amt_raw:
                        try:
                            amt = float(amt_raw.replace(",", "").replace("$", ""))
                        except Exception:
                            amt = None
                    note_rows.append({"text": txt, "amount": amt})
                if note_rows:
                    data["note_rows"] = note_rows
                    print(f"  📝 Applied {len(note_rows)} note row(s)")
            except Exception as e:
                print(f"  ⚠️  Could not load notes file: {e}")

        inv_no = safe_str(data.get("inv_no", label)).replace("/", "_").replace(" ", "_")
        if args.separate:
            generate_separate(data, args.output,
                              packing_lookup=packing_lookup, address_lookup=address_lookup)
        else:
            out_path = os.path.join(args.output, f"{inv_no}.pdf")
            generate_invoice(data, out_path,
                             packing_lookup=packing_lookup, address_lookup=address_lookup)

    if args.all:
        print(f"\nGenerating invoices for {len(groups)} unique containers...")
        for i, (container_no, container_rows) in enumerate(groups.items()):
            print(f"\n[{i+1}/{len(groups)}] {container_no} ({len(container_rows)} line item(s))...")
            _run(container_rows, container_no)

    elif args.container:
        all_container_nos = [args.container.upper().strip()]
        if args.extra_containers:
            all_container_nos += [c.upper().strip() for c in args.extra_containers if c.strip()]

        if len(all_container_nos) == 1:
            container_no   = all_container_nos[0]
            container_rows = groups.get(container_no)
            if not container_rows:
                print(f"❌ Container '{container_no}' not found in sheet")
                sys.exit(1)
            print(f"\nGenerating: {container_no} ({len(container_rows)} line item(s))...")
            _run(container_rows, container_no)
        else:
            merged_rows = []
            labels = []
            for cno in all_container_nos:
                crows = groups.get(cno)
                if not crows:
                    print(f"⚠️  Container '{cno}' not found — skipping")
                    continue
                merged_rows.extend(crows)
                labels.append(cno)
            if not merged_rows:
                print("❌ No containers found")
                sys.exit(1)
            label = "_".join(labels)
            print(f"\nGenerating combined: {labels} ({len(merged_rows)} total line item(s))...")
            _run(merged_rows, label)

    elif args.row is not None:
        idx = args.row - 1
        if idx < 0 or idx >= len(rows):
            print(f"❌ Row {args.row} out of range (1–{len(rows)})")
            sys.exit(1)
        container_col  = COLUMNS["container_no"]
        container_no   = safe_str(
            rows[idx][container_col] if container_col < len(rows[idx]) else ""
        ).upper().strip()
        container_rows = groups.get(container_no, [rows[idx]])
        print(f"\nGenerating row {args.row}: {container_no} ({len(container_rows)} line item(s))...")
        _run(container_rows, container_no)

    else:
        container_no   = list(groups.keys())[0]
        container_rows = groups[container_no]
        print(f"\nGenerating: {container_no} ({len(container_rows)} line item(s))...")
        _run(container_rows, container_no)

    print(f"\n✅ Done! Files saved to: {args.output}/")


if __name__ == "__main__":
    main()

# ═══════════════════════════════════════════════════════════════════
#  PROFORMA INVOICE dc-2 — spec-list layout, aggregated commodities
# ═══════════════════════════════════════════════════════════════════

def _amount_in_words_dc2(amount: float) -> str:
    ones  = ["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine",
             "Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen",
             "Seventeen","Eighteen","Nineteen"]
    tens_ = ["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"]
    def _b1000(n):
        if n==0: return ""
        elif n<20: return ones[n]
        elif n<100: return tens_[n//10]+("" if n%10==0 else " "+ones[n%10])
        else:
            r=_b1000(n%100)
            return ones[n//100]+" Hundred"+(" "+r if r else "")
    n=int(round(amount))
    if n==0: return "Zero"
    parts=[]
    bi=n//1_000_000_000; n%=1_000_000_000
    mi=n//1_000_000;     n%=1_000_000
    th=n//1_000;         n%=1_000
    re_=n
    if bi: parts.append(_b1000(bi)+" Billion")
    if mi: parts.append(_b1000(mi)+" Million")
    if th: parts.append(_b1000(th)+" Thousand")
    if re_: parts.append(_b1000(re_))
    return "US Dollars "+" ".join(parts)+" Only"


def draw_proforma_invoice_dc2(c, data: Dict[str, Any], page_w: float, page_h: float) -> None:
    """
    dc-2 proforma: centered letterhead, spec-list, green total box,
    bank section, memo box, signatures, footer.
    Aggregates qty per commodity across all containers.
    """
    from reportlab.lib.colors import HexColor, white
    from reportlab.lib.units  import mm
    from reportlab.pdfbase    import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # ── Fonts ─────────────────────────────────────────────────────
    _fd = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    def _reg(name, fn):
        p = os.path.join(_fd, fn)
        if os.path.exists(p):
            try: pdfmetrics.registerFont(TTFont(name, p)); return True
            except: pass
        return False
    _hp = _reg("Poppins-Bold",    "poppins_bold.ttf")
    _hd = _reg("DMSans-Regular",  "DMSans-Regular.ttf")
    _hb = _reg("DMSans-Bold",     "DMSans-Bold.ttf")
    _hs = _reg("DMSans-SemiBold", "DMSans-SemiBold.ttf")
    FH  = "Poppins-Bold"    if _hp else "Helvetica-Bold"
    FB  = "DMSans-Regular"  if _hd else "Helvetica"
    FBD = "DMSans-Bold"     if _hb else "Helvetica-Bold"
    FSM = "DMSans-SemiBold" if _hs else "Helvetica-Bold"

    # ── Colours ───────────────────────────────────────────────────
    INK      = HexColor("#1A211B"); INK_MID  = HexColor("#3A423B")
    INK_FAINT= HexColor("#5A625B"); MUTED    = HexColor("#8A938B")
    LIGHTER  = HexColor("#9AA39B"); GREEN_MID= HexColor("#2A7D24")
    GREEN_MN = HexColor("#3DA836"); GREEN_LT = HexColor("#EDF7EB")
    RULE     = HexColor("#E6EBE6"); RULE_MED = HexColor("#C9D0C9")
    MEMO_BDR = HexColor("#D9E0D9"); MEMO_BG  = HexColor("#F4F7F4")
    MEMO_RED = HexColor("#B23B2E")

    L=20*mm; R=page_w-20*mm; W=R-L; y=page_h-14*mm

    # ── Aggregate commodities across containers ───────────────────
    from collections import OrderedDict
    from decimal import Decimal as D, ROUND_HALF_UP
    comms = OrderedDict()
    for cont in data.get("containers", []):
        for item in cont.get("items", []):
            desc = item.get("desc", "")
            if desc not in comms:
                comms[desc] = {"qty": 0.0, "rate": float(item.get("rate", 0))}
            comms[desc]["qty"] += float(item.get("qty", 0))
    total_due = 0.0
    for desc, v in comms.items():
        v["amount"] = float(D(str(v["qty"])) * D(str(v["rate"])))
        total_due  += v["amount"]
    if data.get("total_due"):
        total_due = float(data["total_due"])
    num_containers = len(data.get("containers", []))
    container_label = f"{num_containers} \u00d7 40\u2032 HC"

    # ── Unit detection: MT vs LBS ────────────────────────────────────
    # If ANY raw line item's qty (before aggregation across containers)
    # exceeds 10,000 we treat the whole document as LBS. Checked on the
    # raw per-container items, not the aggregated `comms` totals, so a
    # single small commodity doesn't get relabeled just because it was
    # summed across many containers. Mirrors the same threshold/logic
    # used in the HTML/WeasyPrint dc-2 renderer for consistency.
    qty_unit = "LBS" if any(
        float(item.get("qty", 0) or 0) > 10000
        for cont in data.get("containers", [])
        for item in cont.get("items", [])
    ) else "MT"

    # ── Logo + name inline ────────────────────────────────────────
    logo_sz=12*mm; co_name="Edge Metals"; co_size=20; gap=3*mm
    name_w = c.stringWidth(co_name, FH, co_size)
    pair_w = logo_sz+gap+name_w; lx=page_w/2-pair_w/2; ly=y-logo_sz
    c.setFillColor(GREEN_MN)
    c.roundRect(lx, ly, logo_sz, logo_sz, 2.8*mm, fill=1, stroke=0)
    c.setFillColor(white)
    bw=logo_sz*.58; bh=logo_sz*.115; bx=lx+logo_sz*.225
    c.rect(bx, ly+logo_sz*.615, bw,       bh, fill=1, stroke=0)
    c.rect(bx, ly+logo_sz*.425, bw*.70,   bh, fill=1, stroke=0)
    c.rect(bx, ly+logo_sz*.235, bw,       bh, fill=1, stroke=0)
    c.setFillColor(HexColor("#BBE9B2"))
    c.circle(lx+logo_sz*.85, ly+logo_sz*.82, logo_sz*.09, fill=1, stroke=0)
    c.setFont(FH, co_size); c.setFillColor(INK)
    c.drawString(lx+logo_sz+gap, ly+logo_sz/2-co_size*.35, co_name)
    y = ly-5*mm

    # Address
    c.setFont(FB, 8); c.setFillColor(INK_FAINT)
    c.drawCentredString(page_w/2, y-3*mm, "1848 E 55th Street, Los Angeles, CA 90058, USA")
    y -= 5*mm
    c.drawCentredString(page_w/2, y-3*mm, "+1 (310) 938-2525  \u00b7  +1 (213) 507-5755")
    y -= 5*mm
    email_p="bose@edgemetals.com  \u00b7  "; web_p="www.edgemetals.com"
    fw=c.stringWidth(email_p,FB,8)+c.stringWidth(web_p,FSM,8)
    cx=page_w/2-fw/2
    c.setFont(FB,8); c.setFillColor(INK)
    c.drawString(cx, y-3*mm, email_p)
    c.setFont(FSM,8); c.setFillColor(GREEN_MID)
    c.drawString(cx+c.stringWidth(email_p,FB,8), y-3*mm, web_p)
    y -= 9*mm

    # Title
    title="Proforma Invoice"; tw=c.stringWidth(title,FH,17)
    c.setFont(FH,17); c.setFillColor(INK)
    c.drawCentredString(page_w/2, y-4*mm, title)
    y -= 9*mm
    c.setFillColor(GREEN_MN)
    c.roundRect(page_w/2-tw/2, y, tw, 2, 1, fill=1, stroke=0)
    y -= 7*mm

    # ── Customer + Refs ───────────────────────────────────────────
    col_mid=page_w/2-4*mm; row_top=y
    c.setFont(FB,8); c.setFillColor(MUTED)
    c.drawString(L, row_top, "Customer :")
    lbl_w=c.stringWidth("Customer :  ",FB,8); cx2=L+lbl_w
    consignee=safe_str(data.get("consignee",""))
    # Name intentionally not drawn — address lines only
    ry=row_top-4.5*mm
    for line in data.get("consignee_address",[]):
        c.setFont(FBD,8); c.setFillColor(INK); c.drawString(cx2, ry, line); ry -= 4*mm
    ref_rows=[
        ("Invoice Date :",    format_date(data.get("inv_date",""))),
        ("Invoice No :",      safe_str(data.get("inv_no",""))),
        ("Buyer PO :",        safe_str(data.get("buyer_po",""))),
        ("Purchase Person :", safe_str(data.get("purchase_person",""))),
        ("Prepared By :",     safe_str(data.get("prepared_by",""))),
    ]
    rry=row_top
    for lbl, val in ref_rows:
        c.setFont(FB,8); c.setFillColor(MUTED); c.drawString(col_mid, rry, lbl)
        c.setFont(FSM,8); c.setFillColor(INK); c.drawRightString(R, rry, val)
        rry -= 5*mm
    y = min(ry, rry)-6*mm

    # ── Spec list ─────────────────────────────────────────────────
    c.setStrokeColor(RULE); c.setLineWidth(.5); c.line(L, y+2*mm, R, y+2*mm)
    y -= 4*mm
    LABEL_W=52*mm

    def spec_row(label, value):
        nonlocal y
        c.setFont(FB,8); c.setFillColor(MUTED); c.drawString(L, y, label)
        c.setFont(FSM,8); c.setFillColor(INK);  c.drawString(L+LABEL_W, y, value)
        y -= 5*mm

    for i,(desc,v) in enumerate(comms.items()):
        spec_row("Commodity" if i==0 else "", desc)
    for i,(desc,v) in enumerate(comms.items()):
        spec_row("Rate" if i==0 else "", f"{format_rate(v['rate'])} / {qty_unit}")
    spec_row("Currency",               safe_str(data.get("currency","US Dollar ($)")))
    spec_row("Trade Terms",            safe_str(data.get("trade_terms","")))
    for i,(desc,v) in enumerate(comms.items()):
        qty_disp = f"{v['qty']:,.0f}" if qty_unit == "LBS" else f"{v['qty']:.0f}"
        spec_row("Quantity (Weight)" if i==0 else "", f"{qty_disp} {qty_unit}  ({desc})")
    spec_row("Quantity (Containers)",  container_label)
    spec_row("Packaging",              safe_str(data.get("packaging","Loose")))
    spec_row("Shipment Qty. Allowance",safe_str(data.get("shipment_allowance","+/- 10% on weights")))
    spec_row("Origin of Material",     safe_str(data.get("country_of_origin","USA")))
    spec_row("Ship To",                safe_str(data.get("port_discharge","")))
    spec_row("Payment Term",           safe_str(data.get("payment_term","T/T 100% Against Shipping Documents")))
    spec_row("Documents",              safe_str(data.get("documents","BL \u00b7 Commercial Invoice \u00b7 Packing List \u00b7 Weight Ticket \u00b7 8 Container Loading Pics")))
    y -= 3*mm

    # ── Estimated Total box ───────────────────────────────────────
    box_h=14*mm
    c.setFillColor(GREEN_LT)
    c.roundRect(L, y-box_h, W, box_h, 3*mm, fill=1, stroke=0)
    frt=safe_str(data.get("freight_label","CIF (freight included)"))
    def _qty_disp(q):
        return f"{q:,.0f}" if qty_unit == "LBS" else f"{q:.0f}"

    if len(comms)==1:
        desc0,v0=next(iter(comms.items()))
        subtitle=f"{_qty_disp(v0['qty'])} {qty_unit} \u00d7 {format_rate(v0['rate'])} / {qty_unit}  \u00b7  {frt}"
    else:
        parts=" + ".join(f"{_qty_disp(v['qty'])} {qty_unit} \u00d7 {format_rate(v['rate'])}" for v in comms.values())
        subtitle=f"{parts}  \u00b7  {frt}"
    c.setFont(FBD,7); c.setFillColor(GREEN_MID)
    c.drawString(L+5*mm, y-5*mm, "ESTIMATED TOTAL AMOUNT")
    c.setFont(FB,7.5); c.setFillColor(INK_FAINT)
    c.drawString(L+5*mm, y-9.5*mm, subtitle)
    c.setFont(FH,17); c.setFillColor(INK)
    c.drawRightString(R-5*mm, y-10*mm, f"${total_due:,.2f}")
    y -= box_h+6*mm

    # ── Bank ──────────────────────────────────────────────────────
    c.setFont(FBD,7); c.setFillColor(GREEN_MID); c.drawString(L, y, "BENEFICIARY BANK")
    y -= 5.5*mm
    for lbl,val in [("Beneficiary",safe_str(data.get("bank_beneficiary","Edge Metals Inc."))),
                    ("Bank",       safe_str(data.get("bank_name",""))),
                    ("Account / SWIFT",safe_str(data.get("bank_account_swift","")))]:
        if not val: continue
        c.setFont(FB,8); c.setFillColor(MUTED); c.drawString(L, y, lbl)
        c.setFont(FSM,8); c.setFillColor(INK);  c.drawString(L+LABEL_W, y, val)
        y -= 5*mm
    y -= 4*mm

    # ── Memo ──────────────────────────────────────────────────────
    memo=safe_str(data.get("memo",""))
    if memo:
        memo_h=16*mm
        c.setStrokeColor(MEMO_BDR); c.setLineWidth(.5)
        c.roundRect(L, y-memo_h, W, memo_h, 2.5*mm, fill=0, stroke=1)
        c.setFillColor(MEMO_BG)
        c.roundRect(L+.5, y-6.5*mm, W-1, 6.5*mm, 2.5*mm, fill=1, stroke=0)
        c.rect(L+.5, y-6.5*mm, W-1, 3*mm, fill=1, stroke=0)
        c.setStrokeColor(RULE); c.setLineWidth(.4)
        c.line(L+1, y-6.5*mm, R-1, y-6.5*mm)
        c.setFont(FH,8); c.setFillColor(GREEN_MID)
        c.drawCentredString(page_w/2, y-4.5*mm, "Memo")
        c.setFont(FBD,9); c.setFillColor(MEMO_RED)
        c.drawCentredString(page_w/2, y-12*mm, memo)
        y -= memo_h+6*mm

    # ── Signatures ────────────────────────────────────────────────
    y -= 4*mm; half_w=(W-16*mm)/2
    for side,role,name in [("left","SELLER","Edge Metals Inc."),
                            ("right","BUYER", data.get("consignee_address",[""])[0] if data.get("consignee_address") else "")]:
        sx=L if side=="left" else L+half_w+16*mm; mid=sx+half_w/2
        c.setFont(FBD,7); c.setFillColor(MUTED); c.drawCentredString(mid, y, role)
        c.setFont(FH,10); c.setFillColor(INK);   c.drawCentredString(mid, y-5*mm, name)
        sig_y=y-18*mm
        c.setFont(FB,8.5); c.setFillColor(MUTED); c.drawString(sx, sig_y, "X")
        c.setStrokeColor(RULE_MED); c.setLineWidth(1)
        c.line(sx+5*mm, sig_y, sx+half_w, sig_y)
        c.setFont(FB,7); c.setFillColor(LIGHTER)
        c.drawCentredString(mid, sig_y-4*mm, "Authorized Signature")
    y -= 26*mm

    # ── Footer ────────────────────────────────────────────────────
    y -= 5*mm
    c.setStrokeColor(RULE); c.setLineWidth(.4); c.line(L, y, R, y); y -= 4.5*mm
    c.setFont(FB,7); c.setFillColor(LIGHTER)
    c.drawString(L, y, "This proforma invoice reflects the actual price of the goods described. Inspection welcome \u00b7 weight-discrepancy guarantee.")
    dot_x=R-c.stringWidth("Edge Metals Inc.",FBD,7)-5*mm
    c.setFillColor(GREEN_MN); c.circle(dot_x, y+2, 2, fill=1, stroke=0)
    c.setFont(FBD,7); c.setFillColor(GREEN_MID)
    c.drawString(dot_x+4*mm, y, "Edge Metals Inc.")


def generate_proforma_pdf_dc2(data: Dict[str, Any], output_path: str) -> str:
    """Build and save a dc-2 style proforma invoice PDF."""
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as _canvas
    page_w, page_h = A4
    c = _canvas.Canvas(output_path, pagesize=A4)
    draw_proforma_invoice_dc2(c, data, page_w, page_h)
    c.save()
    return output_path
