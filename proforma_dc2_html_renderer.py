"""
dc-2 Proforma Invoice — HTML/CSS exact-match renderer.

Replaces the ReportLab hand-drawn approximation with a real HTML+CSS
template rendered through WeasyPrint, so gradients, fonts, spacing,
and the radial glow blobs come out pixel-identical to the reference
design instead of being approximated with canvas primitives.

Requirements (install once on the server):
    pip install weasyprint jinja2 --break-system-packages

Font availability used by the template (verify these paths exist on
the production server, or update the @font-face src paths in
proforma_dc2_template.html):
    - Poppins:        /usr/share/fonts/truetype/google-fonts/Poppins-*.ttf
    - IBM Plex Mono:  bundle alongside this file (see FONTS dir below)
    - DM Sans:        NOT available on this build server — substituted
                       with DejaVu Sans in the template. If you have the
                       actual DM Sans .ttf files, drop them next to this
                       script and update the @font-face src in the
                       template for a true exact match on body text.
"""

import os
from datetime import datetime
from typing import Any, Dict
from jinja2 import Environment, FileSystemLoader

_TEMPLATE_DIR = os.path.dirname(os.path.abspath(__file__))
_env = Environment(loader=FileSystemLoader(_TEMPLATE_DIR))

# Seller signature — signature.png lives next to app.py/invoice_gen.py on
# the server. No conversion needed; just check it exists so a missing file
# never breaks PDF generation.
_SIGNATURE_PNG = os.path.join(_TEMPLATE_DIR, "signature.png")


def _signature_filename() -> str:
    return "signature.png" if os.path.exists(_SIGNATURE_PNG) else ""


def _format_date(val) -> str:
    if isinstance(val, datetime):
        return val.strftime("%m/%d/%Y")
    if val:
        s = str(val).strip()
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]:
            try:
                return datetime.strptime(s, fmt).strftime("%m/%d/%Y")
            except Exception:
                continue
        return s
    return ""


def _amount_in_words(amount: float) -> str:
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
            "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def _b1000(n):
        if n == 0:
            return ""
        if n < 20:
            return ones[n]
        if n < 100:
            return tens[n // 10] + ("" if n % 10 == 0 else " " + ones[n % 10])
        r = _b1000(n % 100)
        return ones[n // 100] + " Hundred" + (" " + r if r else "")

    n = int(round(amount))
    if n == 0:
        return "Zero"
    parts = []
    bi, n = divmod(n, 1_000_000_000)
    mi, n = divmod(n, 1_000_000)
    th, n = divmod(n, 1_000)
    if bi:
        parts.append(_b1000(bi) + " Billion")
    if mi:
        parts.append(_b1000(mi) + " Million")
    if th:
        parts.append(_b1000(th) + " Thousand")
    if n:
        parts.append(_b1000(n))
    return "US Dollars " + " ".join(parts) + " Only"


def _format_rate(value) -> str:
    """Format a per-unit rate preserving exactly the precision the person
    typed on the website — 1.47 stays $1.47, 0.4256 stays $0.4256 — never
    rounded to a fixed 0 or 2 decimals. Only floor: 2 decimals minimum, so
    whole numbers still read as $1.00, not $1. Mirrors format_rate() in
    invoice_gen.py; kept as a local copy since this module doesn't import
    invoice_gen.py elsewhere.
    """
    from decimal import Decimal
    d = Decimal(str(value)).normalize()
    if d.as_tuple().exponent > 0:
        d = d.quantize(Decimal(1))
    if d.as_tuple().exponent > -2:
        d = d.quantize(Decimal("0.01"))
    return f"${d:,f}"


def _format_qty(value) -> str:
    """Comma-separated, 2 decimals, always — same treatment regardless of
    whether the number is 21 or 44000. No magnitude-based branching, so a
    quantity displays exactly as entered on the website, just formatted
    for readability."""
    return f"{float(value):,.2f}"


def render_proforma_dc2_html(data: Dict[str, Any]) -> str:
    """Build the filled HTML string from invoice data. Exposed separately
    from PDF generation so the form/UI can preview it in an iframe too."""
    containers = data.get("containers", [])
    total_qty = 0.0
    total_due = 0.0
    # Build a render-only copy of containers with pre-formatted strings per
    # item, so the template does no magnitude-dependent formatting logic —
    # every quantity and rate is displayed exactly as entered on the
    # website (proforma.html), just formatted for readability, with no
    # automatic unit relabeling or precision changes based on value size.
    render_containers = []
    for cont in containers:
        render_items = []
        for item in cont.get("items", []):
            qty = float(item.get("qty", 0))
            rate = float(item.get("rate", 0))
            total_qty += qty
            total_due += qty * rate
            render_items.append({
                "desc":      item.get("desc", ""),
                "qty":       qty,
                "rate":      rate,
                "qty_fmt":   _format_qty(qty),
                "rate_fmt":  _format_rate(rate),
                "amount_fmt": f"${qty*rate:,.2f}",
            })
        render_containers.append({"container_no": cont.get("container_no", ""), "items": render_items})
    if data.get("total_due"):
        total_due = float(data["total_due"])

    addr = data.get("consignee_address", []) or []
    consignee_name = addr[0] if addr else data.get("consignee", "")
    # Everything after the name splits into "contact" lines (start with
    # TEL/FAX/PHONE, case-insensitive) and "address" lines (everything
    # else). This holds regardless of how many lines are in each section,
    # and works the same whether the address came from the Google Doc
    # lookup or was hand-edited in the UI textarea.
    rest = addr[1:] if len(addr) > 1 else []
    address_lines = [l for l in rest if not l.strip().upper().startswith(("TEL", "FAX", "PHONE"))]
    contact_lines = [l for l in rest if l.strip().upper().startswith(("TEL", "FAX", "PHONE"))]
    consignee_address_line = "<br>".join(address_lines)
    consignee_contact_line = "<br>".join(contact_lines)

    trade_terms = data.get("trade_terms", "")
    trade_terms_short = trade_terms.split()[0] if trade_terms else "CIF"

    freight_label = data.get("freight_label", "CIF (freight included)")
    freight_status = "INCLUDED" if "included" in freight_label.lower() or freight_label.upper().startswith("CIF") \
        else ("EXCLUDED" if freight_label else "NOT STATED")

    terms_cells = [
        ("Buyer's Order No & Date", (data.get("buyer_po","") + (" \u00b7 " + _format_date(data.get("buyer_po_date","")) if data.get("buyer_po_date") else "")) or "-"),
        ("Payment Terms",           data.get("payment_term", "T/T 100% Against Shipping Documents")),
        ("Trade Terms",             trade_terms),
        ("Origin",                  data.get("country_of_origin", "USA")),
        ("Pre-Carriage",            "By Sea"),
        ("Port of Loading",         "USA Ports"),
        ("Port of Discharge",       data.get("port_discharge", "")),
        ("Shipment Allowance",      data.get("shipment_allowance", "+/- 10% on weights")),
    ]

    tpl = _env.get_template("proforma_dc2_template.html")
    return tpl.render(
        inv_no=data.get("inv_no", ""),
        inv_date=_format_date(data.get("inv_date", "")),
        reference=data.get("reference", "Proforma & Email Conf."),
        consignee_name=consignee_name,
        consignee_address_line=consignee_address_line,
        consignee_contact_line=consignee_contact_line,
        terms_cells=terms_cells,
        containers=render_containers,
        total_qty=total_qty,
        total_qty_fmt=_format_qty(total_qty),
        total_due=total_due,
        # Unit is whatever the person explicitly selected on the form's
        # MT/LBS dropdown — never inferred from the size of the number.
        # Defaults to "MT" only if the field is missing entirely (e.g. an
        # older saved draft), matching the website's original behavior.
        qty_unit=data.get("qty_unit", "MT"),
        amount_in_words=_amount_in_words(total_due),
        trade_terms_short=trade_terms_short,
        freight_status=freight_status,
        shipment_allowance=data.get("shipment_allowance", "+/- 10% on weights"),
        signature_img=_signature_filename(),
    )


def generate_proforma_pdf_dc2_html(data: Dict[str, Any], output_path: str) -> str:
    """Exact-match dc-2 proforma PDF, rendered through WeasyPrint from
    real HTML/CSS instead of hand-drawn ReportLab primitives."""
    from weasyprint import HTML
    html_str = render_proforma_dc2_html(data)
    HTML(string=html_str, base_url=_TEMPLATE_DIR).write_pdf(output_path)
    return output_path
