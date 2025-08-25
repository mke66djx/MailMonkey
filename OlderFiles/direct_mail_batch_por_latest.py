# direct_mail_batch_por.py  — FULL VERSION
# Implements trust/partnership salutation rule and robust first-name handling
# while preserving original features (combined PDF, mapping CSV, optional singles).

import os
import csv
import re
import sys
import argparse
from typing import Dict, List, Tuple, Optional

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.utils import ImageReader

# ------------------------
# Configuration & Defaults
# ------------------------

POSSIBLE_ADDRESS = [
    "property address", "situs address", "situs_address", "address",
    "mailing address", "mailing_address", "situs", "property_address"
]

# Expanded to include primary-* variants
POSSIBLE_OWNER_FIRST = [
    "owner first name","owner_first_name","first name","firstname","ownerfirst",
    "owner first","owner_first","ownerfirst name",
    "primary first","primary first name","primary_first","primary_first_name"
]

# Expanded to include Primary Name
POSSIBLE_OWNER_NAME  = [
    "primary name","primary_name","owner name","owner","name",
    "owner_full_name","owner full name","owner(s)","owner 1","owner1"
]

POSSIBLE_SIG_IMAGE = ["signature image","signature_image","sig image","sig_image"]

# Styles
styles = getSampleStyleSheet()
body_style = ParagraphStyle(
    "body", parent=styles["Normal"], fontName="Times-Roman",
    fontSize=11, leading=14, alignment=TA_LEFT,
)
greeting_style = ParagraphStyle(
    "greet", parent=styles["Normal"], fontName="Times-Bold",
    fontSize=12, leading=16, spaceAfter=6,
)
sig_lead = ParagraphStyle(
    "sig_lead", parent=styles["Normal"], spaceBefore=18,
    fontName="Times-Roman", fontSize=11,
)
sig_line = ParagraphStyle(
    "sig_line", parent=styles["Normal"], fontName="Times-Roman",
    fontSize=10, leading=12,
)

# Your template (unchanged). We will surgically replace the first line when needed.
LETTER_TEMPLATE = """Dear {OwnerFirstName},

I hope this note finds you well. My name is {YourName}. My dad, and I run a small real estate business buying homes around Sacramento. We're reaching out because we're interested in your home on {StreetName}.

We buy as-is, cover closing costs, and can work on your timeline—fast or slow—whatever works best for you. No pressure and no obligation. If you'd like to explore options, I'd be happy to chat.

Sincerely,

{YourName}
{YourPhone}
{YourEmail}
"""

# ------------------------
# Utilities
# ------------------------

def to_title_case(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).title()

def extract_street_name(address: str) -> str:
    if not address:
        return ""
    parts = address.split(",")
    return parts[0].strip()

# Entity tokens for display cleaning when it's a person (not trusts)
ENTITY_TOKENS = [
    " LLC"," L.L.C."," INC"," INC."," CORP"," CORPORATION",
    " LP"," L.P."," LLP"," L.L.P."," TRUST"," TR"," HOLDINGS"," FUND"
]

def clean_entity_tokens(name: str) -> str:
    if not name:
        return ""
    s = " " + name.upper()
    for tok in ENTITY_TOKENS:
        s = s.replace(tok, "")
    return re.sub(r"\s+", " ", s.strip())

def find_column(headers: List[str], candidates: List[str]) -> Optional[str]:
    hlow = {h.lower(): h for h in headers}
    for c in candidates:
        if c in hlow:
            return hlow[c]
    return None

def _first_nonempty_from_row(row, keys):
    for k in keys:
        if k in row and str(row[k]).strip():
            return str(row[k]).strip()
    return ""

def is_trust_case(row) -> bool:
    """
    True when Primary First/Last and all Secondary* fields are empty AND Primary Name exists.
    In those cases we DO NOT say "Dear ...". We start the letter with "Primary Name,".
    """
    primary_first = _first_nonempty_from_row(row, [
        "Primary First","PRIMARY FIRST","primary first","primary_first","primary first name","primary_first_name"
    ])
    primary_last  = _first_nonempty_from_row(row, [
        "Primary Last","PRIMARY LAST","primary last","primary_last","primary last name","primary_last_name"
    ])
    secondary_name  = _first_nonempty_from_row(row, ["Secondary Name","SECONDARY NAME","secondary name","secondary_name"])
    secondary_first = _first_nonempty_from_row(row, ["Secondary First","SECONDARY FIRST","secondary first","secondary_first"])
    secondary_last  = _first_nonempty_from_row(row, ["Secondary Last","SECONDARY LAST","secondary last","secondary_last"])
    primary_name    = _first_nonempty_from_row(row, ["Primary Name","PRIMARY NAME","primary name","primary_name"])

    return (not (primary_first or primary_last or secondary_name or secondary_first or secondary_last)) and bool(primary_name)

def split_owner_first(owner_first: Optional[str], owner_name: Optional[str]) -> str:
    """
    Robust first-name token for person-to-person greeting.
    - If a first-name field exists, use its first token.
    - Else, if owner_name is "Last, First ..." use First (robust to missing parts).
    - Else, first token of owner_name (after stripping entity tokens).
    Returns empty string if nothing usable (we will then use full name salutation).
    """
    try:
        if owner_first and str(owner_first).strip():
            return str(owner_first).strip().split()[0]

        cleaned = clean_entity_tokens((owner_name or "").strip())
        if not cleaned:
            return ""

        if "," in cleaned:
            parts = [p.strip() for p in cleaned.split(",") if p and p.strip()]
            if len(parts) >= 2:
                tokens = parts[1].split()
                if tokens:
                    return tokens[0]

        tokens = cleaned.split()
        return tokens[0] if tokens else ""
    except Exception:
        return ""

# ------------------------
# Letter Generation
# ------------------------

def personalize_letter(row: Dict[str, str], your_name: str, your_phone: str, your_email: str, template_text: str) -> Tuple[str, str, str, str]:
    headers = list(row.keys())
    col_first = find_column(headers, POSSIBLE_OWNER_FIRST)
    col_name  = find_column(headers, POSSIBLE_OWNER_NAME)
    col_addr  = find_column(headers, POSSIBLE_ADDRESS)

    # First name for "Dear ..." when we have a person
    owner_first_raw = split_owner_first(row.get(col_first or "", ""), row.get(col_name or "", ""))
    owner_first = to_title_case(owner_first_raw).strip()

    # Determine display/full name
    trust = is_trust_case(row)
    if trust:
        primary_name = _first_nonempty_from_row(row, ["Primary Name","PRIMARY NAME","primary name","primary_name"])
        owner_full_raw = primary_name
        owner_display = to_title_case(primary_name) if primary_name else ""
    else:
        owner_full_raw = row.get(col_name or "", "") or ""
        # For people, remove entity tokens for display
        owner_display = to_title_case(clean_entity_tokens(owner_full_raw)) if owner_full_raw else ""

    address = (row.get(col_addr or "", "") or "").strip()
    street = to_title_case(extract_street_name(address)) or "your street"

    # Build letter content
    if trust:
        # Replace the greeting with "Primary Name,"
        adjusted = re.sub(r"^Dear\s*\{OwnerFirstName\},\s*\n+", "{SalutationLine}\n\n", template_text, flags=re.M)
        content = adjusted.format(
            SalutationLine=f"{owner_full_raw},",
            OwnerFirstName=owner_first or owner_display,  # still available if used elsewhere in body
            StreetName=street,
            YourName=your_name,
            YourPhone=your_phone,
            YourEmail=your_email
        )
    elif owner_first:
        # Normal person greeting
        content = template_text.format(
            OwnerFirstName=owner_first,
            StreetName=street,
            YourName=your_name,
            YourPhone=your_phone,
            YourEmail=your_email
        )
    else:
        # No first-name available → use full owner name as salutation (no "Dear")
        display = owner_display or owner_full_raw
        adjusted = re.sub(r"^Dear\s*\{OwnerFirstName\},\s*\n+", "{SalutationLine}\n\n", template_text, flags=re.M)
        content = adjusted.format(
            SalutationLine=f"{display},",
            OwnerFirstName="",  # unused in this path
            StreetName=street,
            YourName=your_name,
            YourPhone=your_phone,
            YourEmail=your_email
        )

    # File stub prefers a first name if present; otherwise use display/full
    stub_name = owner_first or (owner_display or owner_full_raw or "Owner")
    filestub = f"{stub_name.replace(' ', '_')}_{street.replace(' ', '_')}".replace("/", "_")

    # Return owner_display for mapping (what users see on callback)
    return content, filestub, (owner_display or owner_full_raw), address

def letter_story(content: str, sig_image: Optional[str]) -> List:
    story = []
    lines = content.splitlines()
    if lines:
        greet = lines[0].strip()
        story.append(Paragraph(greet, greeting_style))
        story.append(Spacer(1, 6))
        body_text = "\n".join(lines[1:])
    else:
        body_text = content

    for para in body_text.split("\n\n"):
        p = para.strip()
        if p:
            story.append(Paragraph(p, body_style))
            story.append(Spacer(1, 10))

    story.append(Spacer(1, 12))
    story.append(Paragraph("Sincerely,", sig_lead))
    story.append(Spacer(1, 8))

    if sig_image and os.path.isfile(sig_image):
        try:
            ir = ImageReader(sig_image)
            iw, ih = ir.getSize()
            max_w, max_h = 1.8 * inch, 0.45 * inch
            scale = min(max_w / float(iw), max_h / float(ih), 1.0)
            img = RLImage(sig_image, iw * scale, ih * scale)
            img.hAlign = "LEFT"
            story.append(Spacer(1, 2))
            story.append(img)
            story.append(Spacer(1, 3))
        except Exception:
            pass

    return story

# ------------------------
# PDF Writing and Mapping
# ------------------------

def write_combined_pdf(rows: List[Dict[str, str]], out_pdf: str, your_name: str, your_phone: str, your_email: str, template_text: str, sig_image: Optional[str]) -> Tuple[int, List[Dict[str, str]]]:
    mapping = []
    doc = canvas.Canvas(out_pdf, pagesize=LETTER)
    width, height = LETTER

    for idx, row in enumerate(rows, start=1):
        content, filestub, owner_display, prop_address = personalize_letter(row, your_name, your_phone, your_email, template_text)
        story = letter_story(content, sig_image)

        # Render story to canvas
        y = height - 1.25 * inch
        for flow in story:
            if isinstance(flow, Paragraph):
                w, h = flow.wrap(6.5 * inch, y)
                if y - h < 0.5 * inch:
                    doc.showPage()
                    y = height - 1.25 * inch
                    w, h = flow.wrap(6.5 * inch, y)
                flow.drawOn(doc, 1.0 * inch, y - h)
                y -= (h + 6)
            elif isinstance(flow, Spacer):
                y -= flow.height
            elif isinstance(flow, RLImage):
                w = flow.drawWidth
                h = flow.drawHeight
                if y - h < 0.5 * inch:
                    doc.showPage()
                    y = height - 1.25 * inch
                flow.drawOn(doc, 1.0 * inch, y - h)
                y -= (h + 6)

        doc.showPage()

        mapping.append({
            "page": idx,
            "owner": owner_display,
            "property_address": prop_address,
        })

        if idx % 100 == 0:
            print(f"[MAP] {idx} pages...")

    doc.save()
    return len(rows), mapping

def write_mapping_csv(mapping: List[Dict[str, str]], out_csv: str):
    os.makedirs(os.path.dirname(out_csv) or ".", exist_ok=True)
    with open(out_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["page","owner","property_address"])
        w.writeheader()
        w.writerows(mapping)

# ------------------------
# CLI
# ------------------------

def main():
    ap = argparse.ArgumentParser(description="Generate direct mail letters (combined PDF and optional singles) and a mapping CSV.")
    ap.add_argument("--csv", required=True, help="Input CSV (campaign_master.csv)")
    ap.add_argument("--outdir", default="Singles", help="Folder for optional single-letter PDFs")
    ap.add_argument("--combine-out", default="letters_batch.pdf", help="Combined PDF output filename")
    ap.add_argument("--map-out", default="letters_mapping.csv", help="Mapping CSV output filename")
    ap.add_argument("--template-id", type=int, default=101, help="Template ID (for future use)")
    ap.add_argument("--signature-image", default="", help="Optional signature image path")
    ap.add_argument("--skip-singles", action="store_true", help="Skip creating single PDFs")
    ap.add_argument("--name", required=True)
    ap.add_argument("--phone", required=True)
    ap.add_argument("--email", required=True)
    args = ap.parse_args()

    in_csv = args.csv
    outdir = args.outdir
    combine_pdf = args.combine_out
    map_csv = args.map_out
    sig_image = args.signature_image.strip() or ""

    your_name = args.name
    your_phone = args.phone
    your_email = args.email

    if not os.path.isfile(in_csv):
        print(f"[ERROR] CSV not found: {in_csv}")
        sys.exit(1)

    with open(in_csv, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        rows = [{k: (v or "").strip() for k, v in row.items()} for row in r]

    os.makedirs(os.path.dirname(combine_pdf) or ".", exist_ok=True)
    count, mapping = write_combined_pdf(rows, combine_pdf, your_name, your_phone, your_email, LETTER_TEMPLATE, sig_image)
    print(f"[OK] Combined PDF: {combine_pdf}  (pages={count})")

    if not args.skip_singles:
        os.makedirs(outdir, exist_ok=True)
        for i, row in enumerate(rows, start=1):
            content, filestub, owner_display, prop_address = personalize_letter(row, your_name, your_phone, your_email, LETTER_TEMPLATE)
            single_path = os.path.join(outdir, f"{filestub}.pdf")
            try:
                doc = SimpleDocTemplate(
                    single_path, pagesize=LETTER,
                    leftMargin=1.0 * inch, rightMargin=1.0 * inch,
                    topMargin=1.25 * inch, bottomMargin=0.75 * inch
                )
                story = letter_story(content, sig_image)
                doc.build(story)
                print(f"[SAVE] single PDF for {filestub} (page {i})")
            except Exception as e:
                print(f"[SKIP] single PDF for {filestub} (error: {e})")

    write_mapping_csv(mapping, map_csv)
    print(f"[OK] Mapping CSV: {map_csv}")

if __name__ == "__main__":
    main()
