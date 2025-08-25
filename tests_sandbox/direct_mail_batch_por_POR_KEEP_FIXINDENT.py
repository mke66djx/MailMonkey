#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Direct Mail - Singles + Combined PDF + Mapping CSV + Template Selection (by ID) + Optional Signature Image

POR-KEEP VERSION
- Preserves POR templates and formatting.
- Combined PDF defaults to BatchLetterFiles (via --batch-dir).
- Mapping CSV defaults to RefFiles.
- --skip-singles respected so Singles folder isn't created when skipping.
- Salutation rule:
    * If Primary First/Last and all Secondary* are empty AND Primary Name exists,
      start with "Primary Name," (no "Dear").
    * Otherwise use "Dear {OwnerFirstName}," where first name is robustly extracted.
    * No "Neighbor" fallback; if no first name (and not trust case), use full owner name as salutation line (no "Dear").
"""

import os
import re
import csv
import argparse
import random
from typing import Dict, List, Optional, Tuple

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import LETTER as RL_LETTER
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

# ---------------- Built-in templates (POR) ----------------

ORIGINAL_TEMPLATE = """Dear {OwnerFirstName},

I hope this note finds you well. My name is {YourName}. My dad, Albert, and I live in Rancho Cordova and buy/restore homes around Sacramento. We're reaching out because we're interested in purchasing your property on {StreetName}.

Because we purchase directly, you can avoid paying the typical 5–6% realtor fees and the months of uncertainty that often come with listing a property. We pay cash and can close on your timeline, so you won’t have to deal with inspections, repairs, or buyers backing out at the last minute.

We’re also local and personally handle each purchase. We are not wholesalers and don’t make “low-ball” offers just to flip the contract—we actually buy and hold properties ourselves. That means we can be flexible, fair, and work around your timing and needs.

If selling your property on {StreetName} is something you’d consider, we’d appreciate the opportunity to get in touch. There’s no obligation—just seeing if it makes sense for both of us. We can text/email proof of funds and our local title company contact info for your verification.

You can reach me directly at {YourPhone}. I look forward to hearing from you.
"""

RANCHO_LOCAL_CASH_TEMPLATE = """Dear {OwnerFirstName},

My dad and I live in Rancho Cordova and buy/restore homes around Sacramento (we are not agents). If selling your place on {StreetName} is on your mind, we buy with cash (no financing or appraisals) and take it as-is: no repairs, no cleaning, and no showings or open houses. There is no 5-6% realtor commission. We usually cover standard seller closing costs, you pick the move-out date (soon or later), and you can leave behind what you do not want.

Unlike buyers who tie up a property and then assign it to another investor, we buy for ourselves and hold the property. That lets us be flexible on timing and terms and keep the process simple.

If a simple, private sale would help, call or text me at {YourPhone}. I can text proof of funds and our title company contact so you know we are real. Thanks for reading.
"""

BUILTINS: Dict[str, Dict[str, str]] = {
    "orig": {"ref": "101", "text": ORIGINAL_TEMPLATE},
    "rc":   {"ref": "202", "text": RANCHO_LOCAL_CASH_TEMPLATE},
    "101":  {"ref": "101", "text": ORIGINAL_TEMPLATE},
    "202":  {"ref": "202", "text": RANCHO_LOCAL_CASH_TEMPLATE},
}

# ---------------- Template selection ----------------

def load_template_by_id(template_id: str, templates_dir: Optional[str]) -> Tuple[str, str, str]:
    """
    Returns (template_text, template_ref, template_source)
      - template_ref is a string, typically the same as template_id if file-based.
      - template_source is 'file' or 'builtin' (or 'fallback').
    Selection order:
      1) If templates_dir and <id>.txt exists -> load file, ref = id, source='file'
      2) Else if id in BUILTINS -> use builtin, ref = BUILTINS[id]['ref'], source='builtin'
      3) Else fallback to builtin '202' (rc), ref='202', source='fallback'
    """
    tid = str(template_id).strip()
    if templates_dir:
        candidate = os.path.join(templates_dir, f"{tid}.txt")
        if os.path.isfile(candidate):
            with open(candidate, "r", encoding="utf-8") as f:
                return f.read(), tid, "file"
    if tid in BUILTINS:
        b = BUILTINS[tid]
        return b["text"], b["ref"], "builtin"
    b = BUILTINS["202"]
    return b["text"], b["ref"], "fallback"

# ---------------- Helpers ----------------

def generate_ref_code() -> str:
    # Starts with a letter to feel less like a serial number
    return f"L{random.randint(10000, 99999)}"

# Column candidates (include Primary variants)
POSSIBLE_OWNER_FIRST = [
    "owner first name","owner_first_name","first name","firstname","ownerfirst",
    "owner first","owner_first","ownerfirst name",
    "primary first","primary first name","primary_first","primary_first_name"
]
POSSIBLE_OWNER_NAME  = [
    "primary name","primary_name","owner name","owner","name",
    "owner_full_name","owner full name","owner(s)","owner 1","owner1"
]
POSSIBLE_ADDRESS     = [
    "situs address","mailing address","property address","site address","address",
    "situsaddr","situs","situsaddress","property situs","prop address","situs_address"
]

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").strip().lower())

def find_column(headers: List[str], candidates: List[str]) -> Optional[str]:
    norm_map = {h: _norm(h) for h in headers}
    # exact
    for cand in candidates:
        nc = _norm(cand)
        for h, n in norm_map.items():
            if n == nc:
                return h
    # contains
    for cand in candidates:
        nc = _norm(cand)
        for h, n in norm_map.items():
            if nc in n:
                return h
    return None

LLC_TOKENS = r"(ET\s+AL|TRUST|LLC|INC|CO|LP|L\.P\.|LTD)"
def clean_entity_tokens(name: str) -> str:
    return re.sub(r"\b" + LLC_TOKENS + r"\b\.?,?", "", name or "", flags=re.I).strip()

def to_title_case(s: str) -> str:
    if not s:
        return s
    s = s.strip().lower()
    out = []
    m_ord = re.compile(r"^(\d+)(st|nd|rd|th)$")
    for w in s.split():
        m = m_ord.match(w)
        out.append(m.group(1)+m.group(2) if m else w.capitalize())
    return " ".join(out)

def split_owner_first(owner_first: Optional[str], owner_name: Optional[str]) -> str:
    """
    Robust first-name extraction for person-to-person greeting.
    Returns "" if nothing usable (we'll handle the salutation without 'Dear').
    """
    try:
        if owner_first and owner_first.strip():
            return owner_first.strip().split()[0]
        if owner_name and owner_name.strip():
            cleaned = clean_entity_tokens(owner_name)
            if "," in cleaned:
                parts = [p.strip() for p in cleaned.split(",") if p.strip()]
                if len(parts) >= 2:
                    tokens = parts[1].split()
                    if tokens:
                        return tokens[0]
            tokens = cleaned.split()
            return tokens[0] if tokens else ""
        return ""
    except Exception:
        return ""

STREET_TYPE_WORDS = {
    "ave","avenue","blvd","boulevard","cir","circle","ct","court","dr","drive","hwy","highway",
    "ln","lane","pkwy","parkway","pl","place","rd","road","st","street","ter","terrace","way",
    "trl","trail","sq","square"
}

def extract_street_name(full_address: str) -> str:
    if not full_address:
        return "your street"
    s = full_address.strip()
    first_seg = s.split(",")[0]
    first_seg = re.sub(r"\b(apt|unit|#)\s*\w+", "", first_seg, flags=re.I).strip()
    tokens = first_seg.split()
    if not tokens:
        return "your street"
    i = 0
    while i < len(tokens) and re.match(r"^\d+[a-zA-Z]?$", tokens[i]):
        i += 1
    street_tokens = tokens[i:] or tokens
    lower = [t.lower().strip(".") for t in street_tokens]
    end_idx = None
    for idx, tok in enumerate(lower):
        if tok in STREET_TYPE_WORDS:
            end_idx = idx
            break
    core = " ".join(street_tokens[:end_idx+1]) if end_idx is not None else " ".join(street_tokens)
    return core.strip() or "your street"

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

def personalize_letter(row: Dict[str, str], your_name: str, your_phone: str, your_email: str, template_text: str) -> Tuple[str, str, str, str]:
    headers = list(row.keys())
    col_first = find_column(headers, POSSIBLE_OWNER_FIRST)
    col_name  = find_column(headers, POSSIBLE_OWNER_NAME)
    col_addr  = find_column(headers, POSSIBLE_ADDRESS)

    # Determine first-name for "Dear ..."
    owner_first_raw = split_owner_first(row.get(col_first or "", ""), row.get(col_name or "", ""))
    owner_first = to_title_case(owner_first_raw).strip()

    # Display / full
    trust = is_trust_case(row)
    if trust:
        primary_name = _first_nonempty_from_row(row, ["Primary Name","PRIMARY NAME","primary name","primary_name"])
        owner_full_raw = primary_name
        owner_display = to_title_case(primary_name) if primary_name else ""
    else:
        owner_full_raw = row.get(col_name or "", "") or ""
        owner_display = to_title_case(clean_entity_tokens(owner_full_raw)) if owner_full_raw else ""

    address = (row.get(col_addr or "", "") or "").strip()
    street = to_title_case(extract_street_name(address)) or "your street"

    # Compose content based on salutation rule
    if trust:
        adjusted = re.sub(r"^Dear\s*\{OwnerFirstName\},\s*\n+", "{SalutationLine}\n\n", template_text, flags=re.M)
        content = adjusted.format(
            SalutationLine=f"{owner_full_raw},",
            OwnerFirstName=owner_first or owner_display,  # keep available if used elsewhere
            StreetName=street,
            YourName=your_name,
            YourPhone=your_phone,
            YourEmail=your_email
        )
    elif owner_first:
        content = template_text.format(
            OwnerFirstName=owner_first,
            StreetName=street,
            YourName=your_name,
            YourPhone=your_phone,
            YourEmail=your_email
        )
    else:
        display = owner_display or owner_full_raw or "Owner"
        adjusted = re.sub(r"^Dear\s*\{OwnerFirstName\},\s*\n+", "{SalutationLine}\n\n", template_text, flags=re.M)
        content = adjusted.format(
            SalutationLine=f"{display},",
            OwnerFirstName="",  # not used in this path
            StreetName=street,
            YourName=your_name,
            YourPhone=your_phone,
            YourEmail=your_email
        )

    stub_name = owner_first or (owner_display or owner_full_raw or "Owner")
    filestub = f"{stub_name.replace(' ', '_')}_{street.replace(' ', '_')}".replace("/", "_")
    return content, filestub, (owner_display or owner_full_raw), address

# ---------------- Layout (POR formatting) ----------------

def build_story(content: str, sig_image: Optional[str] = None):
    styles = getSampleStyleSheet()

    # Body style
    body = ParagraphStyle(
        "Body", parent=styles["Normal"],
        fontName="Times-Roman", fontSize=12,
        leading=15, spaceAfter=12
    )

    # Signature styles (tight)
    sig_lead = ParagraphStyle(   # "Sincerely,"
        "SigLead", parent=styles["Normal"],
        fontName="Times-Roman", fontSize=12,
        leading=15, spaceBefore=18, spaceAfter=4
    )
    sig_line = ParagraphStyle(   # contact lines
        "SigLine", parent=styles["Normal"],
        fontName="Times-Roman", fontSize=12,
        leading=14, spaceBefore=0, spaceAfter=2
    )

    story: List = []
    # Push content down ~2 inches (POR look)
    story.append(Spacer(1, 2 * inch))

    # Body paragraphs
    for para in [p for p in content.strip().split("\n\n") if p.strip()]:
        story.append(Paragraph(para.replace("\n", " ").strip(), body))

    # Signature block
    story.append(Paragraph("Sincerely,", sig_lead))

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

    for line in ["Ed & Albert Beluli", "916-905-7281", "eabeluli@gmail.com"]:
        story.append(Paragraph(line, sig_line))

    return story

def footer_fn_factory(ref_code: str):
    def footer_fn(canvas, doc):
        canvas.saveState()
        canvas.setFont("Times-Italic", 9)
        canvas.drawRightString(
            doc.pagesize[0] - doc.rightMargin,
            doc.bottomMargin - 12,
            f"Ref: {ref_code}"
        )
        canvas.restoreState()
    return footer_fn

def write_single_letter_pdf(out_dir: str, filestub: str, content: str, ref_code: str,
                            csv_base: str, total_rows: int, template_ref: str,
                            sig_image_path: Optional[str] = None) -> str:
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{filestub}.pdf")
    doc = SimpleDocTemplate(
        path, pagesize=RL_LETTER,
        leftMargin=1 * inch, rightMargin=1 * inch,
        topMargin=1 * inch, bottomMargin=1 * inch
    )
    story = build_story(content, sig_image_path)
    footer = footer_fn_factory(ref_code)

    def _meta_first_page(canvas, doc):
        canvas.setTitle(f"{csv_base} - {total_rows} rows - T{template_ref}")
        canvas.setSubject(f"Source CSV: {csv_base}; Rows: {total_rows}; Template: T{template_ref}")
        footer(canvas, doc)

    doc.build(story, onFirstPage=_meta_first_page, onLaterPages=footer)
    return path

def write_combined_pdf(out_path: str, contents_with_refs: List[Tuple[str, str, str]],
                       csv_base: str, total_rows: int, template_ref: str,
                       sig_image_path: Optional[str] = None) -> None:
    """contents_with_refs: list of (content, filestub, ref_code). One page per item."""
    doc = SimpleDocTemplate(
        out_path, pagesize=RL_LETTER,
        leftMargin=1 * inch, rightMargin=1 * inch,
        topMargin=1 * inch, bottomMargin=1 * inch
    )
    total = len(contents_with_refs)
    story_all: List = []
    ref_codes: List[str] = []

    for i, (content, filestub, ref_code) in enumerate(contents_with_refs, start=1):
        story_all.extend(build_story(content, sig_image_path))
        ref_codes.append(ref_code)
        if i < total:
            story_all.append(PageBreak())

    def footer_fn(canvas, doc):
        pg = int(canvas.getPageNumber())
        code = ref_codes[pg - 1] if 1 <= pg <= len(ref_codes) else "L00000"
        canvas.setTitle(f"{csv_base} - {total_rows} rows - T{template_ref}")
        canvas.setSubject(f"Source CSV: {csv_base}; Rows: {total_rows}; Template: T{template_ref}")
        canvas.saveState()
        canvas.setFont("Times-Italic", 9)
        canvas.drawRightString(
            doc.pagesize[0] - doc.rightMargin,
            doc.bottomMargin - 12,
            f"Ref: {code}"
        )
        canvas.restoreState()

    doc.build(story_all, onFirstPage=footer_fn, onLaterPages=footer_fn)

# ---------------- Main ----------------

def main():
    ap = argparse.ArgumentParser(description="POR mailer: singles + combined PDF + mapping (template by ID)")
    ap.add_argument("--csv", required=True, help="Input CSV with recipient info")
    ap.add_argument("--outdir", required=True, help="Folder for individual PDFs")
    ap.add_argument("--combine-out", required=True, help="Output filename for the combined PDF (e.g., letters_batch.pdf)")
    ap.add_argument("--map-out", default=None, help="(Optional) Output filename for the mapping CSV; default uses <csv>_<N>_T<id>_refs.csv")
    ap.add_argument("--template-id", required=True, help="Template ID (e.g., 101 or 202)")
    ap.add_argument("--templates-dir", default=None, help="Optional folder containing <template-id>.txt files")
    ap.add_argument("--sig-image", default=None, help="Optional path to a signature image (PNG/JPG)")
    ap.add_argument("--name", required=True)
    ap.add_argument("--phone", required=True)
    ap.add_argument("--email", required=True)
    ap.add_argument("--batch-dir", default="BatchLetterFiles", help="Folder for combined batch PDFs (POR default)")
    ap.add_argument("--refs-dir", default="RefFiles", help="Folder for mapping CSVs (POR default)")
    ap.add_argument("--skip-singles", action="store_true", help="Only create combined PDF + mapping CSV (no individual PDFs)")
    args = ap.parse_args()

    # Read rows
    with open(args.csv, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        print("[INFO] No rows found in CSV.")
        return

    total = len(rows)
    csv_base = os.path.splitext(os.path.basename(args.csv))[0]

    # Ensure output directories exist
    os.makedirs(args.batch_dir, exist_ok=True)
    os.makedirs(args.refs_dir, exist_ok=True)

    # Select template
    template_text, template_ref, template_source = load_template_by_id(args.template_id, args.templates_dir)

    # Auto-name combined PDF if using generic
    default_comb_name = f"{csv_base}_{total}_T{template_ref}_letters.pdf"
    comb_base_name = os.path.basename(args.combine_out)
    if comb_base_name.lower() in ("letters_batch.pdf", "combined.pdf"):
        comb_base_name = default_comb_name
    combine_pdf_path = os.path.join(args.batch_dir, comb_base_name)

    # Mapping filename
    if args.map_out:
        map_base_name = os.path.basename(args.map_out)
    else:
        map_base_name = f"{csv_base}_{total}_T{template_ref}_refs.csv"
    map_out_path = os.path.join(args.refs_dir, map_base_name)

    # Singles may be skipped
    if not args.skip_singles:
        os.makedirs(args.outdir, exist_ok=True)

    contents: List[Tuple[str, str, str]] = []
    map_rows: List[Dict[str, str]] = []

    for i, row in enumerate(rows, start=1):
        content, filestub, owner_display, prop_address = personalize_letter(
            row, args.name, args.phone, args.email, template_text
        )
        ref_code = generate_ref_code()
        contents.append((content, filestub, ref_code))

        single_path = ""
        if not args.skip_singles:
            single_path = write_single_letter_pdf(
                args.outdir, filestub, content, ref_code,
                csv_base, total, template_ref, args.sig_image
            )
            print(f"[SAVE] {i}/{total}: {single_path}  (Ref: {ref_code})")
        else:
            print(f"[SKIP] single PDF for {filestub} (Ref: {ref_code})")

        map_rows.append({
            "page": i,
            "owner": owner_display or "",
            "property_address": prop_address or "",
            "ref_code": ref_code,
            "file_stub": filestub,
            "single_pdf": os.path.abspath(single_path) if single_path else "",
            "template_ref": template_ref,
            "template_source": template_source,
        })

    # Combined PDF
    write_combined_pdf(combine_pdf_path, contents, csv_base, total, template_ref, args.sig_image)
    print(f"[COMBINED] {combine_pdf_path}  (pages: {total})  [Template T{template_ref} - {template_source}]")

    # Mapping CSV
    with open(map_out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "page", "owner", "property_address", "ref_code",
                "file_stub", "single_pdf", "template_ref", "template_source"
            ]
        )
        writer.writeheader()
        writer.writerows(map_rows)
    print(f"[MAP] {map_out_path}  (rows: {len(map_rows)})  [Template T{template_ref} - {template_source}]")

if __name__ == "__main__":
    main()
