
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Direct Mail Campaign - Batch PDF Builder (no printing)
- Reads a CSV and produces ONE multi-page PDF.
- Each row becomes a neatly formatted letter page.
- Fixed, stacked signature block:
      Sincerely,
      Ed Beluli
      916-905-7281
      ed.beluli@gmail.com
- Clean layout using ReportLab Platypus:
      1" margins, 12pt Times, 15pt leading, paragraph spacing,
      letter body starts ~2" below top margin.
- Footer on every page: "<Batch> • i of N"

Usage:
  python direct_mail_batch_pdf.py --csv "ShortTest.csv" --out "letters_batch.pdf" --batch-name "Batch 2025-08-21"
"""

import os, re, csv, argparse, datetime
from typing import List, Dict, Optional

# PDF deps
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import LETTER as RL_LETTER
from reportlab.lib.units import inch

# ----------- Letter template (NO signature here) -----------
LETTER_TEMPLATE = """Dear {OwnerFirstName},

I hope this note finds you well. My name is {YourName}, and I’m reaching out because I’m interested in purchasing your property on {StreetName}.

I buy homes directly from owners like you, which means you can avoid paying the typical 5–6% realtor fees and the months of uncertainty that often come with listing a property. I pay cash and can close quickly, so you won’t have to deal with inspections, repairs, or buyers backing out at the last minute.

What makes me different from others who send letters like this is that I’m local and personally handle each purchase. I don’t make “low-ball” offers just to flip the contract—I actually buy and hold properties myself. That means I can be flexible, fair, and work around your timing and needs.

If selling your property on {StreetName} is something you’d consider, I’d love the opportunity to talk with you. There’s no obligation at all—I’m simply interested in seeing if we can find a solution that works well for both of us.

You can reach me directly at {YourPhone}. I look forward to hearing from you.
"""

# ----------- CSV column helpers -----------
POSSIBLE_OWNER_FIRST = ["owner first name","owner_first_name","first name","firstname","ownerfirst","owner first","ownerfirstname"]
POSSIBLE_OWNER_NAME  = ["owner name","owner","name","owner_full_name","owner full name","owner(s)","owner 1","owner1"]
POSSIBLE_ADDRESS     = ["situs address","mailing address","property address","site address","address","situsaddr","situs","situsaddress","property situs","prop address","situs_address"]

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").strip().lower())

def find_column(headers: List[str], candidates: List[str]) -> Optional[str]:
    norm_map = {h: _norm(h) for h in headers}
    # exact normalized match
    for cand in candidates:
        nc = _norm(cand)
        for h, n in norm_map.items():
            if n == nc:
                return h
    # substring contains
    for cand in candidates:
        nc = _norm(cand)
        for h, n in norm_map.items():
            if nc in n:
                return h
    return None

LLC_TOKENS = r"(ET\s+AL|TRUST|LLC|INC|CO|LP|L\.P\.|LTD)"
def clean_entity_tokens(name: str) -> str:
    return re.sub(r"\b" + LLC_TOKENS + r"\b\.?,?", "", name, flags=re.I).strip()

def to_title_case(s: str) -> str:
    if not s:
        return s
    s = s.strip().lower()
    out = []
    for w in s.split():
        m = re.match(r"^(\d+)(st|nd|rd|th)$", w)
        if m:
            out.append(m.group(1)+m.group(2))
        else:
            out.append(w.capitalize())
    return " ".join(out)

def split_owner_first(owner_first: Optional[str], owner_name: Optional[str]) -> str:
    if owner_first and owner_first.strip():
        return owner_first.strip().split()[0]
    if owner_name and owner_name.strip():
        cleaned = clean_entity_tokens(owner_name)
        if "," in cleaned:
            parts = [p.strip() for p in cleaned.split(",")]
            first_guess = parts[1].split()[0] if len(parts) >= 2 else cleaned.split()[0]
        else:
            first_guess = cleaned.split()[0]
        return first_guess
    return "Neighbor"

STREET_TYPE_WORDS = {"ave","avenue","blvd","boulevard","cir","circle","ct","court","dr","drive","hwy","highway",
                     "ln","lane","pkwy","parkway","pl","place","rd","road","st","street","ter","terrace","way",
                     "trl","trail","sq","square","cv","cove","pt","point"}

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
            end_idx = idx; break
    core = " ".join(street_tokens[:end_idx+1]) if end_idx is not None else " ".join(street_tokens)
    return core.strip() or "your street"

# ----------- PDF building -----------
def build_batch_pdf(csv_path: str, out_pdf: str, your_name: str, your_phone: str, your_email: str, batch_name: Optional[str] = None):
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    total = len(rows)
    if total == 0:
        raise SystemExit("No rows found in CSV.")

    os.makedirs(os.path.dirname(out_pdf) or ".", exist_ok=True)

    doc = SimpleDocTemplate(
        out_pdf,
        pagesize=RL_LETTER,
        leftMargin=1*inch, rightMargin=1*inch,
        topMargin=1*inch, bottomMargin=1*inch
    )

    styles = getSampleStyleSheet()
    body = ParagraphStyle("Body", parent=styles["Normal"],
                          fontName="Times-Roman", fontSize=12,
                          leading=15, spaceAfter=12)
    sig = ParagraphStyle("Sig", parent=body, spaceBefore=18)

    story = []

    headers = rows[0].keys()
    col_first = find_column(list(headers), POSSIBLE_OWNER_FIRST)
    col_name  = find_column(list(headers), POSSIBLE_OWNER_NAME)
    col_addr  = find_column(list(headers), POSSIBLE_ADDRESS)

    batch_footer = batch_name or f"Batch {datetime.date.today().isoformat()}"

    def footer_fn(canvas, doc):
        canvas.saveState()
        canvas.setFont("Times-Italic", 9)
        txt = f"{batch_footer} • {canvas.getPageNumber()} of {total}"
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, doc.bottomMargin - 12, txt)
        canvas.restoreState()

    for idx, row in enumerate(rows, start=1):
        owner_first_raw = split_owner_first(row.get(col_first or "", ""), row.get(col_name or "", ""))
        owner_first = to_title_case(owner_first_raw) or "Neighbor"
        address = (row.get(col_addr or "", "") or "").strip()
        street = to_title_case(extract_street_name(address)) or "your street"

        # personalize
        content = LETTER_TEMPLATE.format(
            OwnerFirstName=owner_first,
            StreetName=street,
            YourName=your_name,
            YourPhone=your_phone,
            YourEmail=your_email
        )

        # top spacer to push down
        story.append(Spacer(1, 2*inch))

        # body paragraphs
        for para in [p for p in content.strip().split("\n\n") if p.strip()]:
            story.append(Paragraph(para.replace("\n"," ").strip(), body))

        # fixed stacked signature
        story.append(Spacer(1, 18))
        for line in ["Sincerely,", "Ed Beluli", "916-905-7281", "ed.beluli@gmail.com"]:
            story.append(Paragraph(line, sig))

        # page break between recipients (not after the last page; SimpleDocTemplate will ignore trailing PageBreak)
        if idx < total:
            story.append(PageBreak())

    doc.build(story, onFirstPage=footer_fn, onLaterPages=footer_fn)

def main():
    ap = argparse.ArgumentParser(description="Build one multi-page PDF (one letter per CSV row). No printing.")
    ap.add_argument("--csv", required=True, help="Input CSV")
    ap.add_argument("--out", required=True, help="Output PDF path, e.g., letters_batch.pdf")
    ap.add_argument("--name", default="Ed Beluli", help="Your display name (used in body text)")
    ap.add_argument("--phone", default="916-905-7281", help="Your phone (used in body text)")
    ap.add_argument("--email", default="ed.beluli@gmail.com", help="Your email (used in body text)")
    ap.add_argument("--batch-name", default=None, help="Footer label; defaults to 'Batch YYYY-MM-DD'")
    args = ap.parse_args()

    build_batch_pdf(args.csv, args.out, args.name, args.phone, args.email, args.batch_name)
    print(f"[DONE] Wrote: {args.out}")

if __name__ == "__main__":
    main()
