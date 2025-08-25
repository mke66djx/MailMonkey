
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Direct Mail - Build Singles + One Combined PDF (clean layout)

What it does
- Reads a CSV of recipients.
- Generates *individual* PDFs (one per row) into an output folder.
- Also builds ONE combined multi-page PDF (one page per row) you can print/preview.
- Professional layout using ReportLab Platypus.
- Fixed signature block (Sincerely, Ed Beluli, 916-905-7281, ed.beluli@gmail.com).

Usage (example)
  python direct_mail_batch_and_singles.py \
    --csv "ShortTest.csv" \
    --outdir "output_letters" \
    --combine-out "letters_batch.pdf" \
    --name "Ed Beluli" \
    --phone "916-905-7281" \
    --email "ed.beluli@gmail.com" \
    --batch-name "Batch 2025-08-21"

Dependencies
  pip install reportlab
"""

import os, re, csv, argparse, datetime
from typing import Dict, List, Optional, Tuple

# ReportLab imports
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import LETTER as RL_LETTER
from reportlab.lib.units import inch

# ---------------- Template & helpers ----------------

LETTER_TEMPLATE = """Dear {OwnerFirstName},

I hope this note finds you well. My name is {YourName}, and I’m reaching out because I’m interested in purchasing your property on {StreetName}.

I buy homes directly from owners like you, which means you can avoid paying the typical 5–6% realtor fees and the months of uncertainty that often come with listing a property. I pay cash and can close quickly, so you won’t have to deal with inspections, repairs, or buyers backing out at the last minute.

What makes me different from others who send letters like this is that I’m local and personally handle each purchase. I don’t make “low-ball” offers just to flip the contract—I actually buy and hold properties myself. That means I can be flexible, fair, and work around your timing and needs.

If selling your property on {StreetName} is something you’d consider, I’d love the opportunity to talk with you. There’s no obligation at all—I’m simply interested in seeing if we can find a solution that works well for both of us.

You can reach me directly at {YourPhone}. I look forward to hearing from you.
"""

POSSIBLE_OWNER_FIRST = ["owner first name","owner_first_name","first name","firstname","ownerfirst","owner first","owner_first","ownerfirst name"]
POSSIBLE_OWNER_NAME  = ["owner name","owner","name","owner_full_name","owner full name","owner(s)","owner 1","owner1"]
POSSIBLE_ADDRESS     = ["situs address","mailing address","property address","site address","address","situsaddr","situs","situsaddress","property situs","prop address","situs_address"]

def _norm(s: str) -> str:
    import re
    return re.sub(r"[^a-z0-9]+", " ", (s or "").strip().lower())

def find_column(headers: List[str], candidates: List[str]) -> Optional[str]:
    norm_map = {h: _norm(h) for h in headers}
    # exact normalized match
    for cand in candidates:
        nc = _norm(cand)
        for h, n in norm_map.items():
            if n == nc:
                return h
    # substring fallback
    for cand in candidates:
        nc = _norm(cand)
        for h, n in norm_map.items():
            if nc in n:
                return h
    return None

LLC_TOKENS = r"(ET\s+AL|TRUST|LLC|INC|CO|LP|L\.P\.|LTD)"
def clean_entity_tokens(name: str) -> str:
    import re
    return re.sub(r"\b" + LLC_TOKENS + r"\b\.?,?", "", name, flags=re.I).strip()

def to_title_case(s: str) -> str:
    if not s: return s
    s = s.strip().lower()
    out = []
    import re
    for w in s.split():
        m = re.match(r"^(\d+)(st|nd|rd|th)$", w)
        out.append(m.group(1)+m.group(2) if m else w.capitalize())
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
                     "trl","trail","sq","square"}

def extract_street_name(full_address: str) -> str:
    if not full_address:
        return "your street"
    s = full_address.strip()
    first_seg = s.split(",")[0]
    import re
    first_seg = re.sub(r"\b(apt|unit|#)\s*\w+", "", first_seg, flags=re.I).strip()
    tokens = first_seg.split()
    if not tokens:
        return "your street"
    i = 0
    import re
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

def personalize_letter(row: Dict[str,str], your_name: str, your_phone: str, your_email: str) -> Tuple[str,str]:
    headers = list(row.keys())
    col_first = find_column(headers, POSSIBLE_OWNER_FIRST)
    col_name  = find_column(headers, POSSIBLE_OWNER_NAME)
    col_addr  = find_column(headers, POSSIBLE_ADDRESS)
    owner_first_raw = split_owner_first(row.get(col_first or "", ""), row.get(col_name or "", ""))
    owner_first = to_title_case(owner_first_raw) or "Neighbor"
    address = (row.get(col_addr or "", "") or "").strip()
    street = to_title_case(extract_street_name(address)) or "your street"
    content = LETTER_TEMPLATE.format(
        OwnerFirstName=owner_first,
        StreetName=street,
        YourName=your_name,
        YourPhone=your_phone,
        YourEmail=your_email
    )
    filestub = f"{owner_first.replace(' ','_')}_{street.replace(' ','_')}".replace("/", "_")
    return content, filestub

# ---------------- Layout (ReportLab Platypus) ----------------

def build_story(content: str, footer_text: str, page_idx: int, total_pages: int):
    styles = getSampleStyleSheet()
    body = ParagraphStyle("Body", parent=styles["Normal"],
                          fontName="Times-Roman", fontSize=12,
                          leading=15, spaceAfter=12)
    sig = ParagraphStyle("Sig", parent=body, spaceBefore=18)

    story = []
    # Push body down ~2 inches for balance
    story.append(Spacer(1, 2*inch))

    # Body paragraphs
    for para in [p for p in content.strip().split("\n\n") if p.strip()]:
        story.append(Paragraph(para.replace("\n", " ").strip(), body))

    # Fixed signature block
    story.append(Spacer(1, 18))
    for line in ["Sincerely,", "Ed Beluli", "916-905-7281", "ed.beluli@gmail.com"]:
        story.append(Paragraph(line, sig))

    return story

def footer_fn_factory(footer_text: str, page_idx: int, total_pages: int):
    def footer_fn(canvas, doc):
        canvas.saveState()
        canvas.setFont("Times-Italic", 9)
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin,
                               doc.bottomMargin - 12,
                               f"{footer_text} • {page_idx} of {total_pages}")
        canvas.restoreState()
    return footer_fn

def write_single_letter_pdf(out_dir: str, filestub: str, content: str, footer_text: str, page_idx: int, total_pages: int) -> str:
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{filestub}.pdf")
    doc = SimpleDocTemplate(path, pagesize=RL_LETTER,
                            leftMargin=1*inch, rightMargin=1*inch,
                            topMargin=1*inch, bottomMargin=1*inch)
    story = build_story(content, footer_text, page_idx, total_pages)
    doc.build(story, onFirstPage=footer_fn_factory(footer_text, page_idx, total_pages),
                    onLaterPages=footer_fn_factory(footer_text, page_idx, total_pages))
    return path

def write_combined_pdf(out_path: str, contents: List[Tuple[str, str]], footer_text: str):
    """contents: list of (content, filestub). Writes one page per item."""
    doc = SimpleDocTemplate(out_path, pagesize=RL_LETTER,
                            leftMargin=1*inch, rightMargin=1*inch,
                            topMargin=1*inch, bottomMargin=1*inch)
    total = len(contents)
    story_all = []
    for i, (content, filestub) in enumerate(contents, start=1):
        story_all.extend(build_story(content, footer_text, i, total))
        if i < total:
            story_all.append(PageBreak())
    def footer_fn(canvas, doc):
        # Pick current page number (1-based)
        pg = int(canvas.getPageNumber())
        canvas.saveState()
        canvas.setFont("Times-Italic", 9)
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin,
                               doc.bottomMargin - 12,
                               f"{footer_text} • {pg} of {total}")
        canvas.restoreState()
    doc.build(story_all, onFirstPage=footer_fn, onLaterPages=footer_fn)

# ---------------- Main ----------------

def main():
    ap = argparse.ArgumentParser(description="Build individual letter PDFs + one combined multi-page PDF")
    ap.add_argument("--csv", required=True, help="Input CSV with recipient info")
    ap.add_argument("--outdir", required=True, help="Folder for individual PDFs")
    ap.add_argument("--combine-out", required=True, help="Output path for the combined PDF (e.g., letters_batch.pdf)")
    ap.add_argument("--name", required=True, help="Your name (used inside body text)")
    ap.add_argument("--phone", required=True, help="Your phone (used inside body text)")
    ap.add_argument("--email", required=True, help="Your email (used inside body text)")
    ap.add_argument("--batch-name", default=None, help="Footer batch label; default: 'Batch YYYY-MM-DD'")
    args = ap.parse_args()

    # Read CSV
    with open(args.csv, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    if not rows:
        print("[INFO] No rows found in CSV.")
        return

    footer_text = args.batch_name or f"Batch {datetime.date.today().isoformat()}"

    # Personalize & generate singles
    contents = []
    total = len(rows)
    os.makedirs(args.outdir, exist_ok=True)
    for i, row in enumerate(rows, start=1):
        content, filestub = personalize_letter(row, args.name, args.phone, args.email)
        contents.append((content, filestub))
        # single file
        path = write_single_letter_pdf(args.outdir, filestub, content, footer_text, i, total)
        print(f"[SAVE] {i}/{total}: {path}")

    # Combined
    write_combined_pdf(args.combine_out, contents, footer_text)
    print(f"[COMBINED] {args.combine_out}")

if __name__ == "__main__":
    main()
