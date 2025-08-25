#!/usr/bin/env python3
"""
Direct Mail Campaign - Print Safe PDF Version (Windows-focused)
- Generates personalized PDF letters from a CSV
- Title-cases owner first names + street names
- Adds footers with batch date + sequence ("i of N")
- Submits PDF prints silently via SumatraPDF.exe
- Tracks each job in Windows spooler until it completes (or errors)
- Logs to print_state.csv and supports --resume
"""

import os, re, csv, argparse, time, datetime, subprocess, platform
from typing import Dict, Optional, Tuple, List

# ---------- PDF (ReportLab) ----------
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas

# ---------- Windows printing deps (only used on Windows) ----------
IS_WINDOWS = platform.system().lower() == "windows"
if IS_WINDOWS:
    try:
        import wmi  # pip install wmi
        import win32print  # pip install pywin32
    except Exception:
        # We'll catch missing deps at runtime and explain how to install
        wmi = None
        win32print = None

# ---------- LETTER TEMPLATE ----------
LETTER_TEMPLATE = """Dear {OwnerFirstName},

I hope this note finds you well. My name is {YourName}, and I’m reaching out because I’m interested in purchasing your property on {StreetName}.

I buy homes directly from owners like you, which means you can avoid paying the typical 5–6% realtor fees and the months of uncertainty that often come with listing a property. I pay cash and can close quickly, so you won’t have to deal with inspections, repairs, or buyers backing out at the last minute.

What makes me different from others who send letters like this is that I’m local and personally handle each purchase. I don’t make “low-ball” offers just to flip the contract—I actually buy and hold properties myself. That means I can be flexible, fair, and work around your timing and needs.

If selling your property on {StreetName} is something you’d consider, I’d love the opportunity to talk with you. There’s no obligation at all—I’m simply interested in seeing if we can find a solution that works well for both of us.

You can reach me directly at {YourPhone}. I look forward to hearing from you.

Sincerely,
{YourName}
{YourPhone}
{YourEmail}
"""

# ---------- COLUMN GUESSING ----------
POSSIBLE_OWNER_FIRST = ["owner first name","owner_first_name","first name","firstname","ownerfirst","owner first"]
POSSIBLE_OWNER_NAME  = ["owner name","owner","name","owner_full_name","owner full name","owner(s)"]
POSSIBLE_ADDRESS     = ["situs address","mailing address","property address","site address","address","situsaddr","situs","situsaddress","property situs","prop address"]

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").strip().lower())

def find_column(header: List[str], candidates: List[str]) -> Optional[str]:
    norm_map = {h: _norm(h) for h in header}
    for cand in candidates:
        c = _norm(cand)
        for h, n in norm_map.items():
            if n == c:
                return h
        for h, n in norm_map.items():
            if c in n:
                return h
    return None

# ---------- NAME + STREET ----------
def to_title_case(s: str) -> str:
    """'SUNDAHL DR' -> 'Sundahl Dr'; '34TH AVE' -> '34th Ave'."""
    if not s:
        return s
    s = s.strip().lower()
    out = []
    for w in s.split():
        m = re.match(r"^(\d+)(st|nd|rd|th)$", w)
        if m:
            out.append(m.group(1) + m.group(2))  # keep ordinal suffix lowercase
        else:
            out.append(w.capitalize())
    return " ".join(out)

def split_owner_first(owner_first: Optional[str], owner_name: Optional[str]) -> str:
    if owner_first and owner_first.strip():
        return owner_first.strip().split()[0]
    if owner_name and owner_name.strip():
        cleaned = re.sub(r"\b(ET\s+AL|TRUST|LLC|INC|CO|LP|L\.P\.|LTD|,)\b\.?", "", owner_name, flags=re.I)
        return cleaned.strip().split()[0]
    return "Neighbor"

STREET_TYPE_WORDS = [
    "ave","avenue","blvd","boulevard","cir","circle","ct","court","dr","drive","hwy","highway",
    "ln","lane","pkwy","parkway","pl","place","rd","road","st","street","ter","terrace","way","trl","trail"
]

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
    street_tokens = tokens[i:]
    if not street_tokens:
        return "your street"
    lower_tokens = [t.lower().strip(".") for t in street_tokens]
    end_idx = None
    for idx, tok in enumerate(lower_tokens):
        if tok in STREET_TYPE_WORDS:
            end_idx = idx
            break
    if end_idx is not None:
        street_core = " ".join(street_tokens[:end_idx+1])
    else:
        street_core = " ".join(street_tokens)
    return street_core.strip() if street_core else "your street"

# ---------- PDF GENERATOR ----------
def write_letter_pdf(out_dir: str, filename_stub: str, content: str, footer: str, page_num: int, total_pages: int) -> str:
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{filename_stub}.pdf")
    c = canvas.Canvas(path, pagesize=LETTER)
    width, height = LETTER
    margin = 72  # 1"
    x = margin
    y = height - margin

    # Body text
    c.setFont("Times-Roman", 12)
    max_width = width - 2 * margin
    for para in content.split("\n\n"):
        words = para.split()
        line = ""
        for w in words:
            test = (line + " " + w).strip()
            if c.stringWidth(test, "Times-Roman", 12) <= max_width:
                line = test
            else:
                c.drawString(x, y, line)
                y -= 14
                if y < margin + 54:
                    # footer before page break
                    c.setFont("Times-Italic", 9)
                    c.drawRightString(width - margin, 36, f"{footer} • {page_num} of {total_pages}")
                    c.showPage()
                    c.setFont("Times-Roman", 12)
                    y = height - margin
                line = w
        if line:
            c.drawString(x, y, line)
            y -= 14
        y -= 10  # paragraph spacing

    # Footer
    c.setFont("Times-Italic", 9)
    c.drawRightString(width - margin, 36, f"{footer} • {page_num} of {total_pages}")

    c.showPage()
    c.save()
    return path

# ---------- WINDOWS PRINT TRACK ----------
def printer_online_win(printer_name: str) -> Tuple[bool, str]:
    if not (IS_WINDOWS and win32print):
        return True, "Non-Windows or pywin32 missing; skipping readiness check."
    try:
        h = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(h, 2)
        win32print.ClosePrinter(h)
        status = info["Status"]
        PRINTER_STATUS_PAUSED    = 0x00000001
        PRINTER_STATUS_ERROR     = 0x00000002
        PRINTER_STATUS_PAPER_OUT = 0x00000010
        PRINTER_STATUS_OFFLINE   = 0x00000080
        bad = status & (PRINTER_STATUS_PAUSED | PRINTER_STATUS_ERROR | PRINTER_STATUS_PAPER_OUT | PRINTER_STATUS_OFFLINE)
        if bad:
            return False, f"Printer not ready (status={status})."
        return True, "Printer ready."
    except Exception as e:
        return False, f"Could not query printer: {e}"

def sumatra_print(pdf_path: str, printer_name: str, sumatra_path: str) -> Tuple[bool, str]:
    if not os.path.exists(sumatra_path):
        return False, "SumatraPDF.exe not found."
    cmd = [sumatra_path, "-silent", "-print-to", printer_name, "-exit-on-print", pdf_path]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode == 0:
        return True, "Spool submitted."
    return False, f"Sumatra failed: {r.stderr or r.stdout}"

def find_job_by_document(printer_name: str, doc_basename: str, timeout_s: int = 30):
    if not (IS_WINDOWS and wmi):
        return None
    c = wmi.WMI()
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        for job in c.Win32_PrintJob():
            if job.Document and os.path.basename(job.Document).lower() == doc_basename.lower():
                if job.Name.split(",")[0].lower() == printer_name.lower():
                    return job
        time.sleep(0.5)
    return None

def wait_for_job_complete(job, max_wait_s: int = 600) -> Tuple[bool, str]:
    if not (IS_WINDOWS and wmi):
        # If we can't track, assume success once spooled.
        return True, "Spool submitted (no tracking available)."
    c = wmi.WMI()
    start = time.time()
    job_id = job.JobId
    printer_name = job.Name.split(",")[0]
    while time.time() - start < max_wait_s:
        current = None
        for j in c.Win32_PrintJob():
            if j.JobId == job_id and j.Name.split(",")[0] == printer_name:
                current = j
                break
        if not current:
            return True, "Job finished (removed from queue)."
        try:
            tp = int(current.TotalPages or 0)
            pp = int(current.PagesPrinted or 0)
        except Exception:
            tp, pp = 0, 0
        if tp > 0 and pp >= tp:
            return True, f"Printed {pp}/{tp}."
        js = ((current.JobStatus or "") + " " + (current.Status or "")).strip()
        if "Error" in js or "Offline" in js or "Paper Out" in js:
            return False, f"Spooler issue: {js or 'unknown'}"
        time.sleep(1.0)
    return False, "Timeout waiting for job to complete."

# ---------- STATE LOG / RESUME ----------
STATE_PATH = "print_state.csv"

def read_printed_rows() -> set:
    done = set()
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, newline="", encoding="utf-8") as f:
            r = csv.DictReader(f)
            for row in r:
                if row.get("status") == "printed":
                    done.add(row.get("row_id"))
    return done

def log_state(row_id: str, pdf: str, status: str, msg: str):
    write_header = not os.path.exists(STATE_PATH)
    with open(STATE_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["row_id", "pdf", "status", "msg"])
        if write_header:
            w.writeheader()
        w.writerow({"row_id": row_id, "pdf": pdf, "status": status, "msg": msg})

# ---------- PERSONALIZE ----------
def personalize_letter(row: Dict[str, str], your_name: str, your_phone: str, your_email: str) -> Tuple[str, str]:
    header = list(row.keys())
    col_first = find_column(header, POSSIBLE_OWNER_FIRST)
    col_name  = find_column(header, POSSIBLE_OWNER_NAME)
    col_addr  = find_column(header, POSSIBLE_ADDRESS)

    owner_first_raw = split_owner_first(row.get(col_first or "", ""), row.get(col_name or "", ""))
    owner_first = to_title_case(owner_first_raw)

    address = (row.get(col_addr or "", "") or "").strip()
    street_raw = extract_street_name(address)
    street = to_title_case(street_raw)

    content = LETTER_TEMPLATE.format(
        OwnerFirstName=owner_first or "Neighbor",
        StreetName=street or "your street",
        YourName=your_name,
        YourPhone=your_phone,
        YourEmail=your_email
    )
    filestub = f"{(owner_first or 'Neighbor').replace(' ', '_')}_{street.replace(' ', '_')}".replace("/", "_")
    return content, filestub

# ---------- MAIN ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", required=True, help="Path to input CSV")
    ap.add_argument("--out", required=True, help="Output folder for PDFs")
    ap.add_argument("--name", required=True, help="Sender name")
    ap.add_argument("--phone", required=True, help="Sender phone")
    ap.add_argument("--email", required=True, help="Sender email")
    ap.add_argument("--printer", help="Printer name (exact, from Get-Printer)")
    ap.add_argument("--sumatra", default="SumatraPDF.exe", help="Path to SumatraPDF.exe")
    ap.add_argument("--dry-run", action="store_true", help="Generate PDFs only, no printing")
    ap.add_argument("--resume", action="store_true", help="Skip rows already marked printed in print_state.csv")
    ap.add_argument("--no-halt", action="store_true", help="Continue on failures (default is halt on first failure)")
    args = ap.parse_args()

    # Dependency check (Windows)
    if IS_WINDOWS and (wmi is None or win32print is None):
        print("[WARN] pywin32 and/or wmi not installed. Install with:")
        print("       pip install pywin32 wmi")
        print("       (Tracking disabled until installed.)")

    # Load rows
    with open(args.csv, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    total = len(rows)
    if total == 0:
        print("[INFO] CSV has no data rows.")
        return

    resume_skip = set()
    if args.resume and os.path.exists(STATE_PATH):
        resume_skip = read_printed_rows()
        if resume_skip:
            print(f"[INFO] Resume mode: will skip {len(resume_skip)} already-printed rows.")

    batch_footer = "Batch " + datetime.date.today().isoformat()

    printed = 0
    saved = 0

    # Optional preflight
    if args.printer and not args.dry_run:
        ok, msg = printer_online_win(args.printer)
        if not ok:
            print(f"[ERROR] {msg}")
            return
        print(f"[INFO] {msg}")

    for idx, row in enumerate(rows, start=1):
        row_id = str(idx)
        if row_id in resume_skip:
            continue

        content, filestub = personalize_letter(row, args.name, args.phone, args.email)
        pdf_path = write_letter_pdf(args.out, filestub, content, batch_footer, idx, total)
        saved += 1
        print(f"[SAVE] {pdf_path}")

        if args.dry_run or not args.printer:
            continue

        # Submit + track
        ok, msg = printer_online_win(args.printer)
        if not ok:
            print(f"[ERROR] {msg}")
            log_state(row_id, pdf_path, "fail", f"NOT_READY: {msg}")
            if not args.no_halt: break
            else: continue

        ok, msg = sumatra_print(pdf_path, args.printer, args.sumatra)
        if not ok:
            print(f"[ERROR] {msg}")
            log_state(row_id, pdf_path, "fail", f"SPOOL_FAIL: {msg}")
            if not args.no_halt: break
            else: continue

        job = find_job_by_document(args.printer, os.path.basename(pdf_path), timeout_s=30)
        if job is None:
            # If we can't see the job, this could be a race with the spooler or viewer.
            print("[WARN] Could not locate job in spooler; will assume submitted.")
            ok2, msg2 = True, "Assumed submitted (no job handle)."
        else:
            ok2, msg2 = wait_for_job_complete(job, max_wait_s=600)

        if ok2:
            printed += 1
            print(f"[PRINTED] {msg2}")
            log_state(row_id, pdf_path, "printed", msg2)
        else:
            print(f"[ERROR] {msg2}")
            log_state(row_id, pdf_path, "fail", f"JOB_FAIL: {msg2}")
            if not args.no_halt: break

    print(f"\n[SUMMARY] Rows: {total}, PDFs saved: {saved}, printed: {printed}")
    if args.dry_run:
        print("[SUMMARY] Dry-run: no print jobs were submitted.")

if __name__ == "__main__":
    main()
