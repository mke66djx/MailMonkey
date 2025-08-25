#!/usr/bin/env python3
"""
Direct Mail Campaign Script
- Reads a CSV of property owners
- Generates personalized letters from a template
- (Optionally) prints to a selected printer, with basic status checks

Usage:
  python direct_mail.py --csv "/path/to/list.csv" --out "/path/to/output" \
    --name "Your Name" --phone "555-555-5555" --email "you@example.com" \
    [--printer "Printer Name"] [--dry-run]
"""

import os
import re
import csv
import sys
import argparse
import platform
from typing import Dict, Optional, Tuple

# ---------- CONFIGURABLE LETTER TEMPLATE ----------
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

# ---------- HELPER: COLUMN GUESSING ----------
POSSIBLE_OWNER_FIRST = ["owner first name", "owner_first_name", "first name", "firstname", "ownerfirst", "owner first"]
POSSIBLE_OWNER_NAME  = ["owner name", "owner", "name", "owner_full_name", "owner full name", "owner(s)"]
POSSIBLE_ADDRESS     = ["situs address", "mailing address", "property address", "site address", "address", "situsaddr", "situs", "situsaddress", "property situs", "prop address"]

def normalize(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", s.strip().lower())

def find_column(header: list, candidates: list) -> Optional[str]:
    norm_map = {h: normalize(h) for h in header}
    for cand in candidates:
        c = normalize(cand)
        # exact match
        for h, n in norm_map.items():
            if n == c:
                return h
        # contains
        for h, n in norm_map.items():
            if c in n:
                return h
    return None

# ---------- HELPER: NAME + STREET EXTRACTION ----------
def to_title_case(s: str) -> str:
    """
    Lowercase then Title Case each token; keep numeric ordinals like 34th lowercased suffix.
    Example: 'SUNDAHL DR' -> 'Sundahl Dr', '34TH AVE' -> '34th Ave'
    """
    if not s:
        return s
    s = s.strip().lower()
    out = []
    for w in s.split():
        # handle ordinals: 34th, 21st, 22nd, 23rd
        m = re.match(r"^(\d+)(st|nd|rd|th)$", w)
        if m:
            out.append(m.group(1) + m.group(2))  # digits + lowercase suffix
        else:
            out.append(w.capitalize())
    return " ".join(out)

def split_owner_first(owner_first: Optional[str], owner_name: Optional[str]) -> str:
    """
    Prefer explicit OwnerFirst column if present.
    If only OwnerName/FullName is present, try to take first token.
    """
    if owner_first and owner_first.strip():
        return owner_first.strip().split()[0]
    if owner_name and owner_name.strip():
        # remove common noise like "ET AL", "TRUST", etc. minimally
        cleaned = re.sub(r"\b(ET\s+AL|TRUST|LLC|INC|CO|LP|L\.P\.|LTD|,)\b\.?", "", owner_name, flags=re.I)
        return cleaned.strip().split()[0]
    return "Neighbor"

STREET_TYPE_WORDS = [
    "ave","avenue","blvd","boulevard","cir","circle","ct","court","dr","drive","hwy","highway",
    "ln","lane","pkwy","parkway","pl","place","rd","road","st","street","ter","terrace","way","trl","trail"
]

def extract_street_name(full_address: str) -> str:
    """
    Extract something like 'Main St' from '123 MAIN ST, SACRAMENTO, CA'
    """
    if not full_address:
        return "your street"
    s = full_address.strip()
    first_seg = s.split(",")[0]
    first_seg = re.sub(r"\b(apt|unit|#)\s*\w+", "", first_seg, flags=re.I).strip()

    tokens = first_seg.split()
    if not tokens:
        return "your street"

    # Remove leading house number(s)
    i = 0
    while i < len(tokens) and re.match(r"^\d+[a-zA-Z]?$", tokens[i]):
        i += 1

    street_tokens = tokens[i:]
    if not street_tokens:
        return "your street"

    # If contains known street type, cut up to that
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

# ---------- PRINTER HELPERS (OS-SPECIFIC, OPTIONAL) ----------
def list_printers() -> list:
    system = platform.system().lower()
    printers = []
    try:
        if system == "windows":
            import win32print  # type: ignore
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        else:
            # macOS/Linux via CUPS
            import cups  # type: ignore
            conn = cups.Connection()
            printers = list(conn.getPrinters().keys())
    except Exception:
        pass
    return printers

def check_printer_ready(printer_name: str) -> Tuple[bool, str]:
    """
    Basic readiness: installed, not offline, not paused. Ink/paper levels are vendor-specific and not always exposed.
    """
    system = platform.system().lower()
    try:
        if system == "windows":
            import win32print  # type: ignore
            handle = win32print.OpenPrinter(printer_name)
            try:
                info = win32print.GetPrinter(handle, 2)
                status = info["Status"]
                attributes = info["Attributes"]
                offline = status & 0x00000080  # PRINTER_STATUS_OFFLINE
                error   = status & 0x00000002  # ERROR
                paused  = status & 0x00000001  # PAUSED
                if offline or error or paused:
                    return False, f"Printer status not ready (status={status}, attributes={attributes})."
                return True, "Printer appears ready."
            finally:
                win32print.ClosePrinter(handle)
        else:
            import cups  # type: ignore
            conn = cups.Connection()
            printers = conn.getPrinters()
            if printer_name not in printers:
                return False, "Printer not found."
            state = printers[printer_name].get("printer-state", 3)  # 3=idle,4=printing,5=stopped
            if int(state) == 5:
                return False, "Printer is stopped."
            return True, "Printer appears ready."
    except Exception as e:
        return False, f"Could not verify printer readiness ({e})."

def send_to_printer(printer_name: str, filepath: str) -> Tuple[bool, str]:
    """
    Sends a text or PDF file to printer.
    On Windows: uses ShellExecute 'print' verb for .txt/.pdf if associated app supports printing.
    On macOS/Linux: uses CUPS (lp) via pycups if available; else fallback to 'lp' command.
    """
    system = platform.system().lower()
    try:
        if system == "windows":
            import win32api  # type: ignore
            win32api.ShellExecute(0, "print", filepath, None, ".", 0)
            return True, "Print command sent."
        else:
            try:
                import cups  # type: ignore
                conn = cups.Connection()
                job_id = conn.printFile(printer_name, filepath, os.path.basename(filepath), {})
                return True, f"CUPS job submitted: {job_id}."
            except Exception:
                import subprocess
                r = subprocess.run(["lp", "-d", printer_name, filepath], capture_output=True, text=True)
                if r.returncode == 0:
                    return True, "lp job submitted."
                return False, f"lp failed: {r.stderr}"
    except Exception as e:
        return False, f"Print error: {e}"

# ---------- CORE PIPELINE ----------
def personalize_letter(row: Dict[str, str], your_name: str, your_phone: str, your_email: str) -> Tuple[str, str]:
    # Guess columns
    header = list(row.keys())
    col_first = find_column(header, POSSIBLE_OWNER_FIRST)
    col_name  = find_column(header, POSSIBLE_OWNER_NAME)
    col_addr  = find_column(header, POSSIBLE_ADDRESS)

    owner_first_raw = split_owner_first(row.get(col_first or "", ""), row.get(col_name or "", ""))
    owner_first = to_title_case(owner_first_raw)

    address = row.get(col_addr or "", "").strip()
    street_raw  = extract_street_name(address)
    street = to_title_case(street_raw)

    content = LETTER_TEMPLATE.format(
        OwnerFirstName=owner_first or "Neighbor",
        StreetName=street or "your street",
        YourName=your_name,
        YourPhone=your_phone,
        YourEmail=your_email
    )
    filename_stub = f"{(owner_first or 'Neighbor').replace(' ', '_')}_{street.replace(' ', '_')}".replace("/", "_")
    return content, filename_stub

def write_letter_txt(out_dir: str, filename_stub: str, content: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{filename_stub}.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    return path

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True, help="Path to input CSV")
    parser.add_argument("--out", required=True, help="Directory for generated letters")
    parser.add_argument("--name", required=True, help="Your sender name")
    parser.add_argument("--phone", required=True, help="Your phone number")
    parser.add_argument("--email", required=True, help="Your email address")
    parser.add_argument("--printer", help="Target printer name (optional)")
    parser.add_argument("--dry-run", action="store_true", help="Generate letters but do not send to printer")
    args = parser.parse_args()

    # Printer listing/help
    if args.printer:
        ready, msg = check_printer_ready(args.printer)
        if not ready:
            print(f"[WARN] {msg}")
        else:
            print(f"[INFO] {msg}")
    else:
        printers = list_printers()
        if printers:
            print("[INFO] Available printers:")
            for p in printers:
                print("  -", p)
        else:
            print("[INFO] No printers found (or printer libraries not installed). Proceeding with file generation.")

    # Process CSV
    total = 0
    printed = 0
    saved = 0

    with open(args.csv, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            total += 1
            content, filestub = personalize_letter(row, args.name, args.phone, args.email)
            letter_path = write_letter_txt(args.out, filestub, content)
            saved += 1
            print(f"[SAVE] {letter_path}")
            if args.printer and not args.dry_run:
                ok, msg = send_to_printer(args.printer, letter_path)
                if ok:
                    printed += 1
                    print(f"[PRINT] {msg}")
                else:
                    print(f"[ERROR] Print failed for {letter_path}: {msg}")

    print(f"\n[SUMMARY] Rows processed: {total}, letters saved: {saved}, printed: {printed}")
    if args.dry_run:
        print("[SUMMARY] Dry-run: no print jobs were submitted.")

if __name__ == "__main__":
    main()
