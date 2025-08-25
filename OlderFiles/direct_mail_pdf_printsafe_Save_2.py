
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Direct Mail Campaign - Print Safe PDF Version (Windows-focused, comprehensive, SNMP-proofed)
Author: ChatGPT (for Ed Beluli)

Capabilities
- Generates personalized PDF letters from a CSV (ReportLab)
- Title-cases owner names + street names; extracts core street from full address
- Adds per-letter footer with batch/date + sequence (i of N)
- Prints via SumatraPDF.exe silently to the chosen printer
- Tracks print jobs on Windows: finds the job in spooler and polls until finished
- Polls live PRINTER_INFO_2 status (paused/offline/paper out) during the job
- Optional SNMP proof-of-print: verifies the printer page counter increased
- Resume-friendly state log (print_state.csv) + --resume to skip printed rows
- Flags: --no-halt, --list-printers, --batch-name, --sumatra path
- Preserves prior functionality while fixing PySNMP 7.x usage and call sites.

Windows requirements
  pip install reportlab pywin32 wmi
  (for SNMP proof) pip install pysnmp
  Place SumatraPDF.exe (portable) next to this script (or pass --sumatra path)

macOS/Linux
  - PDF generation works
  - Printing uses 'lp' best-effort (no spool tracking). SNMP works if IP provided.
"""

from __future__ import annotations
import os, re, csv, sys, time, argparse, datetime, subprocess, platform, asyncio
from typing import Dict, Optional, Tuple, List

WIN = platform.system().lower() == "windows"

# ---------- PDF deps ----------
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
except Exception:
    LETTER = None
    canvas = None

# ---------- Windows printing deps (lazy) ----------
if WIN:
    try:
        import wmi  # type: ignore
        import win32print  # type: ignore
    except Exception:
        wmi = None
        win32print = None

# ---------- SNMP (PySNMP 7.x asyncio HLAPI) ----------
# Printer total page counter (Printer-MIB)
OID_PRT_MARKER_LIFE_COUNT = "1.3.6.1.2.1.43.10.2.1.4.1.1"

def _snmp_imports():
    # lazy import so script runs without pysnmp unless SNMP is requested
    from pysnmp.hlapi.v1arch.asyncio import (
        SnmpDispatcher, CommunityData, UdpTransportTarget,
        ObjectType, ObjectIdentity, get_cmd,
    )
    return SnmpDispatcher, CommunityData, UdpTransportTarget, ObjectType, ObjectIdentity, get_cmd

async def _snmp_get_int_async(ip: str, community: str, oid: str, timeout: float = 2.0) -> int | None:
    SnmpDispatcher, CommunityData, UdpTransportTarget, ObjectType, ObjectIdentity, get_cmd = _snmp_imports()
    disp = SnmpDispatcher()
    target = await UdpTransportTarget.create((ip, 161))
    errInd, errStat, errIdx, varBinds = await get_cmd(
        disp,
        CommunityData(community),
        target,
        ObjectType(ObjectIdentity(oid)),
    )
    if errInd or errStat:
        return None
    for ob in varBinds:
        try:
            return int(ob[1])
        except Exception:
            pass
    return None

def snmp_get_page_counter(ip: str, community: str) -> int | None:
    """Synchronous wrapper to read the printer's lifetime page counter via SNMP (v1 community)."""
    try:
        return asyncio.run(_snmp_get_int_async(ip, community, OID_PRT_MARKER_LIFE_COUNT))
    except RuntimeError:
        # if an event loop is already running (e.g., in some hosts), use a new one
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(_snmp_get_int_async(ip, community, OID_PRT_MARKER_LIFE_COUNT))
        finally:
            loop.close()

def wait_for_page_increment(ip: str, community: str, baseline: int, expected_delta: int = 1, max_wait: int = 180) -> tuple[bool, str]:
    """Poll SNMP page counter until it increases by at least expected_delta (or timeout)."""
    start = time.time()
    last = None
    while time.time() - start < max_wait:
        val = snmp_get_page_counter(ip, community)
        if val is not None:
            last = val
            if val >= baseline + expected_delta:
                return True, f"SNMP page counter increased: {baseline} -> {val}."
        time.sleep(2.0)
    if last is None:
        return False, "SNMP counter unavailable (no response)."
    return False, f"SNMP counter did not increase within timeout (baseline {baseline}, last {last})."

# ---------- Letter template & personalization helpers ----------
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

POSSIBLE_OWNER_FIRST = ["owner first name","owner_first_name","first name","firstname","ownerfirst","owner first"]
POSSIBLE_OWNER_NAME  = ["owner name","owner","name","owner_full_name","owner full name","owner(s)","owner 1","owner1"]
POSSIBLE_ADDRESS     = ["situs address","mailing address","property address","site address","address","situsaddr","situs","situsaddress","property situs","prop address","situs_address"]

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").strip().lower())

def find_column(headers: List[str], candidates: List[str]) -> Optional[str]:
    norm_map = {h: _norm(h) for h in headers}
    for cand in candidates:
        nc = _norm(cand)
        for h, n in norm_map.items():
            if n == nc:
                return h
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
                     "trl","trail","sq","square"}

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

# ---------- PDF generation ----------
def ensure_reportlab():
    if LETTER is None or canvas is None:
        raise RuntimeError("ReportLab is required. Install with: pip install reportlab")

def write_letter_pdf(out_dir: str, filename_stub: str, content: str, footer: str, seq_num: int, total: int) -> str:
    ensure_reportlab()
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{filename_stub}.pdf")
    c = canvas.Canvas(path, pagesize=LETTER)
    width, height = LETTER
    margin = 72
    x = margin
    y = height - margin
    c.setFont("Times-Roman", 12)
    max_w = width - 2*margin
    for para in content.split("\n\n"):
        words = para.split()
        line = ""
        for w in words:
            test = (line + " " + w).strip()
            if c.stringWidth(test, "Times-Roman", 12) <= max_w:
                line = test
            else:
                c.drawString(x, y, line)
                y -= 14
                if y < margin + 54:
                    c.setFont("Times-Italic", 9)
                    c.drawRightString(width - margin, 36, f"{footer} • {seq_num} of {total}")
                    c.showPage()
                    c.setFont("Times-Roman", 12)
                    y = height - margin
                line = w
        if line:
            c.drawString(x, y, line)
            y -= 14
        y -= 10
    c.setFont("Times-Italic", 9)
    c.drawRightString(width - margin, 36, f"{footer} • {seq_num} of {total}")
    c.showPage(); c.save()
    return path

# ---------- Windows printing & tracking ----------
def require_windows_modules():
    if not WIN:
        return
    if wmi is None or win32print is None:
        raise RuntimeError("Windows printing requires: pip install pywin32 wmi")

def list_printers_win() -> List[str]:
    require_windows_modules()
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    return [p[2] for p in win32print.EnumPrinters(flags)]

def printer_status_flags(printer_name: str) -> Tuple[bool, str, int]:
    """Return (ok, message, status_bits) from PRINTER_INFO_2.Status"""
    require_windows_modules()
    try:
        h = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(h, 2)
        win32print.ClosePrinter(h)
        status = info["Status"] or 0
        PAUSED=0x1; ERROR=0x2; PAPER_OUT=0x10; OFFLINE=0x80
        bad = status & (PAUSED|ERROR|PAPER_OUT|OFFLINE)
        if bad:
            return False, f"Printer not ready (status={status}).", status
        return True, "Printer ready.", status
    except Exception as e:
        return False, f"Could not query printer: {e}", -1

def sumatra_print(pdf_path: str, printer_name: str, sumatra_path: str) -> Tuple[bool, str]:
    if not os.path.exists(sumatra_path):
        return False, f"SumatraPDF.exe not found at: {sumatra_path}"
    try:
        cmd = [sumatra_path, "-silent", "-print-to", printer_name, "-exit-on-print", pdf_path]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode == 0:
            return True, "Spool submitted."
        return False, f"Sumatra failed: {r.stderr or r.stdout}"
    except Exception as e:
        return False, f"Sumatra error: {e}"

def find_job_by_document(printer_name: str, doc_basename: str, timeout_s: int = 60):
    require_windows_modules()
    c = wmi.WMI()
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        for job in c.Win32_PrintJob():
            if job.Document and os.path.basename(job.Document).lower() == doc_basename.lower():
                if job.Name.split(",")[0].lower() == printer_name.lower():
                    return job
        time.sleep(0.5)
    return None

def wait_for_job_complete(printer_name: str, job, max_wait_s: int = 1200) -> Tuple[bool, str]:
    """Poll the job AND the printer for real-time errors like PAPER OUT."""
    require_windows_modules()
    c = wmi.WMI()
    start = time.time()
    job_id = job.JobId
    pname = job.Name.split(",")[0]
    while time.time() - start < max_wait_s:
        # Check printer live status for paper/toner/offline
        okp, msgp, status_bits = printer_status_flags(printer_name)
        if not okp:
            return False, f"PRINTER_STATUS issue during job: {msgp}"
        # Check job
        current = None
        for j in c.Win32_PrintJob():
            if j.JobId == job_id and j.Name.split(",")[0] == pname:
                current = j; break
        if not current:
            return True, "Finished (removed from queue)."
        try:
            tp = int(current.TotalPages or 0); pp = int(current.PagesPrinted or 0)
        except Exception:
            tp, pp = 0, 0
        if tp > 0 and pp >= tp:
            return True, f"Printed {pp}/{tp}."
        js = ((current.JobStatus or "") + " " + (current.Status or "")).strip()
        if any(k in js for k in ["Error", "Offline", "Paper Out", "Out of paper"]):
            return False, f"Spooler job reported issue: {js or 'unknown'}"
        time.sleep(0.8)
    return False, "Timeout waiting for job to complete."

def print_non_windows(pdf_path: str, printer_name: Optional[str]) -> Tuple[bool, str]:
    cmd = ["lp", pdf_path] if not printer_name else ["lp", "-d", printer_name, pdf_path]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode == 0:
            return True, "Submitted to lp."
        return False, f"lp failed: {r.stderr or r.stdout}"
    except Exception as e:
        return False, f"lp error: {e}"

# ---------- State log ----------
STATE_PATH = "print_state.csv"
def read_printed_rows() -> set[int]:
    done = set()
    if not os.path.exists(STATE_PATH):
        return done
    try:
        with open(STATE_PATH, "r", newline="", encoding="utf-8") as f:
            r = csv.DictReader(f)
            for row in r:
                if str(row.get("status","")).lower() == "printed":
                    try: done.add(int(row.get("row_id","0")))
                    except: pass
    except Exception:
        pass
    return done

def append_state(row_id: int, pdf: str, status: str, message: str) -> None:
    newfile = not os.path.exists(STATE_PATH)
    with open(STATE_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if newfile:
            w.writerow(["row_id","pdf","status","message","timestamp"])
        w.writerow([row_id, pdf, status, message, datetime.datetime.now().isoformat(timespec="seconds")])

# ---------- Personalization ----------
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

# ---------- Dep checks ----------
def ensure_deps_or_exit():
    if LETTER is None or canvas is None:
        print("[ERROR] ReportLab is not installed. Run: pip install reportlab")
        sys.exit(2)
    if WIN and (wmi is None or win32print is None):
        print("[WARN] For full tracking on Windows, install: pip install pywin32 wmi")

# ---------- CLI / Main ----------
def main():
    ap = argparse.ArgumentParser(description="Direct Mail PDF generator & print-safe runner with SNMP proof-of-print")
    ap.add_argument("--csv", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--name", required=True)
    ap.add_argument("--phone", required=True)
    ap.add_argument("--email", required=True)
    ap.add_argument("--printer")
    ap.add_argument("--sumatra", default="SumatraPDF.exe")
    ap.add_argument("--batch-name", default=None)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--resume", action="store_true")
    ap.add_argument("--no-halt", action="store_true")
    ap.add_argument("--list-printers", action="store_true")
    # SNMP proof-of-print
    ap.add_argument("--printer-ip", help="Printer IP address (enables SNMP proof-of-print)")
    ap.add_argument("--snmp-community", default="public", help="SNMP community (default: public)")
    ap.add_argument("--snmp-wait-seconds", type=int, default=180, help="Seconds to wait for page counter to increase")
    args = ap.parse_args()

    ensure_deps_or_exit()

    if args.list_printers:
        if WIN and (wmi is not None and win32print is not None):
            try:
                for n in list_printers_win():
                    print(" -", n)
            except Exception as e:
                print(f"[WARN] Could not list printers: {e}")
        else:
            try:
                r = subprocess.run(["lpstat","-p","-d"], capture_output=True, text=True)
                print(r.stdout or r.stderr)
            except Exception as e:
                print(f"[WARN] Could not query printers: {e}")
        return

    # Load CSV
    try:
        with open(args.csv, "r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        print(f"[ERROR] CSV not found: {args.csv}"); sys.exit(2)

    total = len(rows)
    if total == 0:
        print("[INFO] No rows in CSV."); return

    printed_rows = read_printed_rows() if args.resume else set()
    batch_footer = args.batch_name or f"Batch {datetime.date.today().isoformat()}"

    # Pre-flight on Windows
    if args.printer and not args.dry_run and WIN and wmi is not None and win32print is not None:
        ok, msg, _ = printer_status_flags(args.printer)
        if not ok:
            print(f"[ERROR] {msg}"); sys.exit(2)
        else:
            print(f"[INFO] {msg}")

    # Pre-flight SNMP check (call site #1)
    snmp_enabled = bool(args.printer_ip)
    if snmp_enabled:
        test_counter = snmp_get_page_counter(args.printer_ip, args.snmp_community)
        if test_counter is None:
            print("[WARN] Could not read SNMP page counter from printer. SNMP proof-of-print will be disabled.")
            snmp_enabled = False
        else:
            print(f"[INFO] SNMP page counter at start: {test_counter}")

    successes = 0; failures = 0

    for idx, row in enumerate(rows, start=1):
        if args.resume and idx in printed_rows:
            continue

        content, filestub = personalize_letter(row, args.name, args.phone, args.email)
        try:
            pdf_path = write_letter_pdf(args.out, filestub, content, batch_footer, idx, total)
        except Exception as e:
            failures += 1
            append_state(idx, "", "fail", f"PDF error: {e}")
            print(f"[ERROR] Row {idx}: PDF error: {e}")
            if not args.no_halt: break
            else: continue

        print(f"[SAVE] Row {idx}/{total}: {pdf_path}")
        append_state(idx, pdf_path, "saved", "PDF generated")

        if args.dry_run or not args.printer:
            successes += 1
            continue

        # Per-job SNMP baseline (call site #2) BEFORE sending to printer
        baseline = None
        if snmp_enabled:
            baseline = snmp_get_page_counter(args.printer_ip, args.snmp_community)
            if baseline is not None:
                print(f"[INFO] Baseline page counter before print: {baseline}")
            else:
                print("[WARN] SNMP baseline unavailable for this job.")

        if WIN and (wmi is not None and win32print is not None):
            ok, msg, _ = printer_status_flags(args.printer)
            if not ok:
                failures += 1; append_state(idx, pdf_path, "fail", f"NOT_READY: {msg}")
                print(f"[ERROR] Row {idx}: {msg}")
                if not args.no_halt: break
                else: continue

            ok, msg = sumatra_print(pdf_path, args.printer, args.sumatra)
            if not ok:
                failures += 1; append_state(idx, pdf_path, "fail", f"SPOOL_FAIL: {msg}")
                print(f"[ERROR] Row {idx}: {msg}")
                if not args.no_halt: break
                else: continue

            job = find_job_by_document(args.printer, os.path.basename(pdf_path), timeout_s=90)
            if not job:
                failures += 1; append_state(idx, pdf_path, "fail", "NO_JOB_FOUND")
                print(f"[ERROR] Row {idx}: Spooler did not expose the job.")
                if not args.no_halt: break
                else: continue

            ok, msg = wait_for_job_complete(args.printer, job, max_wait_s=1800)
            if not ok:
                failures += 1; append_state(idx, pdf_path, "fail", f"JOB_FAIL: {msg}")
                print(f"[ERROR] Row {idx}: {msg}")
                if not args.no_halt: break
                else: continue

            # SNMP proof-of-print after job completion
            if snmp_enabled and baseline is not None:
                ok2, msg2 = wait_for_page_increment(args.printer_ip, args.snmp_community, baseline, expected_delta=1, max_wait=args.snmp_wait_seconds)
                if ok2:
                    successes += 1; append_state(idx, pdf_path, "printed", f"{msg} | {msg2}")
                    print(f"[PRINTED] Row {idx}: {msg} | {msg2}")
                else:
                    failures += 1; append_state(idx, pdf_path, "fail", f"NO_PAGE_INCREMENT: {msg2}")
                    print(f"[ERROR] Row {idx}: {msg2}")
                    if not args.no_halt: break
                    else: continue
            else:
                successes += 1; append_state(idx, pdf_path, "printed", msg)
                print(f"[PRINTED] Row {idx}: {msg}")

        else:
            # Non-Windows best-effort
            try:
                r = subprocess.run(["lp", "-d", args.printer, pdf_path] if args.printer else ["lp", pdf_path],
                                   capture_output=True, text=True)
                if r.returncode == 0:
                    if snmp_enabled and baseline is not None:
                        ok2, msg2 = wait_for_page_increment(args.printer_ip, args.snmp_community, baseline, expected_delta=1, max_wait=args.snmp_wait_seconds)
                        if ok2:
                            successes += 1; append_state(idx, pdf_path, "printed", f"lp submitted | {msg2}")
                            print(f"[PRINTED] Row {idx}: lp submitted | {msg2}")
                        else:
                            failures += 1; append_state(idx, pdf_path, "fail", f"NO_PAGE_INCREMENT: {msg2}")
                            print(f"[ERROR] Row {idx}: {msg2}")
                            if not args.no_halt: break
                            else: continue
                    else:
                        successes += 1; append_state(idx, pdf_path, "printed", "lp submitted")
                        print(f"[PRINTED] Row {idx}: lp submitted")
                else:
                    failures += 1; append_state(idx, pdf_path, "fail", f"lp failed: {r.stderr or r.stdout}")
                    print(f"[ERROR] Row {idx}: lp failed: {r.stderr or r.stdout}")
                    if not args.no_halt: break
                    else: continue
            except Exception as e:
                failures += 1; append_state(idx, pdf_path, "fail", f"lp error: {e}")
                print(f"[ERROR] Row {idx}: lp error: {e}")
                if not args.no_halt: break
                else: continue

    print("\n=== RUN SUMMARY ===")
    print(f"Total rows:  {total}")
    print(f"Successes:   {successes}")
    print(f"Failures:    {failures}")
    if args.resume:
        print(f"Resume mode: ON (skipped {len(printed_rows)} already-printed rows)")
    print(f"State log:   {os.path.abspath(STATE_PATH)}")
    print("===================")

if __name__ == "__main__":
    main()
