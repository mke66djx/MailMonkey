#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Direct Mail Campaign - Print Safe PDF Version (Windows-focused, comprehensive)
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

Windows requirements
  pip install reportlab pywin32 wmi
  (for SNMP proof) pip install pysnmp
  Place SumatraPDF.exe (portable) next to this script (or pass --sumatra path)

macOS/Linux
  - PDF generation works
  - Printing uses 'lp' best-effort (no spool tracking). SNMP works if IP provided.
"""

from __future__ import annotations
import os, re, csv, sys, time, argparse, datetime, subprocess, platform
from typing import Dict, Optional, Tuple, List

WIN = platform.system().lower() == "windows"

# PDF deps
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
except Exception:
    LETTER = None
    canvas = None

# Windows printing deps (lazy-checked)
if WIN:
    try:
        import wmi  # type: ignore
        import win32print  # type: ignore
    except Exception:
        wmi = None
        win32print = None

# SNMP deps (lazy load only if requested)
# ---------- SNMP (PySNMP 7.x asyncio HLAPI) ----------
import asyncio

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
    try:
        return asyncio.run(_snmp_get_int_async(ip, community, OID_PRT_MARKER_LIFE_COUNT))
    except RuntimeError:
        # if an event loop is already running
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(_snmp_get_int_async(ip, community, OID_PRT_MARKER_LIFE_COUNT))
        finally:
            loop.close()

def wait_for_page_increment(ip: str, community: str, baseline: int, expected_delta: int = 1, max_wait: int = 180) -> tuple[bool, str]:
    """Poll SNMP page counter until it increases by at least expected_delta (or timeout)."""
    import time
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

# ---------- PERSONALIZATION ----------
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

# ---------- DEP CHECK ----------
def ensure_deps_or_exit():
    if LETTER is None or canvas is None:
        print("[ERROR] ReportLab is not installed. Run: pip install reportlab")
        sys.exit(2)
    if WIN and (wmi is None or win32print is None):
        print("[WARN] For full tracking on Windows, install: pip install pywin32 wmi")

# ---------- MAIN ----------
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

    # Pre-flight SNMP
    snmp_enabled = bool(args.printer_ip)
    if snmp_enabled:
        if not _load_pysnmp():
            print("[ERROR] SNMP requested but pysnmp is not installed. Run: pip install pysnmp")
            sys.exit(2)
        # quick read to validate connectivity
        test_counter = get_page_counter(args.printer_ip, args.snmp_community)
        if test_counter is None:
            print("[WARN] Could not read SNMP page counter from printer. SNMP proof-of-print will be disabled.")
            snmp_enabled = False

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

            # SNMP proof-of-print
            if snmp_enabled:
                baseline = get_page_counter(args.printer_ip, args.snmp_community)
                if baseline is None:
                    successes += 1; append_state(idx, pdf_path, "printed", f"{msg} | SNMP baseline unavailable")
                    print(f"[PRINTED] Row {idx}: {msg} | SNMP baseline unavailable")
                else:
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
                    if snmp_enabled:
                        baseline = get_page_counter(args.printer_ip, args.snmp_community)
                        if baseline is not None:
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
                            successes += 1; append_state(idx, pdf_path, "printed", "lp submitted | SNMP baseline unavailable")
                            print(f"[PRINTED] Row {idx}: lp submitted | SNMP baseline unavailable")
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
