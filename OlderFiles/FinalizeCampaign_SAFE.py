
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# FinalizeCampaign_SAFE.py
# Idempotent finalize step: updates MasterCampaignTracker without double-counting
# when re-running on the same campaign/mapping. Appends to executed log only for
# *new* letters not previously recorded.

import os, csv, argparse, datetime, re, glob, sys
from typing import Dict, Tuple, List, Optional

# ----------------- Script-relative tracker paths -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRACKER_DIR = os.path.join(BASE_DIR, "MasterCampaignTracker")
TRACKER_FILE = os.path.join(TRACKER_DIR, "MasterPropertyCampaignCounter.csv")
ZIP_TALLY_FILE = os.path.join(TRACKER_DIR, "Zip5_LetterTally.csv")

# ----------------- IO helpers -----------------

def ensure_dir_for_file(path: str):
    d = os.path.dirname(path)
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)

def read_csv(path: str):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        rows = [{k:(v or "").strip() for k,v in row.items()} for row in r]
        return rows, list(r.fieldnames or [])

def append_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    ensure_dir_for_file(path)
    file_exists = os.path.isfile(path)
    with open(path, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        if not file_exists:
            w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

def write_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    ensure_dir_for_file(path)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

# ----------------- Normalization -----------------

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def standardize_address(addr: str) -> str:
    return norm_space(addr).upper()

def norm_key(addr: str, owner: str) -> Tuple[str, str]:
    return (standardize_address(addr), norm_space(owner).upper())

def norm_zip(z) -> str:
    s = str(z or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits[:5] if digits else ""

# ----------------- Master lookups -----------------

def load_master_zip_index(master_path: str) -> Dict[str, str]:
    rows, headers = read_csv(master_path)
    low_map = {h.lower(): h for h in headers}
    # Address column (best-effort)
    col_addr = None
    for cand in ["address","situs address","property address","mailing address","site address"]:
        if cand in low_map:
            col_addr = low_map[cand]; break
    if not col_addr:
        for hard in ["Address","Property Address","Situs Address","Mailing Address"]:
            if hard in headers: col_addr = hard; break
    # ZIP columns
    col_zip = None
    for cand in ["zip","zip5","situs zip","property zip","mail zip","mailing zip"]:
        if cand in low_map:
            col_zip = low_map[cand]; break
    if not col_zip:
        for hard in ["ZIP","ZIP5","SITUS ZIP","PROPERTY ZIP","Mail ZIP","MAIL ZIP"]:
            if hard in headers: col_zip = hard; break
    col_mail_zip = None
    for cand in ["mail zip","mailing zip","mail_zip","mail_zip5"]:
        if cand in low_map:
            col_mail_zip = low_map[cand]; break
    if not col_mail_zip:
        for hard in ["Mail ZIP","MAIL ZIP"]:
            if hard in headers: col_mail_zip = hard; break

    idx: Dict[str,str] = {}
    for r in rows:
        addr = standardize_address(r.get(col_addr, "")) if col_addr else ""
        z1 = norm_zip(r.get(col_zip, "")) if col_zip else ""
        z2 = norm_zip(r.get(col_mail_zip, "")) if col_mail_zip else ""
        z = z1 or z2
        if addr and z:
            idx[addr] = z
    return idx

# ----------------- Trackers -----------------

def init_tracker_if_missing():
    if not os.path.isfile(TRACKER_FILE):
        ensure_dir_for_file(TRACKER_FILE)
        write_csv(TRACKER_FILE, [], [
            "PropertyAddress","OwnerName","ZIP5",
            "LastTemplateRef","LastRefCode","LastPage",
            "CampaignName","CampaignNumber",
            "FirstSentDate","LastUpdatedDate","LettersSent"
        ])

def load_tracker_index():
    init_tracker_if_missing()
    rows, _ = read_csv(TRACKER_FILE)
    by_key: Dict[Tuple[str,str], Dict[str,str]] = {}
    for r in rows:
        by_key[norm_key(r.get("PropertyAddress",""), r.get("OwnerName",""))] = r
    return by_key

def save_tracker_index(by_key: Dict[Tuple[str,str], Dict[str,str]]):
    rows = list(by_key.values())
    headers = [
        "PropertyAddress","OwnerName","ZIP5",
        "LastTemplateRef","LastRefCode","LastPage",
        "CampaignName","CampaignNumber",
        "FirstSentDate","LastUpdatedDate","LettersSent"
    ]
    write_csv(TRACKER_FILE, rows, headers)

def init_zip_tally_if_missing():
    if not os.path.isfile(ZIP_TALLY_FILE):
        ensure_dir_for_file(ZIP_TALLY_FILE)
        write_csv(ZIP_TALLY_FILE, [], ["ZIP5","LettersSent","FirstSentDate","LastUpdatedDate"])

def load_zip_tally_index():
    init_zip_tally_if_missing()
    rows, _ = read_csv(ZIP_TALLY_FILE)
    by_zip: Dict[str, Dict[str,str]] = {}
    for r in rows:
        by_zip[r.get("ZIP5","")] = r
    return by_zip

def save_zip_tally_index(by_zip: Dict[str, Dict[str,str]]):
    rows = list(by_zip.values())
    write_csv(ZIP_TALLY_FILE, rows, ["ZIP5","LettersSent","FirstSentDate","LastUpdatedDate"])

# ----------------- Mapping auto-detect -----------------

def autodetect_mapping(camp_dir: str) -> Optional[str]:
    ref_dir = os.path.join(camp_dir, "RefFiles")
    if not os.path.isdir(ref_dir):
        return None
    # Prefer *_refs.csv, newest first
    refs = sorted(glob.glob(os.path.join(ref_dir, "*_refs.csv")), key=os.path.getmtime, reverse=True)
    if refs:
        return refs[0]
    # Fallback
    cand = os.path.join(ref_dir, "letters_mapping.csv")
    return cand if os.path.isfile(cand) else None

# ----------------- Exec log helpers -----------------

EXEC_HEADERS = [
    "RunTimestamp",
    "CampaignName","CampaignNumber","CampaignDir",
    "PropertyAddress","OwnerName","ZIP5",
    "TemplateRef","TemplateSource","RefCode","Page",
    "CombinedPDF","SinglePDF"
]

def read_existing_exec_log(log_path: str) -> List[Dict[str,str]]:
    if not os.path.isfile(log_path):
        return []
    rows, _ = read_csv(log_path)
    return rows

def build_seen_keys(exec_rows: List[Dict[str,str]], campaign_name: str, campaign_number: int):
    """Return sets for idempotency: seen refcodes and seen tuples for this campaign."""
    seen_ref = set()
    seen_tuple = set()
    for r in exec_rows:
        try:
            if str(r.get("CampaignName","")) != str(campaign_name): 
                continue
            if int(str(r.get("CampaignNumber","0"))) != int(campaign_number):
                continue
        except Exception:
            continue
        rc = r.get("RefCode","").strip()
        if rc:
            seen_ref.add(rc)
        key = (
            standardize_address(r.get("PropertyAddress","")),
            (r.get("OwnerName","") or "").strip().upper(),
            str(campaign_number)
        )
        seen_tuple.add(key)
    return seen_ref, seen_tuple

# ----------------- Main -----------------

def main():
    ap = argparse.ArgumentParser(description="SAFE finalize: idempotent updates to trackers. Skips letters already logged for the same campaign.")
    ap.add_argument("--campaign-name", required=True)
    ap.add_argument("--campaign-number", required=True, type=int)
    ap.add_argument("--campaign-dir", required=True, help="Folder containing campaign_master.csv and RefFiles/")
    ap.add_argument("--mapping", help="Path to mapping CSV relative to campaign-dir or absolute; if omitted, auto-detects under RefFiles/")
    ap.add_argument("--dry-run", action="store_true", help="Show what would change without writing")
    ap.add_argument("--force-recount", action="store_true", help="Disable idempotency and count everything again")
    args = ap.parse_args()

    # Resolve paths
    camp_dir = args.campaign_dir
    if not os.path.isabs(camp_dir):
        camp_dir = os.path.abspath(camp_dir)
    master_path = os.path.join(camp_dir, "campaign_master.csv")

    mapping_path = args.mapping
    if mapping_path:
        if not os.path.isabs(mapping_path):
            mapping_path = os.path.join(camp_dir, mapping_path)
    else:
        mapping_path = autodetect_mapping(camp_dir)

    # Validate inputs
    if not os.path.isfile(master_path):
        print(f"[ERROR] Master CSV not found: {master_path}")
        sys.exit(1)
    if not mapping_path or not os.path.isfile(mapping_path):
        print(f"[ERROR] Mapping CSV not found. Pass --mapping or ensure a *_refs.csv/letters_mapping.csv exists in {os.path.join(camp_dir,'RefFiles')}")
        sys.exit(1)

    # Load data
    mapping_rows, _ = read_csv(mapping_path)
    zip_index = load_master_zip_index(master_path)

    # Build executed log rows with ZIP5 from master
    now_iso = datetime.datetime.now().isoformat(timespec="seconds")
    combined_pdf_guess = ""
    comb_dir = os.path.join(camp_dir, "BatchLetterFiles")
    if os.path.isdir(comb_dir):
        pdfs = sorted(glob.glob(os.path.join(comb_dir, "*.pdf")), key=os.path.getmtime, reverse=True)
        if pdfs:
            combined_pdf_guess = pdfs[0]

    # Load existing exec log to enforce idempotency
    log_path = os.path.join(camp_dir, "executed_campaign_log.csv")
    existing_exec = read_existing_exec_log(log_path)
    seen_ref, seen_tuple = build_seen_keys(existing_exec, args.campaign_name, args.campaign_number)

    to_append: List[Dict[str,str]] = []
    for m in mapping_rows:
        prop_addr = norm_space(m.get("property_address","") or m.get("PropertyAddress",""))
        owner = norm_space(m.get("owner","") or m.get("Owner","") or m.get("OwnerName",""))
        page = m.get("page","")
        ref_code = (m.get("ref_code","") or "").strip()
        single_pdf = m.get("single_pdf","")
        template_ref = m.get("template_ref","")
        template_src = m.get("template_source","")
        zip5 = zip_index.get(standardize_address(prop_addr), "")

        if not args.force_recount:
            if ref_code and ref_code in seen_ref:
                # already logged
                continue
            key = (standardize_address(prop_addr), owner.upper(), str(args.campaign_number))
            if key in seen_tuple:
                # already logged (e.g., missing ref code)
                continue

        to_append.append({
            "RunTimestamp": now_iso,
            "CampaignName": str(args.campaign_name),
            "CampaignNumber": str(args.campaign_number),
            "CampaignDir": os.path.basename(camp_dir),
            "PropertyAddress": prop_addr,
            "OwnerName": owner,
            "ZIP5": zip5,
            "TemplateRef": template_ref,
            "TemplateSource": template_src,
            "RefCode": ref_code,
            "Page": str(page),
            "CombinedPDF": combined_pdf_guess,
            "SinglePDF": single_pdf
        })

    print(f"[SUMMARY] Mapping rows: {len(mapping_rows)} | Already logged (skipped): {len(mapping_rows)-len(to_append)} | To add now: {len(to_append)}")

    if args.dry_run:
        print("[DRY RUN] No changes written.")
        sys.exit(0)

    # Append new rows to execution log
    if to_append:
        append_csv(log_path, to_append, EXEC_HEADERS)
        print(f"[OK] Appended {len(to_append)} rows to {log_path}")
    else:
        print("[OK] Nothing new to append to execution log.")

    # Update trackers only for the appended rows
    init_tracker_if_missing()
    tracker = load_tracker_index()
    init_zip_tally_if_missing()
    zip_tally = load_zip_tally_index()
    today = datetime.date.today().isoformat()

    for r in to_append:
        addr = r["PropertyAddress"]; own = r["OwnerName"]; zip5 = r["ZIP5"]
        k = norm_key(addr, own)
        entry = tracker.get(k)
        if entry is None:
            entry = {
                "PropertyAddress": addr,
                "OwnerName": own,
                "ZIP5": zip5,
                "LastTemplateRef": r["TemplateRef"],
                "LastRefCode": r["RefCode"],
                "LastPage": r["Page"],
                "CampaignName": args.campaign_name,
                "CampaignNumber": str(args.campaign_number),
                "FirstSentDate": today,
                "LastUpdatedDate": today,
                "LettersSent": "1"
            }
            tracker[k] = entry
        else:
            if zip5:
                entry["ZIP5"] = zip5
            entry["LastTemplateRef"] = r["TemplateRef"]
            entry["LastRefCode"] = r["RefCode"]
            entry["LastPage"] = r["Page"]
            entry["CampaignName"] = args.campaign_name
            entry["CampaignNumber"] = str(args.campaign_number)
            entry["LastUpdatedDate"] = today
            try:
                entry["LettersSent"] = str(int(entry.get("LettersSent","0")) + 1)
            except Exception:
                entry["LettersSent"] = "1"
            if not entry.get("FirstSentDate"):
                entry["FirstSentDate"] = today

        if zip5:
            zrow = zip_tally.get(zip5)
            if zrow is None:
                zip_tally[zip5] = {
                    "ZIP5": zip5,
                    "LettersSent": "1",
                    "FirstSentDate": today,
                    "LastUpdatedDate": today
                }
            else:
                try:
                    zrow["LettersSent"] = str(int(zrow.get("LettersSent","0")) + 1)
                except Exception:
                    zrow["LettersSent"] = "1"
                if not zrow.get("FirstSentDate"):
                    zrow["FirstSentDate"] = today
                zrow["LastUpdatedDate"] = today

    save_tracker_index(tracker)
    save_zip_tally_index(zip_tally)
    print(f"[OK] Master tracker updated: {TRACKER_FILE}")
    print(f"[OK] ZIP5 tally updated: {ZIP_TALLY_FILE}")

if __name__ == "__main__":
    main()
