#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
FinalizeCampaign.py (ZIP5-FIX)
- Writes executed_campaign_log.csv in the campaign folder (with ZIP5 filled by joining to campaign_master.csv).
- Updates MasterCampaignTracker/MasterPropertyCampaignCounter.csv (includes ZIP5).
- Maintains MasterCampaignTracker/Zip5_LetterTally.csv with rolling counts per ZIP5.

CLI:
  python FinalizeCampaign.py ^
    --campaign-name "Campaign" ^
    --campaign-number 1 ^
    --campaign-dir "Campaign_1_Aug2025" ^
    --mapping "RefFiles\\letters_mapping.csv"

Notes:
- Always run from the repo root (MailCampaigns) OR pass an absolute --campaign-dir.
- This version fixes empty ZIP5 by joining mapping.PropertyAddress -> master.Address.
"""

import os, csv, argparse, datetime, re
from typing import Dict, Tuple, List

TRACKER_DIR = "MasterCampaignTracker"
TRACKER_FILE = os.path.join(TRACKER_DIR, "MasterPropertyCampaignCounter.csv")
ZIP_TALLY_FILE = os.path.join(TRACKER_DIR, "Zip5_LetterTally.csv")

# ----------------- IO helpers -----------------

def ensure_dir(p: str):
    d = os.path.dirname(p)
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)

def read_csv(path: str):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        rows = [{k:(v or "").strip() for k,v in row.items()} for row in r]
        return rows, list(r.fieldnames or [])

def write_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    ensure_dir(path)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

# ----------------- Normalization -----------------

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def standardize_address(addr: str) -> str:
    # Preserve street line as-is except whitespace + uppercase (to match master)
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
    """
    Build { ADDRESS(line1 upper) -> ZIP5 } from campaign_master.csv.
    Prefers 'ZIP', falls back to 'Mail ZIP'.
    """
    rows, headers = read_csv(master_path)
    # pick columns
    col_addr = None
    col_zip = None
    col_mail_zip = None
    low_map = {h.lower(): h for h in headers}
    # best-effort matching
    for cand in ["address","situs address","property address","mailing address","site address"]:
        if cand in low_map:
            col_addr = low_map[cand]; break
    if not col_addr and "Address" in headers:
        col_addr = "Address"
    for cand in ["zip","zip5","situs zip","property zip","mail zip","mailing zip"]:
        if cand in low_map:
            col_zip = low_map[cand]; break
    if not col_zip and "ZIP" in headers:
        col_zip = "ZIP"
    for cand in ["mail zip","mailing zip","mail_zip","mail_zip5"]:
        if cand in low_map:
            col_mail_zip = low_map[cand]; break
    if not col_mail_zip and "Mail ZIP" in headers:
        col_mail_zip = "Mail ZIP"

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

def load_tracker():
    if not os.path.isfile(TRACKER_FILE):
        ensure_dir(TRACKER_FILE)
        write_csv(TRACKER_FILE, [], [
            "PropertyAddress","OwnerName","ZIP5",
            "LastTemplateRef","LastRefCode","LastPage",
            "CampaignName","CampaignNumber",
            "FirstSentDate","LastUpdatedDate","LettersSent"
        ])
    rows, headers = read_csv(TRACKER_FILE)
    # index by (addr, owner)
    by_key: Dict[Tuple[str,str], Dict[str,str]] = {}
    for r in rows:
        k = norm_key(r.get("PropertyAddress",""), r.get("OwnerName",""))
        by_key[k] = r
    return by_key

def save_tracker(by_key: Dict[Tuple[str,str], Dict[str,str]]):
    rows = list(by_key.values())
    headers = [
        "PropertyAddress","OwnerName","ZIP5",
        "LastTemplateRef","LastRefCode","LastPage",
        "CampaignName","CampaignNumber",
        "FirstSentDate","LastUpdatedDate","LettersSent"
    ]
    write_csv(TRACKER_FILE, rows, headers)

def load_zip_tally():
    if not os.path.isfile(ZIP_TALLY_FILE):
        ensure_dir(ZIP_TALLY_FILE)
        write_csv(ZIP_TALLY_FILE, [], ["ZIP5","LettersSent","FirstSentDate","LastUpdatedDate"])
    rows, _ = read_csv(ZIP_TALLY_FILE)
    by_zip: Dict[str, Dict[str,str]] = {}
    for r in rows:
        by_zip[r.get("ZIP5","")] = r
    return by_zip

def save_zip_tally(by_zip: Dict[str, Dict[str,str]]):
    rows = list(by_zip.values())
    write_csv(ZIP_TALLY_FILE, rows, ["ZIP5","LettersSent","FirstSentDate","LastUpdatedDate"])

# ----------------- Main -----------------

def main():
    ap = argparse.ArgumentParser(description="Finalize campaign, write executed log (with ZIP5), and update trackers.")
    ap.add_argument("--campaign-name", required=True)
    ap.add_argument("--campaign-number", required=True, type=int)
    ap.add_argument("--campaign-dir", required=True, help="Folder containing campaign_master.csv and RefFiles/")
    ap.add_argument("--mapping", required=True, help="Path to letters_mapping.csv relative to campaign-dir or absolute")
    args = ap.parse_args()

    # Resolve paths robustly
    camp_dir = args.campaign_dir
    if not os.path.isabs(camp_dir):
        camp_dir = os.path.abspath(camp_dir)
    master_path = os.path.join(camp_dir, "campaign_master.csv")

    # mapping path may be relative to campaign dir
    mapping_path = args.mapping
    if not os.path.isabs(mapping_path):
        mapping_path = os.path.join(camp_dir, mapping_path)

    # Load data
    if not os.path.isfile(master_path):
        raise FileNotFoundError(f"Master CSV not found: {master_path}")
    if not os.path.isfile(mapping_path):
        raise FileNotFoundError(f"Mapping CSV not found: {mapping_path}")

    mapping_rows, _ = read_csv(mapping_path)
    zip_index = load_master_zip_index(master_path)

    # Build executed log rows with ZIP5 from master
    exec_rows: List[Dict[str,str]] = []
    for m in mapping_rows:
        prop_addr = norm_space(m.get("property_address","") or m.get("PropertyAddress",""))
        owner = norm_space(m.get("owner","") or m.get("Owner","") or m.get("OwnerName",""))
        page = m.get("page","")
        ref_code = m.get("ref_code","")
        single_pdf = m.get("single_pdf","")
        template_ref = m.get("template_ref","")
        # ZIP5 join
        zip5 = zip_index.get(standardize_address(prop_addr), "")

        exec_rows.append({
            "CampaignName": str(args.campaign_name),
            "CampaignNumber": str(args.campaign_number),
            "CampaignDir": os.path.basename(camp_dir),
            "PropertyAddress": prop_addr,
            "OwnerName": owner,
            "ZIP5": zip5,
            "TemplateRef": template_ref,
            "RefCode": ref_code,
            "Page": str(page),
            "SinglePDF": single_pdf
        })

    # Save execution log in campaign dir
    log_path = os.path.join(camp_dir, "executed_campaign_log.csv")
    write_csv(log_path, exec_rows, [
        "CampaignName","CampaignNumber","CampaignDir",
        "PropertyAddress","OwnerName","ZIP5","TemplateRef","RefCode","Page","SinglePDF"
    ])
    print(f"[OK] Execution log: {log_path} (rows={len(exec_rows)})")

    # Update trackers
    tracker = load_tracker()
    zip_tally = load_zip_tally()
    today = datetime.date.today().isoformat()

    for r in exec_rows:
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
            entry["ZIP5"] = zip5 or entry.get("ZIP5","")
            entry["LastTemplateRef"] = r["TemplateRef"]
            entry["LastRefCode"] = r["RefCode"]
            entry["LastPage"] = r["Page"]
            entry["CampaignName"] = args.campaign_name
            entry["CampaignNumber"] = str(args.campaign_number)
            entry["LastUpdatedDate"] = today
            # bump count
            try:
                entry["LettersSent"] = str(int(entry.get("LettersSent","0")) + 1)
            except Exception:
                entry["LettersSent"] = "1"
            if not entry.get("FirstSentDate"):
                entry["FirstSentDate"] = today

        # zip tally
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

    # Save trackers
    save_tracker(tracker)
    save_zip_tally(zip_tally)
    print(f"[OK] Master tracker updated: {TRACKER_FILE}")
    print(f"[OK] ZIP5 tally updated: {ZIP_TALLY_FILE}")

if __name__ == "__main__":
    main()
