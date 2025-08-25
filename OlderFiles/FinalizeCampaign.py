
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# FinalizeCampaign.py
# -------------------
# After letters are generated/printed, finalize the campaign:
#   - Write executed_campaign_log.csv in the campaign folder (rich log).
#   - Update MasterCampaignTracker/MasterPropertyCampaignCounter.csv
#     * Tracks (PropertyAddress, OwnerName) pair
#     * Maintains ZIP5, LettersSent counter, last campaign meta
#   - Update MasterCampaignTracker/Zip5_LetterTally.csv (rolling totals per ZIP5)
#
# Usage:
#   python FinalizeCampaign.py --campaign-name "Campaign" --campaign-number 1 --campaign-dir "Campaign_1_Aug2025" --mapping "RefFiles\\letters_mapping.csv"
#   # Or omit --mapping to auto-detect the newest *_refs.csv (or letters_mapping.csv) under RefFiles/
#
import os, csv, argparse, datetime, re, glob
from typing import Dict, Tuple, List, Optional

TRACKER_DIR = "MasterCampaignTracker"
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
    # Build { ADDRESS(line1 upper) -> ZIP5 } from campaign_master.csv.
    # Prefers 'ZIP' or 'ZIP5', falls back to 'Mail ZIP' or 'MAIL ZIP' if needed.
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
    # Fallback to common name
    cand = os.path.join(ref_dir, "letters_mapping.csv")
    return cand if os.path.isfile(cand) else None

# ----------------- Main -----------------

def main():
    ap = argparse.ArgumentParser(description="Finalize campaign, write executed log (with ZIP5), and update trackers.")
    ap.add_argument("--campaign-name", required=True)
    ap.add_argument("--campaign-number", required=True, type=int)
    ap.add_argument("--campaign-dir", required=True, help="Folder containing campaign_master.csv and RefFiles/")
    ap.add_argument("--mapping", help="Path to mapping CSV relative to campaign-dir or absolute; if omitted, auto-detects under RefFiles/")
    args = ap.parse_args()

    # Resolve paths
    camp_dir = args.campaign_dir
    if not os.path.isabs(camp_dir):
        camp_dir = os.path.abspath(camp_dir)
    master_path = os.path.join(camp_dir, "campaign_master.csv")

    # Mapping path
    mapping_path = args.mapping
    if mapping_path:
        if not os.path.isabs(mapping_path):
            mapping_path = os.path.join(camp_dir, mapping_path)
    else:
        mapping_path = autodetect_mapping(camp_dir)

    # Validate inputs
    if not os.path.isfile(master_path):
        raise FileNotFoundError(f"Master CSV not found: {master_path}")
    if not mapping_path or not os.path.isfile(mapping_path):
        raise FileNotFoundError(f"Mapping CSV not found. Pass --mapping or ensure a *_refs.csv/letters_mapping.csv exists in {os.path.join(camp_dir,'RefFiles')}")

    mapping_rows, _ = read_csv(mapping_path)
    zip_index = load_master_zip_index(master_path)

    # Build executed log rows with ZIP5 from master
    now_iso = datetime.datetime.now().isoformat(timespec="seconds")
    exec_rows: List[Dict[str,str]] = []
    combined_pdf_guess = ""  # optional convenience: capture a likely combined path
    comb_dir = os.path.join(camp_dir, "BatchLetterFiles")
    if os.path.isdir(comb_dir):
        pdfs = sorted(glob.glob(os.path.join(comb_dir, "*.pdf")), key=os.path.getmtime, reverse=True)
        if pdfs:
            combined_pdf_guess = pdfs[0]

    for m in mapping_rows:
        prop_addr = norm_space(m.get("property_address","") or m.get("PropertyAddress",""))
        owner = norm_space(m.get("owner","") or m.get("Owner","") or m.get("OwnerName",""))
        page = m.get("page","")
        ref_code = m.get("ref_code","")
        single_pdf = m.get("single_pdf","")
        template_ref = m.get("template_ref","")
        template_src = m.get("template_source","")
        # ZIP5 join
        zip5 = zip_index.get(standardize_address(prop_addr), "")

        exec_rows.append({
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

    # Save execution log in campaign dir
    log_path = os.path.join(camp_dir, "executed_campaign_log.csv")
    write_csv(log_path, exec_rows, [
        "RunTimestamp",
        "CampaignName","CampaignNumber","CampaignDir",
        "PropertyAddress","OwnerName","ZIP5",
        "TemplateRef","TemplateSource","RefCode","Page",
        "CombinedPDF","SinglePDF"
    ])
    print(f"[OK] Execution log: {log_path} (rows={len(exec_rows)})")

    # Update trackers
    tracker = load_tracker_index()
    zip_tally = load_zip_tally_index()
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
                "CampaignNumber": str(args.campaign_number),  # last campaign number
                "FirstSentDate": today,
                "LastUpdatedDate": today,
                "LettersSent": "1"
            }
            tracker[k] = entry
        else:
            # keep latest metadata and bump counts
            if zip5:
                entry["ZIP5"] = zip5
            entry["LastTemplateRef"] = r["TemplateRef"]
            entry["LastRefCode"] = r["RefCode"]
            entry["LastPage"] = r["Page"]
            entry["CampaignName"] = args.campaign_name
            entry["CampaignNumber"] = str(args.campaign_number)  # last campaign number
            entry["LastUpdatedDate"] = today
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
    save_tracker_index(tracker)
    save_zip_tally_index(zip_tally)
    print(f"[OK] Master tracker updated: {TRACKER_FILE}")
    print(f"[OK] ZIP5 tally updated: {ZIP_TALLY_FILE}")

if __name__ == "__main__":
    main()
