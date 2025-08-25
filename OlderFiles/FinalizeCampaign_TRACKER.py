
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FinalizeCampaign_TRACKER
------------------------
Purpose:
  - Append executed campaign rows to the campaign's executed log (with ZIP5 backfill).
  - Update MasterCampaignTracker/MasterPropertyCampaignTracker.csv with a simplified schema:
      PropertyAddress, OwnerName, ZIP5, CampaignCount, FirstSentDt, LastSentDt,
      CampaignNumbers, TemplateIds
  - Rebuild MasterCampaignTracker/Zip5_LetterTally.csv by scanning all executed logs (deduped).
  - Preserve prior functionality: dry-run, force-recount append-all, optional history file.

Run from the MailCampaigns root, e.g.:
  python FinalizeCampaign_TRACKER.py \
    --campaign-dir "Campaign_1_Aug2025" \
    --campaign-name "Campaign" \
    --campaign-number 1

Key behaviors:
  - Mapping auto-detection: looks for letters_mapping.csv in campaign dir or RefFiles/.
  - ZIP5 resolution order: mapping row -> campaign_master.csv -> common ZIP columns -> regex from address.
  - Dedup for executed log (unless --force-recount): avoids re-adding same (owner,address,page,template_ref).
  - Tracker rename: writes to MasterPropertyCampaignTracker.csv (auto-migrates from old counter name if present).
"""

import os, csv, re, sys, argparse, datetime, glob
from typing import Dict, List, Tuple, Optional

MAIL_ROOT_DEFAULT = os.getcwd()

def read_csv(path: str) -> List[Dict[str,str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [{k:(v or "").strip() for k,v in row.items()} for row in r]

def write_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

def append_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    exists = os.path.exists(path)
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        if not exists:
            w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).upper()

def key_pair(addr: str, owner: str) -> Tuple[str,str]:
    return (norm(addr), norm(owner))

ZIP_COLS_MAIL = ["Mail ZIP", "Mailing ZIP", "MAIL ZIP", "Mail ZIP Code", "Mail Zip", "MAIL ZIP CODE"]
ZIP_COLS_GENERIC = ["ZIP", "Zip", "Zip Code", "ZIP CODE"]
ZIP_COLS_SITUS = ["SITUS ZIP", "SITUS ZIP CODE", "Situs Zip", "Property ZIP", "Property Zip"]
ADDR_COLS = ["Mailing Address", "MAILING ADDRESS", "Property Address", "PROPERTY ADDRESS", "Situs Address", "SITUS ADDRESS", "Address", "ADDRESS", "SITUS"]

def extract_zip5_from_text(text: str) -> str:
    m = re.search(r"\b(\d{5})(?:-\d{4})?\b", text or "")
    return m.group(1) if m else ""

def resolve_zip5(row: Dict[str,str], master_idx: Dict[Tuple[str,str], Dict[str,str]]) -> str:
    # 1) ZIP5 already on row?
    if "ZIP5" in row and re.fullmatch(r"\d{5}", row.get("ZIP5","")):
        return row["ZIP5"]
    # Try to look up in campaign_master index using owner/address
    addr = row.get("property_address") or row.get("PropertyAddress") or row.get("Address") or ""
    owner = row.get("owner") or row.get("OwnerName") or row.get("Primary Name") or row.get("Owner") or ""
    k = key_pair(addr, owner)
    if k in master_idx:
        r = master_idx[k]
        z = r.get("ZIP5","")
        if re.fullmatch(r"\d{5}", z): return z
        # mailing zip
        for col in ZIP_COLS_MAIL + ZIP_COLS_GENERIC + ZIP_COLS_SITUS:
            if col in r and r[col].strip():
                z = extract_zip5_from_text(r[col])
                if re.fullmatch(r"\d{5}", z): return z
        # from addresses
        for col in ADDR_COLS:
            if col in r and r[col].strip():
                z = extract_zip5_from_text(r[col])
                if re.fullmatch(r"\d{5}", z): return z
    # 2) try mapping row columns
    for col in ZIP_COLS_MAIL + ZIP_COLS_GENERIC + ZIP_COLS_SITUS:
        if col in row and row[col].strip():
            z = extract_zip5_from_text(row[col])
            if re.fullmatch(r"\d{5}", z): return z
    # 3) regex from address in mapping
    for col in ["property_address","PropertyAddress","Address"] + ADDR_COLS:
        if col in row and row[col].strip():
            z = extract_zip5_from_text(row[col])
            if re.fullmatch(r"\d{5}", z): return z
    return ""

def load_campaign_master_index(camp_dir: str) -> Dict[Tuple[str,str], Dict[str,str]]:
    path = os.path.join(camp_dir, "campaign_master.csv")
    if not os.path.exists(path):
        return {}
    rows = read_csv(path)
    idx = {}
    # attempt to detect address/owner columns
    headers = rows[0].keys() if rows else []
    addr_cols = [c for c in ADDR_COLS if c in headers] + ["PropertyAddress","Address"]
    owner_cols = ["Primary Name","OwnerName","OWNER NAME","Owner","Owner Name","Name"]
    for r in rows:
        addr = ""
        for c in addr_cols:
            if c in r and r[c].strip():
                addr = r[c].strip(); break
        owner = ""
        for c in owner_cols:
            if c in r and r[c].strip():
                owner = r[c].strip(); break
        if addr and owner:
            idx[key_pair(addr, owner)] = r
    return idx

def autodetect_mapping(camp_dir: str) -> Optional[str]:
    candidates = [
        os.path.join(camp_dir, "letters_mapping.csv"),
        os.path.join(camp_dir, "RefFiles", "letters_mapping.csv"),
        os.path.join(camp_dir, "refs", "letters_mapping.csv"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    # fallback: any *mapping*.csv
    found = list(glob.glob(os.path.join(camp_dir, "**", "*mapping*.csv"), recursive=True))
    return found[0] if found else None

def dedup_key_for_log(r: Dict[str,str], campaign_name: str, campaign_number: str) -> Tuple:
    return (
        campaign_name or "",
        campaign_number or "",
        norm(r.get("owner","") or r.get("Owner","") or r.get("OwnerName","")),
        norm(r.get("property_address","") or r.get("PropertyAddress","") or r.get("Address","")),
        r.get("page","") or r.get("Page","") or "",
        r.get("ref_code","") or r.get("RefCode","") or r.get("ref","") or "",
        r.get("template_ref","") or r.get("TemplateRef","") or "",
    )

def read_existing_log_keys(path: str) -> set:
    keys = set()
    if not os.path.exists(path):
        return keys
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        for row in r:
            k = (
                row.get("CampaignName",""),
                row.get("CampaignNumber",""),
                norm(row.get("Owner","")),
                norm(row.get("PropertyAddress","")),
                row.get("Page",""),
                row.get("RefCode",""),
                row.get("TemplateRef",""),
            )
            keys.add(k)
    return keys

def rebuild_zip_tally(root: str, tally_path: str):
    # Scan all executed_campaign_log.csv files
    counts = {}
    seen = set()
    for log_path in glob.glob(os.path.join(root, "*", "executed_campaign_log.csv")):
        try:
            rows = read_csv(log_path)
        except Exception:
            continue
        camp_dir = os.path.dirname(log_path)
        for r in rows:
            k = (
                camp_dir,
                r.get("CampaignName",""),
                r.get("CampaignNumber",""),
                r.get("Page",""),
                norm(r.get("Owner","")),
                norm(r.get("PropertyAddress","")),
                r.get("TemplateRef",""),
                r.get("RefCode",""),
            )
            if k in seen:
                continue
            seen.add(k)
            z = r.get("ZIP5","")
            if not re.fullmatch(r"\d{5}", z):
                continue
            counts[z] = counts.get(z, 0) + 1
    out_rows = [{"ZIP5": z, "LettersSent": str(counts[z])} for z in sorted(counts.keys())]
    write_csv(tally_path, out_rows, ["ZIP5","LettersSent"])

def parse_args():
    ap = argparse.ArgumentParser(description="Finalize a campaign and update MasterPropertyCampaignTracker.")
    ap.add_argument("--campaign-dir", required=True, help="Campaign folder (e.g., Campaign_1_Aug2025)")
    ap.add_argument("--campaign-name", default="", help="Campaign name (metadata; optional)")
    ap.add_argument("--campaign-number", default="", help="Campaign number as string/int (metadata; optional)")
    ap.add_argument("--mapping", default="", help="Path to letters_mapping.csv (auto-detect if omitted)")
    ap.add_argument("--tracker-path", default=os.path.join("MasterCampaignTracker","MasterPropertyCampaignTracker.csv"),
                    help="Path to the master tracker CSV.")
    ap.add_argument("--history-path", default=os.path.join("MasterCampaignTracker","MasterPairHistory.csv"),
                    help="Optional detailed history CSV (only if --write-history).")
    ap.add_argument("--write-history", action="store_true", help="Write per-piece history to history CSV.")
    ap.add_argument("--dry-run", action="store_true", help="Print what would change; write nothing.")
    ap.add_argument("--force-recount", action="store_true", help="Append all mapping rows to executed log even if duplicate keys exist.")
    return ap.parse_args()

def main():
    args = parse_args()
    root = MAIL_ROOT_DEFAULT
    camp_dir = args.campaign_dir
    campaign_name = args.campaign_name
    campaign_number = str(args.campaign_number or "").strip()
    now_date = datetime.date.today().isoformat()

    if not os.path.isdir(camp_dir):
        print(f"[ERROR] Campaign directory not found: {camp_dir}")
        sys.exit(1)

    # Mapping
    mapping_path = args.mapping or autodetect_mapping(camp_dir)
    if not mapping_path or not os.path.exists(mapping_path):
        print("[ERROR] Could not find letters_mapping.csv. Use --mapping to provide a path.")
        sys.exit(1)
    mapping_rows = read_csv(mapping_path)

    # Build index from campaign_master.csv for ZIP backfill
    master_idx = load_campaign_master_index(camp_dir)

    # Prepare executed log path (in the campaign dir)
    executed_log = os.path.join(camp_dir, "executed_campaign_log.csv")
    existing_keys = set() if args.force_recount else read_existing_log_keys(executed_log)

    # Normalize + backfill ZIP + build rows to append
    appended = []
    skip = 0
    for r in mapping_rows:
        owner = r.get("owner") or r.get("Owner") or r.get("OwnerName") or ""
        prop = r.get("property_address") or r.get("PropertyAddress") or r.get("Address") or ""
        if not owner or not prop:
            # skip malformed rows quietly
            continue
        z = resolve_zip5(r, master_idx)
        k = dedup_key_for_log(r, campaign_name, campaign_number)
        if not args.force_recount and k in existing_keys:
            skip += 1
            continue

        appended.append({
            "Date": now_date,
            "CampaignName": campaign_name,
            "CampaignNumber": campaign_number,
            "Page": r.get("page","") or r.get("Page",""),
            "Owner": owner,
            "PropertyAddress": prop,
            "ZIP5": z,
            "TemplateRef": r.get("template_ref","") or r.get("TemplateRef",""),
            "RefCode": r.get("ref_code","") or r.get("RefCode",""),
        })

    print(f"[SUMMARY] Mapping rows: {len(mapping_rows)} | Already logged (skipped): {skip} | To add now: {len(appended)}")

    # Tracker path (rename support): prefer new, else migrate old name
    tracker_path = args.tracker_path
    if (not os.path.exists(tracker_path)) and os.path.exists(os.path.join("MasterCampaignTracker","MasterPropertyCampaignCounter.csv")):
        tracker_path = os.path.join("MasterCampaignTracker","MasterPropertyCampaignTracker.csv")

    # Load tracker
    tracker_rows = []
    if os.path.exists(tracker_path):
        tracker_rows = read_csv(tracker_path)

    # Index tracker by (addr, owner)
    tracker_idx = { key_pair(t.get("PropertyAddress",""), t.get("OwnerName","")): t for t in tracker_rows }

    # Update tracker with appended rows
    for a in appended:
        k = key_pair(a["PropertyAddress"], a["Owner"])
        t = tracker_idx.get(k)
        if not t:
            t = {
                "PropertyAddress": a["PropertyAddress"],
                "OwnerName": a["Owner"],
                "ZIP5": a["ZIP5"],
                "CampaignCount": "1",
                "FirstSentDt": now_date,
                "LastSentDt": now_date,
                "CampaignNumbers": a["CampaignNumber"],
                "TemplateIds": a["TemplateRef"],
            }
            tracker_idx[k] = t
        else:
            # update count
            try:
                c = int(t.get("CampaignCount","0") or "0")
            except ValueError:
                c = 0
            t["CampaignCount"] = str(c + 1)
            # update dates
            if not t.get("FirstSentDt"):
                t["FirstSentDt"] = now_date
            t["LastSentDt"] = now_date
            # update ZIP if missing
            if (not t.get("ZIP5")) and a["ZIP5"]:
                t["ZIP5"] = a["ZIP5"]
            # append lists
            def append_token(s: str, tok: str) -> str:
                s = (s or "").strip()
                tok = (tok or "").strip()
                if not tok:
                    return s
                if not s:
                    return tok
                # avoid exact dup of last token
                parts = [p for p in s.split("|") if p]
                if not parts or parts[-1] != tok:
                    parts.append(tok)
                return "|".join(parts)
            t["CampaignNumbers"] = append_token(t.get("CampaignNumbers",""), a["CampaignNumber"])
            t["TemplateIds"] = append_token(t.get("TemplateIds",""), a["TemplateRef"])

    # Materialize tracker rows list in a stable order
    headers_tracker = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    out_tracker_rows = list(tracker_idx.values())

    if args.dry_run:
        print("[DRY RUN] No changes written.")
        return

    # Write executed log
    log_headers = ["Date","CampaignName","CampaignNumber","Page","Owner","PropertyAddress","ZIP5","TemplateRef","RefCode"]
    if appended:
        append_csv(executed_log, appended, log_headers)
        print(f"[OK] Appended {len(appended)} rows to {executed_log}")
    else:
        print("[OK] Executed log already up to date.")

    # Write tracker
    write_csv(tracker_path, out_tracker_rows, headers_tracker)
    print(f"[OK] Master tracker updated: {tracker_path}")

    # Optional history (kept functionality)
    if args.write_history:
        hist_path = args.history_path
        hist_headers = ["Date","CampaignName","CampaignNumber","Page","Owner","PropertyAddress","ZIP5","TemplateRef","RefCode"]
        append_csv(hist_path, appended, hist_headers)
        print(f"[OK] Pair history appended: {hist_path}")

    # Rebuild ZIP tally from all executed logs
    tally_path = os.path.join("MasterCampaignTracker","Zip5_LetterTally.csv")
    rebuild_zip_tally(root, tally_path)
    print(f"[OK] ZIP5 tally rebuilt: {tally_path}")

if __name__ == "__main__":
    main()
