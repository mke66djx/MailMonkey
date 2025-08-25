
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FinalizeCampaign_TRACKER_STRICT
-------------------------------
Finalize a campaign by updating the executed log, the master tracker, and ZIP5 tallies.
This "strict" version **requires** --campaign-dir and does **not** auto-pick a campaign.
It will derive campaign name/number from the given folder if you don't pass them explicitly.

Deriving rules:
- Campaign name:
    1) If --campaign-name provided, use it.
    2) Else parse "<Name>_<Number>_<MonYYYY>" pattern from folder name → Name.
    3) Else default to the folder's base name.
- Campaign number:
    1) If --campaign-number provided, use it.
    2) Else parse "<Name>_<Number>_<MonYYYY>" pattern from folder name → Number.
    3) Else, if executed_campaign_log.csv exists and has CampaignNumber values → use the most common.
    4) Else, if mapping has CampaignNumber column with a single distinct value → use it.
    5) Else, require explicit --campaign-number.

Mapping file:
- Searched only **within the provided campaign folder**, in this order:
    <camp>/letters_mapping.csv
    <camp>/RefFiles/letters_mapping.csv
    <camp>/**/*mapping*.csv (first match)

Usage (from MailCampaigns root or anywhere):
  python FinalizeCampaign_TRACKER_STRICT.py ^
    --campaign-dir "Campaign_1_Aug2025" ^
    --dry-run

Optional overrides:
  --campaign-name "Campaign"
  --campaign-number 1
  --tracker-path "MasterCampaignTracker/MasterPropertyCampaignTracker.csv"
  --force-recount (append even if keys already present)
  --write-history (also append to MasterCampaignTracker/MasterPairHistory.csv)

Outputs/updates:
- <camp>/executed_campaign_log.csv
- MasterCampaignTracker/MasterPropertyCampaignTracker.csv
- MasterCampaignTracker/Zip5_LetterTally.csv
"""

import os, csv, re, sys, argparse, datetime, glob, collections
from typing import Dict, List, Tuple, Optional

# ---------------- CSV helpers ----------------

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

# ---------------- Normalization & keys ----------------

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).upper()

def key_pair(addr: str, owner: str) -> Tuple[str,str]:
    return (norm(addr), norm(owner))

# ---------------- ZIP resolution ----------------

ZIP_COLS_MAIL = ["Mail ZIP", "Mailing ZIP", "MAIL ZIP", "Mail ZIP Code", "Mail Zip", "MAIL ZIP CODE"]
ZIP_COLS_GENERIC = ["ZIP", "Zip", "Zip Code", "ZIP CODE"]
ZIP_COLS_SITUS = ["SITUS ZIP", "SITUS ZIP CODE", "Situs Zip", "Property ZIP", "Property Zip"]
ADDR_COLS = ["Mailing Address", "MAILING ADDRESS", "Property Address", "PROPERTY ADDRESS", "Situs Address", "SITUS ADDRESS", "Address", "ADDRESS", "SITUS"]

def extract_zip5_from_text(text: str) -> str:
    m = re.search(r"\b(\d{5})(?:-\d{4})?\b", text or "")
    return m.group(1) if m else ""

def resolve_zip5(row: Dict[str,str], master_idx: Dict[Tuple[str,str], Dict[str,str]]) -> str:
    """Prefer ZIP5 from campaign_master, default to parsing any ZIP-like fields or address text."""
    if "ZIP5" in row and re.fullmatch(r"\d{5}", row.get("ZIP5","")):
        return row["ZIP5"]
    # Match by (addr, owner) into campaign_master index
    addr = row.get("property_address") or row.get("PropertyAddress") or row.get("Address") or ""
    owner = row.get("owner") or row.get("Owner") or row.get("OwnerName") or row.get("Primary Name") or ""
    k = key_pair(addr, owner)
    if k in master_idx:
        r = master_idx[k]
        z = r.get("ZIP5","")
        if re.fullmatch(r"\d{5}", z): return z
        # try common ZIP columns in the source row
        for col in ZIP_COLS_SITUS + ZIP_COLS_MAIL + ZIP_COLS_GENERIC:
            if col in r and r[col].strip():
                z = extract_zip5_from_text(r[col])
                if re.fullmatch(r"\d{5}", z):
                    return z
        # scan address-like fields
        for col in ADDR_COLS:
            if col in r and r[col].strip():
                z = extract_zip5_from_text(r[col])
                if re.fullmatch(r"\d{5}", z):
                    return z
    # fallback: parse ZIP from mapping row itself
    for col in ZIP_COLS_SITUS + ZIP_COLS_MAIL + ZIP_COLS_GENERIC:
        if col in row and row[col].strip():
            z = extract_zip5_from_text(row[col])
            if re.fullmatch(r"\d{5}", z):
                return z
    for col in ["property_address","PropertyAddress","Address"] + ADDR_COLS:
        if col in row and row[col].strip():
            z = extract_zip5_from_text(row[col])
            if re.fullmatch(r"\d{5}", z):
                return z
    return ""

def load_campaign_master_index(camp_dir: str) -> Dict[Tuple[str,str], Dict[str,str]]:
    path = os.path.join(camp_dir, "campaign_master.csv")
    if not os.path.exists(path):
        return {}
    rows = read_csv(path)
    idx = {}
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

# ---------------- Campaign metadata (STRICT) ----------------

CAMPAIGN_NAME_NUM_RE = re.compile(r"^(?P<name>.+?)_(?P<num>\d+)_")

def derive_campaign_name_number(camp_dir: str, name_arg: str, num_arg: str, mapping_rows: List[Dict[str,str]]) -> Tuple[str,str]:
    base = os.path.basename(os.path.normpath(camp_dir))
    name = (name_arg or "").strip()
    num  = (num_arg or "").strip()

    # Name
    if not name:
        m = CAMPAIGN_NAME_NUM_RE.match(base)
        if m:
            name = m.group("name")
        else:
            name = base

    # Number
    if not num:
        m = CAMPAIGN_NAME_NUM_RE.match(base)
        if m:
            num = m.group("num")

    # Try executed log (most common number) if still empty
    if not num:
        log_path = os.path.join(camp_dir, "executed_campaign_log.csv")
        if os.path.exists(log_path):
            try:
                nums = [r.get("CampaignNumber","").strip() for r in read_csv(log_path) if (r.get("CampaignNumber","").strip())]
                if nums:
                    num = collections.Counter(nums).most_common(1)[0][0]
            except Exception:
                pass

    # Try mapping column if still empty and uniform
    if not num:
        vals = set()
        for r in mapping_rows:
            v = r.get("CampaignNumber","").strip()
            if v:
                vals.add(v)
        if len(vals) == 1:
            num = list(vals)[0]

    if not num:
        print("[ERROR] Could not determine campaign number. Pass --campaign-number.")
        sys.exit(1)

    return name, num

# ---------------- Mapping discovery (STRICT to camp dir) ----------------

def find_mapping_in_campaign(camp_dir: str) -> Optional[str]:
    candidates = [
        os.path.join(camp_dir, "letters_mapping.csv"),
        os.path.join(camp_dir, "RefFiles", "letters_mapping.csv"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    found = list(glob.glob(os.path.join(camp_dir, "**", "*mapping*.csv"), recursive=True))
    return found[0] if found else None

# ---------------- Dedupe keys for executed log ----------------

def dedup_key_for_log(r: Dict[str,str], campaign_name: str, campaign_number: str) -> Tuple:
    return (
        campaign_name or "",
        campaign_number or "",
        norm(r.get("Owner","") or r.get("owner","") or r.get("OwnerName","")),
        norm(r.get("PropertyAddress","") or r.get("property_address","") or r.get("Address","")),
        r.get("Page","") or r.get("page","") or "",
        r.get("RefCode","") or r.get("ref_code","") or r.get("ref","") or "",
        r.get("TemplateRef","") or r.get("template_ref","") or "",
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

# ---------------- ZIP Tally ----------------

def rebuild_zip_tally(root: str, tally_path: str):
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

# ---------------- Tracker update ----------------

def update_tracker(tracker_path: str, appended_rows: List[Dict[str,str]], today: str):
    tracker_rows = read_csv(tracker_path) if os.path.exists(tracker_path) else []
    tracker_idx = { key_pair(t.get("PropertyAddress",""), t.get("OwnerName","")): t for t in tracker_rows }
    for a in appended_rows:
        k = key_pair(a["PropertyAddress"], a["Owner"])
        t = tracker_idx.get(k)
        if not t:
            t = {
                "PropertyAddress": a["PropertyAddress"],
                "OwnerName": a["Owner"],
                "ZIP5": a["ZIP5"],
                "CampaignCount": "1",
                "FirstSentDt": today,
                "LastSentDt": today,
                "CampaignNumbers": a["CampaignNumber"],
                "TemplateIds": a["TemplateRef"],
            }
            tracker_idx[k] = t
        else:
            try:
                c = int(t.get("CampaignCount","0") or "0")
            except ValueError:
                c = 0
            t["CampaignCount"] = str(c + 1)
            if not t.get("FirstSentDt"):
                t["FirstSentDt"] = today
            t["LastSentDt"] = today
            if (not t.get("ZIP5")) and a["ZIP5"]:
                t["ZIP5"] = a["ZIP5"]
            # Append lists without duplicating adjacent identical values
            def append_tok(s: str, tok: str) -> str:
                s = (s or "").strip()
                tok = (tok or "").strip()
                if not tok: return s
                if not s: return tok
                parts = [p for p in s.split("|") if p]
                if not parts or parts[-1] != tok:
                    parts.append(tok)
                return "|".join(parts)
            t["CampaignNumbers"] = append_tok(t.get("CampaignNumbers",""), a["CampaignNumber"])
            t["TemplateIds"] = append_tok(t.get("TemplateIds",""), a["TemplateRef"])
    headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    write_csv(tracker_path, list(tracker_idx.values()), headers)

# ---------------- Main ----------------

def parse_args():
    ap = argparse.ArgumentParser(description="Finalize a campaign (STRICT) — requires --campaign-dir; no cross-folder inference.")
    ap.add_argument("--campaign-dir", required=True, help="Campaign folder (e.g., Campaign_1_Aug2025). Required.")
    ap.add_argument("--campaign-name", default="", help="Optional name; if omitted, parsed from folder name.")
    ap.add_argument("--campaign-number", default="", help="Optional number; if omitted, parsed or derived from files in the folder.")
    ap.add_argument("--mapping", default="", help="Optional mapping CSV path inside the campaign folder. If omitted, auto-find within that folder only.")
    ap.add_argument("--tracker-path", default=os.path.join("MasterCampaignTracker","MasterPropertyCampaignTracker.csv"),
                    help="Path to the master tracker CSV.")
    ap.add_argument("--write-history", action="store_true", help="Also append to MasterCampaignTracker/MasterPairHistory.csv")
    ap.add_argument("--history-path", default=os.path.join("MasterCampaignTracker","MasterPairHistory.csv"))
    ap.add_argument("--dry-run", action="store_true", help="Print what would change; write nothing.")
    ap.add_argument("--force-recount", action="store_true", help="Append all mapping rows to executed log even if duplicate keys exist.")
    return ap.parse_args()

def main():
    args = parse_args()
    camp_dir = args.campaign_dir
    if not os.path.isdir(camp_dir):
        print(f"[ERROR] --campaign-dir not found: {camp_dir}")
        sys.exit(1)

    # Mapping inside the given folder
    mapping_path = args.mapping or find_mapping_in_campaign(camp_dir)
    if not mapping_path or not os.path.exists(mapping_path):
        print("[ERROR] Could not find letters_mapping.csv inside the campaign folder. Use --mapping to provide a path.")
        sys.exit(1)
    mapping_rows = read_csv(mapping_path)

    # Derive campaign name/number strictly from this folder / provided args / its own files
    campaign_name, campaign_number = derive_campaign_name_number(camp_dir, args.campaign_name, args.campaign_number, mapping_rows)

    # Build index from campaign_master.csv for ZIP backfill
    master_idx = load_campaign_master_index(camp_dir)

    # Prepare executed log path (+ dedupe keys, unless force-recount)
    executed_log = os.path.join(camp_dir, "executed_campaign_log.csv")
    existing_keys = set() if args.force_recount else read_existing_log_keys(executed_log)

    today = datetime.date.today().isoformat()

    # Normalize + backfill ZIP + build rows to append
    appended = []
    skip = 0
    for r in mapping_rows:
        owner = r.get("owner") or r.get("Owner") or r.get("OwnerName") or ""
        prop  = r.get("property_address") or r.get("PropertyAddress") or r.get("Address") or ""
        if not owner or not prop:
            continue
        z = resolve_zip5(r, master_idx)
        k = (
            campaign_name or "",
            campaign_number or "",
            norm(owner), norm(prop),
            r.get("page","") or r.get("Page",""),
            r.get("ref_code","") or r.get("RefCode","") or r.get("ref","") or "",
            r.get("template_ref","") or r.get("TemplateRef","") or "",
        )
        if k in existing_keys and not args.force_recount:
            skip += 1
            continue
        appended.append({
            "Date": today,
            "CampaignName": campaign_name,
            "CampaignNumber": campaign_number,
            "Page": r.get("page","") or r.get("Page",""),
            "Owner": owner,
            "PropertyAddress": prop,
            "ZIP5": z,
            "TemplateRef": r.get("template_ref","") or r.get("TemplateRef",""),
            "RefCode": r.get("ref_code","") or r.get("RefCode",""),
        })

    print(f"[USING] campaign-dir={camp_dir}")
    print(f"[DERIVED] campaign-name={campaign_name!r}  campaign-number={campaign_number!r}")
    print(f"[SUMMARY] Mapping rows: {len(mapping_rows)} | Already logged (skipped): {skip} | To add now: {len(appended)}")

    if args.dry_run:
        print("[DRY RUN] No changes written.")
        return

    # Append executed log
    log_headers = ["Date","CampaignName","CampaignNumber","Page","Owner","PropertyAddress","ZIP5","TemplateRef","RefCode"]
    if appended:
        append_csv(executed_log, appended, log_headers)
        print(f"[OK] Appended {len(appended)} rows to {executed_log}")
    else:
        print("[OK] Executed log already up to date.")

    # Update tracker
    tracker_path = args.tracker_path
    update_tracker(tracker_path, appended, today)
    print(f"[OK] Master tracker updated: {tracker_path}")

    # Optional history
    if args.write_history and appended:
        hist_headers = log_headers
        append_csv(args.history_path, appended, hist_headers)
        print(f"[OK] Pair history appended: {args.history_path}")

    # Rebuild ZIP tally (root = parent directory of the campaign dir)
    root = os.path.abspath(os.path.join(camp_dir, os.pardir))
    tally_path = os.path.join(root, "MasterCampaignTracker","Zip5_LetterTally.csv")
    os.makedirs(os.path.dirname(tally_path), exist_ok=True)
    rebuild_zip_tally(root, tally_path)
    print(f"[OK] ZIP5 tally rebuilt: {tally_path}")

if __name__ == "__main__":
    main()
