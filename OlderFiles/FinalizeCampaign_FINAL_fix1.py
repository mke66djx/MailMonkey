#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# FinalizeCampaign_FINAL.py
#
# Backwards-compatible with your SAFE script. Adds OPTIONAL full-history logging
# without removing any existing behavior.
#
# Default (no new flags) = SAME behavior as SAFE:
#   - Appends to <campaign-dir>\executed_campaign_log.csv (idempotent per campaign).
#   - Updates MasterCampaignTracker\MasterPropertyCampaignCounter.csv (per-pair compact ledger).
#   - Updates MasterCampaignTracker\Zip5_LetterTally.csv.
#
# Optional new features (opt-in):
#   --write-history           -> also append every send to MasterCampaignTracker\MasterPairHistory.csv
#   --history-path <custom>   -> use a custom file for history
#   --mapping <file>          -> override mapping; optional (auto-detects if omitted)
#   --combined-pdf <file>     -> store combined PDF path in history rows
#
# Other flags (unchanged from SAFE):
#   --campaign-name (optional)  --campaign-number (required)  --campaign-dir (required)
#   --project-root "."          --dry-run  --force-recount  --verbose
#
# Usage (examples, Windows cmd.exe):
#   python FinalizeCampaign_FINAL.py ^
#     --campaign-name "Campaign" ^
#     --campaign-number 1 ^
#     --campaign-dir "Campaign_1_Aug2025"
#
#   python FinalizeCampaign_FINAL.py ^
#     --campaign-number 1 ^
#     --campaign-dir "Campaign_1_Aug2025" ^
#     --write-history
#
import os, csv, argparse, re, sys, datetime, glob
from typing import Dict, List, Tuple, Optional

# ---------- CSV helpers ----------
def read_csv(path: str):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        rows = [{k:(v or '').strip() for k,v in row.items()} for row in r]
        return rows, (r.fieldnames or [])

def write_csv(path: str, rows: List[Dict[str,str]], headers: List[str]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

def append_csv(path: str, rows: List[Dict[str,str]], headers: List[str]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    exists = os.path.isfile(path)
    with open(path, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        if not exists:
            w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

# ---------- Normalization ----------
def norm_space(s:str) -> str: return re.sub(r"\s+", " ", (s or "").strip())
def norm_owner(s:str) -> str: return norm_space(s).upper()
def norm_addr(s:str) -> str:  return norm_space(s).upper()

# ---------- Mapping auto-detect ----------
def newest(paths):
    paths = [p for p in paths if os.path.isfile(p)]
    return max(paths, key=lambda p: os.path.getmtime(p)) if paths else ""

def autodetect_mapping(campaign_dir: str, project_root: str) -> str:
    cand = os.path.join(campaign_dir, "letters_mapping.csv")
    if os.path.isfile(cand):
        return cand
    ref_dir1 = os.path.join(campaign_dir, "RefFiles")
    ref_dir2 = os.path.join(project_root, "RefFiles")
    cands = []
    for rd in (ref_dir1, ref_dir2):
        if os.path.isdir(rd):
            cands += glob.glob(os.path.join(rd, "*ref*.csv"))
            cands += glob.glob(os.path.join(rd, "*mapping*.csv"))
    return newest(cands) or ""

# ---------- Mapping normalization ----------
def pick_col(headers, options):
    low = {h.lower():h for h in headers}
    for o in options:
        if o.lower() in low:
            return low[o.lower()]
    for h in headers:
        hl = h.lower()
        for o in options:
            if o.lower() in hl:
                return h
    return None

def extract_zip5(s: str) -> str:
    if not s: return ""
    m = re.search(r"(\d{5})(?:-\d{4})?$", s)
    return m.group(1) if m else ""

def infer_mapping_fields(rows):
    headers = list(rows[0].keys()) if rows else []
    return {
        "owner": pick_col(headers, ["owner","ownername","primary name","name"]),
        "address": pick_col(headers, ["property_address","property address","situs address","address"]),
        "zip5": pick_col(headers, ["zip5","zip","situs zip","mail zip"]),
        "page": pick_col(headers, ["page"]),
        "ref": pick_col(headers, ["ref_code","ref","refcode"]),
        "tref": pick_col(headers, ["template_ref","template","templaterref"]),
    }

# ---------- Master / History headers ----------
MASTER_HEADERS = [
    "PropertyAddress","OwnerName","ZIP5",
    "LettersSent",
    "FirstSentDate","PrevSentDate","LastSentDate",
    "PrevCampaignNumber","LastCampaignNumber",
    "LastTemplateRef","LastRefCode","LastPage"
]

HISTORY_HEADERS = [
    "PropertyAddress","OwnerName","ZIP5",
    "CampaignName","CampaignNumber","SentDate",
    "TemplateRef","RefCode","Page",
    "CampaignDir","MappingPath","CombinedPDFPath","RunId"
]

ZIP_HEADERS = ["ZIP5","LettersSent","FirstSentDate","LastUpdatedDate"]

def load_master(path: str) -> Dict[Tuple[str,str], Dict[str,str]]:
    d = {}
    if os.path.isfile(path):
        rows, _ = read_csv(path)
        for r in rows:
            key = (norm_addr(r.get("PropertyAddress","")), norm_owner(r.get("OwnerName","")))
            d[key] = r
    return d

def load_history_index(path: str) -> Dict[Tuple[str,str], int]:
    idx = {}
    if os.path.isfile(path):
        rows, _ = read_csv(path)
        for r in rows:
            cnum = r.get("CampaignNumber","")
            ref  = r.get("RefCode","")
            addr = norm_addr(r.get("PropertyAddress",""))
            own  = norm_owner(r.get("OwnerName",""))
            if ref:
                idx[("REF", cnum + "||" + ref)] = 1
            idx[("PAIR", cnum + "||" + addr + "||" + own)] = 1
    return idx

def load_executed_index(path: str, cnum: str) -> Dict[Tuple[str,str], int]:
    """Index by REF and by (PAIR) for this campaign's executed log (if exists)."""
    idx = {}
    if os.path.isfile(path):
        rows, _ = read_csv(path)
        fields = infer_mapping_fields(rows) if rows else {"owner": "OwnerName","address":"PropertyAddress","ref":"RefCode"}
        ref_col = fields.get("ref") or "RefCode"
        addr_col = fields.get("address") or "PropertyAddress"
        own_col  = fields.get("owner") or "OwnerName"
        for r in rows:
            ref  = r.get(ref_col, "")
            addr = norm_addr(r.get(addr_col, ""))
            own  = norm_owner(r.get(own_col, ""))
            if ref:
                idx[("REF", cnum + "||" + ref)] = 1
            if addr and own:
                idx[("PAIR", cnum + "||" + addr + "||" + own)] = 1
    return idx

def today_iso() -> str:
    return datetime.date.today().isoformat()

def now_run_id() -> str:
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")

# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser(description="Finalize a campaign (SAFE-compatible, with optional history).")
    ap.add_argument("--campaign-name", default="", help="Optional label for your campaign (kept in per-campaign log and history).")
    ap.add_argument("--campaign-number", required=True, type=int, help="Campaign number (integer).")
    ap.add_argument("--campaign-dir", required=True, help="Campaign folder (e.g., Campaign_1_Aug2025).")
    ap.add_argument("--mapping", default="", help="Mapping CSV path. Optional; auto-detects if omitted.")
    ap.add_argument("--combined-pdf", default="", help="Optional combined PDF path to record in history.")
    ap.add_argument("--project-root", default=".", help="Folder that contains MasterCampaignTracker/. Defaults to current.")
    ap.add_argument("--write-history", action="store_true", help="Also append to MasterPairHistory.csv (optional).")
    ap.add_argument("--history-path", default="", help="Custom path for MasterPairHistory.csv (optional).")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force-recount", action="store_true")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    project_root = os.path.abspath(args.project_root)
    tracker_dir = os.path.join(project_root, "MasterCampaignTracker")
    os.makedirs(tracker_dir, exist_ok=True)

    master_path   = os.path.join(tracker_dir, "MasterPropertyCampaignCounter.csv")
    history_path  = os.path.join(tracker_dir, "MasterPairHistory.csv") if not args.history_path else os.path.abspath(args.history_path)
    ziptally_path = os.path.join(tracker_dir, "Zip5_LetterTally.csv")

    # Mapping (explicit or auto-detect)
    mapping_path = os.path.abspath(args.mapping) if args.mapping else autodetect_mapping(args.campaign_dir, project_root)
    if not mapping_path or not os.path.isfile(mapping_path):
        print(f"[ERROR] Mapping not found. Provide --mapping or place a mapping CSV under '{args.campaign_dir}\\' or 'RefFiles\\'.")
        sys.exit(1)

    # Load mapping
    mrows, mheaders = read_csv(mapping_path)
    if not mrows:
        print("[INFO] Mapping CSV is empty; nothing to do.")
        return
    fields = infer_mapping_fields(mrows)
    if args.verbose:
        print("[FIELDS]", fields)

    # Build idempotency indexes
    cnum_str = str(args.campaign_number)
    executed_log_path = os.path.join(os.path.abspath(args.campaign_dir), "executed_campaign_log.csv")
    exec_idx = load_executed_index(executed_log_path, cnum_str)
    hist_idx = load_history_index(history_path) if args.write_history else {}

    def already_logged(row, addr, owner) -> bool:
        if args.force_recount:
            return False
        ref  = row.get(fields["ref"] or "", "")
        addrN = norm_addr(addr); ownN = norm_owner(owner)
        # Executed log for this same campaign first
        if ref and ("REF", cnum_str + "||" + ref) in exec_idx:
            return True
        if ("PAIR", cnum_str + "||" + addrN + "||" + ownN) in exec_idx:
            return True
        # Optional global history second
        if args.write_history:
            if ref and ("REF", cnum_str + "||" + ref) in hist_idx:
                return True
            if ("PAIR", cnum_str + "||" + addrN + "||" + ownN) in hist_idx:
                return True
        return False

    # Load compact master and zip tallies
    master = load_master(master_path)
    existing_zip = {}
    if os.path.isfile(ziptally_path):
        zrows, _ = read_csv(ziptally_path)
        for zr in zrows:
            z = zr.get("ZIP5","").strip()
            n = int((zr.get("LettersSent","0") or "0"))
            if z:
                existing_zip[z] = existing_zip.get(z, 0) + n

    # Prepare aggregates
    run_id = now_run_id()
    sent_date = today_iso()
    combined_pdf = args.combined_pdf  # recorded in history rows when provided
    to_executed = []
    to_history  = []
    new_zip_counts = {}
    skipped = 0; added = 0

    for r in mrows:
        owner = r.get(fields["owner"] or "", "")
        addr  = r.get(fields["address"] or "", "")
        if not owner or not addr:
            continue

        page = r.get(fields["page"] or "", "")
        ref  = r.get(fields["ref"] or "", "")
        tref = r.get(fields["tref"] or "", "")
        zip5 = r.get(fields["zip5"] or "", "")
        if not zip5:
            zip5 = extract_zip5(addr) or extract_zip5(r.get("ZIP","")) or extract_zip5(r.get("Mail ZIP","")) or extract_zip5(r.get("SITUS ZIP",""))

        if already_logged(r, addr, owner):
            skipped += 1
            continue

        # Per-campaign executed log row
        to_executed.append({
            "DateTime": f"{sent_date} {datetime.datetime.now().strftime('%H:%M:%S')}",
            "CampaignName": args.campaign_name,
            "CampaignNumber": cnum_str,
            "OwnerName": owner,
            "PropertyAddress": addr,
            "ZIP5": zip5,
            "TemplateRef": tref,
            "RefCode": ref,
            "Page": page,
        })

        # Optional global history
        if args.write_history:
            to_history.append({
                "PropertyAddress": addr,
                "OwnerName": owner,
                "ZIP5": zip5,
                "CampaignName": args.campaign_name,
                "CampaignNumber": cnum_str,
                "SentDate": sent_date,
                "TemplateRef": tref,
                "RefCode": ref,
                "Page": page,
                "CampaignDir": os.path.abspath(args.campaign_dir),
                "MappingPath": mapping_path,
                "CombinedPDFPath": args.combined_pdf,
                "RunId": run_id,
            })

        # Update compact master
        key = (norm_addr(addr), norm_owner(owner))
        m = master.get(key)
        if not m:
            master[key] = {
                "PropertyAddress": addr,
                "OwnerName": owner,
                "ZIP5": zip5,
                "LettersSent": "1",
                "FirstSentDate": sent_date,
                "PrevSentDate": "",
                "LastSentDate": sent_date,
                "PrevCampaignNumber": "",
                "LastCampaignNumber": cnum_str,
                "LastTemplateRef": tref,
                "LastRefCode": ref,
                "LastPage": page,
            }
        else:
            prev_last_date = m.get("LastSentDate","")
            prev_last_cnum = m.get("LastCampaignNumber","")
            if prev_last_date:
                m["PrevSentDate"] = prev_last_date
                m["PrevCampaignNumber"] = prev_last_cnum
            m["LastSentDate"] = sent_date
            m["LastCampaignNumber"] = cnum_str
            if tref: m["LastTemplateRef"] = tref
            if ref:  m["LastRefCode"] = ref
            if page: m["LastPage"] = page
            if (not m.get("ZIP5","")) and zip5: m["ZIP5"] = zip5
            try:
                m["LettersSent"] = str(int(m.get("LettersSent","0") or "0") + 1)
            except Exception:
                m["LettersSent"] = "1"

        if zip5:
            new_zip_counts[zip5] = new_zip_counts.get(zip5, 0) + 1
        added += 1

    total = len(mrows)
    print(f"[SUMMARY] Mapping rows: {total} | Already logged (skipped): {skipped} | To add now: {added}")

    if args.dry_run:
        print("[DRY RUN] No changes written.")
        return

    # Write per-campaign executed log (append)
    exec_headers = ["DateTime","CampaignName","CampaignNumber","OwnerName","PropertyAddress","ZIP5","TemplateRef","RefCode","Page"]
    append_csv(executed_log_path, to_executed, exec_headers)
    print(f"[OK] Appended {len(to_executed)} rows to {executed_log_path}")

    # Optional global history
    if args.write_history:
        append_csv(history_path, to_history, HISTORY_HEADERS)
        print(f"[OK] History updated: {history_path} (+{len(to_history)})")

    # Write compact master
    master_rows = []
    for _, row in master.items():
        for h in MASTER_HEADERS:
            row.setdefault(h, "")
        master_rows.append(row)
    write_csv(master_path, master_rows, MASTER_HEADERS)
    print(f"[OK] Master tracker updated: {master_path}")

    # Update ZIP5 tally
    for z, n in new_zip_counts.items():
        if not z: 
            continue
        existing_zip[z] = existing_zip.get(z, 0) + n
    zip_rows = [{"ZIP5": z, "LettersSent": str(n), "FirstSentDate":"", "LastUpdatedDate": today_iso()} for z,n in sorted(existing_zip.items()) if z]
    write_csv(ziptally_path, zip_rows, ZIP_HEADERS)
    print(f"[OK] ZIP5 tally updated: {ziptally_path}")

if __name__ == "__main__":
    main()
