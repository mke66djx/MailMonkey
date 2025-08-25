#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Finalize Campaign (STRICT, idempotent) with correct CampaignCount handling,
ZIP5 backfill, and optional repair/dedupe utilities without removing prior functionality.

Key points:
- Append only NEW rows to executed_campaign_log.csv for a campaign (dedupe by
  (OwnerName, PropertyAddress, CampaignNumber) and by RefCode when available).
- Update MasterCampaignTracker/MasterPropertyCampaignTracker.csv WITHOUT
  over-incrementing: CampaignCount := distinct count of CampaignNumbers.
- Maintain FirstSentDt as the earliest date and LastSentDt as latest.
- Fill ZIP5 from mapping or campaign_master.csv when possible.
- Rebuild MasterCampaignTracker/Zip5_LetterTally.csv from all executed logs.
- Optional: --write-history to also append to MasterPairHistory.csv (off by default).
- Optional: --dedupe-log to de-duplicate an existing executed_campaign_log.csv.
- Optional: --recount-tracker (alias --repair-counts) to recompute CampaignCount
  from CampaignNumbers for all rows (fix past over-increments).
"""

import os
import csv
import re
import argparse
from datetime import datetime
from typing import Dict, List, Tuple, Optional

# ------------- Helpers -------------

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def norm_key(addr: str, owner: str) -> Tuple[str, str]:
    return (norm_space(addr).upper(), norm_space(owner).upper())

def read_csv(path: str) -> List[Dict[str, str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [ {k: (v or "").strip() for k,v in row.items()} for row in r ]

def write_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow(row)

def try_parse_date(s: str) -> Optional[datetime]:
    s = (s or "").strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None

def today_mmddyyyy() -> str:
    now = datetime.now()
    return f"{now.month}/{now.day}/{now.year}"

def find_file(*candidates: str) -> Optional[str]:
    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return None

def infer_campaign_from_dir(campaign_dir: str) -> Tuple[str, Optional[str]]:
    """Returns (campaign_name, campaign_number_str or None) inferred from folder like "Campaign_2_Aug2025"."""
    base = os.path.basename(os.path.normpath(campaign_dir))
    m = re.match(r"^(?P<name>.+?)_(?P<num>\d+)_", base)
    if m:
        return m.group("name"), m.group("num")
    return base, None

def build_zip_index_from_master(campaign_dir: str) -> Dict[Tuple[str,str], str]:
    """Build a dict (addr_norm, owner_norm) -> ZIP5 from campaign_master.csv if present."""
    idx: Dict[Tuple[str,str], str] = {}
    cm_path = os.path.join(campaign_dir, "campaign_master.csv")
    if not os.path.isfile(cm_path):
        return idx
    rows = read_csv(cm_path)

    def get_zip_from_row(r: Dict[str,str]) -> str:
        for key in ("ZIP5","Zip5","ZIP","Zip","Mail ZIP","MAIL ZIP","SITUS ZIP","SITUS ZIP CODE","Mail Zip Code","MAIL ZIP CODE"):
            if key in r and r[key].strip():
                m = re.search(r"(\d{5})", r[key])
                if m: return m.group(1)
        for key in ("Property Address","PROPERTY ADDRESS","Address","ADDRESS","Situs Address","SITUS ADDRESS","PropertyAddress","SITUS","Mailing Address","MAILING ADDRESS"):
            if key in r and r[key].strip():
                m = re.search(r"(\d{5})(?:-\d{4})?$", r[key])
                if m: return m.group(1)
        return ""

    def get_addr(r: Dict[str,str]) -> str:
        for key in ("Property Address","PROPERTY ADDRESS","Address","ADDRESS","Situs Address","SITUS ADDRESS","PropertyAddress","SITUS"):
            if key in r and r[key].strip():
                return r[key]
        return ""

    def get_owner(r: Dict[str,str]) -> str:
        for key in ("Primary Name","PRIMARY NAME","OwnerName","OWNER NAME","Owner","OWNER"):
            if key in r and r[key].strip():
                return r[key]
        # fallback compose
        f = ""; l = ""
        for fk in ("Primary First","PRIMARY FIRST","Owner First","OWNER FIRST","First Name","FIRST NAME"):
            if fk in r and r[fk].strip(): f = r[fk]; break
        for lk in ("Primary Last","PRIMARY LAST","Owner Last","OWNER LAST","Last Name","LAST NAME"):
            if lk in r and r[lk].strip(): l = r[lk]; break
        return (f + " " + l).strip()

    for r in rows:
        z = get_zip_from_row(r)
        a = get_addr(r)
        o = get_owner(r)
        if a and o and z:
            idx[norm_key(a,o)] = z
    return idx

def get_zip_from_row_generic(r: Dict[str,str]) -> str:
    for key in ("ZIP5","Zip5","ZIP","Zip","Mail ZIP","MAIL ZIP","SITUS ZIP","SITUS ZIP CODE","Mail Zip Code","MAIL ZIP CODE"):
        if key in r and r[key].strip():
            m = re.search(r"(\d{5})", r[key])
            if m: return m.group(1)
    for key in ("property_address","Property Address","PROPERTY ADDRESS","Address","ADDRESS","Situs Address","SITUS ADDRESS"):
        if key in r and r[key].strip():
            m = re.search(r"(\d{5})(?:-\d{4})?$", r[key])
            if m: return m.group(1)
    return ""

# ------------- Core -------------

def main():
    ap = argparse.ArgumentParser(description="Finalize a campaign: append executed log (idempotent), update master tracker, rebuild ZIP5 tally.")
    ap.add_argument("--campaign-dir", required=True, help="Path to the campaign folder (e.g., Campaign_2_Aug2025)")
    ap.add_argument("--campaign-name", default=None, help="Override inferred campaign name")
    ap.add_argument("--campaign-number", default=None, help="Override inferred campaign number (int-like string)")
    ap.add_argument("--mapping", default=None, help="Override mapping path; default: <campaign-dir>/RefFiles/letters_mapping.csv or <campaign-dir>/letters_mapping.csv")
    ap.add_argument("--tracker-path", default="MasterCampaignTracker/MasterPropertyCampaignTracker.csv", help="Path to master tracker CSV")
    ap.add_argument("--write-history", action="store_true", help="Also append to MasterCampaignTracker/MasterPairHistory.csv")
    ap.add_argument("--force-recount", action="store_true", help="(Advanced) Force append even if rows already exist (generally NOT recommended)")
    ap.add_argument("--dedupe-log", action="store_true", help="De-duplicate the campaign's executed log by (OwnerName, PropertyAddress, CampaignNumber) and RefCode")
    ap.add_argument("--recount-tracker", action="store_true", help="Recompute CampaignCount for all rows from CampaignNumbers (alias: --repair-counts)")
    ap.add_argument("--repair-counts", action="store_true", help="Alias for --recount-tracker")
    ap.add_argument("--dry-run", action="store_true", help="Show what would change but do not write files")
    args = ap.parse_args()

    campaign_dir = args.campaign_dir
    if not os.path.isdir(campaign_dir):
        print(f"[ERROR] campaign-dir not found: {campaign_dir}")
        return

    mapping_path = args.mapping or find_file(
        os.path.join(campaign_dir, "RefFiles", "letters_mapping.csv"),
        os.path.join(campaign_dir, "letters_mapping.csv")
    )
    if not mapping_path or not os.path.isfile(mapping_path):
        print(f"[ERROR] mapping file not found in {campaign_dir}. Looked for RefFiles/letters_mapping.csv and letters_mapping.csv")
        return

    inferred_name, inferred_num = infer_campaign_from_dir(campaign_dir)
    campaign_name = args.campaign_name or inferred_name
    campaign_number = args.campaign_number or inferred_num
    if not campaign_number:
        print("[WARN] Could not infer campaign-number from folder; please provide --campaign-number")
        return

    mapping_rows = read_csv(mapping_path)
    if not mapping_rows:
        print(f"[ERROR] Mapping file has no rows: {mapping_path}")
        return

    zip_idx = build_zip_index_from_master(campaign_dir)

    executed_log = os.path.join(campaign_dir, "executed_campaign_log.csv")
    if args.dedupe_log and os.path.isfile(executed_log):
        pre = read_csv(executed_log)
        seen_keys = set()
        seen_refs = set()
        deduped = []
        for r in pre:
            addr = r.get("PropertyAddress","") or r.get("property_address","") or r.get("Address","");
            owner = r.get("OwnerName","") or r.get("owner","") or r.get("Owner","");
            refc  = r.get("RefCode","") or r.get("ref_code","");
            campn = r.get("CampaignNumber","");
            key = (norm_key(addr, owner), (campn or "").strip())
            if refc and refc in seen_refs: 
                continue
            if key in seen_keys: 
                continue
            seen_keys.add(key)
            if refc: seen_refs.add(refc)
            deduped.append(r)
        if args.dry_run:
            print(f"[DRY RUN] Would rewrite executed log: {len(pre)} -> {len(deduped)} rows")
        else:
            headers = pre[0].keys() if pre else ["ExecutedDt","CampaignName","CampaignNumber","OwnerName","PropertyAddress","TemplateId","RefCode","ZIP5"]
            write_csv(executed_log, deduped, list(headers))
            print(f"[OK] Deduped executed log: {len(pre)} -> {len(deduped)} rows")

    existing_log = read_csv(executed_log) if os.path.isfile(executed_log) else []

    exist_pair_campaign = set()
    exist_ref = set()
    for r in existing_log:
        addr = r.get("PropertyAddress","") or r.get("property_address","") or r.get("Address","");
        owner = r.get("OwnerName","") or r.get("owner","") or r.get("Owner","");
        refc  = r.get("RefCode","") or r.get("ref_code","");
        campn = r.get("CampaignNumber","");
        exist_pair_campaign.add( (norm_key(addr, owner), (campn or "").strip()) )
        if refc:
            exist_ref.add(refc)

    to_add: List[Dict[str,str]] = []
    for r in mapping_rows:
        owner = r.get("owner","") or r.get("Owner","") or r.get("OwnerName","");
        addr  = r.get("property_address","") or r.get("Property Address","") or r.get("PropertyAddress","") or r.get("Address","");
        refc  = r.get("ref_code","") or r.get("RefCode","");
        templ = r.get("template_ref","") or r.get("template_id","") or r.get("TemplateId","") or r.get("Template","");
        z5    = r.get("ZIP5","") or get_zip_from_row_generic(r);
        if not z5 and (addr and owner):
            z5 = zip_idx.get(norm_key(addr, owner), "");

        key = (norm_key(addr, owner), str(campaign_number).strip())

        if not args.force_recount:
            if key in exist_pair_campaign or (refc and refc in exist_ref):
                continue

        to_add.append({
            "ExecutedDt": today_mmddyyyy(),
            "CampaignName": campaign_name,
            "CampaignNumber": str(campaign_number),
            "OwnerName": owner,
            "PropertyAddress": addr,
            "TemplateId": (templ or "").strip(),
            "RefCode": (refc or "").strip(),
            "ZIP5": (z5 or "").strip(),
        })

    print(f"[SUMMARY] Mapping rows: {len(mapping_rows)} | Already logged (skipped): {len(mapping_rows)-len(to_add)} | To add now: {len(to_add)}")

    if args.dry_run:
        print("[DRY RUN] No changes written.")
        return

    if to_add:
        all_rows = existing_log + to_add
        headers = list({k for row in all_rows for k in row.keys()})
        pref = ["ExecutedDt","CampaignName","CampaignNumber","OwnerName","PropertyAddress","TemplateId","RefCode","ZIP5","page","file_stub","single_pdf","template_source"]
        ordered = [h for h in pref if h in headers] + [h for h in headers if h not in pref]
        write_csv(executed_log, all_rows, ordered)
        print(f"[OK] Appended {len(to_add)} rows to {executed_log}")
    else:
        print("[OK] Nothing to append to executed log.")

    tracker_path = args.tracker_path
    os.makedirs(os.path.dirname(tracker_path), exist_ok=True)
    tracker_rows = read_csv(tracker_path) if os.path.isfile(tracker_path) else []

    idx: Dict[Tuple[str,str], Dict[str,str]] = {}
    for tr in tracker_rows:
        k = norm_key(tr.get("PropertyAddress",""), tr.get("OwnerName",""))
        idx[k] = tr

    by_pair_new: Dict[Tuple[str,str], List[Dict[str,str]]] = {}
    for r in to_add:
        k = norm_key(r["PropertyAddress"], r["OwnerName"])
        by_pair_new.setdefault(k, []).append(r)

    today_str = today_mmddyyyy()

    for k, rows in by_pair_new.items():
        r0 = rows[0]
        addr = r0["PropertyAddress"]
        owner = r0["OwnerName"]
        z5 = r0.get("ZIP5","");

        if k in idx:
            tr = idx[k]
            if (not tr.get("ZIP5","")) and z5:
                tr["ZIP5"] = z5

            existing_cns = [x for x in (tr.get("CampaignNumbers","") or "").split("|") if x]
            existing_ts  = [x for x in (tr.get("TemplateIds","") or "").split("|") if x]
            cn_set = set(existing_cns)
            t_set  = set(existing_ts)

            for rr in rows:
                cn_set.add(rr["CampaignNumber"])
                tid = rr.get("TemplateId","");
                if tid: t_set.add(tid)

            tr["CampaignNumbers"] = "|".join(sorted(cn_set, key=lambda x: int(re.sub(r"[^0-9]", "", x) or "0")))
            tr["TemplateIds"]     = "|".join(sorted(t_set))
            tr["CampaignCount"]   = str(len(cn_set))

            first_dt = try_parse_date(tr.get("FirstSentDt","")) or try_parse_date(tr.get("FirstSentDtUTC",""))
            if not first_dt:
                tr["FirstSentDt"] = today_str
            tr["LastSentDt"] = today_str

            if not tr.get("PropertyAddress","" ).strip(): tr["PropertyAddress"] = addr
            if not tr.get("OwnerName","" ).strip(): tr["OwnerName"] = owner

        else:
            cn_set = {rows[0]["CampaignNumber"]}
            t_set  = set([t for t in [rr.get("TemplateId","") for rr in rows] if t])
            idx[k] = {
                "PropertyAddress": addr,
                "OwnerName": owner,
                "ZIP5": z5,
                "CampaignCount": str(len(cn_set)),
                "FirstSentDt": today_str,
                "LastSentDt": today_str,
                "CampaignNumbers": "|".join(sorted(cn_set, key=lambda x: int(re.sub(r"[^0-9]", "", x) or "0"))),
                "TemplateIds": "|".join(sorted(t_set)),
            }

    if args.recount_tracker or args.repair_counts:
        for tr in idx.values():
            cns = [x for x in (tr.get("CampaignNumbers","") or "").split("|") if x]
            tr["CampaignCount"] = str(len(set(cns)))

    final_rows = list(idx.values())
    headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    extra = [h for h in (final_rows[0].keys() if final_rows else []) if h not in headers]
    write_csv(tracker_path, final_rows, headers + extra)
    print(f"[OK] Master tracker updated: {tracker_path}")

    if args.write_history:
        hist_path = os.path.join(os.path.dirname(tracker_path), "MasterPairHistory.csv")
        hist_rows = read_csv(hist_path) if os.path.isfile(hist_path) else []
        for r in to_add:
            hist_rows.append({
                "ExecutedDt": r["ExecutedDt"],
                "CampaignName": r["CampaignName"],
                "CampaignNumber": r["CampaignNumber"],
                "PropertyAddress": r["PropertyAddress"],
                "OwnerName": r["OwnerName"],
                "TemplateId": r.get("TemplateId",""),
                "RefCode": r.get("RefCode",""),
                "ZIP5": r.get("ZIP5",""),
            })
        if to_add:
            headers_h = ["ExecutedDt","CampaignName","CampaignNumber","PropertyAddress","OwnerName","TemplateId","RefCode","ZIP5"]
            write_csv(hist_path, hist_rows, headers_h)
            print(f"[OK] History appended: {hist_path}")

    rebuild_zip5_tally(os.getcwd())

def rebuild_zip5_tally(root: str):
    tally: Dict[str,int] = {}
    for dirpath, dirnames, filenames in os.walk(root):
        if "executed_campaign_log.csv" in filenames:
            p = os.path.join(dirpath, "executed_campaign_log.csv")
            try:
                rows = read_csv(p)
            except Exception:
                continue
            for r in rows:
                z = (r.get("ZIP5","") or "").strip()
                if z:
                    tally[z] = tally.get(z, 0) + 1

    tracker_dir = os.path.join(root, "MasterCampaignTracker")
    os.makedirs(tracker_dir, exist_ok=True)
    out = os.path.join(tracker_dir, "Zip5_LetterTally.csv")
    rows_out = [ {"ZIP5": z, "Count": tally[z]} for z in sorted(tally.keys()) ]
    write_csv(out, rows_out, ["ZIP5","Count"])
    print(f"[OK] ZIP5 tally rebuilt: {out}")

if __name__ == "__main__":
    main()
