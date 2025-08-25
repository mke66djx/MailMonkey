
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FinalizeCampaign_TRACKER_STRICT_v4_MAILZIP_SEQTEMPLATE.py

Key changes vs v3 MAILZIP:
- TemplateIds are kept as an ORDERED SEQUENCE with duplicates allowed
  (e.g., "202|202" if the same template was used in two campaigns).
- New flag: --rebuild-templates
  Rebuilds TemplateIds (sequence) and CampaignNumbers (unique, sorted) for ALL pairs
  by scanning all executed_campaign_log.csv files under the root.
- Mailing ZIP remains the source of truth for ZIP5 fields and tallies.

Idempotent: won't append duplicate rows to executed log; tracker CampaignCount remains
the count of DISTINCT CampaignNumbers.
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
    base = os.path.basename(os.path.normpath(campaign_dir))
    m = re.match(r"^(?P<name>.+?)_(?P<num>\d+)_", base)
    if m:
        return m.group("name"), m.group("num")
    return base, None

# ---------------- ZIP helpers (MAIL-FIRST) ----------------
def _zip_from_text(s: str) -> str:
    if not s: return ""
    m = re.search(r"(\d{5})(?:-\d{4})?$", str(s).strip())
    return m.group(1) if m else ""

def build_zip_index_from_master(campaign_dir: str) -> Dict[Tuple[str,str], str]:
    """Build (addr_norm, owner_norm) -> ZIP5 from campaign_master.csv, MAIL-FIRST."""
    idx: Dict[Tuple[str,str], str] = {}
    cm_path = os.path.join(campaign_dir, "campaign_master.csv")
    if not os.path.isfile(cm_path):
        return idx
    rows = read_csv(cm_path)

    def get_zip_from_row(r: Dict[str,str]) -> str:
        # 1) Explicit mailing/owner ZIP
        mail_zip_cols = [
            "Mail ZIP","MAIL ZIP","Mail Zip","Mail Zip Code","MAIL ZIP CODE",
            "MAIL ZIP5","Mail ZIP5","MAILING ZIP","MAILING ZIP CODE","MAILING ZIP5",
            "Owner ZIP","OWNER ZIP","Owner Zip","OWNER ZIP5","Owner ZIP5"
        ]
        for k in mail_zip_cols:
            if k in r and r[k].strip():
                z = _zip_from_text(r[k])
                if z: return z
        # 2) Parse from mailing/owner address
        mail_addr_cols = [
            "MAILING ADDRESS","Mailing Address","Mailing Address 1","Mailing Address1",
            "OWNER ADDRESS","Owner Address","OWNER MAILING ADDRESS","Owner Mailing Address"
        ]
        for k in mail_addr_cols:
            if k in r and r[k].strip():
                z = _zip_from_text(r[k]); 
                if z: return z
        # 3) Generic ZIPs
        for k in ("ZIP5","Zip5","ZIP","Zip","Zip Code","ZIP CODE","ZIP CODE 5"):
            if k in r and r[k].strip():
                z = _zip_from_text(r[k]); 
                if z: return z
        # 4) Situs/Property ZIP or parse from property address (last)
        for k in ("SITUS ZIP","SITUS ZIP CODE","SITUS ZIP CODE 5-DIGIT","SITUS ZIP5","Situs ZIP","Situs Zip Code"):
            if k in r and r[k].strip():
                z = _zip_from_text(r[k]); 
                if z: return z
        for k in ("Property Address","PROPERTY ADDRESS","Address","ADDRESS","Situs Address","SITUS ADDRESS","PropertyAddress","SITUS"):
            if k in r and r[k].strip():
                z = _zip_from_text(r[k]); 
                if z: return z
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
    # 1) Mailing/Owner ZIPs
    for k in ("Mail ZIP","MAIL ZIP","Mail Zip","Mail Zip Code","MAIL ZIP CODE","MAIL ZIP5","Mail ZIP5",
              "MAILING ZIP","MAILING ZIP CODE","MAILING ZIP5","Owner ZIP","OWNER ZIP","Owner Zip","OWNER ZIP5","Owner ZIP5"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
    # 2) Mailing/Owner address strings
    for k in ("MAILING ADDRESS","Mailing Address","Mailing Address 1","Mailing Address1",
              "OWNER ADDRESS","Owner Address","OWNER MAILING ADDRESS","Owner Mailing Address"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
    # 3) Generic ZIPs
    for k in ("ZIP5","Zip5","ZIP","Zip","Zip Code","ZIP CODE","ZIP CODE 5"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
    # 4) Situs/Property (last)
    for k in ("SITUS ZIP","SITUS ZIP CODE","SITUS ZIP CODE 5-DIGIT","SITUS ZIP5","Situs ZIP","Situs Zip Code"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
    for k in ("property_address","Property Address","PROPERTY ADDRESS","Address","ADDRESS","Situs Address","SITUS ADDRESS"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
    return ""

# ------------- Core -------------
def main():
    ap = argparse.ArgumentParser(description="Finalize a campaign with mailing-ZIP logic and sequential TemplateIds.")
    ap.add_argument("--campaign-dir", required=True, help="Path to the campaign folder (e.g., Campaign_2_Aug2025)")
    ap.add_argument("--campaign-name", default=None, help="Override inferred campaign name")
    ap.add_argument("--campaign-number", default=None, help="Override inferred campaign number (int-like string)")
    ap.add_argument("--mapping", default=None, help="Override mapping path; default: <campaign-dir>/RefFiles/letters_mapping.csv or <campaign-dir>/letters_mapping.csv")
    ap.add_argument("--tracker-path", default="MasterCampaignTracker/MasterPropertyCampaignTracker.csv", help="Path to master tracker CSV")
    ap.add_argument("--write-history", action="store_true", help="Also append to MasterCampaignTracker/MasterPairHistory.csv (not used in this minimal version)")
    ap.add_argument("--force-recount", action="store_true", help="(Advanced) Force append even if rows already exist (generally NOT recommended)")
    ap.add_argument("--dedupe-log", action="store_true", help="De-duplicate the campaign's executed log by (OwnerName, PropertyAddress, CampaignNumber) and RefCode")
    ap.add_argument("--recount-tracker", action="store_true", help="Recompute CampaignCount for all rows from CampaignNumbers (alias: --repair-counts)")
    ap.add_argument("--repair-counts", action="store_true", help="Alias for --recount-tracker")
    ap.add_argument("--rebuild-templates", action="store_true", help="Rebuild TemplateIds (sequence, duplicates allowed) and CampaignNumbers for ALL pairs from executed logs after finalize")
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

    # Build index of mailing ZIPs from campaign_master (MAIL-FIRST)
    zip_idx = build_zip_index_from_master(campaign_dir)

    executed_log = os.path.join(campaign_dir, "executed_campaign_log.csv")
    if args.dedupe_log and os.path.isfile(executed_log):
        pre = read_csv(executed_log)
        seen_keys = set()
        seen_refs = set()
        deduped = []
        for r in pre:
            addr = r.get("PropertyAddress","") or r.get("property_address","") or r.get("Address","")
            owner = r.get("OwnerName","") or r.get("owner","") or r.get("Owner","")
            refc  = r.get("RefCode","") or r.get("ref_code","")
            campn = r.get("CampaignNumber","")
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
        addr = r.get("PropertyAddress","") or r.get("property_address","") or r.get("Address","")
        owner = r.get("OwnerName","") or r.get("owner","") or r.get("Owner","")
        refc  = r.get("RefCode","") or r.get("ref_code","")
        campn = r.get("CampaignNumber","")
        exist_pair_campaign.add( (norm_key(addr, owner), (campn or "").strip()) )
        if refc:
            exist_ref.add(refc)

    to_add: List[Dict[str,str]] = []
    for r in mapping_rows:
        owner = r.get("owner","") or r.get("Owner","") or r.get("OwnerName","")
        addr  = r.get("property_address","") or r.get("Property Address","") or r.get("PropertyAddress","") or r.get("Address","")
        refc  = r.get("ref_code","") or r.get("RefCode","")
        templ = r.get("template_ref","") or r.get("template_id","") or r.get("TemplateId","") or r.get("Template","")
        # MAIL-FIRST ZIP resolution
        z5    = r.get("ZIP5","") or get_zip_from_row_generic(r)
        if not z5 and (addr and owner):
            z5 = zip_idx.get(norm_key(addr, owner), "")

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
        z5 = r0.get("ZIP5","")

        if k in idx:
            tr = idx[k]
            if (not tr.get("ZIP5","")) and z5:
                tr["ZIP5"] = z5

            # CampaignNumbers: keep UNIQUE, numeric-sorted
            existing_cns = [x for x in (tr.get("CampaignNumbers","") or "").split("|") if x]
            cn_set = set(existing_cns)
            for rr in rows:
                cn_set.add(rr["CampaignNumber"])
            tr["CampaignNumbers"] = "|".join(sorted(cn_set, key=lambda x: int(re.sub(r"[^0-9]", "", x) or "0")))
            tr["CampaignCount"]   = str(len(cn_set))

            # TemplateIds: SEQUENCE, allow duplicates; append in arrival order
            existing_ts = [x for x in (tr.get("TemplateIds","") or "").split("|") if x]
            for rr in rows:
                tid = rr.get("TemplateId","")
                if tid:
                    existing_ts.append(tid)
            tr["TemplateIds"] = "|".join(existing_ts)

            first_dt = try_parse_date(tr.get("FirstSentDt","")) or try_parse_date(tr.get("FirstSentDtUTC",""))
            if not first_dt:
                tr["FirstSentDt"] = today_str
            tr["LastSentDt"] = today_str

            if not tr.get("PropertyAddress","" ).strip(): tr["PropertyAddress"] = addr
            if not tr.get("OwnerName","" ).strip(): tr["OwnerName"] = owner

        else:
            cn_set = {rows[0]["CampaignNumber"]}
            ts_seq = [t for t in [rr.get("TemplateId","") for rr in rows] if t]  # allow duplicates
            idx[k] = {
                "PropertyAddress": addr,
                "OwnerName": owner,
                "ZIP5": z5,
                "CampaignCount": str(len(cn_set)),
                "FirstSentDt": today_str,
                "LastSentDt": today_str,
                "CampaignNumbers": "|".join(sorted(cn_set, key=lambda x: int(re.sub(r"[^0-9]", "", x) or "0"))),
                "TemplateIds": "|".join(ts_seq),
            }

    # Optionally repair counts
    if args.recount_tracker or args.repair_counts:
        for tr in idx.values():
            cns = [x for x in (tr.get("CampaignNumbers","") or "").split("|") if x]
            tr["CampaignCount"] = str(len(set(cns)))

    final_rows = list(idx.values())
    headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    extra = [h for h in (final_rows[0].keys() if final_rows else []) if h not in headers]
    write_csv(tracker_path, final_rows, headers + extra)
    print(f"[OK] Master tracker updated: {tracker_path}")

    # Rebuild mailing ZIP tally from executed logs
    rebuild_zip5_tally(os.getcwd())

    # Optional: rebuild TemplateIds (sequence) & CampaignNumbers from ALL logs
    if args.rebuild_templates:
        rebuild_templates_and_campaigns_from_logs(tracker_path, os.getcwd())

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

def rebuild_templates_and_campaigns_from_logs(tracker_path: str, root: str):
    """Reconstruct TemplateIds (sequence, duplicates allowed) and CampaignNumbers (unique, sorted)
    for ALL pairs by scanning every executed_campaign_log.csv under root.
    """
    logs: List[Dict[str,str]] = []
    for dirpath, dirnames, filenames in os.walk(root):
        if "executed_campaign_log.csv" in filenames:
            p = os.path.join(dirpath, "executed_campaign_log.csv")
            try:
                rows = read_csv(p)
                logs.extend(rows)
            except Exception:
                continue

    # Build pair -> list of (campaignNumber:int, executedDt, templateId)
    agg: Dict[Tuple[str,str], List[Tuple[int, datetime, str]]] = {}
    for r in logs:
        addr = r.get("PropertyAddress","") or r.get("Address","") or r.get("property_address","")
        owner = r.get("OwnerName","") or r.get("Owner","") or r.get("owner","")
        if not addr or not owner: 
            continue
        key = norm_key(addr, owner)
        cn_raw = (r.get("CampaignNumber","") or "").strip()
        try:
            cn = int(re.sub(r"[^0-9]", "", cn_raw) or "0")
        except Exception:
            cn = 0
        dt = try_parse_date(r.get("ExecutedDt","")) or datetime.min
        tid = (r.get("TemplateId","") or "").strip()
        agg.setdefault(key, []).append((cn, dt, tid))

    # Reduce -> CampaignNumbers unique sorted; TemplateIds = sequence in ascending CN (ties by date)
    out_rows: Dict[Tuple[str,str], Dict[str,str]] = {}

    for key, items in agg.items():
        items.sort(key=lambda t: (t[0], t[1]))  # by campaign number then date
        cn_unique = []
        seen = set()
        tmpl_seq = []
        for cn, dt, tid in items:
            if cn not in seen:
                seen.add(cn)
                cn_unique.append(str(cn))
            if tid:
                tmpl_seq.append(tid)

        out_rows[key] = {
            "CampaignNumbers": "|".join(cn_unique),
            "TemplateIds": "|".join(tmpl_seq),
            "CampaignCount": str(len(cn_unique)),
        }

    # Load tracker and update the fields while preserving other columns
    tracker_rows = read_csv(tracker_path) if os.path.isfile(tracker_path) else []
    idx = { norm_key(r.get("PropertyAddress",""), r.get("OwnerName","")): r for r in tracker_rows }

    for key, vals in out_rows.items():
        if key in idx:
            tr = idx[key]
            tr["CampaignNumbers"] = vals["CampaignNumbers"]
            tr["TemplateIds"]     = vals["TemplateIds"]
            tr["CampaignCount"]   = vals["CampaignCount"]
        else:
            # If pair appears in logs but not in tracker, create a minimal row
            addr, owner = key
            idx[key] = {
                "PropertyAddress": addr.title(),
                "OwnerName": owner.title(),
                "ZIP5": "",
                "CampaignCount": vals["CampaignCount"],
                "FirstSentDt": "", "LastSentDt": "",
                "CampaignNumbers": vals["CampaignNumbers"],
                "TemplateIds": vals["TemplateIds"],
            }

    final_rows = list(idx.values())
    headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    extra = [h for h in (final_rows[0].keys() if final_rows else []) if h not in headers]
    write_csv(tracker_path, final_rows, headers + extra)
    print(f"[OK] Rebuilt TemplateIds (sequence) and CampaignNumbers from logs: {tracker_path}")

if __name__ == "__main__":
    main()
