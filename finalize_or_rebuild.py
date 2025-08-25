
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py

This is the full v5 script (normal finalize + disaster-recovery rebuild)
with a Windows-safe date formatter (no use of "%-m/%-d/%Y").

Features kept from v5:
- Normal finalize for a single campaign (idempotent append to executed log;
  mailing-ZIP logic; tracker update with CampaignNumbers UNIQUE and
  TemplateIds as a SEQUENCE allowing duplicates; ZIP5 tally rebuild).
- --rebuild-all / --reindex-all: scan ALL campaign folders to rebuild
  MasterPropertyCampaignTracker.csv and Zip5_LetterTally.csv from executed logs.
- --write-marker: drop a marker file in the campaign folder.
- --marker-required / --marker-name: only treat folders with the marker as campaigns.
- --rebuild-templates: refresh TemplateIds (sequence) and CampaignNumbers (unique)
  across ALL pairs by scanning all logs after a normal finalize.

Change from v5:
- Replaced Unix-only strftime("%-m/%-d/%Y") with cross‑platform
  formatting via fmt_mdy(dt) -> "M/D/YYYY".
"""

import os, csv, re, argparse
from datetime import datetime
from typing import Dict, List, Tuple, Optional

# ------------------------------ Helpers ------------------------------

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def norm_key(addr: str, owner: str) -> Tuple[str, str]:
    return (norm_space(addr).upper(), norm_space(owner).upper())

def read_csv(path: str) -> List[Dict[str, str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [{k:(v or "").strip() for k,v in row.items()} for row in r]

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

def fmt_mdy(dt: Optional[datetime]) -> str:
    if not dt:
        return ""
    try:
        # Treat year 1 (datetime.min) as empty
        if dt.year <= 1:
            return ""
        return f"{dt.month}/{dt.day}/{dt.year}"
    except Exception:
        return ""

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

def build_zip_index_from_master(campaign_dir: str) -> Dict[Tuple[str,str], str]:
    """Build (addr_norm, owner_norm) -> ZIP5 from campaign_master.csv, MAIL-FIRST."""
    idx: Dict[Tuple[str,str], str] = {}
    cm_path = os.path.join(campaign_dir, "campaign_master.csv")
    if not os.path.isfile(cm_path):
        return idx
    rows = read_csv(cm_path)

    def get_zip_from_row(r: Dict[str,str]) -> str:
        z = get_zip_from_row_generic(r)
        if z: return z
        # As last resort, property fields
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

# ------------------------------ Core ------------------------------

def append_executed_and_update_tracker(args) -> None:
    """Normal finalize for a single campaign (v4 semantics + marker write)."""
    campaign_dir = args.campaign_dir
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

    # Tracker update (sequence templates; unique campaigns)
    tracker_path = args.tracker_path
    os.makedirs(os.path.dirname(tracker_path), exist_ok=True)
    tracker_rows = read_csv(tracker_path) if os.path.isfile(tracker_path) else []
    idx: Dict[Tuple[str,str], Dict[str,str]] = { norm_key(r.get("PropertyAddress",""), r.get("OwnerName","")): r for r in tracker_rows }

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
            # campaigns
            existing_cns = [x for x in (tr.get("CampaignNumbers","") or "").split("|") if x]
            cn_set = set(existing_cns)
            for rr in rows:
                cn_set.add(rr["CampaignNumber"])
            tr["CampaignNumbers"] = "|".join(sorted(cn_set, key=lambda x: int(re.sub(r"[^0-9]", "", x) or "0")))
            tr["CampaignCount"]   = str(len(cn_set))
            # templates (sequence, allow duplicates)
            existing_ts = [x for x in (tr.get("TemplateIds","") or "").split("|") if x]
            for rr in rows:
                tid = rr.get("TemplateId","")
                if tid:
                    existing_ts.append(tid)
            tr["TemplateIds"] = "|".join(existing_ts)
            # dates
            tr["FirstSentDt"] = tr.get("FirstSentDt","") or today_str
            tr["LastSentDt"] = today_str
            # ensure columns
            tr["PropertyAddress"] = tr.get("PropertyAddress","") or addr
            tr["OwnerName"] = tr.get("OwnerName","") or owner
        else:
            cn_set = {rows[0]["CampaignNumber"]}
            ts_seq = [t for t in [rr.get("TemplateId","") for rr in rows] if t]
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

    final_rows = list(idx.values())
    headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    extra = [h for h in (final_rows[0].keys() if final_rows else []) if h not in headers]
    write_csv(tracker_path, final_rows, headers + extra)
    print(f"[OK] Master tracker updated: {tracker_path}")

    # Tally
    rebuild_zip5_tally(args.root)

    # Write marker if requested
    if args.write_marker:
        marker = os.path.join(campaign_dir, args.marker_name)
        try:
            with open(marker, "w", encoding="utf-8") as f:
                f.write("")
            print(f"[OK] Marker written: {marker}")
        except Exception as e:
            print(f"[WARN] Could not write marker: {marker} ({e})")

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

def discover_campaign_folders(root: str, marker_required: bool, marker_name: str) -> List[str]:
    """Find campaign folders by walking for executed_campaign_log.csv.
       If marker_required=True, only accept folders that contain marker_name file as well.
    """
    found = set()
    for dirpath, dirnames, filenames in os.walk(root):
        if "executed_campaign_log.csv" in filenames:
            folder = dirpath
            if marker_required and not os.path.isfile(os.path.join(folder, marker_name)):
                continue
            found.add(folder)
    # Return sorted for determinism
    return sorted(found)

def rebuild_tracker_from_all(args) -> None:
    """Scan all campaign folders and rebuild the tracker & tallies from scratch."""
    root = args.root
    folders = discover_campaign_folders(root, args.marker_required, args.marker_name)
    if not folders:
        print(f"[WARN] No campaign folders found under: {root}")
        return

    print(f"[INFO] Found {len(folders)} campaign folders.")
    # Aggregate data across all logs
    agg: Dict[Tuple[str,str], Dict[str,object]] = {}
    for folder in folders:
        # prepare ZIP index from that folder's campaign_master for backfill
        zip_idx = build_zip_index_from_master(folder)
        log_path = os.path.join(folder, "executed_campaign_log.csv")
        try:
            rows = read_csv(log_path)
        except Exception:
            continue
        for r in rows:
            addr = r.get("PropertyAddress","") or r.get("Address","") or r.get("property_address","")
            owner = r.get("OwnerName","") or r.get("Owner","") or r.get("owner","")
            if not addr or not owner:
                continue
            key = norm_key(addr, owner)
            z5 = (r.get("ZIP5","") or "").strip()
            if not z5:
                # Try to backfill from mapping row (if present) or master index
                z5 = get_zip_from_row_generic(r) or (zip_idx.get(key, ""))

            cn_raw = (r.get("CampaignNumber","") or "").strip()
            try:
                cn = int(re.sub(r"[^0-9]", "", cn_raw) or "0")
            except Exception:
                cn = 0
            dt = try_parse_date(r.get("ExecutedDt","")) or None
            tid = (r.get("TemplateId","") or "").strip()

            rec = agg.setdefault(key, {
                "PropertyAddress": addr,
                "OwnerName": owner,
                "ZIP5": z5,
                "CampaignNumbers": set(),
                "TemplateIds": [],
                "FirstSentDt": None,
                "LastSentDt": None,
            })
            # keep a nice-cased address/owner if we see one later
            if not rec["PropertyAddress"]: rec["PropertyAddress"] = addr
            if not rec["OwnerName"]: rec["OwnerName"] = owner
            if not rec["ZIP5"] and z5: rec["ZIP5"] = z5

            rec["CampaignNumbers"].add(str(cn))
            if tid:
                rec["TemplateIds"].append(tid)

            if dt:
                if rec["FirstSentDt"] is None or dt < rec["FirstSentDt"]:
                    rec["FirstSentDt"] = dt
                if rec["LastSentDt"] is None or dt > rec["LastSentDt"]:
                    rec["LastSentDt"] = dt

    # Build final rows
    final_rows = []
    for key, rec in agg.items():
        cns = sorted({x for x in rec["CampaignNumbers"] if x and x != "0"}, key=lambda x: int(x))
        first = fmt_mdy(rec["FirstSentDt"])
        last = fmt_mdy(rec["LastSentDt"])
        final_rows.append({
            "PropertyAddress": rec["PropertyAddress"],
            "OwnerName": rec["OwnerName"],
            "ZIP5": rec["ZIP5"],
            "CampaignCount": str(len(cns)),
            "FirstSentDt": first,
            "LastSentDt": last,
            "CampaignNumbers": "|".join(cns),
            "TemplateIds": "|".join(rec["TemplateIds"]),  # sequence, allow dups
        })

    tracker_path = args.tracker_path
    headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    write_csv(tracker_path, final_rows, headers)
    print(f"[OK] Rebuilt tracker from scratch: {tracker_path} (rows={len(final_rows)})")

    # Rebuild ZIP tally
    rebuild_zip5_tally(root)

# ------------------------------ CLI ------------------------------

def main():
    ap = argparse.ArgumentParser(description="Finalize a single campaign OR rebuild the tracker by scanning all campaign folders.")
    # Normal finalize params
    ap.add_argument("--campaign-dir", help="Path to the campaign folder (e.g., Campaign_2_Aug2025)")
    ap.add_argument("--campaign-name", default=None, help="Override inferred campaign name")
    ap.add_argument("--campaign-number", default=None, help="Override inferred campaign number (int-like string)")
    ap.add_argument("--mapping", default=None, help="Override mapping path; default: <campaign-dir>/RefFiles/letters_mapping.csv or <campaign-dir>/letters_mapping.csv")
    ap.add_argument("--tracker-path", default="MasterCampaignTracker/MasterPropertyCampaignTracker.csv", help="Path to master tracker CSV")
    ap.add_argument("--force-recount", action="store_true", help="(Advanced) Force append even if rows already exist (generally NOT recommended)")
    ap.add_argument("--dry-run", action="store_true", help="Show what would change but do not write files")
    ap.add_argument("--write-marker", action="store_true", help="Write a campaign marker file into --campaign-dir after finalize")
    # Disaster recovery / global rebuild
    ap.add_argument("--rebuild-all", action="store_true", help="Scan all campaign folders under --root and rebuild the master tracker + tallies from scratch")
    ap.add_argument("--reindex-all", action="store_true", help="Alias of --rebuild-all")
    ap.add_argument("--root", default=".", help="Root folder to scan for campaign folders (default: current directory)")
    ap.add_argument("--marker-required", action="store_true", help="Only treat folders with a marker file as campaigns")
    ap.add_argument("--marker-name", default="CAMPAIGN.TAG", help="Name of the marker file (default: CAMPAIGN.TAG)")
    # Template/campaign rebuild (like v4)
    ap.add_argument("--rebuild-templates", action="store_true", help="Rebuild TemplateIds (sequence) and CampaignNumbers for ALL pairs from executed logs after finalize")
    args = ap.parse_args()

    if args.rebuild_all or args.reindex_all:
        return rebuild_tracker_from_all(args)

    if not args.campaign_dir:
        print("[ERROR] --campaign-dir is required for normal finalize. Or use --rebuild-all to rebuild from all folders.")
        return

    append_executed_and_update_tracker(args)

    if args.rebuild_templates:
        # reuse the all-folders builder to refresh sequences + counts across everything
        rebuild_tracker_from_all(args)

if __name__ == "__main__":
    main()
