
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_fixdate.py

This is a patched version of v5 that avoids %-m/%-d strftime codes (not supported on Windows).
Dates are formatted cross‑platform as M/D/YYYY using dt.month/dt.day/dt.year.
"""

import os, csv, re, argparse
from datetime import datetime
from typing import Dict, List, Tuple, Optional

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
    for k in ("Mail ZIP","MAIL ZIP","Mail Zip","Mail Zip Code","MAIL ZIP CODE","MAIL ZIP5","Mail ZIP5",
              "MAILING ZIP","MAILING ZIP CODE","MAILING ZIP5","Owner ZIP","OWNER ZIP","Owner Zip","OWNER ZIP5","Owner ZIP5"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
    for k in ("MAILING ADDRESS","Mailing Address","Mailing Address 1","Mailing Address1",
              "OWNER ADDRESS","Owner Address","OWNER MAILING ADDRESS","Owner Mailing Address"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
    for k in ("ZIP5","Zip5","ZIP","Zip","Zip Code","ZIP CODE","ZIP CODE 5"):
        if k in r and r[k].strip():
            z = _zip_from_text(r[k])
            if z: return z
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
    idx: Dict[Tuple[str,str], str] = {}
    cm_path = os.path.join(campaign_dir, "campaign_master.csv")
    if not os.path.isfile(cm_path):
        return idx
    rows = read_csv(cm_path)

    def get_zip_from_row(r: Dict[str,str]) -> str:
        z = get_zip_from_row_generic(r)
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

# ------------------------------ Core pieces (same as v5) ------------------------------

def discover_campaign_folders(root: str, marker_required: bool, marker_name: str) -> List[str]:
    found = set()
    for dirpath, dirnames, filenames in os.walk(root):
        if "executed_campaign_log.csv" in filenames:
            if marker_required and not os.path.isfile(os.path.join(dirpath, marker_name)):
                continue
            found.add(dirpath)
    return sorted(found)

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

def rebuild_tracker_from_all(args) -> None:
    root = args.root
    folders = discover_campaign_folders(root, args.marker_required, args.marker_name)
    if not folders:
        print(f"[WARN] No campaign folders found under: {root}")
        return
    print(f"[INFO] Found {len(folders)} campaign folders.")

    from collections import defaultdict
    agg: Dict[Tuple[str,str], Dict[str,object]] = {}
    for folder in folders:
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
            z5 = (r.get("ZIP5","") or "").strip() or get_zip_from_row_generic(r) or zip_idx.get(key, "")

            cn_raw = (r.get("CampaignNumber","") or "").strip()
            try:
                cn = int(re.sub(r"[^0-9]", "", cn_raw) or "0")
            except Exception:
                cn = 0
            dt = try_parse_date(r.get("ExecutedDt",""))
            if not dt:
                # try to infer from folder name timestamp? leave None
                dt = None
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

            if not rec["ZIP5"] and z5: rec["ZIP5"] = z5
            rec["CampaignNumbers"].add(str(cn))
            if tid:
                rec["TemplateIds"].append(tid)

            if dt:
                if rec["FirstSentDt"] is None or dt < rec["FirstSentDt"]:
                    rec["FirstSentDt"] = dt
                if rec["LastSentDt"] is None or dt > rec["LastSentDt"]:
                    rec["LastSentDt"] = dt

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
            "TemplateIds": "|".join(rec["TemplateIds"]),
        })

    tracker_path = args.tracker_path
    headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDt","LastSentDt","CampaignNumbers","TemplateIds"]
    write_csv(tracker_path, final_rows, headers)
    print(f"[OK] Rebuilt tracker from scratch: {tracker_path} (rows={len(final_rows)})")

    rebuild_zip5_tally(root)

def main():
    ap = argparse.ArgumentParser(description="Rebuild tracker (disaster recovery) using mailing ZIP logic. Minimal patch for Windows date formatting.")
    ap.add_argument("--rebuild-all", action="store_true", help="Scan all campaign folders under --root and rebuild the tracker + tallies from scratch")
    ap.add_argument("--root", default=".", help="Root folder to scan for campaign folders (default: current directory)")
    ap.add_argument("--tracker-path", default="MasterCampaignTracker/MasterPropertyCampaignTracker.csv", help="Path to master tracker CSV")
    ap.add_argument("--marker-required", action="store_true", help="Only treat folders with a marker file as campaigns")
    ap.add_argument("--marker-name", default="CAMPAIGN.TAG", help="Name of the marker file (default: CAMPAIGN.TAG)")

    # Keep normal finalize out of this patch to keep the fix small;
    # you can continue to use v5 for normal finalize. This script is your recovery path.
    args = ap.parse_args()

    if args.rebuild_all:
        return rebuild_tracker_from_all(args)

    print("[ERROR] Use --rebuild-all for this recovery script.")

if __name__ == "__main__":
    main()
