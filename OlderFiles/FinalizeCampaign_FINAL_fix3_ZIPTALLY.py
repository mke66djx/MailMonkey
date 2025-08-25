
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FinalizeCampaign - FINAL (fix3) with robust ZIP5 tally rebuild

What’s new vs. previous fix:
- ZIP5 tally is now rebuilt from scratch every run by scanning *all* executed_campaign_log.csv files
  under the project, *deduplicating* entries per (address+owner+campaign+ref/page). This prevents double
  counting when you re-run finalize (even with --force-recount).
- ZIP5 backfill logic remains (mapping -> campaign_master -> address regex).
- All prior options/features are preserved (dry run, force-recount, history, combined-pdf metadata, etc.).
"""

import os, sys, csv, re, argparse, glob, datetime as dt
from pathlib import Path
from typing import Dict, List, Tuple, Optional

# ----------------- Helpers -----------------

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def norm_key(addr: str, owner: str) -> str:
    return f"{norm_space(addr).upper()}||{norm_space(owner).upper()}"

def zip5_from_str(s: str) -> str:
    if not s:
        return ""
    m = re.search(r"(\d{5})(?:-\d{4})?", str(s))
    return m.group(1) if m else ""

ZIP_CANDIDATES = [
    "ZIP5","Zip5","ZIP","Zip","MAIL ZIP","Mail ZIP","MailZip","MAILZIP",
    "SITUS ZIP","Situs ZIP","SitusZip","SITUSZIP","ZIP CODE","Zip Code","ZipCode","ZIPCODE"
]

ADDR_CANDIDATES = [
    "PropertyAddress","PROPERTY ADDRESS","Property Address","Situs Address","SITUS ADDRESS",
    "Address","ADDRESS","Mailing Address","MAILING ADDRESS","PROPERTY_ADDRESS","SITUS_ADDRESS"
]

OWNER_CANDIDATES = [
    "OwnerName","OWNER NAME","Owner Name","Owner","OWNER","Owner(s)","OWNER(S)",
    "Primary Name","PRIMARY NAME","Mail Owner"
]

def get_first_present(row: Dict[str,str], keys: List[str]) -> str:
    for k in keys:
        if k in row and str(row[k]).strip():
            return str(row[k]).strip()
    # try case-insensitive fallback
    low = {k.lower():k for k in row.keys()}
    for k in keys:
        lk = k.lower()
        if lk in low and str(row[low[lk]]).strip():
            return str(row[low[lk]]).strip()
    return ""

def detect_addr_owner(row: Dict[str,str]) -> Tuple[str,str]:
    addr = get_first_present(row, ADDR_CANDIDATES)
    own  = get_first_present(row, OWNER_CANDIDATES)
    return addr, own

def zip_from_row(row: Dict[str,str]) -> str:
    for k in ZIP_CANDIDATES:
        if k in row and str(row[k]).strip():
            z = zip5_from_str(row[k])
            if z: return z
    # case-insensitive fallback
    low = {k.lower():k for k in row.keys()}
    for k in ZIP_CANDIDATES:
        lk = k.lower()
        if lk in low and str(row[low[lk]]).strip():
            z = zip5_from_str(row[low[lk]])
            if z: return z
    # try address text
    addr = get_first_present(row, ADDR_CANDIDATES)
    return zip5_from_str(addr)

def read_csv_rows(path: Path) -> List[Dict[str,str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [{k: (v or "").strip() for k,v in row.items()} for row in r]

def write_csv_rows(path: Path, rows: List[Dict[str,str]], headers: List[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

def append_csv_rows(path: Path, rows: List[Dict[str,str]], headers: List[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    exists = path.exists()
    with path.open("a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        if not exists:
            w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})

def newest_matching(paths: List[Path]) -> Optional[Path]:
    paths = [p for p in paths if p.exists()]
    if not paths: return None
    return max(paths, key=lambda p: p.stat().st_mtime)

# -------------- Mapping & master index --------------

def auto_find_mapping(camp_dir: Path, project_root: Path) -> Optional[Path]:
    # 1) in campaign dir
    for name in ("letters_mapping.csv","mapping.csv"):
        p = camp_dir / name
        if p.exists(): return p
    # 2) in subfolders RefFiles
    candidates = list((camp_dir/"RefFiles").glob("*.csv"))
    candidates += [p for p in camp_dir.glob("*.csv") if re.search(r"(map|ref)", p.name, re.I)]
    if candidates:
        return newest_matching(candidates)
    # 3) project root RefFiles
    pr = project_root / "RefFiles"
    if pr.exists():
        c2 = [p for p in pr.glob("*.csv") if re.search(r"(map|ref)", p.name, re.I)]
        if c2:
            return newest_matching(c2)
    return None

def build_master_index(camp_dir: Path) -> Dict[str, Dict[str,str]]:
    idx = {}
    cm = camp_dir / "campaign_master.csv"
    if not cm.exists():
        return idx
    for r in read_csv_rows(cm):
        addr, own = detect_addr_owner(r)
        if addr and own:
            k = norm_key(addr, own)
            if k not in idx:
                z = zip_from_row(r)
                idx[k] = {"ZIP5": z}
    return idx

def resolve_zip_for(mapping_or_exec_row: Dict[str,str], master_idx: Dict[str,Dict[str,str]]) -> str:
    # 1) from row itself
    z = zip_from_row(mapping_or_exec_row)
    if z: return z
    # 2) from master by addr+owner
    addr, own = detect_addr_owner(mapping_or_exec_row)
    if addr and own:
        k = norm_key(addr, own)
        m = master_idx.get(k, {})
        z = m.get("ZIP5","")
        if z: return z
    # 3) regex already tried in zip_from_row; give up
    return ""

# -------------- Tally Rebuild --------------

def rebuild_zip_tally(project_root: Path) -> Path:
    tracker_dir = project_root / "MasterCampaignTracker"
    out_path = tracker_dir / "Zip5_LetterTally.csv"
    tracker_dir.mkdir(parents=True, exist_ok=True)

    # Find all executed logs under project_root/Campaign_*/
    exec_logs: List[Path] = []
    for p in project_root.iterdir():
        if p.is_dir() and re.match(r"^Campaign_\d+_[A-Za-z]{3}\d{4}$", p.name):
            log = p / "executed_campaign_log.csv"
            if log.exists():
                exec_logs.append(log)

    # Aggregate (dedup per key)
    counts: Dict[str,int] = {}
    seen: set = set()

    for log in exec_logs:
        camp_dir = log.parent
        master_idx = build_master_index(camp_dir)
        rows = read_csv_rows(log)
        # Map column names
        if not rows:
            continue
        hdrs = {h.lower(): h for h in rows[0].keys()}
        col_page = hdrs.get("page") or hdrs.get("pagenum") or hdrs.get("pageno")
        col_owner = hdrs.get("owner") or hdrs.get("ownername")
        col_addr  = hdrs.get("property_address") or hdrs.get("address") or hdrs.get("propertyaddress")
        col_ref   = hdrs.get("ref") or hdrs.get("ref_code") or hdrs.get("refcode")
        col_cnum  = hdrs.get("campaignnumber") or hdrs.get("campaign_number") or "CampaignNumber"

        for r in rows:
            addr = r.get(col_addr or "", "") or get_first_present(r, ADDR_CANDIDATES)
            owner= r.get(col_owner or "", "") or get_first_present(r, OWNER_CANDIDATES)
            cnum = r.get(col_cnum or "", "")
            page = r.get(col_page or "", "")
            refc = r.get(col_ref or "", "")
            key = (norm_key(addr, owner), str(cnum), str(refc or page))
            if key in seen:
                continue
            seen.add(key)

            z = r.get("ZIP5","") or resolve_zip_for(r, master_idx)
            if not z:
                continue
            counts[z] = counts.get(z, 0) + 1

    rows_out = [{"ZIP5": z, "Letters": str(n)} for z, n in sorted(counts.items())]
    write_csv_rows(out_path, rows_out, ["ZIP5","Letters"])
    return out_path

# -------------- Finalize core --------------

def finalize(args):
    project_root = Path(args.project_root or ".").resolve()
    camp_dir = Path(args.campaign_dir).resolve()
    if not camp_dir.exists():
        print(f"[ERROR] Campaign dir not found: {camp_dir}")
        sys.exit(2)

    # mapping path
    mapping_path = Path(args.mapping) if args.mapping else auto_find_mapping(camp_dir, project_root)
    if not mapping_path or not mapping_path.exists():
        print("[ERROR] Could not locate mapping CSV. Use --mapping to specify it explicitly.")
        sys.exit(2)
    mapping_rows = read_csv_rows(mapping_path)
    if not mapping_rows:
        print("[ERROR] Mapping CSV is empty.")
        sys.exit(2)

    # detect basic columns in mapping
    hdrs = {h.lower(): h for h in mapping_rows[0].keys()}
    col_page = hdrs.get("page") or hdrs.get("pagenum") or hdrs.get("pageno")
    col_owner = hdrs.get("owner") or hdrs.get("ownername")
    col_addr  = hdrs.get("property_address") or hdrs.get("address") or hdrs.get("propertyaddress")
    col_ref   = hdrs.get("ref") or hdrs.get("ref_code") or hdrs.get("refcode")
    col_tref  = hdrs.get("template_ref") or hdrs.get("templ") or hdrs.get("template")
    col_tsrc  = hdrs.get("template_source") or None

    today = dt.date.today().isoformat()
    master_idx = build_master_index(camp_dir)

    # Build executed rows with ZIP5
    executed_rows = []
    for r in mapping_rows:
        addr = r.get(col_addr or "", "") or get_first_present(r, ADDR_CANDIDATES)
        owner= r.get(col_owner or "", "") or get_first_present(r, OWNER_CANDIDATES)
        zip5 = resolve_zip_for(r, master_idx)
        executed_rows.append({
            "PropertyAddress": addr,
            "OwnerName": owner,
            "ZIP5": zip5,
            "Page": r.get(col_page or "", ""),
            "RefCode": r.get(col_ref or "", ""),
            "TemplateRef": r.get(col_tref or "", ""),
            "TemplateSource": r.get(col_tsrc or "", ""),
            "CampaignNumber": str(args.campaign_number),
            "CampaignName": args.campaign_name or "",
            "ExecutedDate": today,
            "CombinedPDF": args.combined_pdf or "",
        })

    # Paths
    exec_log = camp_dir / "executed_campaign_log.csv"
    tracker_dir = project_root / "MasterCampaignTracker"
    tracker_dir.mkdir(parents=True, exist_ok=True)
    counter_path = tracker_dir / "MasterPropertyCampaignCounter.csv"
    history_path = Path(args.history_path).resolve() if args.history_path else (tracker_dir / "MasterPairHistory.csv")

    # Idempotency / filter if not force-recount
    to_write_exec = executed_rows
    if not args.force_recount and exec_log.exists():
        existing = read_csv_rows(exec_log)
        seen = set((norm_key(x.get("PropertyAddress",""), x.get("OwnerName","")), x.get("CampaignNumber","")) for x in existing)
        to_write_exec = [x for x in executed_rows if (norm_key(x["PropertyAddress"], x["OwnerName"]), x["CampaignNumber"]) not in seen]

    print(f"[SUMMARY] Mapping rows: {len(mapping_rows)} | Already logged (skipped): {len(executed_rows)-len(to_write_exec)} | To add now: {len(to_write_exec)}")

    if args.dry_run:
        print("[DRY RUN] No changes written.")
        # Even on dry-run, show where we would rebuild the tally from:
        sim_tally = project_root / "MasterCampaignTracker" / "Zip5_LetterTally.csv"
        print(f"[DRY RUN] Would rebuild ZIP5 tally into: {sim_tally}")
        return

    # Append to executed log (ensuring ZIP5 column exists)
    exec_headers = ["PropertyAddress","OwnerName","ZIP5","Page","RefCode","TemplateRef","TemplateSource","CampaignNumber","CampaignName","ExecutedDate","CombinedPDF"]
    append_csv_rows(exec_log, to_write_exec, exec_headers)
    print(f"[OK] Appended {len(to_write_exec)} rows to {exec_log}")

    # Update master counter (with ZIP5)
    counter_rows = []
    if counter_path.exists():
        counter_rows = read_csv_rows(counter_path)
    # Build index
    counter_idx = {norm_key(r.get("PropertyAddress",""), r.get("OwnerName","")): r for r in counter_rows}

    for x in to_write_exec:
        k = norm_key(x["PropertyAddress"], x["OwnerName"])
        cur = counter_idx.get(k)
        if cur is None:
            cur = {
                "PropertyAddress": x["PropertyAddress"],
                "OwnerName": x["OwnerName"],
                "ZIP5": x["ZIP5"],
                "CampaignCount": "1",
                "FirstSentDate": today,
                "PrevSentDate": "",
                "LastSentDate": today,
                "PrevCampaignNumber": "",
                "LastCampaignNumber": x["CampaignNumber"],
                "LastTemplateRef": x["TemplateRef"],
                "LastRefCode": x["RefCode"],
                "LastPage": x["Page"],
            }
            counter_idx[k] = cur
        else:
            # increment
            cnt = int(cur.get("CampaignCount","0") or "0") + 1
            cur["CampaignCount"] = str(cnt)
            # dates
            cur["PrevSentDate"] = cur.get("LastSentDate","") or ""
            cur["LastSentDate"] = today
            # campaigns
            cur["PrevCampaignNumber"] = cur.get("LastCampaignNumber","") or ""
            cur["LastCampaignNumber"] = x["CampaignNumber"]
            # zip refresh if missing
            if not (cur.get("ZIP5","") or "").strip():
                cur["ZIP5"] = x["ZIP5"]
            # last refs
            cur["LastTemplateRef"] = x["TemplateRef"]
            cur["LastRefCode"] = x["RefCode"]
            cur["LastPage"] = x["Page"]

    # Write counter
    counter_headers = ["PropertyAddress","OwnerName","ZIP5","CampaignCount","FirstSentDate","PrevSentDate","LastSentDate","PrevCampaignNumber","LastCampaignNumber","LastTemplateRef","LastRefCode","LastPage"]
    merged_rows = list(counter_idx.values())
    write_csv_rows(counter_path, merged_rows, counter_headers)
    print(f"[OK] Master tracker updated: {counter_path}")

    # Robust ZIP5 tally (rebuild from scratch, dedup, backfill)
    out_tally = rebuild_zip_tally(project_root)
    print(f"[OK] ZIP5 tally rebuilt: {out_tally}")

    # Optional history
    if args.write_history:
        hist_headers = ["ExecutedDate","CampaignNumber","CampaignName","PropertyAddress","OwnerName","ZIP5","Page","RefCode","TemplateRef","TemplateSource","CombinedPDF"]
        append_csv_rows(history_path, to_write_exec, hist_headers)
        print(f"[OK] History appended: {history_path}")

# -------------- CLI --------------

def main():
    ap = argparse.ArgumentParser(description="Finalize campaign: update executed log, master tracker (with ZIP5), and rebuild a deduped ZIP tally. Optional history log.")
    ap.add_argument("--campaign-number", type=int, required=True)
    ap.add_argument("--campaign-dir", required=True)
    ap.add_argument("--campaign-name", default="")
    ap.add_argument("--mapping", default=None, help="Mapping CSV path (owner/address/page/ref/template). Auto-detected if omitted.")
    ap.add_argument("--project-root", default=".", help="Root folder containing MasterCampaignTracker/")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force-recount", action="store_true")
    ap.add_argument("--write-history", action="store_true")
    ap.add_argument("--history-path", default=None)
    ap.add_argument("--combined-pdf", default="", help="Optional: stored as metadata in history and executed log")
    args = ap.parse_args()
    finalize(args)

if __name__ == "__main__":
    main()
