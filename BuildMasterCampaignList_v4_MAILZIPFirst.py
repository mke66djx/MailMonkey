
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BuildMasterCampaignList_v4_MAILZIPFirst.py

Change summary vs your current BuildMasterCampaignList_V4.py:
- ZIP5 now strictly means the *mailing/owner ZIP*, not the property/situs ZIP.
- Presort and postage estimation are driven by mailing ZIP.
- get_zip5_from_row() rewritten to prefer MAIL/OWNER fields first, then fall back.

Everything else is unchanged (tracker read-only, prior filters, strict-150 packing,
schema mirroring, reports, etc.).
"""

import os, sys, csv, re, argparse, datetime, random, collections
from typing import List, Dict, Tuple, Optional

TRACKER_DIR = "MasterCampaignTracker"
TRACKER_FILE = os.path.join(TRACKER_DIR, "MasterPropertyCampaignTracker.csv")

# ------------------------- Helpers -------------------------
def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def norm_key(addr: str, owner: str) -> Tuple[str, str]:
    return (norm_space(addr).upper(), norm_space(owner).upper())

def get_zip5_from_text(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip()
    s = re.sub(r"\.0$", "", s)  # handle 95835.0
    m = re.search(r"(\d{5})(?:-\d{4})?$", s)
    return m.group(1) if m else ""

def get_zip5_from_row(row: Dict[str,str], addr: str) -> str:
    """
    IMPORTANT: Prefer MAILING/OWNER ZIP over property/situs ZIP.
    Order of precedence:
      1) Explicit mailing/owner ZIP columns
      2) Parse from mailing/owner address columns
      3) Generic ZIP5/ZIP columns (if they *clearly* refer to mailing)
      4) Property/Situs ZIP columns or parse from property address (last resort)
    """
    # 1) Explicit mailing/owner ZIP columns
    mail_zip_cols = [
        "Mail ZIP","MAIL ZIP","Mail Zip","Mail Zip Code","MAIL ZIP CODE",
        "MAIL ZIP5","Mail ZIP5","MAILING ZIP","MAILING ZIP CODE","MAILING ZIP5",
        "Owner ZIP","OWNER ZIP","Owner Zip","OWNER ZIP5","Owner ZIP5"
    ]
    for c in mail_zip_cols:
        if c in row and str(row[c]).strip():
            z = get_zip5_from_text(row[c])
            if z:
                return z

    # 2) Parse from mailing/owner address columns
    mail_addr_cols = [
        "MAILING ADDRESS","Mailing Address","Mailing Address 1","Mailing Address1",
        "OWNER ADDRESS","Owner Address","OWNER MAILING ADDRESS","Owner Mailing Address"
    ]
    for c in mail_addr_cols:
        if c in row and str(row[c]).strip():
            z = get_zip5_from_text(row[c])
            if z:
                return z

    # 3) Generic ZIP columns (only if present and likely mailing)
    generic_cols = ["ZIP5","Zip5","ZIP","Zip","Zip Code","ZIP CODE","ZIP CODE 5"]
    for c in generic_cols:
        if c in row and str(row[c]).strip():
            z = get_zip5_from_text(row[c])
            if z:
                return z

    # 4) Property/Situs ZIP columns or parse from property/situs address (last)
    situs_cols = [
        "SITUS ZIP","SITUS ZIP CODE","SITUS ZIP CODE 5-DIGIT","SITUS ZIP5",
        "Situs ZIP","Situs Zip Code"
    ]
    for c in situs_cols:
        if c in row and str(row[c]).strip():
            z = get_zip5_from_text(row[c])
            if z:
                return z

    # Parse from property/situs address string (addr param is usually property address)
    z = get_zip5_from_text(addr)
    if z:
        return z

    return ""

def read_csv_rows(path: str) -> List[Dict[str,str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [{k:(v or "").strip() for k,v in row.items()} for row in r]

def read_csv_rows_headers(path: str) -> Tuple[List[Dict[str,str]], List[str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        rows = [{k:(v or "").strip() for k,v in row.items()} for row in r]
        headers = r.fieldnames or []
        return rows, headers

def write_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        w.writerows(rows)

def campaign_folder(campaign_name: str, campaign_number: int, when: Optional[datetime.date]=None) -> str:
    when = when or datetime.date.today()
    mo_yr = when.strftime("%b%Y")  # e.g., Aug2025
    base = f"{campaign_name}_{campaign_number}_{mo_yr}"
    os.makedirs(base, exist_ok=True)
    return base

# ------------------------- Tracker (READ-ONLY) -------------------------
def read_tracker() -> Dict[Tuple[str,str], Dict[str,str]]:
    d: Dict[Tuple[str,str], Dict[str,str]] = {}
    if not os.path.exists(TRACKER_FILE):
        return d
    rows = read_csv_rows(TRACKER_FILE)
    for r in rows:
        k = norm_key(r.get("PropertyAddress",""), r.get("OwnerName",""))
        d[k] = r
    return d

def parse_last_campaign_number(info: Dict[str,str]) -> int:
    nums = (info.get("CampaignNumbers","") or "").strip()
    if nums:
        parts = re.split(r"[|,]\s*", nums)
        ints = []
        for p in parts:
            try:
                ints.append(int(p))
            except Exception:
                pass
        if ints:
            return ints[-1]
    try:
        return int((info.get("LastCampaignNumber") or "0").strip() or 0)
    except Exception:
        return 0

# ------------------------- Column detection -------------------------
def detect_addr_owner_from_source_row(row: Dict[str,str]) -> Tuple[str,str]:
    # Address detection (keep property/situs first for the LETTER body)
    addr = ""
    lmap = {k.lower():k for k in row}
    addr_candidates = [
        "PropertyAddress","PROPERTY ADDRESS","PROPERTY_ADDRESS",
        "SITUS ADDRESS","SITUS_ADDRESS","SITUS",
        "MAILING ADDRESS","MAILING_ADDRESS",
        "ADDRESS","ADDRESS 1","ADDRESS1","STREET ADDRESS",
        "Situs Address","Mailing Address","Property Address"
    ]
    for c in addr_candidates:
        if c in row and row[c].strip():
            addr = row[c].strip(); break
    if not addr and "address" in lmap:
        addr = row[lmap["address"]]

    # Owner detection
    own = ""
    owner_candidates = [
        "OwnerName","OWNER NAME","OWNER","OWNER(S)","OWNER 1","OWNER1","OWNER NAME 1",
        "Primary Name","PRIMARY NAME","Mail Owner","OWNER NAME(S)"
    ]
    for c in owner_candidates:
        if c in row and row[c].strip():
            own = row[c].strip(); break

    if not own:
        for fkey, lkey in (
            ("Primary First","Primary Last"),
            ("PRIMARY FIRST","PRIMARY LAST"),
            ("Owner First","Owner Last"),
            ("OWNER FIRST","OWNER LAST"),
            ("First Name","Last Name"),
            ("FIRST NAME","LAST NAME"),
        ):
            f = (row.get(fkey,"") or "").strip()
            l = (row.get(lkey,"") or "").strip()
            if f or l:
                own = norm_space(f"{f} {l}")
                break

    if not own and "owner" in lmap:
        own = row[lmap["owner"]].strip()

    return addr, own

def detect_addr_owner_from_selected_row(row: Dict[str,str]) -> Tuple[str,str]:
    addr = row.get("PropertyAddress") or row.get("ADDRESS") or row.get("Address") or row.get("SITUS ADDRESS") or ""
    own  = row.get("OwnerName") or row.get("OWNER NAME") or row.get("Primary Name") or row.get("PRIMARY NAME") or ""
    return addr, own

# ------------------------- Prior filters -------------------------
def passes_prior_rules(k: Tuple[str,str], tracker: Dict[Tuple[str,str],Dict[str,str]],
                       prior_exact: Optional[int], prior_max: Optional[int],
                       min_gap: int, current_campaign_number: int) -> bool:
    info = tracker.get(k)
    if not info:
        if prior_exact is not None:
            return prior_exact == 0
        if prior_max is not None:
            return 0 <= prior_max
        return True

    try:
        cnt = int((info.get("CampaignCount") or "0").strip() or 0)
    except Exception:
        cnt = 0

    last = parse_last_campaign_number(info)

    if prior_exact is not None and cnt != prior_exact:
        return False
    if prior_max is not None and cnt > prior_max:
        return False
    if min_gap > 0 and last > 0:
        if last > (current_campaign_number - min_gap):
            return False
    return True

# ------------------------- Selection logic -------------------------
def pick_optimized(candidates: List[Dict[str,str]], target: int, strict_150: bool) -> List[Dict[str,str]]:
    if target <= 0: return []
    by_zip5: Dict[str, List[Dict[str,str]]] = collections.defaultdict(list)
    for r in candidates:
        by_zip5[r.get("ZIP5","")].append(r)

    buckets = sorted(by_zip5.items(), key=lambda kv: (len(kv[1]), kv[0] != ""), reverse=True)

    chosen: List[Dict[str,str]] = []
    if strict_150:
        for z5, bucket in buckets:
            if len(chosen) >= target: break
            random.shuffle(bucket)
            take_n = (len(bucket) // 150) * 150
            if take_n == 0: continue
            chosen.extend(bucket[:min(take_n, target - len(chosen))])
            by_zip5[z5] = bucket[min(take_n, len(bucket)):]

        if len(chosen) < target:
            leftovers = sorted(by_zip5.items(), key=lambda kv: len(kv[1]), reverse=True)
            for z5, bucket in leftovers:
                if len(chosen) >= target: break
                random.shuffle(bucket)
                for row in bucket:
                    if len(chosen) >= target: break
                    chosen.append(row)
    else:
        for z5, bucket in buckets:
            if len(chosen) >= target: break
            random.shuffle(bucket)
            for row in bucket:
                if len(chosen) >= target: break
                chosen.append(row)

    if len(chosen) < target:
        remaining = [r for r in candidates if r not in chosen]
        by_zip3: Dict[str, List[Dict[str,str]]] = collections.defaultdict(list)
        for r in remaining:
            z3 = (r.get("ZIP5","") or "")[:3]
            by_zip3[z3].append(r)
        zip3_buckets = sorted(by_zip3.items(), key=lambda kv: len(kv[1]), reverse=True)
        for z3, bucket in zip3_buckets:
            if len(chosen) >= target: break
            random.shuffle(bucket)
            for row in bucket:
                if len(chosen) >= target: break
                chosen.append(row)

    return chosen[:target]

# ------------------------- Postage estimate -------------------------
def estimate_postage(chosen: List[Dict[str,str]], rate_5: float, rate_3: float, rate_aadc: float) -> Dict[str, float]:
    by_zip5 = collections.Counter(r.get("ZIP5","") for r in chosen)
    five_digit = 0
    leftovers_by_zip3 = collections.Counter()
    for z5, c in by_zip5.items():
        trays = c // 150
        five_digit += trays * 150
        rem = c - trays * 150
        z3 = (z5 or "")[:3]
        leftovers_by_zip3[z3] += rem

    three_digit = 0
    aadc = 0
    for z3, total in leftovers_by_zip3.items():
        trays = total // 150
        three_digit += trays * 150
        aadc += total - trays * 150

    total_pieces = len(chosen)
    cost_5 = five_digit * rate_5
    cost_3 = three_digit * rate_3
    cost_a = aadc * rate_aadc
    total_cost = cost_5 + cost_3 + cost_a
    avg = (total_cost / total_pieces) if total_pieces else 0.0
    return {
        "five_digit": five_digit,
        "three_digit": three_digit,
        "aadc": aadc,
        "cost_5": cost_5,
        "cost_3": cost_3,
        "cost_a": cost_a,
        "total_cost": total_cost,
        "avg": avg,
    }

# ------------------------- Main -------------------------
def main():
    ap = argparse.ArgumentParser(description="Build USPS-optimized campaign master list (mailing ZIP-first presort).")
    ap.add_argument("--campaign-name", required=True)
    ap.add_argument("--campaign-number", type=int, required=True)
    ap.add_argument("--target-size", type=int, required=True)

    ap.add_argument("--mandatory", nargs="+", required=True, help="Up to 4 mandatory CSV files")
    ap.add_argument("--optional", nargs="*", default=[], help="Up to 2 optional pool CSV files")

    # Prior filters (exclusive: choose either exact or max)
    ap.add_argument("--prior-exact", type=int, help="Only include address/owner pairs with exactly N prior campaigns (0 = never mailed)")
    ap.add_argument("--prior-max", type=int, help="Only include address/owner pairs with ≤ M prior campaigns")
    ap.add_argument("--min-gap", type=int, default=0, help="Require last campaign number be ≥ this many behind current")

    # Optimization
    ap.add_argument("--strict-150", dest="strict_150", action="store_true", help="Prefer packing 5-digit ZIP trays in multiples of 150")
    ap.add_argument("--rate-5digit", type=float, default=0.244)
    ap.add_argument("--rate-3digit", type=float, default=0.275)
    ap.add_argument("--rate-aadc", type=float, default=0.330)

    # Debug
    ap.add_argument("--debug", action="store_true")

    args = ap.parse_args()
    if args.prior_exact is not None and args.prior_max is not None:
        print("[ERROR] Use either --prior-exact OR --prior-max (not both).")
        sys.exit(1)
    if len(args.mandatory) > 4:
        print("[ERROR] Max 4 mandatory lists allowed.")
        sys.exit(1)
    if len(args.optional) > 2:
        print("[ERROR] Max 2 optional lists allowed.")
        sys.exit(1)

    # Load tracker (READ-ONLY)
    tracker = read_tracker()

    # Ingest + normalize + filter
    seen_keys = set()
    all_candidates: List[Dict[str,str]] = []

    stats = {
        "MAND": {"kept":0,"deduped":0,"dropped_prior":0,"missing_addr":0,"missing_owner":0},
        "POOL": {"kept":0,"deduped":0,"dropped_prior":0,"missing_addr":0,"missing_owner":0},
    }

    def process_rows(rows: List[Dict[str,str]], bucket: str):
        for r in rows:
            addr, own = detect_addr_owner_from_source_row(r)
            if not addr:
                stats[bucket]["missing_addr"] += 1
                continue
            if not own:
                stats[bucket]["missing_owner"] += 1
                continue
            k = norm_key(addr, own)
            if k in seen_keys:
                stats[bucket]["deduped"] += 1
                continue
            # prior filters
            if not passes_prior_rules(k, tracker, args.prior_exact, args.prior_max, args.min_gap, args.campaign_number):
                stats[bucket]["dropped_prior"] += 1
                continue
            # ZIP5 (MAILING-FIRST)
            z5 = get_zip5_from_row(r, addr)
            row = {
                "PropertyAddress": norm_space(addr),
                "OwnerName": norm_space(own),
                "ZIP5": z5,
                "_src_row": r,  # keep for schema mapping later
            }
            all_candidates.append(row)
            seen_keys.add(k)
            stats[bucket]["kept"] += 1

    # Mandatory
    for p in args.mandatory:
        rows = read_csv_rows(p)
        if args.debug:
            print(f"[DEBUG] Reading mandatory: {p} (rows={len(rows)})")
        process_rows(rows, "MAND")

    # Safety: if mandatory exceeds target
    mand_kept = stats["MAND"]["kept"]
    if mand_kept > args.target_size:
        print(f"[ERROR] Mandatory lists exceed target after filtering ({mand_kept} > {args.target_size}). Refine inputs.")
        sys.exit(1)

    # Optional pools
    for p in args.optional:
        rows = read_csv_rows(p)
        if args.debug:
            print(f"[DEBUG] Reading optional: {p} (rows={len(rows)})")
        process_rows(rows, "POOL")

    if args.debug:
        kept_m = stats["MAND"]["kept"]; kept_p = stats["POOL"]["kept"]
        print("[DEBUG] Summary after ingest:")
        print(f"  MAND kept={kept_m}  deduped={stats['MAND']['deduped']}  dropped_prior={stats['MAND']['dropped_prior']}  missing_addr={stats['MAND']['missing_addr']}  missing_owner={stats['MAND']['missing_owner']}")
        print(f"  POOL kept={kept_p}  deduped={stats['POOL']['deduped']}  dropped_prior={stats['POOL']['dropped_prior']}  missing_addr={stats['POOL']['missing_addr']}  missing_owner={stats['POOL']['missing_owner']}")
        print(f"  TOTAL candidates={len(all_candidates)}")

    # Choose optimized selection
    chosen = pick_optimized(all_candidates, args.target_size, args.strict_150)

    # Sort chosen for USPS-friendly order (mailing ZIP)
    chosen.sort(key=lambda r: ((r.get("ZIP5") or "ZZZZZ"), r.get("PropertyAddress",""), r.get("OwnerName","")))

    # Reports
    by_zip5 = collections.Counter((r.get("ZIP5","") or "(none)") for r in chosen)
    presort_rows = [{"ZIP5": z5, "Count": c} for z5, c in by_zip5.most_common()]
    by_zip3 = collections.Counter(((z5 if z5!="(none)" else "")[:3]) for z5 in by_zip5.keys())
    presort_rows3 = []
    for z3 in sorted(by_zip3.keys()):
        if z3 is None: z3 = ""
        total = sum(by_zip5[z5] for z5 in by_zip5 if (z5 if z5!="(none)" else "")[:3] == z3)
        est_zip5_buckets = sum(1 for z5 in by_zip5 if (z5 if z5!="(none)" else "")[:3] == z3)
        presort_rows3.append({"ZIP3": z3 or "(none)", "EstZIP5Buckets": est_zip5_buckets, "TotalPieces": total})

    # Output folder
    camp_dir = campaign_folder(args.campaign_name, args.campaign_number)
    master_path = os.path.join(camp_dir, "campaign_master.csv")
    presort_path = os.path.join(camp_dir, "presort_report.csv")
    presort_zip3_path = os.path.join(camp_dir, "presort_zip3_summary.csv")
    postage_path = os.path.join(camp_dir, "postage_estimate.csv")

    # Write presort reports
    write_csv(presort_path, presort_rows, ["ZIP5","Count"])
    write_csv(presort_zip3_path, presort_rows3, ["ZIP3","EstZIP5Buckets","TotalPieces"])

    # Postage estimate
    est = estimate_postage(chosen, args.rate_5digit, args.rate_3digit, args.rate_aadc)
    postage_rows = [
        {"Tier":"5digit","Pieces":est["five_digit"],"Rate":args.rate_5digit,"Cost":round(est["cost_5"],2)},
        {"Tier":"3digit","Pieces":est["three_digit"],"Rate":args.rate_3digit,"Cost":round(est["cost_3"],2)},
        {"Tier":"AADC","Pieces":est["aadc"],"Rate":args.rate_aadc,"Cost":round(est["cost_a"],2)},
        {"Tier":"total","Pieces":len(chosen),"Rate":"","Cost":round(est["total_cost"],2)},
        {"Tier":"AveragePerPiece","Pieces":"","Rate":"","Cost":round(est["avg"],4)},
    ]
    write_csv(postage_path, postage_rows, ["Tier","Pieces","Rate","Cost"])

    # --------- Determine template header from the first mandatory file ----------
    template_headers: List[str] = []
    header_source_path = None
    # Try mandatory in order
    for p in args.mandatory:
        try:
            _, hdrs = read_csv_rows_headers(p)
            if hdrs:
                template_headers = hdrs
                header_source_path = p
                break
        except Exception:
            continue
    # If still none, try optional
    if not template_headers:
        for p in args.optional:
            try:
                _, hdrs = read_csv_rows_headers(p)
                if hdrs:
                    template_headers = hdrs
                    header_source_path = p
                    break
            except Exception:
                continue

    # If still none, fall back to minimal schema
    use_minimal = False
    if not template_headers:
        use_minimal = True
        template_headers = ["Address","Primary Name"]  # minimal fall-back; downstream may need more

    # --------- Build index (addr, owner) -> *source* row (first match wins) ----------
    index: Dict[Tuple[str,str], Dict[str,str]] = {}
    for p in (args.mandatory + args.optional):
        try:
            rows = read_csv_rows(p)
        except Exception:
            continue
        for r in rows:
            a, o = detect_addr_owner_from_source_row(r)
            if not a or not o:
                continue
            k = norm_key(a, o)
            if k not in index:
                index[k] = r

    # --------- Emit rows in chosen order using the template header ----------
    out_rows: List[Dict[str,str]] = []
    missing = 0
    for sel in chosen:
        a = sel.get("PropertyAddress","")
        o = sel.get("OwnerName","")
        k = norm_key(a, o)

        src = index.get(k) or sel.get("_src_row", {})

        if not src:
            missing += 1
            src = {}

        new_row = {}
        if use_minimal:
            # minimal fallback
            new_row = {"Address": a, "Primary Name": o}
        else:
            # copy in template order
            for col in template_headers:
                if col in src:
                    new_row[col] = src.get(col, "")
                elif col in sel:
                    new_row[col] = sel.get(col, "")
                else:
                    new_row[col] = ""

            # ensure Address / Primary Name have values
            for addr_like in ("Address","ADDRESS","Property Address","PROPERTY ADDRESS","Situs Address","SITUS ADDRESS","Mailing Address","MAILING ADDRESS"):
                if addr_like in new_row and not new_row[addr_like].strip():
                    new_row[addr_like] = a
            for owner_like in ("Primary Name","PRIMARY NAME","OwnerName","OWNER NAME","OWNER","OWNER(S)"):
                if owner_like in new_row and not new_row[owner_like].strip():
                    new_row[owner_like] = o

        out_rows.append(new_row)

    write_csv(master_path, out_rows, template_headers)

    # Console summary
    print(f"[OK] Created campaign folder: {camp_dir}")
    print(f"[OK] Master list: {master_path}  (rows={len(chosen)})")
    print(f"[OK] Presort ZIP5: {presort_path}")
    print(f"[OK] Presort ZIP3: {presort_zip3_path}")
    print(f"[OK] Postage estimate: {postage_path}")
    if header_source_path:
        print(f"[INFO] Master schema mirrored from: {header_source_path}")
    if missing and args.debug:
        print(f"[DEBUG] {missing} rows had no exact (addr,owner) source match; backfilled from selection.")

if __name__ == "__main__":
    main()
