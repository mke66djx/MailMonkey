#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
BuildMasterCampaignList_v2_DEBUG.py
- Merges up to 4 mandatory + up to 2 optional lists into a USPS-optimized campaign master list.
- Filters against MasterCampaignTracker/MasterPropertyCampaignCounter.csv by prior campaign count and min gap.
- Address standardization (lightweight) and de-duplication by (PropertyAddress, OwnerName).
- ZIP5-first picking with optional STRICT 150 mode for 5-digit trays.
- Explicitly SORTS final output by ZIP5 -> PropertyAddress -> OwnerName for print order.
- Adds --debug to print why rows were skipped and show source-by-source stats.
- Broader header detection:
    * Address: Address, Property Address, Situs Address, Mailing Address, etc.
    * Owner:   OwnerName, Primary Name, or combine Primary First + Primary Last.

Outputs: <CampaignName>_<Number>_<MonYYYY>/
  - campaign_master.csv
  - presort_report.csv
  - presort_zip3_summary.csv
  - postage_estimate.csv
"""

import os, sys, csv, re, argparse, datetime, random, collections
from typing import List, Dict, Tuple, Optional

TRACKER_DIR = "MasterCampaignTracker"
TRACKER_FILE = os.path.join(TRACKER_DIR, "MasterPropertyCampaignCounter.csv")

# ---------------- Address utilities ----------------

STREET_ABBR = {
    "STREET":"ST", "ST":"ST",
    "AVENUE":"AVE", "AVE":"AVE",
    "BOULEVARD":"BLVD", "BLVD":"BLVD",
    "DRIVE":"DR", "DR":"DR",
    "COURT":"CT", "CT":"CT",
    "LANE":"LN", "LN":"LN",
    "ROAD":"RD", "RD":"RD",
    "TERRACE":"TER", "TER":"TER",
    "PARKWAY":"PKWY", "PKWY":"PKWY",
    "HIGHWAY":"HWY", "HWY":"HWY",
    "PLACE":"PL", "PL":"PL",
    "CIRCLE":"CIR", "CIR":"CIR",
    "TRAIL":"TRL", "TRL":"TRL",
    "WAY":"WAY", "SQUARE":"SQ", "SQ":"SQ"
}

def norm_space(s: str) -> str:
    return re.sub(r"\\s+", " ", (s or "").strip())

def standardize_address(addr: str) -> str:
    """Light standardization for keying; NOT CASS-certified."""
    if not addr: return ""
    a = norm_space(addr).upper()
    # Remove unit markers for keying
    a = re.sub(r"\\b(APT|UNIT|STE|SUITE|#)\\s*\\w+\\b", "", a)
    a = norm_space(a)
    # Normalize common street types at last token
    parts = a.split(" ")
    if parts:
        last = parts[-1].rstrip(".")
        if last in STREET_ABBR:
            parts[-1] = STREET_ABBR[last]
    return " ".join(parts)

def norm_key(addr: str, owner: str) -> Tuple[str, str]:
    return (standardize_address(addr), norm_space(owner).upper())

def get_zip5(addr: str) -> str:
    m = re.search(r"(\\d{5})(?:-\\d{4})?$", addr or "")
    return m.group(1) if m else ""

def get_zip3(zip5: str) -> str:
    return zip5[:3] if zip5 else ""

# ---------------- CSV IO ----------------

def read_csv(path: str) -> List[Dict[str,str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [ {k: (v or "").strip() for k,v in row.items()} for row in r ]

def write_csv(path: str, rows: List[Dict[str,str]], headers: List[str]):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        w.writerows(rows)

def ensure_tracker():
    if not os.path.exists(TRACKER_FILE):
        os.makedirs(TRACKER_DIR, exist_ok=True)
        with open(TRACKER_FILE, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f); w.writerow(
                ["PropertyAddress","OwnerName","CampaignCount","LastCampaignNumber","FirstSeenCampaign","FirstSeenDate","LastUpdatedDate"]
            )

def read_tracker() -> Dict[Tuple[str,str], Dict[str,str]]:
    ensure_tracker()
    rows = read_csv(TRACKER_FILE)
    d = {}
    for r in rows:
        k = norm_key(r.get("PropertyAddress",""), r.get("OwnerName",""))
        d[k] = r
    return d

# ---------------- Field detection ----------------

ADDR_CANDIDATES = [
    "PropertyAddress","PROPERTY ADDRESS","PROPERTY_ADDRESS",
    "SITUS ADDRESS","SITUS_ADDRESS","SITUS",
    "MAILING ADDRESS","MAILING_ADDRESS","MAILING ADDRESS 1",
    "ADDRESS","ADDRESS 1","ADDRESS1","PROPERTY ADDR","STREET ADDRESS",
]

OWNER_CANDIDATES = [
    "OwnerName","OWNER NAME","OWNER","OWNER(S)","OWNER 1","OWNER1","OWNER NAME 1",
    "Primary Name","PRIMARY NAME",
]

FIRST_LAST_PATTERNS = [
    ("Primary First","Primary Last"),
    ("PRIMARY FIRST","PRIMARY LAST"),
    ("Owner First","Owner Last"),
    ("OWNER FIRST","OWNER LAST"),
]

def required_cols(row: Dict[str,str]) -> Tuple[str,str]:
    # Address
    addr = ""
    lmap = {k.lower():k for k in row}
    for c in ADDR_CANDIDATES:
        if c in row and row[c].strip():
            addr = row[c].strip(); break
    if not addr and "address" in lmap:
        addr = row[lmap["address"]]

    # Owner
    own = ""
    for c in OWNER_CANDIDATES:
        if c in row and row[c].strip():
            own = row[c].strip(); break
    if not own:
        # Try First + Last combos
        for fkey, lkey in FIRST_LAST_PATTERNS:
            if fkey in row or lkey in row:
                first = row.get(fkey, "").strip()
                last  = row.get(lkey, "").strip()
                if first or last:
                    own = norm_space(f"{first} {last}")
                    break
    if not own and "owner" in lmap:
        own = row[lmap["owner"]]

    return addr, own

# ---------------- Merge, filter, optimize ----------------

def campaign_folder(campaign_name: str, campaign_number: int, when: Optional[datetime.date]=None) -> str:
    when = when or datetime.date.today()
    mo_yr = when.strftime("%b%Y")  # e.g., Aug2025
    base = f"{campaign_name}_{campaign_number}_{mo_yr}"
    os.makedirs(base, exist_ok=True)
    return base

def passes_prior_rules(k, tracker, prior_exact, prior_max, min_gap, current_campaign_number) -> bool:
    info = tracker.get(k)
    if not info:
        if prior_exact is not None:
            return prior_exact == 0
        if prior_max is not None:
            return 0 <= prior_max
        return True
    cnt = int((info.get("CampaignCount") or "0").strip() or 0)
    last = int((info.get("LastCampaignNumber") or "0").strip() or 0)
    if prior_exact is not None and cnt != prior_exact:
        return False
    if prior_max is not None and cnt > prior_max:
        return False
    if min_gap > 0 and last > 0:
        if last > (current_campaign_number - min_gap):
            return False
    return True

def pick_zip5_strict_150(candidates: List[Dict[str,str]], target: int) -> List[Dict[str,str]]:
    # Group by ZIP5
    by_zip5 = collections.defaultdict(list)
    for r in candidates:
        z5 = get_zip5(r["PropertyAddress"])
        by_zip5[z5].append(r)

    # Sort buckets by size desc (ignore empty ZIP last)
    ordered = sorted(by_zip5.items(), key=lambda kv: (len(kv[1]), kv[0] != ""), reverse=True)

    chosen = []
    # Stage 1: take floor(n/150)*150 from each big bucket
    for z5, bucket in ordered:
        if len(chosen) >= target: break
        n = len(bucket)
        take = (n // 150) * 150
        if take == 0: continue
        random.shuffle(bucket)
        for r in bucket[:min(take, target - len(chosen))]:
            chosen.append(r)

    if len(chosen) >= target:
        return chosen[:target]

    # Stage 2: try to round up promising buckets to 150 if close (e.g., >=100)
    remaining = [r for r in candidates if r not in chosen]
    by_zip5_remaining = collections.defaultdict(list)
    for r in remaining:
        by_zip5_remaining[get_zip5(r["PropertyAddress"])].append(r)

    promising = sorted(((z5, lst) for z5,lst in by_zip5_remaining.items() if len(lst) >= 100),
                       key=lambda kv: len(kv[1]), reverse=True)
    for z5, lst in promising:
        need = 150
        already = sum(1 for r in chosen if get_zip5(r["PropertyAddress"]) == z5)
        take = max(0, min(need - already, len(lst), target - len(chosen)))
        if take > 0:
            random.shuffle(lst)
            chosen.extend(lst[:take])
        if len(chosen) >= target: break

    if len(chosen) >= target:
        return chosen[:target]

    # Stage 3: fill remaining by largest ZIP5 then ZIP3
    remaining = [r for r in candidates if r not in chosen]
    by_zip5_rem2 = collections.defaultdict(list)
    for r in remaining:
        by_zip5_rem2[get_zip5(r["PropertyAddress"])].append(r)
    rem_order = sorted(by_zip5_rem2.items(), key=lambda kv: len(kv[1]), reverse=True)
    for z5, lst in rem_order:
        random.shuffle(lst)
        for r in lst:
            if len(chosen) >= target: break
            chosen.append(r)
        if len(chosen) >= target: break

    # If still short, fill by ZIP3
    if len(chosen) < target:
        remaining = [r for r in candidates if r not in chosen]
        by_zip3 = collections.defaultdict(list)
        for r in remaining:
            by_zip3[get_zip3(get_zip5(r["PropertyAddress"]))].append(r)
        rem3 = sorted(by_zip3.items(), key=lambda kv: len(kv[1]), reverse=True)
        for z3, lst in rem3:
            random.shuffle(lst)
            for r in lst:
                if len(chosen) >= target: break
                chosen.append(r)
            if len(chosen) >= target: break

    return chosen[:target]

def pick_zip5_general(candidates: List[Dict[str,str]], target: int) -> List[Dict[str,str]]:
    if target <= 0: return []
    by_zip5 = collections.defaultdict(list)
    for r in candidates:
        by_zip5[get_zip5(r["PropertyAddress"])].append(r)
    zip5_buckets = sorted(by_zip5.items(), key=lambda kv: (len(kv[1]), kv[0] != ""), reverse=True)
    chosen = []
    for z5, bucket in zip5_buckets:
        random.shuffle(bucket)
        for row in bucket:
            if len(chosen) >= target: break
            chosen.append(row)
        if len(chosen) >= target: break
    if len(chosen) >= target:
        return chosen[:target]
    remaining = [r for r in candidates if r not in chosen]
    by_zip3 = collections.defaultdict(list)
    for r in remaining:
        by_zip3[get_zip3(get_zip5(r["PropertyAddress"]))].append(r)
    zip3_buckets = sorted(by_zip3.items(), key=lambda kv: len(kv[1]), reverse=True)
    for z3, bucket in zip3_buckets:
        random.shuffle(bucket)
        for row in bucket:
            if len(chosen) >= target: break
            chosen.append(row)
        if len(chosen) >= target: break
    return chosen[:target]

# ---------------- Postage estimator ----------------

def estimate_blended_cost(zip5_counts: Dict[str,int],
                          rate_5: float, rate_3: float, rate_aadc: float):
    total_pieces = sum(zip5_counts.values())
    five_digit = 0
    three_digit = 0

    leftovers_by_zip5 = {}
    for z5, c in zip5_counts.items():
        five_chunks = (c // 150) * 150
        five_digit += five_chunks
        leftovers_by_zip5[z5] = c - five_chunks

    by_zip3_left = collections.defaultdict(int)
    for z5, c in leftovers_by_zip5.items():
        by_zip3_left[z5[:3] if z5 else ""] += c
    for z3, c in by_zip3_left.items():
        three_digit += (c // 150) * 150

    used = five_digit + three_digit
    aadc = max(0, total_pieces - used)
    cost = five_digit*rate_5 + three_digit*rate_3 + aadc*rate_aadc
    avg = (cost/total_pieces) if total_pieces else 0.0
    breakdown = {
        "5digit": {"pieces": five_digit, "rate": rate_5, "cost": five_digit*rate_5},
        "3digit": {"pieces": three_digit, "rate": rate_3, "cost": three_digit*rate_3},
        "AADC":   {"pieces": aadc,      "rate": rate_aadc, "cost": aadc*rate_aadc},
        "total":  {"pieces": total_pieces, "cost": cost, "avg": avg}
    }
    return breakdown

# ---------------- Main ----------------

def main():
    ap = argparse.ArgumentParser(description="Build USPS-optimized master list, filtered by prior campaign history.")
    ap.add_argument("--campaign-name", required=True)
    ap.add_argument("--campaign-number", type=int, required=True)
    ap.add_argument("--target-size", type=int, required=True)

    ap.add_argument("--mandatory", nargs="+", required=True, help="1-4 CSVs required")
    ap.add_argument("--optional", nargs="*", default=[], help="0-2 CSVs optional pools")

    # Prior history filters (choose ONE of prior-exact or prior-max)
    ap.add_argument("--prior-exact", type=int, help="Only include entries with exactly N prior campaigns (0=never mailed)")
    ap.add_argument("--prior-max", type=int, help="Only include entries with ≤ M prior campaigns")
    ap.add_argument("--min-gap", type=int, default=0, help="Require last campaign be ≥ this many campaign numbers ago")

    # Optimization flags
    ap.add_argument("--strict-150", action="store_true", help="Favor multiples of 150 per ZIP5 where possible")

    # Postage rates
    ap.add_argument("--rate-5digit", type=float, default=0.244)
    ap.add_argument("--rate-3digit", type=float, default=0.275)
    ap.add_argument("--rate-aadc", type=float, default=0.330)

    # Debug
    ap.add_argument("--debug", action="store_true", help="Print verbose row-diagnostic stats")

    args = ap.parse_args()
    if args.prior_exact is not None and args.prior_max is not None:
        print("[ERROR] Use either --prior-exact OR --prior-max, not both.")
        sys.exit(1)

    # Load and validate lists
    if len(args.mandatory) > 4:
        print("[ERROR] Max 4 mandatory lists allowed."); sys.exit(1)
    if len(args.optional) > 2:
        print("[ERROR] Max 2 optional lists allowed."); sys.exit(1)

    tracker = read_tracker()
    seen = set()
    all_candidates = []

    def add_rows(src_rows: List[Dict[str,str]], tag: str, stats: dict):
        for i, r in enumerate(src_rows, 1):
            addr, own = required_cols(r)
            if not addr:
                stats["missing_addr"] += 1
                if args.debug and (i <= 3):
                    print(f"[DEBUG] {tag}: row {i} missing address")
                continue
            if not own:
                stats["missing_owner"] += 1
                if args.debug and (i <= 3):
                    print(f"[DEBUG] {tag}: row {i} missing owner")
                continue
            k = norm_key(addr, own)
            if k in seen:
                stats["deduped"] += 1
                continue
            if not passes_prior_rules(k, tracker, args.prior_exact, args.prior_max, args.min_gap, args.campaign_number):
                stats["dropped_prior"] += 1
                continue
            row = {
                "PropertyAddress": norm_space(addr),
                "OwnerName": norm_space(own),
                "ZIP5": get_zip5(addr),
            }
            # carry a few common columns if present
            for extra in ("OwnerFirstName","SITUS CITY","SITUS ZIP","MAIL CITY","MAIL ZIP","Phone","Email"):
                if extra in r:
                    row[extra] = r[extra]
            all_candidates.append(row)
            seen.add(k)
            stats["kept"] += 1

    # Read files
    mand_stats = {"missing_addr":0,"missing_owner":0,"dropped_prior":0,"deduped":0,"kept":0}
    pool_stats = {"missing_addr":0,"missing_owner":0,"dropped_prior":0,"deduped":0,"kept":0}

    mand_rows = []
    for p in args.mandatory:
        if not os.path.exists(p):
            print(f"[ERROR] Mandatory list not found: {p}")
            sys.exit(1)
        rows = read_csv(p); mand_rows.extend(rows)
        if args.debug:
            print(f"[DEBUG] Reading mandatory: {p} (rows={len(rows)})")
    add_rows(mand_rows, "MAND", mand_stats)

    if len(all_candidates) > args.target_size:
        print(f"[ERROR] Mandatory lists exceed target ({len(all_candidates)} > {args.target_size}). Refine your inputs.")
        sys.exit(1)

    pool_rows = []
    for p in args.optional:
        if not p: continue
        if not os.path.exists(p):
            print(f"[WARN] Optional list not found (skipping): {p}")
            continue
        rows = read_csv(p); pool_rows.extend(rows)
        if args.debug:
            print(f"[DEBUG] Reading optional: {p} (rows={len(rows)})")
    add_rows(pool_rows, "POOL", pool_stats)

    if args.debug:
        print("[DEBUG] Summary after ingest:")
        print(f"  MAND kept={mand_stats['kept']}  deduped={mand_stats['deduped']}  dropped_prior={mand_stats['dropped_prior']}  missing_addr={mand_stats['missing_addr']}  missing_owner={mand_stats['missing_owner']}")
        print(f"  POOL kept={pool_stats['kept']}  deduped={pool_stats['deduped']}  dropped_prior={pool_stats['dropped_prior']}  missing_addr={pool_stats['missing_addr']}  missing_owner={pool_stats['missing_owner']}")
        print(f"  TOTAL candidates={len(all_candidates)}")

    # Pick optimized set
    if args.strict_150:
        chosen = pick_zip5_strict_150(all_candidates, args.target_size)
    else:
        chosen = pick_zip5_general(all_candidates, args.target_size)

    # Sort final selection
    def _get_z5(row):
        try:
            return get_zip5(row.get("PropertyAddress", ""))
        except Exception:
            return ""
    chosen = sorted(chosen, key=lambda r: (_get_z5(r), r.get("PropertyAddress",""), r.get("OwnerName","")))

    # Presort profile
    by_zip5 = collections.Counter(get_zip5(r["PropertyAddress"]) for r in chosen)
    presort_rows = [{"ZIP5": z5 or "(none)", "Count": c} for z5, c in by_zip5.most_common()]

    by_zip3 = {}
    for z5, cnt in by_zip5.items():
        z3 = get_zip3(z5)
        rec = by_zip3.setdefault(z3, {"EstZIP5Buckets":0, "TotalPieces":0})
        rec["EstZIP5Buckets"] += 1 if z5 else 0
        rec["TotalPieces"] += cnt
    presort_rows3 = [{"ZIP3": z3 or "(none)", **vals} for z3, vals in sorted(by_zip3.items())]

    # Outputs
    camp_dir = campaign_folder(args.campaign_name, args.campaign_number)
    master_path = os.path.join(camp_dir, "campaign_master.csv")
    presort_path = os.path.join(camp_dir, "presort_report.csv")
    presort_zip3_path = os.path.join(camp_dir, "presort_zip3_summary.csv")
    postage_path = os.path.join(camp_dir, "postage_estimate.csv")

    # Save master
    headers = ["PropertyAddress","OwnerName","ZIP5","OwnerFirstName","SITUS CITY","SITUS ZIP","MAIL CITY","MAIL ZIP","Phone","Email"]
    write_csv(master_path, chosen, headers)

    # Save presort details
    write_csv(presort_path, presort_rows, ["ZIP5","Count"])
    write_csv(presort_zip3_path, presort_rows3, ["ZIP3","EstZIP5Buckets","TotalPieces"])

    # Postage estimate
    breakdown = estimate_blended_cost(by_zip5, args.rate_5digit, args.rate_3digit, args.rate_aadc)
    with open(postage_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Tier","Pieces","Rate","Cost"])
        for tier in ("5digit","3digit","AADC","total"):
            row = breakdown[tier]
            if tier == "total":
                w.writerow([tier, row["pieces"], "", f"{row['cost']:.2f}"])
                w.writerow(["AveragePerPiece","", "", f"{row['avg']:.4f}"])
            else:
                w.writerow([tier, row["pieces"], f"{row['rate']:.3f}", f"{row['cost']:.2f}"])

    print(f"[OK] Created campaign folder: {camp_dir}")
    print(f"[OK] Master list: {master_path}  (rows={len(chosen)})")
    print(f"[OK] Presort ZIP5: {presort_path}")
    print(f"[OK] Presort ZIP3: {presort_zip3_path}")
    print(f"[OK] Postage estimate: {postage_path}")
    print(f"[SUMMARY] Estimated blended avg: ${breakdown['total']['avg']:.4f}  "
          f"(5digit={breakdown['5digit']['pieces']}, 3digit={breakdown['3digit']['pieces']}, AADC={breakdown['AADC']['pieces']})")

if __name__ == "__main__":
    main()
