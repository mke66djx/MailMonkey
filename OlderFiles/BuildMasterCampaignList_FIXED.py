
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
BuildMasterCampaignList.py
- Merges up to 4 mandatory + up to 2 optional lists into a USPS-optimized campaign master list.
- Filters against MasterCampaignTracker/MasterPropertyCampaignCounter.csv by prior campaign count and min gap.
- Address standardization (lightweight) and de-duplication by (PropertyAddress, OwnerName).
- ZIP5-first picking with optional STRICT 150 mode for 5-digit trays.
- Outputs campaign folder: <Name>_<Number>_<MonYYYY>/
    - campaign_master.csv
    - presort_report.csv (ZIP5 counts)
    - presort_zip3_summary.csv
    - postage_estimate.csv (with blended cost estimate)

Usage example (PowerShell):
    python BuildMasterCampaignList.py `
      --campaign-name "Delinquencies_Sac" `
      --campaign-number 7 `
      --target-size 5000 `
      --mandatory "PropertyLists/vacant.csv" "PropertyLists/foreclosures.csv" `
      --optional "PropertyLists/non_owner_occupied.csv" `
      --prior-exact 0 `
      --min-gap 2 `
      --strict-150
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
    return re.sub(r"\s+", " ", (s or "").strip())

def to_title_case(s: str) -> str:
    s = (s or "").strip().lower()
    out = []
    m_ord = re.compile(r"^(\d+)(st|nd|rd|th)$")
    for w in s.split():
        m = m_ord.match(w)
        out.append(m.group(1)+m.group(2) if m else w.capitalize())
    return " ".join(out)

def standardize_address(addr: str) -> str:
    """Very light standardization suitable for de-dupe keys; NOT CASS-certified."""
    if not addr: return ""
    a = norm_space(addr).upper()
    # Normalize unit markers (drop for keying)
    a = re.sub(r"\b(APT|UNIT|STE|SUITE|#)\s*\w+\b", "", a)
    a = norm_space(a)
    # Standardize common street types at end tokens
    parts = a.split(" ")
    if parts:
        last = parts[-1]
        if last in STREET_ABBR:
            parts[-1] = STREET_ABBR[last]
        elif last.rstrip(".") in STREET_ABBR:
            parts[-1] = STREET_ABBR[last.rstrip(".")]
    a = " ".join(parts)
    return a

def norm_key(addr: str, owner: str) -> Tuple[str, str]:
    # address + owner normalized as key (upper for robustness)
    return (standardize_address(addr), norm_space(owner).upper())

def get_zip5(addr: str) -> str:
    m = re.search(r"(\d{5})(?:-\d{4})?$", addr or "")
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

def read_tracker() -> Dict[Tuple[str,str], Dict[str,str]]:
    if not os.path.exists(TRACKER_FILE):
        os.makedirs(TRACKER_DIR, exist_ok=True)
        with open(TRACKER_FILE, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f); w.writerow(
                ["PropertyAddress","OwnerName","CampaignCount","LastCampaignNumber","FirstSeenCampaign","FirstSeenDate","LastUpdatedDate"]
            )
        return {}
    rows = read_csv(TRACKER_FILE)
    d = {}
    for r in rows:
        k = norm_key(r.get("PropertyAddress",""), r.get("OwnerName",""))
        d[k] = r
    return d

# ---------------- Field detection ----------------

def required_cols(row: Dict[str,str]) -> Tuple[str,str]:
    # Try common column names for address/owner
    candidates_addr = ["PropertyAddress","SITUS ADDRESS","SITUS_ADDRESS","ADDRESS","MAILING ADDRESS","SITUS","MAILING_ADDRESS","MAILING ADDRESS 1"]
    candidates_owner = ["OwnerName","OWNER NAME","OWNER","OWNER(S)","OWNER 1","OWNER1","OWNER NAME 1"]
    addr = ""; own = ""
    lmap = {k.lower():k for k in row}
    for c in candidates_addr:
        if c in row and row[c].strip():
            addr = row[c].strip(); break
    if not addr and "address" in lmap:
        addr = row[lmap["address"]]
    for c in candidates_owner:
        if c in row and row[c].strip():
            own = row[c].strip(); break
    if not own and "owner" in lmap:
        own = row[lmap["owner"]]
    return addr, own

# ---------------- Merge, filter, optimize ----------------

def merge_sources(paths: List[str]) -> List[Dict[str,str]]:
    out = []
    for p in paths:
        if not p: continue
        out.extend(read_csv(p))
    return out

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

    # Sort buckets by size descending (ignore empty ZIP last)
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

    # Prefer buckets with size >= 100 to bring them to 150 if possible
    promising = sorted(((z5, lst) for z5,lst in by_zip5_remaining.items() if len(lst) >= 100),
                       key=lambda kv: len(kv[1]), reverse=True)
    for z5, lst in promising:
        need = 150
        # how many already chosen from that z5?
        already = sum(1 for r in chosen if get_zip5(r["PropertyAddress"]) == z5)
        take = max(0, min(need - already, len(lst), target - len(chosen)))
        if take > 0:
            random.shuffle(lst)
            chosen.extend(lst[:take])
        if len(chosen) >= target: break

    if len(chosen) >= target:
        return chosen[:target]

    # Stage 3: fill remaining by largest ZIP5 then ZIP3
    # Add the rest by ZIP5 concentration
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
                          rate_5: float, rate_3: float, rate_aadc: float) -> Tuple[float, Dict[str,Dict[str,float]]]:
    """
    Rough estimator:
      - ZIP5 with count >= 150 => 5-digit tranche billed at rate_5; remainder spills to lower tiers.
      - Make a second pass to aggregate small ZIP5s by ZIP3; buckets with >=150 go at rate_3.
      - Remaining pieces at AADC.
    Returns total_cost, breakdown dict.
    """
    total_pieces = sum(zip5_counts.values())
    five_digit = 0
    three_digit = 0
    aadc = 0

    # First pass: lock in 5-digit chunks
    leftovers_by_zip5 = {}
    for z5, c in zip5_counts.items():
        five_chunks = (c // 150) * 150
        five_digit += five_chunks
        leftovers_by_zip5[z5] = c - five_chunks

    # Second: roll leftovers by ZIP3
    by_zip3_left = collections.defaultdict(int)
    for z5, c in leftovers_by_zip5.items():
        z3 = z5[:3] if z5 else ""
        by_zip3_left[z3] += c
    for z3, c in by_zip3_left.items():
        three_chunks = (c // 150) * 150
        three_digit += three_chunks

    # The rest: AADC
    used = five_digit + three_digit
    aadc = max(0, total_pieces - used)

    cost = five_digit*rate_5 + three_digit*rate_3 + aadc*rate_aadc
    breakdown = {
        "5digit": {"pieces": five_digit, "rate": rate_5, "cost": five_digit*rate_5},
        "3digit": {"pieces": three_digit, "rate": rate_3, "cost": three_digit*rate_3},
        "AADC":   {"pieces": aadc,      "rate": rate_aadc, "cost": aadc*rate_aadc},
        "total":  {"pieces": total_pieces, "cost": cost, "avg": (cost/total_pieces if total_pieces else 0.0)}
    }
    return cost, breakdown

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

    # Postage rates (Commercial Marketing Mail letters, automation) – override if needed
    ap.add_argument("--rate-5digit", type=float, default=0.244)
    ap.add_argument("--rate-3digit", type=float, default=0.275)
    ap.add_argument("--rate-aadc", type=float, default=0.33)

    args = ap.parse_args()
    if args.prior_exact is not None and args.prior_max is not None:
        print("[ERROR] Use either --prior-exact OR --prior-max, not both.")
        sys.exit(1)

    # Load and validate lists
    if len(args.mandatory) > 4:
        print("[ERROR] Max 4 mandatory lists allowed."); sys.exit(1)
    if len(args.optional) > 2:
        print("[ERROR] Max 2 optional lists allowed."); sys.exit(1)

    mand = []
    for p in args.mandatory:
        if not os.path.exists(p):
            print(f"[ERROR] Mandatory list not found: {p}")
            sys.exit(1)
        mand.extend(read_csv(p))

    pools = []
    for p in args.optional:
        if not p: continue
        if not os.path.exists(p):
            print(f"[WARN] Optional list not found (skipping): {p}")
            continue
        pools.extend(read_csv(p))

    tracker = read_tracker()
    seen = set()
    all_candidates = []

    def add_rows(src_rows: List[Dict[str,str]]):
        for r in src_rows:
            addr, own = required_cols(r)
            if not addr or not own: 
                continue
            k = norm_key(addr, own)
            if k in seen:
                continue
            if not passes_prior_rules(k, tracker, args.prior_exact, args.prior_max, args.min_gap, args.campaign_number):
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

    add_rows(mand)
    if len(all_candidates) > args.target_size:
        print(f"[ERROR] Mandatory lists exceed target ({len(all_candidates)} > {args.target_size}). Refine your inputs.")
        sys.exit(1)
    add_rows(pools)

    # Pick optimized set
    if args.strict_150:
        chosen = pick_zip5_strict_150(all_candidates, args.target_size)
    else:
        chosen = pick_zip5_general(all_candidates, args.target_size)

    
    # Sort final selection for USPS-friendly print order: ZIP5 -> PropertyAddress -> OwnerName
    def _get_z5(row):
        try:
            return get_zip5(row.get("PropertyAddress", ""))
        except Exception:
            return ""
    chosen = sorted(chosen, key=lambda r: (_get_z5(r), r.get("PropertyAddress",""), r.get("OwnerName","")))

    # Build presort profiles
    by_zip5 = collections.Counter(get_zip5(r["PropertyAddress"]) for r in chosen)
    presort_rows = [{"ZIP5": z5 or "(none)", "Count": c} for z5, c in by_zip5.most_common()]
    # zip3 summary: total pieces and number of non-empty ZIP5s per ZIP3
    by_zip3 = {}
    for z5, cnt in by_zip5.items():
        z3 = get_zip3(z5)
        rec = by_zip3.setdefault(z3, {"EstZIP5Buckets":0, "TotalPieces":0})
        rec["EstZIP5Buckets"] += 1 if z5 else 0
        rec["TotalPieces"] += cnt
    presort_rows3 = [{"ZIP3": z3 or "(none)", **vals} for z3, vals in sorted(by_zip3.items())]

    # Campaign folder + outputs
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
    _, breakdown = estimate_blended_cost(by_zip5, args.rate_5digit, args.rate_3digit, args.rate_aadc)
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
