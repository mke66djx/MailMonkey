
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
BuildMasterCampaignList_v2_DEBUG_ZIPFIX.py
- Same as v2_DEBUG, but ZIP5 extraction will also look at common ZIP columns
  (e.g., ZIP, Zip, SITUS ZIP, Situs Zip Code, Mail ZIP) when the address line
  doesn't include a ZIP code.
"""

import os, sys, csv, re, argparse, datetime, random, collections
from typing import List, Dict, Tuple, Optional

TRACKER_DIR = "MasterCampaignTracker"
TRACKER_FILE = os.path.join(TRACKER_DIR, "MasterPropertyCampaignCounter.csv")

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
    import re
    return re.sub(r"\s+", " ", (s or "").strip())

def standardize_address(addr: str) -> str:
    if not addr: return ""
    a = norm_space(addr).upper()
    a = re.sub(r"\b(APT|UNIT|STE|SUITE|#)\s*\w+\b", "", a)
    a = norm_space(a)
    parts = a.split(" ")
    if parts:
        last = parts[-1].rstrip(".")
        if last in STREET_ABBR:
            parts[-1] = STREET_ABBR[last]
    return " ".join(parts)

def norm_key(addr: str, owner: str) -> Tuple[str, str]:
    return (standardize_address(addr), norm_space(owner).upper())

def get_zip5_from_text(text: str) -> str:
    if not text: return ""
    m = re.search(r"(\d{5})(?:-\d{4})?$", str(text))
    return m.group(1) if m else ""

def extract_zip5_from_row(row: Dict[str,str], addr: str) -> str:
    # Try from the address string first
    z = get_zip5_from_text(addr)
    if z: return z
    # Try common ZIP fields
    zip_keys = [
        "ZIP","Zip","zip",
        "Zip Code","ZIP CODE","ZipCode","ZIPCODE",
        "SITUS ZIP","SITUS ZIP CODE","Situs Zip","Situs Zip Code",
        "MAIL ZIP","Mail ZIP","Mail Zip","Mail Zip Code",
        "MAILING ZIP","Mailing ZIP","Mailing Zip","Mailing Zip Code",
        "PROPERTY ZIP","Property ZIP","Property Zip"
    ]
    for k in zip_keys:
        if k in row and str(row[k]).strip():
            val = str(row[k]).strip()
            val = re.sub(r"\.0$","", val)  # handle 95834.0
            z = get_zip5_from_text(val)
            if z: return z
    return ""

def get_zip3(zip5: str) -> str:
    return zip5[:3] if zip5 else ""

def read_csv(path: str):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [ {k: (v or "").strip() for k,v in row.items()} for row in r ]

def write_csv(path: str, rows, headers):
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

def read_tracker():
    ensure_tracker()
    rows = read_csv(TRACKER_FILE)
    d = {}
    for r in rows:
        key = (standardize_address(r.get("PropertyAddress","")), (r.get("OwnerName","") or "").strip().upper())
        d[key] = r
    return d

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

def required_cols(row):
    addr = ""
    lmap = {k.lower():k for k in row}
    for c in ADDR_CANDIDATES:
        if c in row and row[c].strip():
            addr = row[c].strip(); break
    if not addr and "address" in lmap:
        addr = row[lmap["address"]]

    own = ""
    for c in OWNER_CANDIDATES:
        if c in row and row[c].strip():
            own = row[c].strip(); break
    if not own:
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

def campaign_folder(campaign_name: str, campaign_number: int, when: Optional[datetime.date]=None) -> str:
    when = when or datetime.date.today()
    mo_yr = when.strftime("%b%Y")
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

def pick_zip5_strict_150(candidates, target):
    by_zip5 = collections.defaultdict(list)
    for r in candidates:
        by_zip5[r.get("ZIP5","")].append(r)
    ordered = sorted(by_zip5.items(), key=lambda kv: (len(kv[1]), kv[0] != ""), reverse=True)
    chosen = []
    for z5, bucket in ordered:
        if len(chosen) >= target: break
        n = len(bucket); take = (n // 150) * 150
        if take == 0: continue
        random.shuffle(bucket)
        chosen.extend(bucket[:min(take, target - len(chosen))])
    if len(chosen) >= target:
        return chosen[:target]
    remaining = [r for r in candidates if r not in chosen]
    by_zip5_remaining = collections.defaultdict(list)
    for r in remaining:
        by_zip5_remaining[r.get("ZIP5","")].append(r)
    promising = sorted(((z5, lst) for z5,lst in by_zip5_remaining.items() if len(lst) >= 100),
                       key=lambda kv: len(kv[1]), reverse=True)
    for z5, lst in promising:
        need = 150
        already = sum(1 for r in chosen if r.get("ZIP5","") == z5)
        take = max(0, min(need - already, len(lst), target - len(chosen)))
        if take > 0:
            random.shuffle(lst)
            chosen.extend(lst[:take])
        if len(chosen) >= target: break
    if len(chosen) >= target:
        return chosen[:target]
    remaining = [r for r in candidates if r not in chosen]
    by_zip5_rem2 = collections.defaultdict(list)
    for r in remaining:
        by_zip5_rem2[r.get("ZIP5","")].append(r)
    rem_order = sorted(by_zip5_rem2.items(), key=lambda kv: len(kv[1]), reverse=True)
    for z5, lst in rem_order:
        random.shuffle(lst)
        for r in lst:
            if len(chosen) >= target: break
            chosen.append(r)
        if len(chosen) >= target: break
    if len(chosen) < target:
        remaining = [r for r in candidates if r not in chosen]
        by_zip3 = collections.defaultdict(list)
        for r in remaining:
            by_zip3[(r.get("ZIP5","") or "")[:3]].append(r)
        rem3 = sorted(by_zip3.items(), key=lambda kv: len(kv[1]), reverse=True)
        for z3, lst in rem3:
            random.shuffle(lst)
            for r in lst:
                if len(chosen) >= target: break
                chosen.append(r)
            if len(chosen) >= target: break
    return chosen[:target]

def pick_zip5_general(candidates, target):
    if target <= 0: return []
    by_zip5 = collections.defaultdict(list)
    for r in candidates:
        by_zip5[r.get("ZIP5","")].append(r)
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
        by_zip3[(r.get("ZIP5","") or "")[:3]].append(r)
    zip3_buckets = sorted(by_zip3.items(), key=lambda kv: len(kv[1]), reverse=True)
    for z3, bucket in zip3_buckets:
        random.shuffle(bucket)
        for row in bucket:
            if len(chosen) >= target: break
            chosen.append(row)
        if len(chosen) >= target: break
    return chosen[:target]

def estimate_blended_cost(zip5_counts, rate_5, rate_3, rate_aadc):
    total_pieces = sum(zip5_counts.values())
    five_digit = 0; three_digit = 0
    leftovers_by_zip5 = {}
    for z5, c in zip5_counts.items():
        five_chunks = (c // 150) * 150
        five_digit += five_chunks
        leftovers_by_zip5[z5] = c - five_chunks
    by_zip3_left = collections.defaultdict(int)
    for z5, c in leftovers_by_zip5.items():
        by_zip3_left[(z5 or "")[:3]] += c
    for z3, c in by_zip3_left.items():
        three_digit += (c // 150) * 150
    used = five_digit + three_digit
    aadc = max(0, total_pieces - used)
    cost = five_digit*rate_5 + three_digit*rate_3 + aadc*rate_aadc
    avg = (cost/total_pieces) if total_pieces else 0.0
    return {
        "5digit": {"pieces": five_digit, "rate": rate_5, "cost": five_digit*rate_5},
        "3digit": {"pieces": three_digit, "rate": rate_3, "cost": three_digit*rate_3},
        "AADC":   {"pieces": aadc,      "rate": rate_aadc, "cost": aadc*rate_aadc},
        "total":  {"pieces": total_pieces, "cost": cost, "avg": avg}
    }

def main():
    ap = argparse.ArgumentParser(description="Build USPS-optimized master list, filtered by prior campaign history.")
    ap.add_argument("--campaign-name", required=True)
    ap.add_argument("--campaign-number", type=int, required=True)
    ap.add_argument("--target-size", type=int, required=True)
    ap.add_argument("--mandatory", nargs="+", required=True, help="1-4 CSVs required")
    ap.add_argument("--optional", nargs="*", default=[], help="0-2 CSVs optional pools")
    ap.add_argument("--prior-exact", type=int, help="Only include entries with exactly N prior campaigns (0=never mailed)")
    ap.add_argument("--prior-max", type=int, help="Only include entries with ≤ M prior campaigns")
    ap.add_argument("--min-gap", type=int, default=0, help="Require last campaign be ≥ this many campaign numbers ago")
    ap.add_argument("--strict-150", action="store_true", help="Favor multiples of 150 per ZIP5 where possible")
    ap.add_argument("--rate-5digit", type=float, default=0.244)
    ap.add_argument("--rate-3digit", type=float, default=0.275)
    ap.add_argument("--rate-aadc", type=float, default=0.330)
    ap.add_argument("--debug", action="store_true", help="Print verbose row-diagnostic stats")
    args = ap.parse_args()

    if args.prior_exact is not None and args.prior_max is not None:
        print("[ERROR] Use either --prior-exact OR --prior-max, not both."); sys.exit(1)
    if len(args.mandatory) > 4:
        print("[ERROR] Max 4 mandatory lists allowed."); sys.exit(1)
    if len(args.optional) > 2:
        print("[ERROR] Max 2 optional lists allowed."); sys.exit(1)

    # Load tracker and ingest rows
    def read_tracker():
        if not os.path.exists(TRACKER_FILE):
            os.makedirs(TRACKER_DIR, exist_ok=True)
            with open(TRACKER_FILE, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f); w.writerow(
                    ["PropertyAddress","OwnerName","CampaignCount","LastCampaignNumber","FirstSeenCampaign","FirstSeenDate","LastUpdatedDate"]
                )
        rows = read_csv(TRACKER_FILE)
        d = {}
        for r in rows:
            key = (standardize_address(r.get("PropertyAddress","")), (r.get("OwnerName","") or "").strip().upper())
            d[key] = r
        return d

    tracker = read_tracker()
    seen = set()
    all_candidates = []

    mand_stats = {"missing_addr":0,"missing_owner":0,"dropped_prior":0,"deduped":0,"kept":0}
    pool_stats = {"missing_addr":0,"missing_owner":0,"dropped_prior":0,"deduped":0,"kept":0}

    def add_rows(src_rows, tag, stats):
        for i, r in enumerate(src_rows, 1):
            addr, own = required_cols(r)
            if not addr:
                stats["missing_addr"] += 1
                if args.debug and i <= 3:
                    print(f"[DEBUG] {tag}: row {i} missing address")
                continue
            if not own:
                stats["missing_owner"] += 1
                if args.debug and i <= 3:
                    print(f"[DEBUG] {tag}: row {i} missing owner")
                continue
            k = (standardize_address(addr), (own or "").strip().upper())
            if k in seen:
                stats["deduped"] += 1
                continue
            if not passes_prior_rules(k, tracker, args.prior_exact, args.prior_max, args.min_gap, args.campaign_number):
                stats["dropped_prior"] += 1
                continue
            row = {
                "PropertyAddress": norm_space(addr),
                "OwnerName": norm_space(own),
            }
            row["ZIP5"] = extract_zip5_from_row(r, addr)
            all_candidates.append(row)
            seen.add(k); stats["kept"] += 1

    # Read files
    mand_rows = []
    for pth in args.mandatory:
        if not os.path.exists(pth):
            print(f"[ERROR] Mandatory list not found: {pth}"); sys.exit(1)
        rows = read_csv(pth); mand_rows.extend(rows)
        if args.debug: print(f"[DEBUG] Reading mandatory: {pth} (rows={len(rows)})")
    add_rows(mand_rows, "MAND", mand_stats)

    if len(all_candidates) > args.target_size:
        print(f"[ERROR] Mandatory lists exceed target ({len(all_candidates)} > {args.target_size}). Refine your inputs."); sys.exit(1)

    pool_rows = []
    for pth in args.optional:
        if not pth: continue
        if not os.path.exists(pth):
            print(f"[WARN] Optional list not found (skipping): {pth}"); continue
        rows = read_csv(pth); pool_rows.extend(rows)
        if args.debug: print(f"[DEBUG] Reading optional: {pth} (rows={len(rows)})")
    add_rows(pool_rows, "POOL", pool_stats)

    if args.debug:
        print("[DEBUG] Summary after ingest:")
        print(f"  MAND kept={mand_stats['kept']}  deduped={mand_stats['deduped']}  dropped_prior={mand_stats['dropped_prior']}  missing_addr={mand_stats['missing_addr']}  missing_owner={mand_stats['missing_owner']}")
        print(f"  POOL kept={pool_stats['kept']}  deduped={pool_stats['deduped']}  dropped_prior={pool_stats['dropped_prior']}  missing_addr={pool_stats['missing_addr']}  missing_owner={pool_stats['missing_owner']}")
        print(f"  TOTAL candidates={len(all_candidates)}")

    # Choose and sort
    if args.strict_150:
        chosen = pick_zip5_strict_150(all_candidates, args.target_size)
    else:
        chosen = pick_zip5_general(all_candidates, args.target_size)
    chosen = sorted(chosen, key=lambda r: (r.get("ZIP5",""), r.get("PropertyAddress",""), r.get("OwnerName","")))

    # Presort profile
    by_zip5 = collections.Counter([r.get("ZIP5","") for r in chosen])
    presort_rows = [{"ZIP5": z5 or "(none)", "Count": c} for z5, c in by_zip5.most_common()]
    by_zip3 = {}
    for z5, cnt in by_zip5.items():
        z3 = (z5 or "")[:3]
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

    headers = ["PropertyAddress","OwnerName","ZIP5"]
    write_csv(master_path, chosen, headers)
    write_csv(presort_path, presort_rows, ["ZIP5","Count"])
    write_csv(presort_zip3_path, presort_rows3, ["ZIP3","EstZIP5Buckets","TotalPieces"])

    # Postage estimate
    rate_5 = args.rate_5digit; rate_3 = args.rate_3digit; rate_aadc = args.rate_aadc
    total_pieces = sum(by_zip5.values())
    five = three = 0
    leftovers = {}
    for z5, c in by_zip5.items():
        take = (c // 150) * 150
        five += take; leftovers[z5] = c - take
    by_z3_left = collections.defaultdict(int)
    for z5, c in leftovers.items():
        by_z3_left[(z5 or "")[:3]] += c
    for z3, c in by_z3_left.items():
        three += (c // 150) * 150
    aadc = max(0, total_pieces - (five + three))
    cost = five*rate_5 + three*rate_3 + aadc*rate_aadc
    avg = cost/total_pieces if total_pieces else 0.0

    with open(postage_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Tier","Pieces","Rate","Cost"])
        w.writerow(["5digit", five, f"{rate_5:.3f}", f"{five*rate_5:.2f}"])
        w.writerow(["3digit", three, f"{rate_3:.3f}", f"{three*rate_3:.2f}"])
        w.writerow(["AADC", aadc, f"{rate_aadc:.3f}", f"{aadc*rate_aadc:.2f}"])
        w.writerow(["total", total_pieces, "", f"{cost:.2f}"])
        w.writerow(["AveragePerPiece","","", f"{avg:.4f}"])

    print(f"[OK] Created campaign folder: {camp_dir}")
    print(f"[OK] Master list: {master_path}  (rows={len(chosen)})")
    print(f"[OK] Presort ZIP5: {presort_path}")
    print(f"[OK] Presort ZIP3: {presort_zip3_path}")
    print(f"[OK] Postage estimate: {postage_path}")
    print(f"[SUMMARY] Estimated blended avg: ${avg:.4f}  (5digit={five}, 3digit={three}, AADC={aadc})")

if __name__ == "__main__":
    main()
