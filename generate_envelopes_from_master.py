#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Generate a collated PDF of #10 envelopes (9.5" x 4.125") from your campaign CSV,
including a USPS permit imprint (indicia) box in the top-right.

Key features (unchanged):
- FROM/TO default size: 14 pt (same font), Indicia: 11 pt.
- Indicia auto-fit (nudged up/right) and cap-height alignment of FROM to box top.
- Recipient block nudges: --recipient-offset-x / --recipient-offset-y (inches).
- Safe margins configurable: --safe-top/--safe-right/--safe-left/--safe-bottom.

NEW (optional, off by default so nothing breaks):
- **Bin separators for envelopes**: insert a one-page separator **before each USPS bin**
  (5-digit / 3-digit / AADC) based on mailing ZIPs, matching the letters tool’s behavior.
  Enable with `--bin-separators`. Also writes a manifest CSV by default to `RefFiles\envelope_bin_manifest.csv`
  (override via `--bin-manifest-out`).

Default output:
- If --out omitted, writes next to the CSV under "BatchEnvelopeFiles/envelopes_batch.pdf".
"""

import argparse
import csv
import os
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics

# ---- Page size for #10 envelope ----
W = 9.5 * inch
H = 4.125 * inch

# ---- Layout defaults ----
BASE_FONT = "Helvetica"
BASE_SIZE = 14  # FROM/TO default 14pt
REC_X = 3.75 * inch
REC_Y = (H / 2) + 0.40 * inch
RET_MARGIN_L = 0.50 * inch
RET_MARGIN_T = 0.50 * inch

# Safe margins (defaults; can be overridden via flags in main())
LEFT_SAFE = 0.50 * inch
RIGHT_SAFE = W - 0.50 * inch
BOTTOM_SAFE = 0.50 * inch
TOP_SAFE = H - 0.50 * inch

END_SIZE = 9
END_OFFSET = 6  # px below return block before endorsement

# ---- Indicia (permit imprint) defaults ----
INDICIA_FONT_DEFAULT = "Helvetica"
INDICIA_SIZE_DEFAULT = 11
INDICIA_PAD_IN_DEFAULT = 0.06
INDICIA_OFFSET_X_DEFAULT = 0.06  # inches, nudge right
INDICIA_OFFSET_Y_DEFAULT = 0.06  # inches, nudge up


def title_case_name(s: str) -> str:
    if not s:
        return s
    letters = [c for c in s if c.isalpha()]
    if letters and sum(1 for c in letters if c.isupper()) / len(letters) > 0.6:
        s = s.lower().title()
        for fx in ("LLC", "LP", "LLP", "INC", "TRUST", "REVOCABLE", "FAMILY"):
            s = s.replace(fx.title(), fx)
    return s


def get_primary_name(row):
    p_name = (row.get("Primary Name") or "").strip()
    if not p_name:
        first = (row.get("Primary First") or "").strip()
        last = (row.get("Primary Last") or "").strip()
        p_name = f"{first} {last}".strip()
    return title_case_name(p_name)


def get_secondary_name(row):
    s_name = (row.get("Secondary Name") or "").strip()
    if not s_name:
        s_first = (row.get("Secondary First") or "").strip()
        s_last = (row.get("Secondary Last") or "").strip()
        s_name = f"{s_first} {s_last}".strip()
    return title_case_name(s_name)


def compose_name_line(row):
    p = get_primary_name(row)
    s = get_secondary_name(row)
    if p and s:
        return f"{p} & {s}"
    return p or s


def pick_addr(row, key_mail, key_prop):
    v = (row.get(key_mail) or "").strip()
    return v if v else (row.get(key_prop) or "").strip()


def to_recipient_lines(row):
    name = compose_name_line(row)
    addr1 = pick_addr(row, "Mail Address", "Address")
    city = pick_addr(row, "Mail City", "City")
    state = pick_addr(row, "Mail State", "State")
    zip5 = pick_addr(row, "Mail ZIP", "ZIP")

    lines = []
    if name:
        lines.append(name)
    if addr1:
        lines.append(addr1)
    csz = " ".join([f"{city}, {state}".replace(" ,", ",").strip(), str(zip5).strip()]).strip()
    if csz:
        lines.append(csz)
    return [ln for ln in lines if ln]


def parse_return_arg(retarg: str):
    if not retarg:
        return []
    return [seg.strip() for seg in retarg.split("|") if seg.strip()]


def layout_indicia(c: canvas.Canvas, indicia_class, permit, city, state,
                   fit, pad_in, offset_x_in, offset_y_in, indicia_font, indicia_size,
                   box_w_in, box_h_in):
    """
    Compute indicia box geometry without drawing; return layout + text metrics.
    """
    lines = [indicia_class, "U.S. POSTAGE", "PAID", permit, city, state]
    lead = indicia_size + 2  # dynamic leading based on font size

    c.setFont(indicia_font, indicia_size)
    widths = [c.stringWidth(t, indicia_font, indicia_size) for t in lines]
    max_w = max(widths)
    total_h = len(lines) * lead - (lead - indicia_size)

    pad = pad_in * inch
    if fit:
        w = max_w + 2 * pad
        h = total_h + 2 * pad
    else:
        w = box_w_in * inch
        h = box_h_in * inch

    # Start at top-right inside safe margins, then apply offsets
    x = RIGHT_SAFE - w + (offset_x_in * inch)
    y = TOP_SAFE - h + (offset_y_in * inch)

    # Clamp to safe area
    x = min(max(LEFT_SAFE, x), RIGHT_SAFE - w)
    y = min(max(BOTTOM_SAFE, y), TOP_SAFE - h)

    return {
        "lines": lines,
        "widths": widths,
        "w": w,
        "h": h,
        "x": x,
        "y": y,
        "pad": pad,
        "lead": lead,
        "size": indicia_size,
        "font": indicia_font,
        "total_text_h": total_h,
        "top": y + h,
    }


def draw_indicia_from_layout(c: canvas.Canvas, L, line_width_pt: float):
    c.setLineWidth(line_width_pt)
    c.rect(L["x"], L["y"], L["w"], L["h"])
    c.setFont(L["font"], L["size"])
    # Center block vertically
    start_y = L["y"] + (L["h"] - L["total_text_h"]) / 2 + (len(L["lines"]) - 1) * L["lead"]
    for i, text in enumerate(L["lines"]):
        tw = L["widths"][i]
        tx = L["x"] + (L["w"] - tw) / 2
        ty = start_y - i * L["lead"]
        c.drawString(tx, ty, text)


def draw_return(c: canvas.Canvas, return_lines, endorsement, font_name, font_size, start_y=None):
    leading = font_size + 2
    x = RET_MARGIN_L
    y = start_y if start_y is not None else (H - RET_MARGIN_T)
    c.setFont(font_name, font_size)
    for line in return_lines:
        c.drawString(x, y, line)
        y -= leading
    if endorsement:
        y -= END_OFFSET
        c.setFont(font_name, END_SIZE)
        c.drawString(x, max(y, BOTTOM_SAFE), endorsement)


def draw_recipient(c: canvas.Canvas, lines, font_name, font_size, offset_x_in: float = 0.0, offset_y_in: float = 0.0):
    """offset_x_in: +right / -left; offset_y_in: +down / -up (in inches)."""
    leading = font_size + 2
    base_x = max(REC_X, LEFT_SAFE)
    base_y = min(REC_Y, TOP_SAFE - font_size)
    x = base_x + (offset_x_in * inch)
    y = base_y - (offset_y_in * inch)  # positive is DOWN
    # clamp to safe area
    x = min(max(LEFT_SAFE, x), RIGHT_SAFE - 1.0*inch)  # keep some right margin
    y = max(BOTTOM_SAFE + font_size, min(y, TOP_SAFE - font_size))
    c.setFont(font_name, font_size)
    for line in lines:
        if not line:
            continue
        c.drawString(x, max(y, BOTTOM_SAFE), line)
        y -= leading


def default_out_path(csv_path: Path) -> Path:
    base_dir = csv_path.parent
    out_dir = base_dir / "BatchEnvelopeFiles"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / "envelopes_batch.pdf"


# -------------- ZIP helpers (mailing-first) --------------

def _zip_from_text(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip()
    if s.endswith(".0"):
        s = s[:-2]
    import re
    m = re.search(r"(\d{5})(?:-\d{4})?$", s)
    return m.group(1) if m else ""

def get_zip_from_row_generic(r: dict) -> str:
    # 1) Mailing/Owner ZIPs
    for k in ("Mail ZIP","MAIL ZIP","Mail Zip","Mail Zip Code","MAIL ZIP CODE","MAIL ZIP5","Mail ZIP5",
              "MAILING ZIP","MAILING ZIP CODE","MAILING ZIP5","Owner ZIP","OWNER ZIP","Owner Zip","OWNER ZIP5","Owner ZIP5"):
        if k in r and str(r[k]).strip():
            z = _zip_from_text(r[k])
            if z: return z
    # 2) Mailing/Owner address strings
    for k in ("MAILING ADDRESS","Mailing Address","Mailing Address 1","Mailing Address1",
              "OWNER ADDRESS","Owner Address","OWNER MAILING ADDRESS","Owner Mailing Address"):
        if k in r and str(r[k]).strip():
            z = _zip_from_text(r[k])
            if z: return z
    # 3) Generic ZIPs
    for k in ("ZIP5","Zip5","ZIP","Zip","Zip Code","ZIP CODE","ZIP CODE 5"):
        if k in r and str(r[k]).strip():
            z = _zip_from_text(r[k])
            if z: return z
    # 4) Situs/Property (last)
    for k in ("SITUS ZIP","SITUS ZIP CODE","SITUS ZIP CODE 5-DIGIT","SITUS ZIP5","Situs ZIP","Situs Zip Code"):
        if k in r and str(r[k]).strip():
            z = _zip_from_text(r[k])
            if z: return z
    # parse from address fields
    for k in ("property_address","Property Address","PROPERTY ADDRESS","Address","ADDRESS","Situs Address","SITUS ADDRESS","PropertyAddress","SITUS"):
        if k in r and str(r[k]).strip():
            z = _zip_from_text(r[k])
            if z: return z
    return ""


# -------------- Bin planner (same logic as letters) --------------

def plan_bins_by_order(zips):
    """
    Compute bin spans without reordering the input.
    Input: zips = list of ZIP5 (one per envelope) in the exact order they will be printed.
    Output: list of bins with fields:
      - start (1-based letters index), end (inclusive), type ('5digit'|'3digit'|'aadc'), group (ZIP5 or ZIP3), count, id
    Assumptions: the input is roughly sorted by ZIP5 (builder does this), so ZIP3 ranges are contiguous.
    """
    from collections import Counter, defaultdict

    n = len(zips)
    counts_z5 = Counter(zips)
    trays_total = {z5: (counts_z5[z5] // 150) for z5 in counts_z5}
    trays_assigned = defaultdict(int)
    piece_in_current_tray = defaultdict(int)

    bins = []

    open_z3_start = None
    open_z3_count = 0
    open_z3_group = None

    current_5_start = None
    current_5_z5 = None

    for i, z5 in enumerate(zips, start=1):
        z3 = (z5 or "")[:3]

        # If ZIP3 changed and we had an open Z3 bin, close it as AADC
        if open_z3_count > 0 and open_z3_group is not None and z3 != open_z3_group:
            bins.append({
                "start": open_z3_start,
                "end": i - 1,
                "type": "aadc",
                "group": open_z3_group,
                "count": open_z3_count,
            })
            open_z3_start = None
            open_z3_count = 0
            open_z3_group = None

        # 5-digit assignment first
        if trays_assigned[z5] < trays_total[z5]:
            if current_5_start is None or current_5_z5 != z5:
                current_5_start = i
                current_5_z5 = z5
                piece_in_current_tray[z5] = 0
            piece_in_current_tray[z5] += 1
            if piece_in_current_tray[z5] == 150:
                bins.append({
                    "start": current_5_start,
                    "end": i,
                    "type": "5digit",
                    "group": z5,
                    "count": 150,
                })
                trays_assigned[z5] += 1
                current_5_start = None
                current_5_z5 = None
                piece_in_current_tray[z5] = 0
        else:
            # Leftover contributes to ZIP3/AADC
            if open_z3_count == 0:
                open_z3_start = i
                open_z3_group = z3
            open_z3_count += 1
            if open_z3_count == 150:
                bins.append({
                    "start": open_z3_start,
                    "end": i,
                    "type": "3digit",
                    "group": z3,
                    "count": 150,
                })
                open_z3_start = None
                open_z3_count = 0
                open_z3_group = None

    # Close any open z3 bin as AADC
    if open_z3_count > 0 and open_z3_group is not None:
        bins.append({
            "start": open_z3_start,
            "end": n,
            "type": "aadc",
            "group": open_z3_group,
            "count": open_z3_count,
        })

    # Assign IDs
    for idx, b in enumerate(bins, start=1):
        b["id"] = idx
    return bins


# -------------- Bin separator drawing (envelope-sized page) --------------

def draw_bin_separator_envelope(c: canvas.Canvas, bin_info: dict, csv_base: str):
    """
    Render a 1-page envelope-sized "BIN SEPARATOR" card with the bin details.
    This is intentionally simple and high-contrast.
    """
    c.setLineWidth(2)
    pad = 0.2 * inch
    c.rect(pad, pad, W - 2*pad, H - 2*pad)

    # Big header
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(W/2, H - 0.7*inch, "TRAY / BIN SEPARATOR")

    # Subheader
    label = {"5digit":"5-DIGIT ZIP","3digit":"3-DIGIT ZIP","aadc":"AADC"}.get(str(bin_info.get("type","")).lower(), str(bin_info.get("type","")).upper())
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H - 1.1*inch, label)

    # Details
    c.setFont("Helvetica", 12)
    group = str(bin_info.get("group",""))
    group_line = f"ZIP5: {group}" if bin_info.get("type")=="5digit" else (f"ZIP3: {group}" if bin_info.get("type")=="3digit" else f"ZIP3 group: {group}")
    lines = [
        group_line,
        f"Pieces in this bin: {int(bin_info.get('count',0))}",
        f"Letters index: {int(bin_info.get('start',0))}–{int(bin_info.get('end',0))}",
        f"Source: {csv_base}",
        f"Bin ID: {int(bin_info.get('id',0))}",
    ]
    y = H/2 + 0.1*inch
    for L in lines:
        c.drawCentredString(W/2, y, L)
        y -= 0.25*inch


def main():
    ap = argparse.ArgumentParser(description="Generate a single PDF with #10 envelopes from a CSV.")
    ap.add_argument("--csv", required=True, help="Path to input CSV.")
    ap.add_argument("--out", help="Output PDF path. Default: CSV folder/BatchEnvelopeFiles/envelopes_batch.pdf")
    ap.add_argument("--return", dest="return_block", default="", help="Return address as 'Line1|Line2|Line3'.")
    ap.add_argument("--endorsement", default="", help="Optional line under return, e.g. 'ADDRESS SERVICE REQUESTED'.")
    ap.add_argument("--limit", type=int, default=0, help="Only render first N rows (0 = all).")

    # Bin separators (optional)
    ap.add_argument("--bin-separators", action="store_true", help="Insert a 1-page separator BEFORE each USPS bin (5-digit/3-digit/AADC)")
    ap.add_argument("--bin-manifest-out", default=None, help="Optional manifest CSV path; default: <CSV dir>\\RefFiles\\envelope_bin_manifest.csv when --bin-separators is set")

    # Font controls (FROM/TO both use the same unless overridden)
    ap.add_argument("--font", default=BASE_FONT, help="Font name for FROM/TO (default Helvetica).")
    ap.add_argument("--size", type=int, default=BASE_SIZE, help="Font size for FROM/TO (default 14).")
    ap.add_argument("--return-size", type=int, default=None, help="Optional override for FROM size.")
    ap.add_argument("--recipient-size", type=int, default=None, help="Optional override for TO size.")

    # Indicia options
    ap.add_argument("--no-indicia", action="store_true", help="Disable the permit indicia box.")
    ap.add_argument("--no-indicia-fit", action="store_true", help="Disable auto-fit; use explicit width/height.")
    ap.add_argument("--indicia-w", type=float, default=1.6, help="Indicia width in inches (if not auto-fit).")
    ap.add_argument("--indicia-h", type=float, default=1.2, help="Indicia height in inches (if not auto-fit).")
    ap.add_argument("--indicia-class", default="PRSRT STD", help="Top line in indicia.")
    ap.add_argument("--permit", default="PMT #360", help="Permit line.")
    ap.add_argument("--permit-city", default="Carmichael", help="City line.")
    ap.add_argument("--permit-state", default="CA", help="State line.")
    ap.add_argument("--indicia-pad", type=float, default=INDICIA_PAD_IN_DEFAULT, help="Padding (inches) if auto-fit.")
    ap.add_argument("--indicia-offset-x", type=float, default=INDICIA_OFFSET_X_DEFAULT, help="Nudge indicia right (+)/left (-) in inches.")
    ap.add_argument("--indicia-offset-y", type=float, default=INDICIA_OFFSET_Y_DEFAULT, help="Nudge indicia up (+)/down (-) in inches.")
    ap.add_argument("--indicia-linewidth", type=float, default=1.0, help="Border line width (points).")
    ap.add_argument("--indicia-font", default=INDICIA_FONT_DEFAULT, help="Font for indicia text (default Helvetica).")
    ap.add_argument("--indicia-size", type=float, default=INDICIA_SIZE_DEFAULT, help="Font size for indicia text (default 11).")

    # Alignment controls
    ap.add_argument("--no-align-return-top", action="store_true", help="Do not align return first line to indicia top.")
    ap.add_argument("--align-mode", choices=["baseline", "cap"], default="cap",
                    help="Return alignment method to indicia top: baseline (same Y as box top) or cap (top of letters). Default cap.")

    # Recipient position nudges
    ap.add_argument("--recipient-offset-x", type=float, default=0.0, help="Move TO block right (+) / left (−) in inches.")
    ap.add_argument("--recipient-offset-y", type=float, default=0.0, help="Move TO block down (+) / up (−) in inches.")

    # Safe margin overrides (inches)
    ap.add_argument("--safe-top", type=float, default=0.5, help="Top safe margin in inches (default 0.5).")
    ap.add_argument("--safe-right", type=float, default=0.5, help="Right safe margin in inches (default 0.5).")
    ap.add_argument("--safe-left", type=float, default=0.5, help="Left safe margin in inches (default 0.5).")
    ap.add_argument("--safe-bottom", type=float, default=0.5, help="Bottom safe margin in inches (default 0.5).")

    args = ap.parse_args()

    # Override safe margins globally for this run
    global LEFT_SAFE, RIGHT_SAFE, TOP_SAFE, BOTTOM_SAFE
    LEFT_SAFE = args.safe_left * inch
    RIGHT_SAFE = W - (args.safe_right * inch)
    TOP_SAFE = H - (args.safe_top * inch)
    BOTTOM_SAFE = args.safe_bottom * inch

    csv_path = Path(args.csv)
    out_path = Path(args.out) if args.out else default_out_path(csv_path)

    font_name = args.font
    return_size = args.return_size if args.return_size is not None else args.size
    recipient_size = args.recipient_size if args.recipient_size is not None else args.size

    return_lines = parse_return_arg(args.return_block)
    endorsement = args.endorsement.strip()

    # Read all rows first (so we can plan bins if requested)
    with csv_path.open(newline="", encoding="utf-8-sig") as f:
        rdr = csv.DictReader(f)
        all_rows = [row for row in rdr]

    # Apply limit (only affects actual envelopes; separators are derived from these rows)
    if args.limit and args.limit > 0:
        all_rows = all_rows[:args.limit]

    # Pre-compute ZIP5 list (mailing-first)
    zips = [get_zip_from_row_generic(r) or "" for r in all_rows]

    # If enabled, compute bins (1-based indices into all_rows)
    bins = []
    if args.bin_separators and all_rows:
        bins = plan_bins_by_order(zips)

    # Prepare canvas
    out_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(out_path))
    c.setPageSize((W, H))

    def draw_one_envelope(row):
        lines = to_recipient_lines(row)
        if len(lines) < 2:
            return False

        # Layout indicia (even if we'll skip drawing, to compute alignment)
        L = None
        if not args.no_indicia:
            L = layout_indicia(
                c,
                args.indicia_class,
                args.permit,
                args.permit_city,
                args.permit_state,
                fit=not args.no_indicia_fit,
                pad_in=args.indicia_pad,
                offset_x_in=args.indicia_offset_x,
                offset_y_in=args.indicia_offset_y,
                indicia_font=args.indicia_font,
                indicia_size=args.indicia_size,
                box_w_in=args.indicia_w,
                box_h_in=args.indicia_h,
            )

        # Align return first line to top of indicia box
        if L and not args.no_align_return_top:
            if args.align_mode == "baseline":
                return_start_y = L["top"]
            else:
                ascent_pt = pdfmetrics.getAscent(font_name) * return_size / 1000.0
                return_start_y = L["top"] - ascent_pt
        else:
            return_start_y = H - RET_MARGIN_T

        draw_return(c, return_lines, endorsement, font_name, return_size, start_y=return_start_y)
        draw_recipient(
            c, lines, font_name, recipient_size,
            offset_x_in=args.recipient_offset_x,
            offset_y_in=args.recipient_offset_y
        )

        if L:
            draw_indicia_from_layout(c, L, line_width_pt=args.indicia_linewidth)

        c.showPage()
        return True

    # Draw pages
    env_count = 0
    sep_count = 0
    csv_base = os.path.splitext(os.path.basename(args.csv))[0]

    if bins:
        # Build manifest rows
        manifest_rows = []
        combined_page_cursor = 1
        for b in bins:
            # Separator page first
            draw_bin_separator_envelope(c, b, csv_base)
            c.showPage()
            sep_count += 1

            # Envelopes for this bin (use slice of all_rows by 1-based indices)
            start_idx = int(b["start"]) - 1
            end_idx = int(b["end"]) - 1
            for i in range(start_idx, end_idx + 1):
                if draw_one_envelope(all_rows[i]):
                    env_count += 1

            combined_start = combined_page_cursor  # separator page
            combined_end = combined_start + int(b["count"])  # sep + envelopes
            manifest_rows.append({
                "BinId": b["id"],
                "Type": b["type"],
                "Group": b["group"],
                "Pieces": b["count"],
                "LettersStart": b["start"],
                "LettersEnd": b["end"],
                "CombinedStart": combined_start,
                "CombinedEnd": combined_end,
            })
            combined_page_cursor = combined_end + 1

        # Write manifest if requested / default path
        if args.bin_separators:
            manifest_path = args.bin_manifest_out
            if not manifest_path:
                refs_dir = csv_path.parent / "RefFiles"
                refs_dir.mkdir(parents=True, exist_ok=True)
                manifest_path = refs_dir / "envelope_bin_manifest.csv"
            with open(manifest_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=["BinId","Type","Group","Pieces","LettersStart","LettersEnd","CombinedStart","CombinedEnd"])
                writer.writeheader()
                writer.writerows(manifest_rows)
            print(f"[BINS] Manifest: {manifest_path}  (bins: {len(bins)})")
    else:
        # Original behavior: one envelope per row, no separators
        for row in all_rows:
            if draw_one_envelope(row):
                env_count += 1

    c.save()
    if bins:
        print(f"Created {out_path} with {env_count} envelopes (+ {sep_count} bin separators).")
    else:
        print(f"Created {out_path} with {env_count} envelopes.")

if __name__ == "__main__":
    main()
