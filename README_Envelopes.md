# Envelopes — Address & Permit Indicia Generator

This adds a fast, scriptable way to generate **#10 envelope PDFs** from your `campaign_master.csv`, with a USPS **permit indicia** box (PRSRT STD / U.S. POSTAGE / PAID / PMT #360 / Carmichael / CA) in the top‑right.

> Script: `generate_envelopes_from_master.py` (lives in the MailMonkey root)

---

## What it does
- Reads your campaign CSV and lays out **one envelope per page** (9.5″ × 4.125″).
- Uses **Mail \*** columns first; falls back to property address if missing.
- Builds recipient name as `Primary & Secondary` (using `Primary Name` etc.).
- Prints **return (FROM)** and **recipient (TO)** in the **same font & size** by default.
- Draws a **permit indicia** box top‑right; **auto‑fit** to text is ON by default.
- Aligns the **FROM first line** to the **top of the indicia box** using cap‑height so they are visually level.

---

## Install
```bat
pip install reportlab
```

---

## Inputs → column mapping
Priority for each field:
- **Name:** `Primary Name` | (`Primary First` + `Primary Last`) [+ ` & ` + (`Secondary Name` | (`Secondary First` + `Secondary Last`))]
- **Address:** `Mail Address` | `Address`
- **City:** `Mail City` | `City`
- **State:** `Mail State` | `State`
- **ZIP:** `Mail ZIP` | `ZIP`

---

## Return address & endorsement
Pass your return block as pipe‑separated lines:
```bat
--return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827"
--endorsement "ADDRESS SERVICE REQUESTED"   (optional)
```

---

## Output location
By default (no `--out`):  
`<campaign folder>\BatchEnvelopeFiles\envelopes_batch.pdf` (folder is created next to your CSV).

Override with `--out PATH` if you prefer a different location.

---

## Permit indicia (top‑right box)
Defaults (edit via flags):
- Lines: `PRSRT STD` • `U.S. POSTAGE` • `PAID` • `PMT #360` • `Carmichael` • `CA`
- **Auto‑fit box** tightly around text (on by default).
- Tweak padding/size/position if needed.
- Return alignment: cap‑height aligned to **top of box** by default (`--align-mode cap`).

Flags:
```
--no-indicia             Disable the box entirely
--no-indicia-fit         Use fixed size instead of auto‑fit
--indicia-w 1.6          Box width (in) if not auto‑fit
--indicia-h 1.2          Box height (in) if not auto‑fit
--indicia-pad 0.06       Inner padding (in) when auto‑fit (default 0.06)
--indicia-offset-x 0.06  Nudge right (+) / left (−) in inches
--indicia-offset-y 0.06  Nudge up (+) / down (−) in inches
--indicia-linewidth 1.0  Box border thickness (pt)
--indicia-font Helvetica Font for indicia text
--indicia-size 11        Font size for indicia text
--indicia-class "PRSRT STD"
--permit "PMT #360"
--permit-city "Carmichael"
--permit-state "CA"
--align-mode cap         Return alignment: cap|baseline (default cap)
--no-align-return-top    Turn off alignment to box top entirely
```

> **Make the box smaller:** keep auto‑fit on and lower `--indicia-pad` (e.g., `0.04`).

---

## Fonts & sizes (FROM/TO)
- Default font: **Helvetica**.
- Default size: **14 pt** for both FROM and TO (indicia defaults to 11 pt).

Flags:
```
--font Helvetica       # Font family for both FROM and TO
--size 14              # Shared size for FROM & TO
--return-size 13       # Override FROM only
--recipient-size 15    # Override TO only
```

---

## Usage examples

**From the campaign folder (standard run):**
```bat
cd "C:\Users\Edit Beluli\Desktop\MailMonkey\Campaign_1_Aug2025"
python ..\generate_envelopes_from_master.py ^
  --csv "campaign_master.csv" ^
  --return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827" ^
  --endorsement "ADDRESS SERVICE REQUESTED"
```

**Bigger text + tighter indicia box:**
```bat
python ..\generate_envelopes_from_master.py ^
  --csv "campaign_master.csv" ^
  --return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827" ^
  --size 16 ^
  --indicia-size 12 ^
  --indicia-pad 0.04
```

**Fixed-size indicia, thinner border, nudged up/right:**
```bat
python ..\generate_envelopes_from_master.py ^
  --csv "campaign_master.csv" ^
  --return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827" ^
  --no-indicia-fit --indicia-w 1.6 --indicia-h 1.0 ^
  --indicia-linewidth 0.6 --indicia-offset-x 0.10 --indicia-offset-y 0.12
```

**Disable indicia (test prints):**
```bat
python ..\generate_envelopes_from_master.py ^
  --csv "campaign_master.csv" ^
  --return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827" ^
  --no-indicia
```

**One-liner with nudges (no carets):**
```bat
python ..\generate_envelopes_from_master.py --csv "campaign_master.csv" --return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827" --endorsement "ADDRESS SERVICE REQUESTED" --indicia-offset-x 0.10 --indicia-offset-y 0.12
```

**Version the output file (don’t overwrite):**
```bat
python ..\generate_envelopes_from_master.py ^
  --csv "campaign_master.csv" ^
  --return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827" ^
  --out "BatchEnvelopeFiles\envelopes_batch_v2.pdf"
```

---

## Windows CMD caret tips (to avoid errors)
- The caret `^` must be the **last character on the line** (no trailing spaces).
- All flags (like `--indicia-offset-y`) must be part of the **same** `python ...` command.
- Running a flag **by itself** will produce:  
  `'--indicia-offset-y' is not recognized as an internal or external command...`

Example (correct multi‑line):
```bat
python ..\generate_envelopes_from_master.py ^
  --csv "campaign_master.csv" ^
  --return "Albert Beluli|9626 Knickers Ct|Sacramento, CA 95827" ^
  --endorsement "ADDRESS SERVICE REQUESTED" ^
  --indicia-offset-x 0.10 ^
  --indicia-offset-y 0.12
```

---

## Printing tips (Brother laser, #10)
- In the driver: **Paper Size = #10 Envelope**; choose the **envelope/manual tray**.
- Load with flap closed, print face per printer icon.
- Turn **duplex OFF**; if smudging, enable **Thick/Envelope** fuser mode.
- Preview a handful first: `--limit 5`.

---

## Troubleshooting
- **Nothing prints / PDF looks blank** → Make sure ReportLab is installed: `pip install reportlab`.
- **Weird fonts** → Stick with Helvetica (bundled); or install the desired font and switch with `--font`.
- **Box too big** → Lower `--indicia-pad` (auto‑fit) or use `--no-indicia-fit` with smaller `--indicia-w/--indicia-h`.
- **Address lines wrap** → Slightly reduce `--size` or move recipient block left by editing `REC_X` inside the script.
