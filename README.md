# MailMonkey – README (High-Level Guide + How-Tos)

*Last updated: Aug 24, 2025*

---

## 1) What this project does (high-level)

**MailMonkey** is a small toolchain that turns raw property/owner CSV lists into print-ready direct-mail letters, and maintains a clean history of who you mailed and when. It’s designed to be **idempotent** (safe to re-run) and **recoverable** (you can rebuild the tracker from past campaign folders at any time).

### Key ideas
- **Builder** selects who to mail (from one or more CSV lists), using rules like “never mailed”, “exactly 3 priors”, or “last mailed ≥ 30 days ago”. It outputs a campaign folder with a `campaign_master.csv` and presort reports.
- **Generator** makes the letters (combined PDF + optional singles) from `campaign_master.csv` and logs which template ID you used.
- **Finalizer** appends a per-campaign `executed_campaign_log.csv`, updates the Master tracker (one row per unique `PropertyAddress + OwnerName`), and rebuilds a ZIP5 tally. It’s idempotent and can also rebuild everything from the campaign folders (disaster recovery).
- **Mailing ZIP first**: USPS cares about the owner/mailing ZIP, not the property/situs ZIP. All presort and tallies favor the mailing ZIP.

---

## 2) Directory map (what lives where)

**Root example:** `C:\Users\Edit Beluli\Desktop\MailMonkey`

```text
BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP.py   # Builder with time filters (recommended)
BuildMasterCampaignList_v4_MAILZIPFirst.py           # Builder without time filters
direct_mail_batch_por_POR_KEEP_FIXINDENT.py          # Letter generator
FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py  # Finalizer + recovery (recommended)

MasterCampaignTracker\
  MasterPropertyCampaignTracker.csv   # 1 row per (PropertyAddress, OwnerName)
  Zip5_LetterTally.csv               # Totals by mailing ZIP5

PropertyLists\                        # Your source CSVs
  Foreclosure_08_2025.csv
  PropertyTaxDelinquentList_08_2025.csv
  LienList_ZipCodes_08_2025.csv

Campaign_1_Aug2025\                   # One folder per campaign
  campaign_master.csv                 # Built by the Builder
  presort_report.csv                  # ZIP5 counts
  presort_zip3_summary.csv            # ZIP3 summary
  postage_estimate.csv                # Cost estimate by presort tier
  letters_mapping.csv  (or RefFiles\letters_mapping.csv)
  executed_campaign_log.csv           # Finalizer appends rows here
  BatchLetterFiles\
  Singles\
  CAMPAIGN.TAG                        # (optional) marker file

OlderFiles\                           # Archived legacy scripts (optional)
```

> **Note:** The generator writes `letters_mapping.csv` either directly in the campaign folder or under `RefFiles\letters_mapping.csv` (both are supported by the finalizer).

---

## 3) Script overview (what each script does)

### A) Builder — `BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP.py` *(recommended)*
**Purpose:** Pick a USPS-friendly set of records from your lists. Uses “mailing ZIP first” logic and supports time-based filters.

**Outputs into** `Campaign_<N>_<MonYYYY>\`:
- `campaign_master.csv`
- `presort_report.csv` (ZIP5 counts)
- `presort_zip3_summary.csv`
- `postage_estimate.csv`

**Key filters**
- `--prior-exact N` → only rows with exactly **N** prior campaigns  
- `--prior-max M` → only rows with **≤ M** prior campaigns  
- `--min-gap G` → campaign-number gap (e.g., if current is 7 and `G=1`, anyone last in 7 is excluded; last in ≤6 is OK)

**Time filters (new)**
- `--min-days-since-last D` → require `LastSentDt ≥ D` days ago  
- `--last-sent-before DATE` → require `LastSentDt < DATE` (`YYYY-MM-DD` or `MM/DD/YYYY`)  
- `--missing-last-sent {fail|include}` *(default `fail`)* → when time filters are set and `LastSentDt` is missing, exclude (`fail`) or allow (`include`)

**USPS optimization**
- `--strict-150` packs ZIP5s to favor trays in multiples of 150 before filling.  
- `postage_estimate.csv` uses configurable rates (`--rate-5digit`, `--rate-3digit`, `--rate-aadc`).

---

### B) Builder — `BuildMasterCampaignList_v4_MAILZIPFirst.py`
Same as above **without** the time-based filters. Still uses mailing ZIP first and supports `--prior-exact`, `--prior-max`, `--min-gap`, `--strict-150`.

---

### C) Generator — `direct_mail_batch_por_POR_KEEP_FIXINDENT.py`
**Purpose:** Generate the combined letters PDF and per-row mapping referencing your chosen template ID.

**Outputs into the current** `Campaign_<N>_<MonYYYY>\` **folder:**
- `BatchLetterFiles\...letters_batch.pdf` (combined PDF)
- `letters_mapping.csv` *(or `RefFiles\letters_mapping.csv`)*
- `Singles\` with individual PDFs unless you pass `--skip-singles`

**Template IDs**
- Pass `--template-id 101` (or any other ID you’ve defined). The mapping and logs record the numeric ID you used.

---

### D) Finalizer + Recovery — `FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py` *(recommended)*
**Purpose:** Append the campaign’s executed log, update the master tracker (idempotent), rebuild ZIP5 tally, and support full recovery by scanning all campaign folders.

**Tracker schema** *(per `(PropertyAddress, OwnerName)`)*  
- `ZIP5` (mailing ZIP)  
- `CampaignNumbers` → unique, sorted list like `1|2|5`  
- `CampaignCount` → count of unique campaign numbers  
- `TemplateIds` → sequence allowing duplicates (e.g., `101|101|303`)  
- `FirstSentDt`, `LastSentDt`

**Disaster-recovery**
- `--rebuild-all` *(or `--reindex-all`)* scans every folder containing `executed_campaign_log.csv` and rebuilds the tracker and ZIP tally from scratch. Uses each folder’s `campaign_master.csv` to backfill missing mailing ZIPs.
- **Markers:** `--write-marker` creates an empty `CAMPAIGN.TAG` file in the campaign folder; then `--marker-required --marker-name CAMPAIGN.TAG` tells recovery to only trust marked folders.

**Idempotence**
- Finalizer won’t re-append the same rows if you re-run (dedupes by `(OwnerName+PropertyAddress+CampaignNumber)` and `RefCode` when present). `CampaignCount` recomputes from unique `CampaignNumbers`.

---

## 4) Typical workflows

### Fresh two-campaign test (same pairs twice, mailing ZIP first)

**1) Campaign 1 – build (never mailed):**
```bat
cd "C:\Users\Edit Beluli\Desktop\MailMonkey"
python BuildMasterCampaignList_v4_MAILZIPFirst.py ^
  --campaign-name "Campaign" ^
  --campaign-number 1 ^
  --target-size 5000 ^
  --mandatory "PropertyLists\Foreclosure_08_2025.csv" "PropertyLists\PropertyTaxDelinquentList_08_2025.csv" ^
  --optional "PropertyLists\LienList_ZipCodes_08_2025.csv" ^
  --prior-exact 0 ^
  --strict-150 ^
  --debug
```

**2) Generate letters (combined only):**
```bat
cd "C:\Users\Edit Beluli\Desktop\MailMonkey\Campaign_1_Aug2025"
python ..\direct_mail_batch_por_POR_KEEP_FIXINDENT.py ^
  --csv "campaign_master.csv" ^
  --outdir "Singles" ^
  --combine-out "letters_batch.pdf" ^
  --map-out "letters_mapping.csv" ^
  --template-id 101 ^
  --skip-singles ^
  --sig-image "..\sig_ed.png" ^
  --name "Ed & Albert Beluli" ^
  --phone "916-905-7281" ^
  --email "eabeluli@gmail.com"
```

**3) Finalize:**
```bat
cd "C:\Users\Edit Beluli\Desktop\MailMonkey"
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --campaign-dir "Campaign_1_Aug2025" ^
  --write-marker
```

**4) Campaign 2 – build (exactly once mailed):**
```bat
python BuildMasterCampaignList_v4_MAILZIPFirst.py ^
  --campaign-name "Campaign" ^
  --campaign-number 2 ^
  --target-size 5000 ^
  --mandatory "PropertyLists\Foreclosure_08_2025.csv" "PropertyLists\PropertyTaxDelinquentList_08_2025.csv" ^
  --optional "PropertyLists\LienList_ZipCodes_08_2025.csv" ^
  --prior-exact 1 ^
  --strict-150 ^
  --debug
```

**5) Generate & finalize:**
```bat
cd "C:\Users\Edit Beluli\Desktop\MailMonkey\Campaign_2_Aug2025"
python ..\direct_mail_batch_por_POR_KEEP_FIXINDENT.py ^
  --csv "campaign_master.csv" ^
  --outdir "Singles" ^
  --combine-out "letters_batch.pdf" ^
  --map-out "letters_mapping.csv" ^
  --template-id 101 ^
  --skip-singles ^
  --sig-image "..\sig_ed.png" ^
  --name "Ed & Albert Beluli" ^
  --phone "916-905-7281" ^
  --email "eabeluli@gmail.com"

cd "C:\Users\Edit Beluli\Desktop\MailMonkey"
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --campaign-dir "Campaign_2_Aug2025" ^
  --write-marker
```

**Result:** In the tracker, shared pairs will show `CampaignNumbers = 1|2`, `CampaignCount = 2`, and `TemplateIds` reflecting the template sequence (e.g., `101|101`).

---

### Re-mail same cohort with different templates
Build for `--prior-exact 2`, then 3, etc., and pass a new `--template-id` each time. The tracker’s `TemplateIds` stores the sequence (e.g., `101|101|303|404`).

### Time-based re-mailing
*(e.g., “exactly 4 priors AND last sent ≥ 30 days ago”)*

```bat
python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP.py ^
  --campaign-name "Campaign" ^
  --campaign-number 7 ^
  --target-size 5000 ^
  --mandatory "PropertyLists\Foreclosure_08_2025.csv" "PropertyLists\PropertyTaxDelinquentList_08_2025.csv" ^
  --optional "PropertyLists\LienList_ZipCodes_08_2025.csv" ^
  --prior-exact 4 ^
  --min-days-since-last 30 ^
  --strict-150 ^
  --debug
```
*(Or use `--last-sent-before 2025-07-01`.)*

---

## 5) Finalizer usage patterns (with examples)

**Normal finalize**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --campaign-dir "Campaign_5_Aug2025" ^
  --write-marker
```

**Dry run (preview)**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --campaign-dir "Campaign_5_Aug2025" ^
  --dry-run
```

**If your mapping CSV is in a nonstandard location**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --campaign-dir "Campaign_5_Aug2025" ^
  --mapping "Campaign_5_Aug2025\letters_mapping.csv"
```

**Disaster recovery – rebuild tracker & ZIP tally from all campaign folders**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --rebuild-all ^
  --root "C:\Users\Edit Beluli\Desktop\MailMonkey" ^
  --tracker-path "MasterCampaignTracker\MasterPropertyCampaignTracker.csv"
```

**Same, but only consider folders with a marker**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --rebuild-all ^
  --root "C:\Users\Edit Beluli\Desktop\MailMonkey" ^
  --marker-required ^
  --marker-name "CAMPAIGN.TAG" ^
  --tracker-path "MasterCampaignTracker\MasterPropertyCampaignTracker.csv"
```

**Refresh template sequences & unique campaign numbers (global)**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --campaign-dir "Campaign_6_Aug2025" ^
  --rebuild-templates
```

---

## 6) Data flow & file details

**From lists to print**
1. `PropertyLists/*.csv` → **Builder** filters to `campaign_master.csv`.
2. `campaign_master.csv` → **Generator** outputs `letters_batch.pdf` + `letters_mapping.csv`.
3. `letters_mapping.csv` → **Finalizer** appends to `executed_campaign_log.csv`, updates Master tracker, rebuilds ZIP tally.

**Important CSV fields**
- `(PropertyAddress, OwnerName)` → the unique identity of a row across the system.  
- `ZIP5` → mailing ZIP; finalizer backfills from `campaign_master.csv` when needed.  
- `TemplateId` → numeric ID you pass when generating letters; stored per-send.  
- `CampaignNumbers` / `CampaignCount` → maintained on the tracker, derived from all executed logs.

---

## 7) Troubleshooting

- **Builder outputs 0 rows with `--prior-exact N`** → Your tracker may not reflect those priors yet. Run finalize for earlier campaigns, or use `--rebuild-all` to reconstruct the tracker from existing campaign folders.
- **Finalize says “mapping file not found”** → Make sure you’ve run the generator for that campaign and that `letters_mapping.csv` exists in the campaign folder (or `RefFiles\letters_mapping.csv`). If it’s elsewhere, pass `--mapping` with the full path.
- **Dates on Windows** → The finalizer uses a Windows-safe formatter (no `%-m/%-d`).
- **Marker file shows up** → Expected when using `--write-marker`. It’s intentionally empty; presence is the signal.
- **“src refspec main does not match any” on first push** → You didn’t commit locally yet. `git add .` → `git commit -m "Initial commit"` → push again.
- **Remote has a README and rejects your push** → `git pull --rebase origin main` then `git push`.

---

## 8) Housekeeping tips

- Archive legacy scripts to `OlderFiles\` to keep your root clean.
- Free up space safely by deleting heavy outputs only (keep logs!):
  ```bat
  rmdir /s /q "Campaign_1_Aug2025\BatchLetterFiles"
  rmdir /s /q "Campaign_1_Aug2025\Singles"
  ```
- Never delete a campaign folder’s `executed_campaign_log.csv` if you want to preserve history/rebuild capability.

---

## 9) Quick command cheat-sheet

**Never mailed (Campaign 1)**
```bat
python BuildMasterCampaignList_v4_MAILZIPFirst.py ^
  --campaign-name "Campaign" ^ --campaign-number 1 ^ --target-size 5000 ^ 
  --mandatory "PropertyLists\Foreclosure_08_2025.csv" "PropertyLists\PropertyTaxDelinquentList_08_2025.csv" ^
  --optional "PropertyLists\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 0 --strict-150 --debug
```

**Time-gapped (Exactly 4 priors & last sent ≥ 30 days)**
```bat
python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP.py ^
  --campaign-name "Campaign" ^ --campaign-number 7 ^ --target-size 5000 ^
  --mandatory "PropertyLists\Foreclosure_08_2025.csv" "PropertyLists\PropertyTaxDelinquentList_08_2025.csv" ^
  --optional "PropertyLists\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 4 --min-days-since-last 30 --strict-150 --debug
```

**Generate letters (combined only)**
```bat
python ..\direct_mail_batch_por_POR_KEEP_FIXINDENT.py ^
  --csv "campaign_master.csv" ^ --outdir "Singles" ^ --combine-out "letters_batch.pdf" ^
  --map-out "letters_mapping.csv" ^ --template-id 101 --skip-singles ^
  --sig-image "..\sig_ed.png" --name "Ed & Albert Beluli" --phone "916-905-7281" --email "eabeluli@gmail.com"
```

**Finalize (normal)**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --campaign-dir "Campaign_7_Aug2025" --write-marker
```

**Rebuild all (disaster recovery)**
```bat
python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^
  --rebuild-all ^ --root "C:\Users\Edit Beluli\Desktop\MailMonkey" ^
  --tracker-path "MasterCampaignTracker\MasterPropertyCampaignTracker.csv"
```

---

## 10) Glossary

- **Mapping** – CSV produced by the generator that lists every mail piece, the selected template ID, and optional per-row refs.  
- **Executed log** – Per-campaign CSV that the finalizer appends to (source of truth for history).  
- **Tracker** – Master roll-up (one row per `(PropertyAddress, OwnerName)`). Safe to delete; always recoverable.  
- **Marker** – Empty file (`CAMPAIGN.TAG` by default) used to positively identify a folder as an intentional campaign when rebuilding.
