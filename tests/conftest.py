
import os, shutil, subprocess, re
from pathlib import Path
import csv
import pytest

REQUIRED = [
    "BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP_NAMECASE.py",
    "direct_mail_batch_por_POR_KEEP_FIXINDENT.py",
    "FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py",
]

@pytest.fixture(scope="session")
def project_root() -> Path:
    return Path(__file__).resolve().parents[1]

@pytest.fixture()
def sandbox(project_root: Path):
    sb = project_root / "tests_sandbox"
    if sb.exists():
        shutil.rmtree(sb)
    sb.mkdir()

    (sb / "PropertyLists").mkdir(parents=True, exist_ok=True)

    for fname in REQUIRED:
        src = project_root / fname
        if not src.exists():
            pytest.skip(f"Required script missing: {fname} (place it in project root)")
        shutil.copy2(src, sb / fname)

    (sb / "sig_ed.png").write_bytes(b"\x89PNG\r\n\x1a\n")

    # Fixture rows with varying mailing ZIP columns
    rows = [
        {"Property Address":"1112 MCCLAREN DR","Primary Name":"O'NEIL JR LLC","MAILING ADDRESS":"1112 MCCLAREN DR, Granite Bay, CA 95746","MAIL ZIP":"95746","SITUS ADDRESS":"1112 MCCLAREN DR, Granite Bay, CA 99999","SITUS ZIP":"99999"},
        {"Property Address":"1200 MCCLAREN DR","Primary Name":"SMITH-JOHNSON LP","MAILING ADDRESS":"1200 MCCLAREN DR, Pasadena, CA 91117","MAIL ZIP":"91117","SITUS ADDRESS":"1200 MCCLAREN DR, Pasadena, CA 00000","SITUS ZIP":"00000"},
        {"Property Address":"42 MAIN ST","Primary Name":"MCDONALD INC","MAILING ADDRESS":"42 MAIN ST, Sacramento, CA 95835","MAIL ZIP":"95835","SITUS ADDRESS":"42 MAIN ST, Sacramento, CA 95834","SITUS ZIP":"95834"},
        {"Property Address":"7 OAK AVE","Primary Name":"MARY JANE","MAILING ADDRESS":"7 OAK AVE, Folsom, CA 95630","Owner ZIP":"95630","SITUS ADDRESS":"7 OAK AVE, Folsom, CA 95630","SITUS ZIP":"95630"},
        {"Property Address":"9 PINE RD","Primary Name":"ACME CORP","MAILING ADDRESS":"9 PINE RD, Davis, CA 95616","MAIL ZIP":"95616","SITUS ADDRESS":"9 PINE RD, Davis, CA 95618","SITUS ZIP":"95618"},
        {"Property Address":"55 RIVER DR","Primary Name":"DOE SR","MAILING ADDRESS":"55 RIVER DR, Elk Grove, CA 95757","Owner ZIP5":"95757","SITUS ADDRESS":"55 RIVER DR, Elk Grove, CA 95758","SITUS ZIP":"95758"},
    ]
    # Header = union of all keys (avoid DictWriter extras errors)
    all_fields = set()
    for r in rows:
        all_fields.update(r.keys())
    fieldnames = list(all_fields)

    for name in ("Foreclosure_08_2025.csv","PropertyTaxDelinquentList_08_2025.csv","LienList_ZipCodes_08_2025.csv"):
        path = sb / "PropertyLists" / name
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)

    yield sb

def run(cmd, cwd: Path):
    proc = subprocess.run(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, shell=True)
    return proc.returncode, proc.stdout

def read_csv(path: Path):
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))

# ----------------- Helpers used by tests -----------------
def _campaign_dir(base: Path, n: int) -> Path:
    import datetime as dt
    return base / f"Campaign_{n}_{dt.date.today().strftime('%b%Y')}"

def _extract_zip5(row: dict) -> str:
    # Prefer explicit owner/mail ZIP columns
    for k in ("MAIL ZIP","Mail ZIP","Owner ZIP5","Owner ZIP","ZIP5","ZIP"):
        if k in row and row[k]:
            z = str(row[k]).strip()
            z = re.sub(r"\.0$", "", z)
            m = re.search(r"(\d{5})", z)
            if m: return m.group(1)
    # Try parsing from a mailing address string
    for k in ("MAILING ADDRESS","Mailing Address","MAILING ZIP"):
        v = row.get(k, "") or ""
        m = re.search(r"(\d{5})(?:-\d{4})?$", v)
        if m: return m.group(1)
    return ""

def synthesize_mapping_from_master(camp_dir: Path, template_id: int):
    'Create letters_mapping.csv based on campaign_master.csv if the generator did not produce one.'
    master = camp_dir / "campaign_master.csv"
    if not master.exists():
        return False
    rows = read_csv(master)
    if not rows:
        return False

    # Best-effort columns
    def pick(row, choices):
        for c in choices:
            if c in row and row[c].strip():
                return row[c].strip()
        return ""

    out = []
    for r in rows:
        owner = pick(r, ["Primary Name","OwnerName","OWNER NAME","PRIMARY NAME"])
        addr  = pick(r, ["Property Address","PropertyAddress","ADDRESS","SITUS ADDRESS","Situs Address","Mailing Address"])
        zip5  = _extract_zip5(r)
        out.append({"OwnerName": owner, "PropertyAddress": addr, "ZIP5": zip5, "TemplateId": str(template_id)})

    # Write to root and RefFiles (alternate location)
    (camp_dir / "RefFiles").mkdir(exist_ok=True, parents=True)
    for target in ("letters_mapping.csv", os.path.join("RefFiles","letters_mapping.csv")):
        with (camp_dir / target).open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["OwnerName","PropertyAddress","ZIP5","TemplateId"])
            w.writeheader()
            w.writerows(out)
    return True
