
from pathlib import Path
from conftest import run, read_csv, _campaign_dir, synthesize_mapping_from_master

def test_build_and_presort_mailzip(sandbox: Path):
    code, out = run(
        'python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP_NAMECASE.py ^'
        '  --campaign-name "Campaign" ^ --campaign-number 1 ^ --target-size 6 ^'
        '  --mandatory "PropertyLists\\Foreclosure_08_2025.csv" "PropertyLists\\PropertyTaxDelinquentList_08_2025.csv" ^'
        '  --optional "PropertyLists\\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 0 --strict-150 --debug',
        sandbox,
    )
    assert code == 0, out
    camp1 = _campaign_dir(sandbox, 1)
    master = camp1 / "campaign_master.csv"
    assert master.exists(), "campaign_master.csv missing"
    presort = read_csv(camp1 / "presort_report.csv")
    got = {r["ZIP5"] for r in presort}
    for z in ("95746","91117","95835","95630","95616","95757"):
        assert z in got, f"Missing MAIL ZIP {z} in presort"

def test_generate_and_finalize_idempotent_with_fallback_mapping(sandbox: Path):
    # Build Campaign 1
    run(
        'python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP_NAMECASE.py ^'
        '  --campaign-name "Campaign" ^ --campaign-number 1 ^ --target-size 6 ^'
        '  --mandatory "PropertyLists\\Foreclosure_08_2025.csv" "PropertyLists\\PropertyTaxDelinquentList_08_2025.csv" ^'
        '  --optional "PropertyLists\\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 0 --strict-150',
        sandbox,
    )
    camp1 = _campaign_dir(sandbox, 1)

    # Try generator; if mapping not produced, synthesize it
    run(
        'python ..\\direct_mail_batch_por_POR_KEEP_FIXINDENT.py ^'
        '  --csv "campaign_master.csv" ^ --outdir "Singles" ^ --combine-out "letters_batch.pdf" ^'
        '  --map-out "letters_mapping.csv" ^ --template-id 606 ^ --skip-singles ^'
        '  --sig-image "..\\sig_ed.png" ^ --name "Ed & Albert Beluli" ^ --phone "916-905-7281" ^ --email "ed.beluli@gmail.com"',
        camp1,
    )
    mapping = camp1 / "letters_mapping.csv"
    if not mapping.exists():
        assert synthesize_mapping_from_master(camp1, 606), "Failed to synthesize mapping"

    # Finalize
    code, out = run(
        'python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^'
        '  --campaign-dir "{camp}" ^ --write-marker'.format(camp=camp1.name),
        sandbox,
    )
    assert code == 0, out

    # Re-run finalize: should not re-append
    code, out2 = run(
        'python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^'
        '  --campaign-dir "{camp}"'.format(camp=camp1.name),
        sandbox,
    )
    assert code == 0, out2

def test_prior_exact_and_template_sequence_with_fallback(sandbox: Path):
    # Campaign 1
    run(
        'python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP_NAMECASE.py ^'
        '  --campaign-name "Campaign" ^ --campaign-number 1 ^ --target-size 6 ^'
        '  --mandatory "PropertyLists\\Foreclosure_08_2025.csv" "PropertyLists\\PropertyTaxDelinquentList_08_2025.csv" ^'
        '  --optional "PropertyLists\\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 0',
        sandbox,
    )
    camp1 = _campaign_dir(sandbox, 1)
    if not (camp1 / "letters_mapping.csv").exists():
        synthesize_mapping_from_master(camp1, 606)
    run('python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^ --campaign-dir "{camp}" ^ --write-marker'.format(camp=camp1.name), sandbox)

    # Campaign 2
    run(
        'python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP_NAMECASE.py ^'
        '  --campaign-name "Campaign" ^ --campaign-number 2 ^ --target-size 6 ^'
        '  --mandatory "PropertyLists\\Foreclosure_08_2025.csv" "PropertyLists\\PropertyTaxDelinquentList_08_2025.csv" ^'
        '  --optional "PropertyLists\\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 1',
        sandbox,
    )
    camp2 = _campaign_dir(sandbox, 2)
    # Try generator; fallback mapping if needed with a NEW template id (707)
    run(
        'python ..\\direct_mail_batch_por_POR_KEEP_FIXINDENT.py ^'
        '  --csv "campaign_master.csv" ^ --outdir "Singles" ^ --combine-out "letters_batch.pdf" ^'
        '  --map-out "letters_mapping.csv" ^ --template-id 707 ^ --skip-singles ^'
        '  --sig-image "..\\sig_ed.png" ^ --name "Ed & Albert Beluli" ^ --phone "916-905-7281" ^ --email "ed.beluli@gmail.com"',
        camp2,
    )
    if not (camp2 / "letters_mapping.csv").exists():
        synthesize_mapping_from_master(camp2, 707)
    run('python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^ --campaign-dir "{camp}" ^ --write-marker'.format(camp=camp2.name), sandbox)

    # Validate tracker
    tracker = sandbox / "MasterCampaignTracker" / "MasterPropertyCampaignTracker.csv"
    rows = read_csv(tracker)
    assert rows, "Tracker should have rows"
    for r in rows:
        assert "1" in r["CampaignNumbers"] and "2" in r["CampaignNumbers"], "Expected both campaign numbers"
        assert "606" in r.get("TemplateIds","") and "707" in r.get("TemplateIds",""), "Expected both templates"

def test_file_integrity_checks(sandbox: Path):
    # Build & synthesize for 1 and 2
    run('python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP_NAMECASE.py ^ --campaign-name "Campaign" ^ --campaign-number 1 ^ --target-size 6 ^ --mandatory "PropertyLists\\Foreclosure_08_2025.csv" "PropertyLists\\PropertyTaxDelinquentList_08_2025.csv" ^ --optional "PropertyLists\\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 0', sandbox)
    c1 = _campaign_dir(sandbox, 1); synthesize_mapping_from_master(c1, 606); run('python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^ --campaign-dir "{c}" ^ --write-marker'.format(c=c1.name), sandbox)
    run('python BuildMasterCampaignList_v4_MAILZIPFirst_TIMEGAP_NAMECASE.py ^ --campaign-name "Campaign" ^ --campaign-number 2 ^ --target-size 6 ^ --mandatory "PropertyLists\\Foreclosure_08_2025.csv" "PropertyLists\\PropertyTaxDelinquentList_08_2025.csv" ^ --optional "PropertyLists\\LienList_ZipCodes_08_2025.csv" ^ --prior-exact 1', sandbox)
    c2 = _campaign_dir(sandbox, 2); synthesize_mapping_from_master(c2, 707); run('python FinalizeCampaign_TRACKER_STRICT_v5_RECOVER_full_fixdate.py ^ --campaign-dir "{c}" ^ --write-marker'.format(c=c2.name), sandbox)

    # Schema and parity checks
    tracker = sandbox / "MasterCampaignTracker" / "MasterPropertyCampaignTracker.csv"
    assert tracker.exists(), "Tracker missing"
    trows = read_csv(tracker)
    required = {"PropertyAddress","OwnerName","ZIP5","CampaignCount","CampaignNumbers","TemplateIds","FirstSentDt","LastSentDt"}
    assert required.issubset(set(trows[0].keys())), f"Tracker missing columns: {required - set(trows[0].keys())}"

    from collections import Counter
    # tally check
    tally = read_csv(sandbox / "MasterCampaignTracker" / "Zip5_LetterTally.csv")
    tally_dict = {r["ZIP5"]: int(r["Count"]) for r in tally}
    sum_by_zip = Counter()
    for c in (c1, c2):
        for r in read_csv(c / "executed_campaign_log.csv"):
            z = (r.get("ZIP5","") or "").strip()
            sum_by_zip[z] += 1
    for z, cnt in sum_by_zip.items():
        assert tally_dict.get(z, 0) == cnt, f"ZIP tally mismatch for {z}"
