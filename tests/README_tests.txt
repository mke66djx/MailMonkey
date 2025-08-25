MailCampaigns – Test Suite (v4)
=================================
What changed in v4
- Tests are self-contained: each test builds its own campaign(s).
- If the generator does not emit letters_mapping.csv, tests will synthesize a minimal mapping from campaign_master.csv so finalize/rebuild paths are still exercised.
- Keeps the v3 header-union CSV fixture fix.

Install & run
1) Delete any existing `tests` folder under your MailCampaigns root.
2) Unzip this so you have `...\MailCampaigns\tests\...`
3) Run:  tests\run_tests.bat
