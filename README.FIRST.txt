============================================================
  T3 Compute — Quick Start
============================================================

STEP 1 — Run the tool
-----------------------
Double-click t3_compute.exe.

On first launch, a setup wizard will guide you through
creating your config files. It takes about two minutes.

STEP 2 — Complete the setup wizard
------------------------------------
The wizard will ask you for:

  1. Tax year (e.g. 2025)
  2. Folder where your CDS T3 statements are saved
  3. Path to your ACB spreadsheet
  4. Fund tickers you hold (e.g. VBAL ZCN CPD VRE)
  5. Which brokerage account held each fund

If you don't have an ACB spreadsheet yet, the wizard will
copy the template (acb_worksheet_template.xlsx) for you.
Open it and read the README sheet before entering data.

STEP 3 — Download your CDS T3 statements
-----------------------------------------
Download from: https://ctbsext.posttrade.cds.ca/ctbsExt/

  2025 and later : save as  VBAL_T3_2025.pdf  in <base_dir>\2025\
  2024 and earlier: save as  VBAL_T3_2024.xls  in <base_dir>\2024\

STEP 4 — Enter your Buy/Sell transactions
------------------------------------------
Open your ACB spreadsheet and make sure all Buy/Sell
transactions for the tax year are entered before running
the pipeline.

STEP 5 — Run the pipeline
--------------------------
From the main menu, run each step in order:

  Step 1 — Parse T3 source files
  Step 2 — Update ACB spreadsheet
  Step 3 — Build assets
  Step 4 — Compute T3 totals

Or choose "Run all steps" to run them in sequence.

Your T3 results will be saved to:
  <base_dir>\<year>\T3_results_<year>.txt

============================================================
  Need help?
============================================================

Full documentation:
  https://gitlab.com/vdimitrov_73/t3_compute

Report an issue:
  https://gitlab.com/vdimitrov_73/t3_compute/-/issues

SmartScreen warning ("Windows protected your PC"):
  Click "More info" then "Run anyway". This appears because
  the exe is not signed with a paid EV code signing certificate
  (~$500 USD/year). To avoid this entirely, install from the
  Microsoft Store instead (Canadian Store only — search T3 Compute).
  The tool is open source — you can inspect every line of code
  on GitLab before running it.

============================================================
  Disclaimer
============================================================

This tool is for personal convenience only. It is not tax
advice. Always cross-check computed totals against your
official T3 slips and consult a tax professional if in doubt.

============================================================
