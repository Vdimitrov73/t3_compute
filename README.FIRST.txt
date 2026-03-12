============================================================
  T3 Compute — Quick Start
============================================================

STEP 1 — Copy and rename the config files
------------------------------------------
In this folder you will find three template files. Copy each
one and remove the ".template" from the name:

  config.template.json          →  config.json
  account_periods.template.json →  account_periods.json
  funds.template.json           →  funds.json

STEP 2 — Edit config.json
--------------------------
Open config.json in Notepad and update these three fields:

  "tax_year"        — the year you are processing (e.g. "2025")
  "base_dir"        — the folder where your CDS T3 files are
                      (e.g. "C:\\Users\\You\\Documents\\Tax")
  "acb_spreadsheet" — full path to your ACB spreadsheet
                      (e.g. "C:\\Users\\You\\Documents\\acb_worksheet.xlsx")

Use double backslashes \\ in all file paths.

STEP 3 — Copy and fill in your ACB spreadsheet
-----------------------------------------------
Copy acb_worksheet_template.xlsx and rename it (e.g. acb_worksheet.xlsx).
Open it and read the README sheet inside before entering data.
One sheet per fund, named exactly after the ticker (e.g. VBAL, ZCN).

STEP 4 — Edit account_periods.json
------------------------------------
This tells the tool which brokerage account held each fund on
each distribution record date. See the template for the format.
If you only have one account, one entry per fund is enough.

STEP 5 — Download your CDS T3 statements
-----------------------------------------
Download from: https://ctbsext.posttrade.cds.ca/ctbsExt/

  2025 and later : save as  VBAL_T3_2025.pdf  in <base_dir>\2025\
  2024 and earlier: save as  VBAL_T3_2024.xls  in <base_dir>\2024\

STEP 6 — Run the tool
-----------------------
Double-click t3_compute.exe and follow the interactive menu.

============================================================
  Need help?
============================================================

Full documentation:
  https://github.com/Vdimitrov73/t3_compute

Report an issue:
  https://github.com/Vdimitrov73/t3_compute/issues

SmartScreen warning ("Windows protected your PC"):
  Click "More info" then "Run anyway". This appears because
  the exe is not signed with a paid EV code signing certificate
  (~$500 USD/year). The tool is open source — you can inspect
  every line of code on GitHub before running it.

============================================================
  Disclaimer
============================================================

This tool is for personal convenience only. It is not tax
advice. Always cross-check computed totals against your
official T3 slips and consult a tax professional if in doubt.

============================================================
