# Installation

Choose the option that works best for you.

---

## Option A — Microsoft Store (recommended)

**[🪟 Get it from the Microsoft Store](https://apps.microsoft.com/search?query=T3+Compute&hl=en-US&gl=CA)**

The easiest option. Installs cleanly on Windows 10/11 with no SmartScreen
warning, no Python required, and updates automatically.

> Available in the Canadian Microsoft Store only. If you don't see it,
> check **Settings → Time & Language → Region** is set to Canada.

---

## Option B — ZIP bundle (no Python required)

**[⬇ Download latest release](https://gitlab.com/vdimitrov_73/t3_compute/-/releases/permalink/latest)**

1. Download and unzip anywhere (e.g. `C:\T3Compute\`)
2. Double-click `t3_compute.exe`
3. Follow the first-run setup wizard

> **Windows SmartScreen warning?** Click **"More info"** then **"Run anyway"**.
> This warning appears because the executable is not signed with a paid EV
> certificate (~$500 USD/year). See [docs/BUILD_EXE.md](docs/BUILD_EXE.md)
> for details, or install from the Microsoft Store to avoid it entirely.

---

## Option C — Python + setup.bat

1. Install **Python 3.9 or later** from [python.org/downloads](https://www.python.org/downloads/)
   - Check **"Add Python to PATH"** on the first screen before clicking Install
2. Download and extract the **[source ZIP](https://gitlab.com/vdimitrov_73/t3_compute/-/archive/main/t3_compute-main.zip)**
3. Double-click **`setup.bat`** — it installs all required packages automatically
4. Run `python run_t3.py`

---

## Option D — Manual Python setup (developers)

```
git clone https://gitlab.com/vdimitrov_73/t3_compute.git
cd t3_compute
pip install pdfplumber openpyxl xlrd
python run_t3.py
```

---

## First-Time Setup Wizard

On first launch, the wizard walks you through creating your config files:

**Step 1 — Tax year**
Enter the year you are processing (e.g. `2025`).

**Step 2 — Base folder**
The folder where your CDS T3 statements live and where output will be written.
The pipeline expects statements in `<base_dir>\<year>\`.

**Step 3 — ACB spreadsheet path**
Full path to your ACB tracking spreadsheet. If you don't have one yet,
the wizard copies `acb_worksheet_template.xlsx` as a starting point.

**Step 4 — Fund tickers**
Space-separated list of fund tickers you hold (e.g. `VBAL ZCN CPD VRE`).

**Step 5 — Brokerage accounts**
Which brokerage account held each fund, and whether any funds transferred
between accounts during the year.

After setup, the wizard writes three config files:
- `config.json` — tax year, paths, fund list
- `account_periods.json` — brokerage account periods per fund
- `funds.json` — per-fund PDF parsing hints

These files are excluded from git via `.gitignore` and will never be
accidentally committed. You can edit them manually at any time.

> ⚠️ **Before running Step 2 for the first time**, use `--dry-run`
> to preview what ROC rows will be inserted into your ACB spreadsheet
> without making any changes:
> ```
> python run_t3.py --step 2 --dry-run
> ```
> Step 2 also creates a timestamped backup automatically, but it's
> good practice to preview first.

---

## ACB Spreadsheet Setup

Copy `acb_worksheet_template.xlsx` and rename it (e.g. `acb_worksheet.xlsx`).
Read the **README sheet** inside the file before entering data.

Key rules:
- One sheet per fund, named exactly after the ticker (e.g. `VBAL`, `ZCN`)
- Sheet names must match the fund symbols in `config.json`
- Column layout must match the template — do not add or remove columns
  unless you also update `col_indices` in `config.json`
- Enter all Buy/Sell transactions before running the pipeline
- Do not sort or insert ROC rows manually — use Step 2 to do this safely

---

## CDS T3 Statements

Download from: https://ctbsext.posttrade.cds.ca/ctbsExt/

- **2025 and later:** save as `<FUND>_T3_<year>.pdf` in `<base_dir>\<year>\`
  (e.g. `VBAL_T3_2025.pdf`)
- **2024 and earlier:** save as `<FUND>_T3_<year>.xls` in `<base_dir>\<year>\`
  (e.g. `VBAL_T3_2024.xls`)

The correct parser is selected automatically based on `tax_year` in `config.json`.
