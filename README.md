# T3 Compute Pipeline

A Python pipeline for Canadian ETF investors that reads CDS T3 tax PDFs,
computes per-account T3 slip totals, and generates ACB (Adjusted Cost Base)
helper entries.

Built for investors who:
- Hold Canadian ETFs that issue T3 slips (e.g. Vanguard, BMO, iShares)
- Track their own ACB in a spreadsheet
- Receive CDS Innovations T3 statements and want to automate the math

---

## Prerequisites

- **Python 3.9 or later** — [python.org/downloads](https://www.python.org/downloads/)
- The following Python packages (install once):

```
pip install pdfplumber openpyxl
```

> If your ACB spreadsheet is in the older `.xls` format (rare), also run:
> `pip install xlrd==1.2.0`

---

## Files in this folder

| File | Purpose |
|---|---|
| `run_t3.py` | Main entry point — interactive menu or CLI |
| `parse_t3_pdfs.py` | Step 1 — parses CDS T3 PDFs into JSON |
| `build_assets.py` | Step 2 — reads ACB spreadsheet → per-account share JSONs |
| `compute_t3.py` | Step 3 — computes final T3 slip totals |
| `config.template.json` | Template configuration — copy to `config.json` and edit |
| `account_periods.template.json` | Template account periods — copy to `account_periods.json` and edit |
| `funds.template.json` | Template fund metadata — copy to `funds.json` and edit |
| `acb_worksheet_template.xlsx` | Template ACB spreadsheet — copy and fill in your data |
| `README.md` | This file |
| `LICENSE` | MIT License |

---

## Setup (one time)

### 1. Clone or download the repository

```
git clone https://github.com/YOUR_USERNAME/t3_pipeline.git
cd t3_pipeline
```

Or download the ZIP from GitHub and extract it.

### 2. Install dependencies

```
pip install pdfplumber openpyxl
```

### 3. Create your config files from the templates

```
copy config.template.json config.json
copy account_periods.template.json account_periods.json
copy funds.template.json funds.json
```

> On macOS/Linux use `cp` instead of `copy`.

These three files contain your personal data and are excluded from git
via `.gitignore` — they will never be accidentally committed.

### 4. Copy and fill in your ACB spreadsheet

Copy `acb_worksheet_template.xlsx` and rename it (e.g. `acb_worksheet.xlsx`).
Read the **README sheet** inside the file carefully before entering data.

Key rules:
- One sheet per fund, named exactly after the ticker (e.g. `VBAL`, `ZCN`)
- Sheet names must match the fund symbols in `config.json`
- Column layout must match the template — do not add or remove columns
  unless you also update `col_indices` in `config.json`
- The script only reads **Col A** (date) and **Col G** (share balance)
- You are responsible for maintaining correct share balances and ACB values

### 5. Edit config.json

Open `config.json` and update these three fields:

```json
{
  "tax_year": "2025",
  "base_dir": "C:\\Users\\YourName\\Documents\\Tax Documents",
  "acb_spreadsheet": "C:\\path\\to\\your\\acb_worksheet.xlsx"
}
```

- `tax_year` — the tax year you are processing
- `base_dir` — folder where your T3 PDFs live and where output will be written
- `acb_spreadsheet` — full path to your ACB spreadsheet

The pipeline expects your CDS T3 PDFs to be named like `VBAL_T3_2025.pdf`
and placed in `<base_dir>\2025\`.

Output folders `distributions\` and `assets\` will be created automatically
inside `<base_dir>\2025\`.

### 6. Edit account_periods.json

This file tells the pipeline which brokerage account held each fund on each
distribution record date. Add one entry per fund per account period:

```json
{
  "2025": {
    "VBAL": [
      {"start": "2025-01-01", "account": "My Broker"},
      {"start": "2025-07-01", "account": "New Broker"}
    ],
    "ZCN": [
      {"start": "2025-01-01", "account": "My Broker"}
    ]
  }
}
```

- `start` — the date from which this account was active for this fund
- `account` — a label for the brokerage account (used in output filenames
  and T3 results — use whatever name is meaningful to you)
- If a fund was always in one account for the whole year, just one entry
  with `"start"` set to before the first distribution record date is enough
- Each year add a new year key (e.g. `"2026": { ... }`) — do not delete
  prior years

### 7. Edit funds.json (rarely needed)

Most funds use `null` (auto-detect from PDF). Only change this if the script
fails to parse a fund's PDF correctly:

```json
{
  "VBAL": {"calc_method_override": null},
  "ZCN":  {"calc_method_override": "PERCENT"}
}
```

Valid values: `null` (auto-detect), `"RATE"`, or `"PERCENT"`.
Check the CDS T3 PDF header — it will say either RATE or PERCENT.

---

## Running the pipeline

### Interactive mode (recommended)

```
python run_t3.py
```

A menu appears. Choose a step or run all three in sequence.
You will be prompted for any year or fund overrides.

### CLI mode

```
python run_t3.py --help                        Show all options
python run_t3.py --all                         Run all three steps
python run_t3.py --step 1                      Parse PDFs only
python run_t3.py --step 2                      Build assets only
python run_t3.py --step 3                      Compute T3 only
python run_t3.py --all --year 2026             Override tax year
python run_t3.py --step 1 --funds VBAL ZCN    Process specific funds only
```

---

## Pipeline steps explained

### Step 1 — Parse T3 PDFs

Download the CDS T3 PDFs for your funds from:
https://ctbsext.posttrade.cds.ca/ctbsExt/

**Input:**  `<base_dir>\<year>\<FUND>_T3_<year>.pdf` for each fund

**Output:**
- `<base_dir>\<year>\distributions\<FUND>.json` — per-unit distribution
  amounts for each T3 box and record date
- `<base_dir>\<year>\distributions\<FUND>_ACB_<year>.xlsx` — ROC rows
  ready to paste into your ACB spreadsheet

Run this once when CDS releases the T3 PDFs (usually February).

### Step 2 — Build Assets

**Input:**
- Distribution JSONs from Step 1
- Your ACB spreadsheet

**Output:** `<base_dir>\<year>\assets\<AccountName>_<year>.json` — one file
per brokerage account showing how many shares you held on each record date

Run this after your ACB spreadsheet is finalized for the year.

### Step 3 — Compute T3

**Input:**
- Distribution JSONs from Step 1
- Account JSONs from Step 2

**Output:** `<base_dir>\<year>\T3_results_<year>.txt` — T3 slip totals
broken down by account, with full per-distribution detail

Boxes computed: 21, 23, 25, 26, 32, 34, 39, 42, 49, 50, 51

---

## Annual checklist (each tax season)

1. Download the CDS T3 PDFs for your funds from:
   https://ctbsext.posttrade.cds.ca/ctbsExt/
   Save them in `<base_dir>\<year>\` named `<FUND>_T3_<year>.pdf`
   (e.g. `VBAL_T3_2026.pdf`)
2. Update `tax_year` in `config.json`
3. Add a new year block to `account_periods.json`
4. Verify gross-up rates in `config.json` — CRA occasionally changes these.
   Check [canada.ca](https://www.canada.ca/en/revenue-agency.html) if unsure.
   The script will warn you if the year is unrecognized.
5. Run the pipeline

---

## Understanding the output

### T3 box reference

| Box | Field | Notes |
|-----|-------|-------|
| 21 | Capital Gains | Taxable — 50% inclusion rate |
| 23 | Non-Eligible Dividends | Actual amount |
| 25 | Foreign Non-Business Income | Taxable as regular income |
| 26 | Other Income | Taxable as regular income |
| 32 | Taxable Non-Eligible Dividends | Box 23 × 1.15 |
| 34 | Foreign Tax Paid | Deductible / creditable |
| 39 | Non-Eligible Dividend Tax Credit | Box 32 × 0.090301 |
| 42 | Return of Capital | Not taxable — reduces your ACB |
| 49 | Eligible Dividends | Actual amount |
| 50 | Taxable Eligible Dividends | Box 49 × 1.38 |
| 51 | Eligible Dividend Tax Credit | Box 50 × 0.150198 |

### ACB Excel output (from Step 1)

The Excel file contains ROC rows to paste into your ACB spreadsheet:
- **Positive price/share** — regular ROC, reduces your ACB
- **Negative price/share** — phantom (non-cash) distribution, increases
  your ACB even though no cash was received

---

## Troubleshooting

**"No T3 JSON files found"**
Run Step 1 before Step 2.

**"Sheet 'VBAL' not found"**
The ACB spreadsheet does not have a sheet named `VBAL`. Sheet names must
match the fund tickers exactly.

**"Could not find share balance for VBAL on 2025-04-01"**
No row in the VBAL sheet has a date on or before 2025-04-01. Check that
your ACB spreadsheet has transactions entered before this record date.

**PDF parsing warning: "No distributions found"**
The PDF layout may differ from expected. Check that:
- The file is named exactly `<FUND>_T3_<year>.pdf`
- It is the CDS (Canadian Depository for Securities) T3 statement, not
  the broker's version
- Try setting `calc_method_override` in `funds.json` to `"RATE"` or
  `"PERCENT"` explicitly

**Gross-up rate warning**
The script detected an unrecognized tax year or a mismatch between your
config rates and the known values. Verify current rates at
[canada.ca](https://www.canada.ca/en/revenue-agency.html) and update
`tax_rates` in `config.json`.

---

## Contributing

Contributions are welcome. If you find a bug, have a feature request, or
want to add support for additional fund families:

1. [Open an issue](https://github.com/YOUR_USERNAME/t3_pipeline/issues)
   describing the problem or suggestion
2. Fork the repository
3. Create a branch: `git checkout -b fix/your-fix-name`
4. Make your changes and test them
5. Submit a pull request with a clear description of what changed and why

Please do not commit any personal financial data, real PDF files, or
populated config files.

---

## Disclaimer

This tool is provided for personal use and convenience only. It is not
financial or tax advice. Always cross-check the computed T3 totals against
the official T3 slips issued by your fund provider and consult a qualified
tax professional if in doubt. The authors accept no liability for errors
in tax filings made using this tool.

---

## License

MIT License — see [LICENSE](LICENSE) for full text.
