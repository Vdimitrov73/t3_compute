# T3 Compute

A free tool for Canadian ETF investors that reads CDS Innovations T3 statements,
computes per-account T3 slip totals, and inserts ROC entries directly into your
ACB spreadsheet.

Built for investors who:
- Hold Canadian ETFs (Vanguard, BMO, iShares, etc.) in non-registered accounts
- Track their own ACB in a spreadsheet
- Want to automate the T3 math from CDS Innovations source files

---

## Quick Start (Windows)

**Option A — Microsoft Store (recommended):** no SmartScreen warning
**Option B — ZIP bundle:** [download latest release](https://gitlab.com/vdimitrov_73/t3_compute/-/releases/permalink/latest), unzip anywhere

1. Run **T3 Compute** from the Start Menu (Store) or double-click `t3_compute.exe` (ZIP)
2. Follow the first-run setup wizard — it creates all config files automatically
3. Download your CDS T3 statements and place them in the folder the wizard set up
4. Run the 4-step pipeline from the menu

That's it. See [INSTALL.md](INSTALL.md) if you need more detail.

---

## How It Works

```
Your CDS T3 PDF or XLS file
           │
           ▼
   Step 1 — Parse ──────────────► distributions/<FUND>.json
                                   distributions/<FUND>_ACB_<year>.xlsx
           │
           ▼
   Step 2 — Update ACB ─────────► acb_worksheet.xlsx
             (ROC rows inserted)
           │
           ▼
   Step 3 — Build Assets ───────► assets/<Account>_<year>.json
           │
           ▼
   Step 4 — Compute T3 ─────────► T3_results_<year>.txt
           │
           ▼
   Export — T3 HTML ───────────► T3_<year>[_<FUND>].html
(one CRA-layout slip per brokerage account)
```

---

## Download

**[🪟 Microsoft Store](https://apps.microsoft.com/search?query=T3+Compute&hl=en-US&gl=CA)** — recommended, no SmartScreen warning, Canadian Store only

**[⬇ Latest ZIP release](https://gitlab.com/vdimitrov_73/t3_compute/-/releases/permalink/latest)** — standalone `.exe`, no Python required

**Source:** [GitLab](https://gitlab.com/vdimitrov_73/t3_compute) (primary) | [GitHub](https://github.com/Vdimitrov73/t3_compute) (mirror)

---

## Features

- Parses CDS T3 PDFs (2025+) and XLS files (2024 and earlier)
- Computes all T3 boxes: 21, 23, 25, 26, 32, 34, 39, 42, 49, 50, 51
- Inserts ROC and phantom distribution rows directly into your ACB spreadsheet
- Exports a print-ready CRA-layout T3 HTML slip per brokerage account
- Filter export by fund — generate one slip matching exactly what each broker sent
- Supports multiple brokerages and mid-year account transfers
- First-run setup wizard — no manual JSON editing required
- Fully local — no internet connection, no data collection

---

## Basic Usage

### Interactive menu (recommended)

```
t3_compute.exe
```

or

```
python run_t3.py
```

The first launch runs the setup wizard. After that it goes straight to the menu.

### CLI

```
python run_t3.py --all                           Run all four steps
python run_t3.py --step 2 --dry-run              Preview ACB insertions
python run_t3.py --all --year 2024               Override tax year
python run_t3.py --step 1 --funds VBAL ZCN       Process specific funds only
python run_t3.py --step 4 --export               Compute T3 and export HTML (all funds)
python run_t3.py --step 4 --export --funds VBAL  Export VBAL only → T3_2025_VBAL.html
python run_t3.py --all --export --funds VBAL ZCN Run all steps + export VBAL and ZCN
python run_t3.py --help                          Show all options
```

---

## File Reference

| File | Purpose |
|---|---|
| `run_t3.py` | Main entry point — interactive menu or CLI |
| `parse_t3_pdfs.py` | Step 1 — parses CDS T3 PDFs (2025+) or XLS (2024 and earlier) |
| `update_acb.py` | Step 2 — inserts ROC rows into your ACB spreadsheet |
| `build_assets.py` | Step 3 — reads ACB spreadsheet → per-account share JSONs |
| `compute_t3.py` | Step 4 — computes final T3 slip totals |
| `setup.bat` | Windows: installs Python dependencies automatically |
| `config.template.json` | Copy to `config.json` and edit |
| `account_periods.template.json` | Copy to `account_periods.json` and edit |
| `funds.template.json` | Copy to `funds.json` and edit |
| `acb_worksheet_template.xlsx` | Copy and fill in your Buy/Sell transactions |

---

## Annual Checklist

1. Download CDS T3 statements from https://ctbsext.posttrade.cds.ca/ctbsExt/
2. Update `tax_year` in `config.json`
3. Add a new year block to `account_periods.json`
4. Enter all Buy/Sell transactions for the year into the ACB spreadsheet
5. Run Step 1 → Step 2 → Step 3 → Step 4
6. Export T3 HTML — one CRA-layout slip per brokerage account, ready to print

---

## T3 Box Reference

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

---

## Documentation

- [INSTALL.md](INSTALL.md) — installation options and first-time setup
- [README.FIRST.txt](README.FIRST.txt) — quick start guide included in the ZIP bundle
- [docs/PIPELINE.md](docs/PIPELINE.md) — detailed pipeline step documentation
- [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) — common errors and fixes
- [docs/BUILD_EXE.md](docs/BUILD_EXE.md) — how to build the `.exe` yourself

---

## Contributing

1. [Open an issue](https://gitlab.com/vdimitrov_73/t3_compute/-/issues) describing the problem or suggestion
2. Fork the repository on [GitLab](https://gitlab.com/vdimitrov_73/t3_compute)
3. Create a branch: `git checkout -b fix/your-fix-name`
4. Make your changes and test them
5. Submit a merge request with a clear description of what changed and why

Please do not commit personal financial data, real PDF files, or populated config files.

---

## Disclaimer

This tool is for personal convenience only. It is not tax advice. Always
cross-check computed totals against your official T3 slips and consult a
tax professional if in doubt.

---

## License

MIT License — see [LICENSE](LICENSE) for full text.
