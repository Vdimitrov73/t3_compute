# Pipeline Steps Explained

## Step 1 — Parse T3 Source Files

Download the CDS T3 statements for your funds from:
https://ctbsext.posttrade.cds.ca/ctbsExt/

**For 2025 and later (PDF format):**
Save as `<FUND>_T3_<year>.pdf` in `<base_dir>\<year>\` (e.g. `VBAL_T3_2025.pdf`)

**For 2024 and earlier (XLS format):**
Save as `<FUND>_T3_<year>.xls` in `<base_dir>\<year>\` (e.g. `VBAL_T3_2024.xls`)

The correct parser is selected automatically based on `tax_year` in `config.json`.

**Output:**
- `<base_dir>\<year>\distributions\<FUND>.json` — per-unit distribution
  amounts for each T3 box and record date
- `<base_dir>\<year>\distributions\<FUND>_ACB_<year>.xlsx` — ROC rows
  as a standalone Excel file if you prefer to copy/paste manually

Run this once when CDS releases the T3 PDFs (usually February).

---

## Step 2 — Update ACB Spreadsheet

**Input:**
- Distribution JSONs from Step 1
- Your ACB spreadsheet

**What it does:**
- Creates a timestamped backup before touching the file
- Inserts ROC rows for the configured tax year in date order
- Skips rows that already exist (exact duplicate detection)
- Warns you if a ROC row exists on the same date and sign with a different
  amount, so you can review it manually before deciding what to do
- Rewrites all formula chains after each insertion
- Verifies formula integrity and rolls back automatically if anything fails
- Only touches rows within the configured `tax_year` — prior year data
  is never modified

Use `--dry-run` to preview exactly what would be inserted before committing:

```
python run_t3.py --step 2 --dry-run
```

**Note on phantom distributions:** Some funds issue both a regular ROC
(positive amount, reduces ACB) and a phantom non-cash ROC (negative amount,
increases ACB) on the same date (e.g. Dec-30). Step 2 handles both
automatically — they are treated as two distinct rows and both are inserted.

Run Step 2 after Step 1 and after you have entered all Buy/Sell transactions
for the year into the ACB spreadsheet.

---

## Step 3 — Build Assets

**Input:**
- Distribution JSONs from Step 1
- Your ACB spreadsheet (after Step 2 has been run)

**Output:** `<base_dir>\<year>\assets\<AccountName>_<year>.json` — one file
per brokerage account showing how many shares you held on each record date

Share balances are computed directly from your Buy/Sell transactions, so you
do not need to open the spreadsheet in Excel between Step 2 and Step 3. If
the computed balance disagrees with the cached value in Col G, a warning will
be shown — this usually means a transaction is missing or entered incorrectly.

---

## Step 4 — Compute T3

**Input:**
- Distribution JSONs from Step 1
- Account JSONs from Step 3

**Output:** `<base_dir>\<year>\T3_results_<year>.txt` — T3 slip totals
broken down by account, with full per-distribution detail

Boxes computed: 21, 23, 25, 26, 32, 34, 39, 42, 49, 50, 51
