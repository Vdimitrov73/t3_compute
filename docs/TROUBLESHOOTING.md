# Troubleshooting

**"No T3 JSON files found"**
Run Step 1 before Steps 2, 3, or 4.

**"Sheet 'VBAL' not found"**
The ACB spreadsheet does not have a sheet named `VBAL`. Sheet names must
match the fund tickers exactly.

**"Could not find share balance for VBAL on 2025-04-01"**
No Buy or Sell row in the VBAL sheet has a date on or before 2025-04-01.
Make sure your ACB spreadsheet has transactions entered before this record date.

**Share balance mismatch warning in Step 3**
The balance computed from your Buy/Sell rows disagrees with the cached value
in Col G. This usually means a transaction is missing, entered on the wrong
date, or Col G was manually edited. Check your spreadsheet before relying on
the Step 4 results.

**Step 2 warning: "ROC row exists with different amount"**
A ROC row already exists on that date with the same sign but a different
per-unit amount. Step 2 skips it to be safe — review the
`<FUND>_ACB_<year>.xlsx` file from Step 1 and correct the row manually.

**Step 2 rolled back automatically**
A formula integrity check failed after insertion. Your spreadsheet has been
restored from the backup automatically. Open an issue on GitLab with the
error message and we will investigate.

**Step 1 warning: "No distributions found"**
The file layout may differ from expected. Check that:
- The file is named exactly `<FUND>_T3_<year>.pdf` (2025+) or `<FUND>_T3_<year>.xls` (2024 and earlier)
- It is the CDS Innovations T3 statement, not the broker's version
- Try setting `calc_method_override` in `funds.json` to `"RATE"` or `"PERCENT"` explicitly

**Gross-up rate warning**
The script detected an unrecognized tax year or a mismatch with known CRA
rates. Verify current rates at
[canada.ca](https://www.canada.ca/en/revenue-agency.html) and update
`tax_rates` in `config.json`.

**Results differ from my broker T3 slip by $0.01**
Check `broker_rounding` in `config.json`. For 2025+ PDF data, `true`
(default) gives the closest match to broker slips. For 2024 and earlier
XLS data, `false` (accumulate mode) may produce better results.
