# Changelog

All notable changes to T3 Compute are listed here, newest first.

---
## v1.6.1 — 2026-04-06
- Fix: frozen .exe no longer spawns a background console window on launch
  (PyInstaller --windowed flag added to build)

## v1.6.0 — 2026-04-06
- Added: Windows GUI (t3_gui.py) — Tkinter shell replacing the
  interactive CLI menu; all pipeline logic unchanged
- Added: SetupWizard GUI dialog — 4-page visual replacement for
  the CLI first-run wizard; doubles as an edit dialog for existing config
- Added: run_t3.py --gui flag to launch GUI from source install
- Changed: frozen .exe now launches GUI by default instead of CLI menu
- Added: T3 HTML export — generates a CRA-layout T3 slip per brokerage
  account, styled with the official 5-column grid, bilingual labels,
  and print-ready page breaks
- Added: generate_t3_html() in run_t3.py — recomputes T3 from source
  JSONs at export time; never reads stale results
- Added: --export CLI flag — generate T3 HTML after --step 4 or --all
- Added: --funds filter for export — e.g. --export --funds VBAL generates
  T3_2025_VBAL.html matching the exact slip TD sent for that fund only
- Added: GUI Export T3 HTML button — respects year and funds override
  fields; opens generated file in browser automatically
- Added: Session guard in interactive menu — E (Export) only appears
  after Step 4 completes successfully in the current session; stale
  results files from prior runs do not unlock it
- Added: Session state captures exact year and funds used in Step 4
  so E exports the identical selection without re-prompting

## v1.5.1 — 2026-03-31
- Removed: `broker_rounding` config option and all LRM code — accumulate
  mode with fund-boundary rounding is now the only mode. Results match
  broker T3 slips across all tested brokers (TD DI, Wealthsimple, Disnat)
  for both 2024 and 2025 data. No config changes required — the key is
  simply ignored if present in older configs.
- Fix: cross-fund float accumulation corrected — each fund's subtotal is
  now rounded to 2dp before being added to the account total, eliminating
  a $0.01 discrepancy when the same box appears in multiple funds
  
## v1.5.0 — 2026-03-28
- Added: ANSI colour output for all interactive menus
- Fix: PERCENTAGE-method ETFs (ZCN, etc.) dominant component rounded to 5dp
- Fix: genuine_constraint guard prevents LRM from firing on PERCENTAGE funds
  whose components already sum to the cash total, eliminating penny adjustments
- Fix: SAFE_LRM_THRESHOLD (0.05) — LRM skipped when no component is near the
  rounding midpoint, preventing false corrections on clean distributions
- Per-unit precision in breakdown lines: 4dp -> 5dp

## v1.4.9 — 2026-03-26
- Change: `broker_rounding` default set to `true` (largest_remainder method)
  for all new installations and for users without a `broker_rounding` key in
  their config. This matches the observed behaviour of Canadian brokers for
  2025 PDF-based CDS data.
- Note: users processing 2024 or earlier (XLS-based CDS data) should set
  `"broker_rounding": false` in config.json for closer accumulate-mode results.

## v1.4.8 — 2026-03-22
- Fix: year and fund overrides in interactive mode no longer carry over
  into subsequent menu iterations (pressing Enter now correctly resets
  to config default)
- Fix: Box 50/51 and Box 32/39 now computed from rounded intermediate
  values at each step, eliminating a $0.01 cascade error on derived boxes
- Fix: Step 4 (Compute T3) now supports --funds override, matching
  steps 1–3 behaviour
- Fix: interactive menu option labels aligned for consistency

## v1.4.7 — 2026-03-18
 - Fix: use FOLDERID_Documents to bypass MSIX filesystem virtualization

## v1.4.6 — 2026-03-13
- Fix: GitLab CI YAML syntax error in release job

## v1.4.5 — 2026-03-13
- Fix: GitLab CI release job failing to attach artifacts

## v1.4.4 — 2026-03-13
- GitLab set as primary repository; GitHub remains as mirror
- Added GitLab CI workflow (equivalent to GitHub Actions build)

## v1.4.3 — 2026-03-12
- Fix: ACB spreadsheet default path in setup wizard now points to
  base_dir (e.g. Documents\Tax Documents) instead of AppData

## v1.4.2 — 2026-03-12
- Fix: config files now stored in %LOCALAPPDATA%\T3Compute\ for
  .exe and MSIX users so they persist across relaunches
- Python and setup.bat users are unaffected — no behaviour change

## v1.4.1 — 2026-03-12
- Fix: "Done" message in update_acb.py now always prints before exit
- Fix: clean_dist() no longer overwrites nonCashCapitalGains from XLS path
- Fix: standardized output_txt filename across all steps
- Fix: transfer date in setup wizard now validated with date.fromisoformat()
- Fix: menu box alignment in interactive mode

## v1.4.0 — 2026-03-12
- Added: first-run setup wizard — guides users through creating all
  config files on first launch; no manual JSON editing required
- Wizard covers tax year, base folder, ACB path, fund tickers,
  and brokerage account periods including mid-year transfers

## v1.3.6 — 2026-03-12
- Fix: Option B download link corrected to point to source ZIP

## v1.3.5 — 2026-03-12
- Fix: Option B instructions updated to reference source code ZIP

## v1.3.4 — 2026-03-12
- setup.bat added as explicit release asset in GitHub Actions workflow

## v1.3.3 — 2026-03-12
- README improvements

## v1.3.2 — 2026-03-12
- Release ZIP now bundles exe with config templates, ACB template,
  README.FIRST.txt, README.md, and BUILD_EXE.md

## v1.3.1 — 2026-03-12
- setup.bat added as downloadable asset in CI release workflow

## v1.3.0 — 2026-03-12
- Added: GitHub Actions CI/CD workflow — builds and releases .exe
  automatically on version tags
- Added: setup.bat — double-click to install Python dependencies
- Added: BUILD_EXE.md — explains SmartScreen warning and how to
  build the .exe yourself from source

## v1.2.5 — 2026-03-11
- Added: XLS format parsing for CDS statements from 2024 and earlier
  (CDS switched from XLS to PDF starting with the 2025 tax year)

## v1.2.2 — 2026-03-11
- Improved warning and error handling across all pipeline steps

## v1.2.0 — 2026-03-11
- Fix: partial phantom ROC handling corrected
- Fix: duplicate ROC row insertion now detected and skipped
- Fix: share balance calculation after Step 4 corrected
- Added: formulas and dropdowns to ACB worksheet template

## v1.0.0 — 2026-03-10
- Initial public release
- Supports CDS T3 PDF format (2025+)
- Computes T3 boxes 21, 23, 25, 26, 32, 34, 39, 42, 49, 50, 51
- Supports multiple brokerages and mid-year account transfers
