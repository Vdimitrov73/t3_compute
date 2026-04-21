# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Running the Pipeline
- **Interactive mode (recommended)**: `python run_t3.py` - launches menu-driven interface
- **CLI mode**: 
  - `python run_t3.py --all` - run all four pipeline steps
  - `python run_t3.py --step 1` - parse T3 PDFs only
  - `python run_t3.py --step 2` - update ACB spreadsheet (use `--dry-run` first to preview)
  - `python run_t3.py --step 3` - build assets from ACB spreadsheet
  - `python run_t3.py --step 4` - compute T3 slip totals
  - `python run_t3.py --step 4 --export` - compute T3 and export HTML slips
- **GUI mode**: Double-click `t3_compute.exe` or run `python run_t3.py --gui`

### Configuration
- Configuration files are stored in:
  - Script directory (when running as Python script)
  - `Documents\\T3Compute\\` (when running as frozen .exe from Microsoft Store)
- Key config files:
  - `config.json` - tax year, paths, fund list, tax rates
  - `account_periods.json` - brokerage account periods per fund per year
  - `funds.json` - per-fund calculation method overrides
- Use `--config`, `--year`, and `--funds` flags to override settings

### Testing
- Run `python test_pipeline.py` to execute the full test suite
- Individual step testing: run each step with `--dry-run` flag where available
- Verify outputs in `<base_dir>/<year>/` subdirectories:
  - `distributions/` - parsed T3 data (JSON) and ACB helper files (Excel)
  - `assets/` - per-account share balance JSONs
  - Root year folder - T3 results text file and HTML exports

## Code Style
- Follow existing patterns — no major refactoring without discussion
- Keep `.gitlab-ci.yml` unchanged except `--windowed` flag for PyInstaller
- Tests must pass: `python test_pipeline.py` before commits

## Common Pitfalls
- Step 2 creates ACB backup but **never** modifies formulas — only data rows
- `--funds` filter applies to all steps when specified
- Frozen exe config is in `Documents\T3Compute\` — not current directory
- Configuration loading differs between script and frozen exe modes

## Testing Notes
- test_pipeline.py uses only in-memory fixtures — no real PDFs or spreadsheets needed
- Step 1 cannot be integration-tested without real CDS T3 files; mock the PDF parser
- xlrd is required for 2024 and earlier (XLS) tests but not available in all environments

## Build Notes
- AppxManifest.xml patching belongs ONLY in the MSIX packaging workflow
- Never modify .gitlab-ci.yml or build_exe.yml to patch AppxManifest
- PyInstaller entry point is run_t3.main_gui() for the GUI exe
- Config files route to %LOCALAPPDATA%\T3Compute\ only in frozen exe mode

## Precision Rules (do not change without running regression tests)
- Box 49 must be rounded to 2dp before deriving Box 50/51
- Box 23 must be rounded to 2dp before deriving Box 32/39
- Each fund's subtotals are rounded to 2dp before being added to account totals
- Per-unit distribution values use 7dp internally; never round these in compute_t3.py
- The XLS parser (2024 and earlier) uses 7dp rounding for PERCENT-method conversions

## Known Intentional Behaviour
- Per-box discrepancies of $0.01 vs broker T3 slips are CDS institutional precision artefacts — do not attempt to "fix" by changing rounding logic without regression data
- update_acb.py auto-confirms the prior-year prompt when called from the GUI (builtins.input is patched in t3_gui.py _run_worker)
- build_assets.py reads the ACB spreadsheet with data_only=False intentionally — formula strings are needed because newly inserted rows have uncached values

## Repository
- Primary: GitLab (vdimitrov_73/t3_compute)
- Mirror: GitHub (Vdimitrov73/t3_compute) via dual-push
- CI/CD builds .exe on version tags; never manually edit build_exe.yml to trigger releases

## Code Architecture

### Pipeline Design
The application follows a strict 4-step data pipeline:
1. **parse_t3_pdfs.py** - Extracts distribution data from CDS T3 PDFs/XLS → JSON + Excel helpers
2. **update_acb.py** - Inserts Return of Capital (ROC) rows into user's ACB spreadsheet
3. **build_assets.py** - Calculates share balances per account on each record date
4. **compute_t3.py** - Computes final T3 slip totals per account using distribution data and share balances

### Data Flow
```
CDS T3 PDF/XLS → [Step 1] → distributions/<FUND>.json
                                                              ↓
ACB spreadsheet → [Step 2] → (modified ACB with ROC rows)
                                                              ↓
[Step 3] → assets/<Account>_<year>.json
                                                              ↓
[Step 4] → T3_results_<year>.txt → (optional) → T3_<year>[_[FUND]].html
```

### Key Components
- **run_t3.py** - Main orchestrator with interactive menu, CLI args, and setup wizard
- **t3_colors.py** - Colorized terminal output utilities
- **t3_gui.py** - Graphical interface (launched via `--gui` or double-clicking .exe)
- Configuration loading - Handles both script and frozen executable paths appropriately
- Modular design - Each pipeline step can be run independently or in sequence

### File Organization
- Root directory: Main scripts, config templates, documentation
- `docs/` - Detailed documentation for each pipeline step and build process
- Generated directories (created at runtime):
  - `<base_dir>/<year>/distributions/` - Parsed T3 data
  - `<base_dir>/<year>/assets/` - Account share balances
  - `<base_dir>/<year>/` - Final T3 results and HTML exports

### Safety Features
- Automatic timestamped backups before modifying ACB spreadsheet
- Dry-run modes for previewing changes
- Duplicate detection to prevent double-processing
- Year-boundary enforcement (only modifies current tax year data)
- Formula integrity verification after spreadsheet modifications
- Automatic rollback on failure