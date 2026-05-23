"""
run_t3.py
---------
Wrapper / menu for the T3 tax pipeline.

Interactive mode (no arguments):
    python run_t3.py

CLI mode:
    python run_t3.py --step 1               # parse PDFs only
    python run_t3.py --step 2               # update ACB spreadsheet
    python run_t3.py --step 2 --dry-run     # preview ACB insertions only
    python run_t3.py --step 3               # build assets only
    python run_t3.py --step 4               # compute T3 only
    python run_t3.py --all                  # run all four steps in sequence
    python run_t3.py --all --year 2026      # override tax year
    python run_t3.py --step 1 --funds VBAL  # override fund list
    python run_t3.py --gui                  # run the pipeline in GUI mode
    python run_t3.py --export               # export T3 HTML forms
    python run_t3.py --help                 # show this help

Pipeline steps:
    Step 1 — parse_t3_pdfs.py
             Reads CDS T3 PDFs and outputs one distribution JSON per fund
             plus an Excel ACB helper file.

    Step 2 — update_acb.py
             Reads distribution JSONs from Step 1 and inserts ROC rows
             directly into your ACB spreadsheet in date order.
             Creates a timestamped backup before making any changes.
             Use --dry-run to preview without modifying the file.

    Step 3 — build_assets.py
             Reads the ACB spreadsheet and the distribution JSONs, then
             outputs one JSON per brokerage account showing share balances
             on each record date.

    Step 4 — compute_t3.py
             Reads the distribution JSONs and account JSONs, computes
             T3 slip totals per account, and saves a results text file.

    Step 5 - Export your T3 slips (optional)
             After Step 4 completes, use "Export T3 HTML". This generates
             a print-ready CRA-layout T3 slip for each brokerage account.

Config files:
    Python script (python run_t3.py / setup.bat):
        Same directory as this script.
    Bundled .exe (PyInstaller / Microsoft Store):
        Documents\\T3Compute\\
        e.g. C:\\Users\\<you>\\Documents\\T3Compute\\

    config.json           — tax year, paths, fund list, tax rates
    account_periods.json  — brokerage account periods per fund per year
    funds.json            — per-fund PDF parsing hints
"""

import argparse
import sys
import os
import json
import shutil
from datetime import date
from pathlib import Path

try:
    from t3_colors import use_color, c, BOLD, DIM, CYAN, GREEN, YELLOW, RED
except ImportError:
    def use_color(): return False
    def c(code, text): return text
    BOLD = DIM = CYAN = GREEN = YELLOW = RED = ""

# ── Config directory ─────────────────────────────────────────────────────────

def get_config_dir(script_dir):
    """
    Return the directory where config files should be read/written.
    MSIX/Store (frozen exe): Documents\\T3Compute\\
    Plain Python script    : same directory as the script (script_dir)
    """
    if getattr(sys, "frozen", False):
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", wintypes.BYTE * 8),
            ]

        fid = GUID(
            0xFDD39AD0, 0x238F, 0x46AF,
            (0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7)
        )
        path_ptr = wintypes.LPWSTR()
        ret = ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(fid), 0, None, ctypes.byref(path_ptr)
        )
        if ret == 0 and path_ptr.value:
            base_path = Path(path_ptr.value)
            ctypes.windll.ole32.CoTaskMemFree(path_ptr)
        else:
            base_path = Path.home() / "Documents"

        config_dir = base_path / "T3Compute"
        config_dir.mkdir(parents=True, exist_ok=True)
        return str(config_dir)

    return script_dir

# ── First-run setup wizard ────────────────────────────────────────────────────

# Known CRA tax rates — stable since 2019
_KNOWN_TAX_RATES = {
    "eligible_div_grossup":        1.38,
    "eligible_div_tax_credit":     0.150198,
    "non_eligible_div_grossup":    1.15,
    "non_eligible_div_tax_credit": 0.090301,
}

def _prompt(msg, default=None):
    """Prompt the user, showing default in brackets. Returns stripped input."""
    suffix = f" [{default}]" if default else ""
    val = input(f"  {msg}{suffix}: ").strip()
    return val if val else default

def _prompt_yn(msg, default="N"):
    val = input(f"  {msg} [Y/N] (default {default}): ").strip().upper()
    return (val if val in ("Y", "N") else default) == "Y"

def _separator():
    print()
    print("  " + "─" * 56)
    print()

def run_first_time_setup(script_dir, config_dir, config_path, periods_path, funds_path):
    """
    Interactive wizard that creates config.json, account_periods.json,
    and funds.json for a new user. Only called when at least one of these
    files is missing.
    """
    def row(content):
        return c(BOLD + YELLOW, "║") + c(BOLD + CYAN, content) + c(BOLD + YELLOW, "║")
    
    print()
    print(c(BOLD + YELLOW, "  ╔══════════════════════════════════════════════════════╗"))
    print("  " + row(      "             T3 Compute — First-Time Setup            "))
    print(c(BOLD + YELLOW, "  ╚══════════════════════════════════════════════════════╝"))
    print()
    print(c(BOLD + CYAN, "  Welcome! This wizard will create your config files."))
    print(c(BOLD + CYAN, "  You can edit them manually at any time after setup."))
    print()

    # ── 1. Tax year ───────────────────────────────────────────────────────────
    _separator()
    print(c(BOLD + CYAN, "  STEP 1 of 5 — Tax Year"))
    print()
    print(c(BOLD + CYAN, "  Which tax year are you processing?"))
    print(c(BOLD + CYAN, "  (CDS releases T3 statements in February for the prior year)"))
    print()
    while True:
        year = _prompt("Tax year (e.g. 2025)")
        if year and year.isdigit() and 2019 <= int(year) <= date.today().year:
            break
        print(c(BOLD + YELLOW, "  Please enter a valid 4-digit year."))

    # ── 2. Base directory ─────────────────────────────────────────────────────
    _separator()
    print(c(BOLD + CYAN, "  STEP 2 of 5 — Folder for T3 source files and output"))
    print()
    print("  This is where your CDS T3 statements live and where the")
    print("  pipeline will write its output files.")
    print(f"  The pipeline expects statements in: <base_dir>\\{year}\\")
    print()
    default_dir = os.path.join(os.path.expanduser("~"), "Documents", "Tax Documents")
    while True:
        base_dir = _prompt("Base folder path", default=default_dir)
        if base_dir:
            break
        print("  Please enter a folder path.")

    # ── 3. ACB spreadsheet ────────────────────────────────────────────────────
    _separator()
    print(c(BOLD + CYAN, "  STEP 3 of 5 — ACB Spreadsheet"))
    print()
    print("  Full path to your ACB tracking spreadsheet (.xlsx).")
    print("  If you haven't created it yet, copy acb_worksheet_template.xlsx")
    print("  and fill in your Buy/Sell transactions first.")
    print()
    default_acb = os.path.join(base_dir, "acb_worksheet.xlsx")
    while True:
        acb_path = _prompt("ACB spreadsheet path", default=default_acb)
        if acb_path:
            break
        print("  Please enter a file path.")

    # ── 4. Funds ──────────────────────────────────────────────────────────────
    _separator()
    print(c(BOLD + CYAN, "  STEP 4 of 5 — Funds"))
    print()
    print("  Enter the ticker symbols of the ETFs you hold, separated by spaces.")
    print("  These must match the sheet names in your ACB spreadsheet exactly.")
    print("  Example: VBAL CPD VRE ZCN")
    print()
    while True:
        funds_input = _prompt("Fund tickers")
        funds = [f.strip().upper() for f in funds_input.split() if f.strip()] if funds_input else []
        if funds:
            break
        print("  Please enter at least one fund ticker.")
    print()
    print(f"  Funds: {', '.join(funds)}")

    # ── 5. Brokerage accounts ─────────────────────────────────────────────────
    _separator()
    print(c(BOLD + CYAN, "  STEP 5 of 5 — Brokerage Accounts"))
    print()
    print("  The pipeline needs to know which brokerage held each fund")
    print("  on each distribution record date, so it can split the T3")
    print("  totals correctly between your accounts.")
    print()
    print("  How many brokerage accounts did you hold these funds in")
    print(f"  during {year}?")
    print()
    while True:
        try:
            num_accounts = int(_prompt("Number of accounts", default="1"))
            if num_accounts >= 1:
                break
        except (TypeError, ValueError):
            pass
        print("  Please enter a number.")

    accounts = []
    for i in range(num_accounts):
        name = _prompt(f"Account {i+1} name (e.g. TDDI, Wealthsimple, DISNAT)")
        accounts.append(name.strip() if name else f"Account{i+1}")

    # For each fund, assign account periods
    periods = {year: {}}
    print()
    print("  Now assign account periods for each fund.")
    print("  If a fund stayed in one account all year, just press Enter")
    print("  to accept the default (all year in first account).")
    print()

    for fund in funds:
        print(f"  --- {fund} ---")
        if num_accounts == 1:
            periods[year][fund] = [{"start": f"{year}-01-01", "account": accounts[0]}]
            print(f"  → All year in {accounts[0]}")
        else:
            fund_periods = []
            print(f"  Which account held {fund} at the start of {year}?")
            for idx, acc in enumerate(accounts):
                print(f"    {idx+1}. {acc}")
            while True:
                try:
                    choice = int(_prompt("Account number", default="1")) - 1
                    if 0 <= choice < len(accounts):
                        break
                except (TypeError, ValueError):
                    pass
                print("  Please enter a valid number.")

            fund_periods.append({"start": f"{year}-01-01", "account": accounts[choice]})

            first_transfer = True
            while True:
                if first_transfer:
                    moved = _prompt_yn(f"  Did {fund} transfer to a different account during {year}?")
                    first_transfer = False
                else:
                    moved = _prompt_yn(f"  Did {fund} transfer again to another account?")
                if not moved:
                    break
                while True:
                    transfer_date = _prompt("  Transfer date (YYYY-MM-DD)")
                    try:
                        date.fromisoformat(transfer_date)
                        break
                    except (ValueError, TypeError):
                        print("  Invalid date format — please enter as YYYY-MM-DD (e.g. 2025-08-11).")
                print(f"  Which account did {fund} move to?")
                for idx, acc in enumerate(accounts):
                    print(f"    {idx+1}. {acc}")
                while True:
                    try:
                        choice = int(_prompt("Account number", default="1")) - 1
                        if 0 <= choice < len(accounts):
                            break
                    except (TypeError, ValueError):
                        pass
                    print("  Please enter a valid number.")
                fund_periods.append({"start": transfer_date, "account": accounts[choice]})

            periods[year][fund] = fund_periods
        print()

    # ── Write config.json ─────────────────────────────────────────────────────
    config = {
        "tax_year": year,
        "base_dir": base_dir,
        "acb_spreadsheet": acb_path,
        "funds": funds,
        "tax_rates": _KNOWN_TAX_RATES,
        "col_indices": {"date": 0, "share_balance": 6},
    }
    with open(config_path, "w") as f:
        json.dump(config, f, indent=2)

    # ── Write account_periods.json ────────────────────────────────────────────
    # Merge with existing file if it exists (to preserve prior years)
    existing_periods = {}
    if os.path.exists(periods_path):
        try:
            with open(periods_path) as f:
                existing_periods = json.load(f)
        except Exception:
            pass
    existing_periods.update(periods)
    with open(periods_path, "w") as f:
        json.dump(existing_periods, f, indent=2)

    # ── Write funds.json (always — create with defaults, preserve existing overrides) ──
    if os.path.exists(funds_path):
        try:
            with open(funds_path) as f:
                existing_funds = json.load(f)
        except Exception:
            existing_funds = {}
    else:
        existing_funds = {}
    # Add any new funds with null override; never overwrite existing entries
    for fund in funds:
        if fund not in existing_funds:
            existing_funds[fund] = {"calc_method_override": None}
    with open(funds_path, "w") as f:
        json.dump(existing_funds, f, indent=2)

    # ── Copy ACB template if spreadsheet doesn't exist yet ───────────────────
    template_src = os.path.join(script_dir, "acb_worksheet_template.xlsx")
    acb_copied = False
    if not os.path.exists(acb_path) and os.path.exists(template_src):
        try:
            os.makedirs(os.path.dirname(os.path.abspath(acb_path)), exist_ok=True)
            shutil.copy2(template_src, acb_path)
            acb_copied = True
        except Exception:
            pass

    # ── Summary ───────────────────────────────────────────────────────────────
    _separator()
    print(c(GREEN,  "  ✓  Setup complete! Config files written:"))
    print(c(GREEN,  f"     config.json            tax year {year}, {len(funds)} fund(s)"))
    print(c(GREEN,  f"     account_periods.json   {num_accounts} account(s)"))
    print(c(GREEN,  f"     funds.json             {', '.join(funds)}"))
    if acb_copied:
        print(c(GREEN,  f"     acb_worksheet.xlsx     copied from template to {acb_path}"))
    print()
    print("  Next steps before running the pipeline:")
    print()
    use_pdf = int(year) >= 2025
    ext = "PDF" if use_pdf else "XLS"
    print(f"  1. Download your CDS T3 {ext} files from:")
    print("     https://ctbsext.posttrade.cds.ca/ctbsExt/")
    print(f"     Save them in: {os.path.join(base_dir, year)}")
    for fund in funds:
        fname = f"{fund}_T3_{year}.{'pdf' if use_pdf else 'xls'}"
        print(f"     → {fname}")
    print()
    if acb_copied:
        print(f"  2. Open {acb_path}")
        print("     Add your Buy/Sell transactions for each fund")
        print("     (one sheet per fund, named exactly after the ticker)")
    else:
        print(f"  2. Make sure {acb_path} has your Buy/Sell transactions")
    print()
    print("  3. Run the pipeline — Step 1 first, then 2, 3, 4 in order")
    print()
    print("  You can edit config.json at any time to change the tax year,")
    print("  paths, or fund list.")
    _separator()

    input("  Press Enter to continue to the main menu...")


def check_and_run_setup(script_dir, config_dir, args):
    """
    Check if any config files are missing. If so, run the first-time wizard.
    Only called in interactive mode — CLI mode (--step / --all) skips this
    entirely to avoid blocking automation.
    """
    config_path  = os.path.join(config_dir, "config.json")
    periods_path = os.path.join(config_dir, "account_periods.json")
    funds_path   = os.path.join(config_dir, "funds.json")

    # Keep args.config in sync so run_step() picks up the right path
    args.config = config_path

    missing = []
    if not os.path.exists(config_path):
        missing.append("config.json")
    if not os.path.exists(periods_path):
        missing.append("account_periods.json")
    if not os.path.exists(funds_path):
        missing.append("funds.json")

    if not missing:
        return   # all present, nothing to do

    def row(content):
        return c(BOLD + YELLOW, "│") + c(BOLD + CYAN, content) + c(BOLD + YELLOW, "│")
    print()
    print(c(BOLD + YELLOW, "  ┌───────────────────────────────────────────────────────┐"))
    print("  " + row(      "     Welcome to T3 Compute!                            "))
    print("  " + row(      "     Config files not found — starting setup wizard... "))
    print(c(BOLD + YELLOW, "  └───────────────────────────────────────────────────────┘"))
    print()
    print(c(YELLOW, f"  Missing: {', '.join(missing)}"))

    run_first_time_setup(script_dir, config_dir, config_path, periods_path, funds_path)


# ── Step metadata ─────────────────────────────────────────────────────────────

STEPS = {
    1: {
        "name":   "Parse T3 PDFs",
        "module": "parse_t3_pdfs",
        "desc":   "Reads CDS T3 PDFs → distribution JSONs + ACB Excel files",
    },
    2: {
        "name":   "Update ACB Spreadsheet",
        "module": "update_acb",
        "desc":   "Inserts ROC rows from Step 1 into your ACB spreadsheet",
    },
    3: {
        "name":   "Build Assets",
        "module": "build_assets",
        "desc":   "Reads ACB spreadsheet + distribution JSONs → per-account JSONs",
    },
    4: {
        "name":   "Compute T3",
        "module": "compute_t3",
        "desc":   "Reads distribution + account JSONs → T3 slip totals",
    },
}

# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        prog="run_t3.py",
        description="T3 tax pipeline wrapper — parse PDFs, build assets, compute T3 slips, update ACB.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python run_t3.py                            Interactive menu
  python run_t3.py --all                      Run all four steps
  python run_t3.py --step 1                   Parse PDFs only
  python run_t3.py --step 2                   Update ACB spreadsheet
  python run_t3.py --step 2 --dry-run         Preview ACB insertions only
  python run_t3.py --step 3                   Build assets only
  python run_t3.py --step 4                   Compute T3 only
  python run_t3.py --all --year 2026          Run all steps for a different year
  python run_t3.py --step 1 --funds VBAL ZCN  Run step 1 for specific funds
  python run_t3.py --step 4 --export                # all funds
  python run_t3.py --step 4 --export --funds VBAL   # VBAL only → T3_2024_VBAL.html
  python run_t3.py --all --export --funds VBAL ZCN  # run all + export VBAL ZCN
        """,
    )
    parser.add_argument(
        "--step",
        type=int,
        choices=[1, 2, 3, 4],
        help="Run a single pipeline step (1, 2, 3, or 4)",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Run all four steps in sequence",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Path to config.json (default: config.json in current directory)",
    )
    parser.add_argument(
        "--year",
        help="Override the tax year in config.json",
    )
    parser.add_argument(
        "--funds",
        nargs="+",
        help="Override the fund list (e.g. --funds VBAL ZCN)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Step 2 only: preview ROC insertions without modifying the ACB spreadsheet",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the graphical interface",
    )
    parser.add_argument(
        "--export",
        action="store_true",
        help="After Step 4, generate T3_<year>.html in base_dir/<year>/",
    )
    return parser.parse_args()

# ── Step runner ───────────────────────────────────────────────────────────────

def build_step_argv(args, step_num):
    """
    Build sys.argv for a given step so the step's own parse_args() picks up
    the overrides passed to the wrapper without requiring any changes to the
    individual scripts.
    """
    argv = [STEPS[step_num]["module"] + ".py"]
    argv += ["--config", args.config]
    if args.year:
        argv += ["--year", args.year]
    # --funds is supported by all steps
    if args.funds and step_num in (1, 2, 3, 4):
        argv += ["--funds"] + args.funds
    # --dry-run is only for step 2
    if step_num == 2 and hasattr(args, "dry_run") and args.dry_run:
        argv += ["--dry-run"]
    return argv

def run_step(step_num, args):
    """Import and execute the main() of the requested step."""
    info = STEPS[step_num]
    print()
    print(c(BOLD + YELLOW, "=" * 60))
    print(c(BOLD + YELLOW, f"  STEP {step_num}: {info['name']}"))
    print(c(BOLD + YELLOW, "=" * 60))

    original_argv = sys.argv
    sys.argv = build_step_argv(args, step_num)

    try:
        if step_num == 1:
            import parse_t3_pdfs
            parse_t3_pdfs.main()
        elif step_num == 2:
            import update_acb
            update_acb.main()
        elif step_num == 3:
            import build_assets
            build_assets.main()
        elif step_num == 4:
            import compute_t3
            compute_t3.main()
    except SystemExit as e:
        if e.code == 2:
            # Exit code 2 = completed with warnings (e.g. missing PDFs in Step 1)
            print(c(YELLOW, f"\n  Step {step_num} completed with warnings — see above."))
            return 1  # signal warnings to caller
        if e.code != 0:
            print(c(RED, f"\n  Step {step_num} exited with code {e.code}."))
            raise
    except Exception as e:
        print(c(RED, f"\n  ERROR in step {step_num}: {e}"))
        raise
    finally:
        sys.argv = original_argv

    print(c(GREEN, f"\n  Step {step_num} completed successfully."))
    return 0  # signal success

# ── Interactive menu ──────────────────────────────────────────────────────────

def show_menu(export_available=False):
    def row(content):
        return c(BOLD + YELLOW, "║") + c(BOLD + CYAN, content) + c(BOLD + YELLOW, "║")
    print()
    print(c(BOLD + YELLOW, "╔═══════════════════════════════════════════════════════════╗"))
    print(row(             "                    T3 Tax Pipeline                        "))
    print(c(BOLD + YELLOW, "╠═══════════════════════════════════════════════════════════╣"))
    print(row(             "  1  Parse T3 PDFs           Reads CDS PDFs → JSONs + ACB  "))
    print(row(             "  2  Update ACB Spreadsheet  Inserts ROC rows from Step 1  "))
    print(row(             "  3  Build Assets            Reads ACB spreadsheet → JSONs "))
    print(row(             "  4  Compute T3              Reads JSONs → T3 slip totals  "))
    print(row(             "  A  Run all steps           Execute all steps in sequence "))
    if export_available:
        print(row(         "  E  Export T3 HTML          Save T3 HTML from last run    "))
    print(row(             "  Q  Quit                                                  "))
    print(c(BOLD + YELLOW, "╚═══════════════════════════════════════════════════════════╝"))
    print()

def interactive_mode(args):
    """Present a menu and let the user choose which step(s) to run."""
    print()
    print("T3 Tax Pipeline — Interactive Mode")
    print("(Run with --help for CLI usage)")

    session_export_ready = False
    session_export_year  = None
    session_export_funds = None
    
    while True:
        show_menu(export_available=session_export_ready)
        valid = "1 / 2 / 3 / 4 / A / E / Q" if session_export_ready else "1 / 2 / 3 / 4 / A / Q"
        choice = input(c(BOLD + CYAN, f"  Enter choice [{valid}]: ")).strip().upper()

        if choice == "Q":
            print(c(BOLD + CYAN, "Exiting."))
            break

        elif choice == "E":
            if not session_export_ready:
                print(c(BOLD + YELLOW, "\n  ⚠  Run Step 4 (Compute T3) first before exporting."))
                continue
            try:
                out = generate_t3_html(args.config, session_export_year, session_export_funds)
                print(c(GREEN, f"\n  ✔  T3 HTML saved to: {out}"))
            except Exception as ex:
                print(c(RED, f"\n  ERROR: {ex}"))
            continue

        elif choice == "A":
            steps_to_run = [1, 2, 3, 4]
        elif choice in ("1", "2", "3", "4"):
            steps_to_run = [int(choice)]
        else:
            print(c(BOLD + YELLOW,"  Invalid choice, please try again."))
            continue

        # Reset overrides each iteration so pressing Enter truly means "use config default"
        args.year  = None
        args.funds = None

        # Optional year override
        year_input = input(
            c(BOLD + CYAN, "  Tax year override? (press Enter to use config default): ")
        ).strip()
        if year_input:
            args.year = year_input

        # Optional fund override for all steps
        if any(s in steps_to_run for s in (1, 2, 3, 4)):
            funds_input = input(
                c(BOLD + CYAN, "  Fund override? (e.g. VBAL ZCN, or Enter for all): ")
            ).strip()
            if funds_input:
                args.funds = funds_input.split()

        # Dry-run option for step 2
        if 2 in steps_to_run:
            dry = input(
                c(BOLD + CYAN, "  Dry-run for Step 2? Preview insertions without modifying file [Y/N]: ")
            ).strip().upper()
            args.dry_run = (dry == "Y")

        # Run selected steps
        step4_ran_ok = False
        try:
            for step_num in steps_to_run:
                rc = run_step(step_num, args)
                if step_num == 4 and rc is not None:
                    step4_ran_ok = True
        except (SystemExit, Exception):
            print("\nPipeline stopped due to an error.")

        # Only unlock E if Step 4 completed successfully this session
        if step4_ran_ok:
            session_export_ready = True
            session_export_year  = args.year   # may be None — that's fine
            session_export_funds = args.funds  # may be None — exports all funds
            print(c(GREEN, "\n  ✔  E (Export T3 HTML) is now available in the menu."))

        print()
        again = input(c(BOLD + CYAN, "  Run another step? [Y / N]: ")).strip().upper()
        if again != "Y":
            print(c(BOLD + CYAN, "Exiting."))
            break

# ── T3 HTML export ────────────────────────────────────────────────────────────

def generate_t3_html(config_path, year_override=None, funds_override=None):
    import compute_t3 as ct3
    cfg = ct3.load_config(config_path)
    if year_override:
        cfg["tax_year"] = year_override
        ct3._set_derived_paths(cfg, year_override)
    year     = cfg["tax_year"]
    funds    = ct3.load_distributions(cfg["dist_dir"])

    # Filter to only the requested funds
    if funds_override:
        unknown = [f for f in funds_override if f not in funds]
        if unknown:
            raise ValueError(f"Fund(s) not found in distributions: {', '.join(unknown)}")
        funds = {k: v for k, v in funds.items() if k in funds_override}

    accounts = ct3.load_assets(cfg["assets_dir"])
    results, _ = ct3.compute_t3(funds, accounts, cfg["tax_rates"])

    # Use fund names in filename if filtered
    fund_suffix = f"_{'_'.join(sorted(funds_override))}" if funds_override else ""
    out_dir  = os.path.join(cfg["base_dir"], year)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"T3_{year}{fund_suffix}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(_build_t3_html(results, year))
    return out_path

def _do_export(args):
    """CLI wrapper for the HTML export — called after Step 4."""
    # Guard: verify Step 4 output exists
    try:
        with open(args.config) as f:
            cfg = json.load(f)
        year     = (args.year if hasattr(args, "year") and args.year
                    else cfg.get("tax_year", ""))
        base_dir = cfg.get("base_dir", "")
        txt      = os.path.join(base_dir, year, f"T3_results_{year}.txt")
        if not os.path.exists(txt):
            print(c(BOLD + YELLOW, "\n  ⚠  T3_results file not found — run Step 4 first."))
            print(c(BOLD + YELLOW, f"     Expected: {txt}"))
            return
    except Exception as ex:
        print(c(RED, f"\n  ERROR reading config: {ex}"))
        return

    print()
    print(c(BOLD + YELLOW, "=" * 60))
    print(c(BOLD + YELLOW, " EXPORT: Generating T3 HTML slip..."))
    print(c(BOLD + YELLOW, "=" * 60))
    try:
        yr    = args.year
        funds = args.funds or None
        out   = generate_t3_html(args.config, yr, funds)
        print(c(GREEN, f"\n  ✔  T3 HTML saved to: {out}"))
    except Exception as ex:
        print(c(RED, f"\n  ERROR generating T3 HTML: {ex}"))

def _build_t3_html(results, year):
    """Build CRA-layout T3 slip HTML using the official 5-column grid. Pure stdlib."""
    from datetime import date as _date

    def v(totals, field):
        val = totals.get(field, 0.0)
        return f"${val:,.2f}" if val else ""

    def other_rows(totals):
        extras = [
            ("25", "foreignNonBusinessIncome"),
            ("34", "foreignNonBusinessIncomeTaxPaid"),
            ("42", "returnOfCapital"),
        ]
        filled = [(box, v(totals, field)) for box, field in extras if v(totals, field)]
        while len(filled) < 4:
            filled.append(("", ""))
        rows_html = ""
        for i in range(0, len(filled), 2):
            b1, a1 = filled[i]
            b2, a2 = filled[i + 1] if i + 1 < len(filled) else ("", "")
            rows_html += f"""<tr>
                <td>{b1}</td><td>{a1}</td>
                <td>{b2}</td><td>{a2}</td>
            </tr>"""
        return rows_html

    forms_html = ""
    for account in sorted(results.keys()):
        t = results[account]
        forms_html += f"""
    <div class="account-block">
    <div class="form-container">
        <div class="header">
            <div>
                <strong>Canada Revenue Agency</strong><br>
                <strong>Agence du revenu du Canada</strong>
            </div>
            <div style="text-align:center;">
                <h1 style="font-size:14px;margin:0;">Statement of Trust Income Allocations and Designations</h1>
                <h2 style="font-size:12px;margin:0;font-style:italic;">État des revenus de fiducie (répartitions et attributions)</h2>
                <div style="margin-top:4px;font-size:11px;color:#555;">Account: <strong>{account}</strong></div>
            </div>
            <div style="text-align:right;">
                <strong style="font-size:20px;">T3</strong><br>
                Protected B when completed<br>
                <span class="fr">Protégé B une fois rempli</span>
            </div>
        </div>

        <div style="display:flex;justify-content:flex-end;margin-bottom:5px;">
            <div style="border:1px solid #000;padding:5px;">
                <strong>Year / <span class="fr">Année</span></strong>:
                <input type="text" style="width:60px;border:none;border-bottom:1px solid #000;"
                       value="{year}">
            </div>
        </div>

        <div class="grid-main">
            <!-- Row 1: Eligible dividends -->
            <div class="box span-2">
                <span class="box-num">49</span>
                <div class="box-label">Actual amount of eligible dividends<br>
                    <span class="fr">Montant réel des dividendes déterminés</span></div>
                <input type="text" class="input-area" value="{v(t,'actualAmountOfEligibleDividends')}">
            </div>
            <div class="box span-2">
                <span class="box-num">50</span>
                <div class="box-label">Taxable amount of eligible dividends<br>
                    <span class="fr">Montant imposable des dividendes déterminés</span></div>
                <input type="text" class="input-area" value="{v(t,'taxableAmountOfEligibleDividends')}">
            </div>
            <div class="box">
                <span class="box-num">51</span>
                <div class="box-label">Dividend tax credit for eligible dividends<br>
                    <span class="fr">Crédit d'impôt pour dividendes déterminés</span></div>
                <input type="text" class="input-area" value="{v(t,'dividendTaxCreditForEligibleDividends')}">
            </div>

            <!-- Row 2: Non-eligible dividends -->
            <div class="box span-2">
                <span class="box-num">23</span>
                <div class="box-label">Actual amount of dividends other than eligible dividends<br>
                    <span class="fr">Montant réel des dividendes autres que des dividendes déterminés</span></div>
                <input type="text" class="input-area" value="{v(t,'actualAmountOfNonEligibleDividends')}">
            </div>
            <div class="box span-2">
                <span class="box-num">32</span>
                <div class="box-label">Taxable amount of dividends other than eligible dividends<br>
                    <span class="fr">Montant imposable des dividendes autres que des dividendes déterminés</span></div>
                <input type="text" class="input-area" value="{v(t,'taxableAmountOfNonEligibleDividends')}">
            </div>
            <div class="box">
                <span class="box-num">39</span>
                <div class="box-label">Dividend tax credit for dividends other than eligible dividends<br>
                    <span class="fr">Crédit d'impôt pour dividendes autres que des dividendes déterminés</span></div>
                <input type="text" class="input-area" value="{v(t,'dividendTaxCreditForNonEligibleDividends')}">
            </div>

            <!-- Row 3: Capital gains / Other income -->
            <div class="box span-2">
                <span class="box-num">21</span>
                <div class="box-label">Capital gains<br>
                    <span class="fr">Gains en capital</span></div>
                <input type="text" class="input-area" value="{v(t,'capitalGains')}">
            </div>
            <div class="box span-2">
                <span class="box-num">30</span>
                <div class="box-label">Capital gains eligible for deduction<br>
                    <span class="fr">Gains en capital admissibles pour déduction</span></div>
                <input type="text" class="input-area" value="">
            </div>
            <div class="box">
                <span class="box-num">26</span>
                <div class="box-label">Other income<br>
                    <span class="fr">Autres revenus</span></div>
                <input type="text" class="input-area" value="{v(t,'otherIncome')}">
            </div>

            <!-- Row 4: Other info / Footnotes / Trust year end -->
            <div class="box span-2">
                <div class="box-label" style="margin-left:0;">
                    Other information — Box 25 Foreign Income · Box 34 Foreign Tax · Box 42 ROC<br>
                    <span class="fr">Autres renseignements</span>
                </div>
                <table class="other-info-table">
                    <tr>
                        <th>Box/Case</th><th>Amount/Montant</th>
                        <th>Box/Case</th><th>Amount/Montant</th>
                    </tr>
                    {other_rows(t)}
                </table>
            </div>
            <div class="box span-2">
                <div class="box-label" style="margin-left:0;">
                    Footnotes — Notes de bas de page
                </div>
                <div style="height:50px;"></div>
            </div>
            <div class="box">
                <div class="box-label" style="margin-left:0;">
                    Trust year end<br>
                    <span class="fr">Fin de l'année de la fiducie</span>
                </div>
                <div style="margin-top:5px;">
                    Year/Année: <input type="text" style="width:40px;border:none;border-bottom:1px solid #000;" value="{year}"><br>
                    Month/Mois: <input type="text" style="width:30px;border:none;border-bottom:1px solid #000;" value="12">
                </div>
            </div>

            <!-- Row 5: Recipient / Trust name -->
            <div class="box span-3">
                <div class="box-label" style="margin-left:0;">
                    Recipient's name (last name first) and address<br>
                    <span class="fr">Nom, prénom et adresse du bénéficiaire</span>
                </div>
                <div style="height:40px;"></div>
            </div>
            <div class="box span-2">
                <div class="box-label" style="margin-left:0;">
                    Trust's name and address<br>
                    <span class="fr">Nom et adresse de la fiducie</span>
                </div>
                <div style="height:40px;"></div>
            </div>

            <!-- Row 6: ID fields -->
            <div class="box">
                <span class="box-num">12</span>
                <div class="box-label">Recipient identification number<br>
                    <span class="fr">Numéro d'identification du bénéficiaire</span></div>
                <input type="text" class="input-area">
            </div>
            <div class="box">
                <span class="box-num">14</span>
                <div class="box-label">Account number<br>
                    <span class="fr">Numéro de compte</span></div>
                <input type="text" class="input-area" value="{account}">
            </div>
            <div class="box">
                <span class="box-num">16</span>
                <div class="box-label">Report code<br>
                    <span class="fr">Code du type de feuillet</span></div>
                <input type="text" class="input-area">
            </div>
            <div class="box">
                <span class="box-num">18</span>
                <div class="box-label">Beneficiary code<br>
                    <span class="fr">Code du bénéficiaire</span></div>
                <input type="text" class="input-area">
            </div>
            <div class="box" style="display:flex;align-items:center;justify-content:center;background:#f9f9f9;">
                <strong style="font-size:14px;">T3 ({year})</strong>
            </div>
        </div>

        <div class="footer-note">
            <div>For details, see next pages.<br>
                <span class="fr">Lisez aussi les renseignements aux pages suivantes.</span>
            </div>
            <div style="text-align:right;color:#888;">
                Generated by T3 Compute — personal reference only
            </div>
        </div>
    </div>
    </div>"""

    today = _date.today().strftime("%B %d, %Y")
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>T3 Statement {year}</title>
    <style>
        @media print {{
            .noprint {{ display: none; }}
            .account-block {{ break-before: page; }}
            .account-block:first-of-type {{ break-before: auto; }}
            body {{ margin: 8mm; }}
            .form-container {{ box-shadow: none; border: 1px solid #000; }}
        }}
        body {{ font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;
                font-size:11px; color:#000; margin:20px; background:#f4f4f4; }}
        .noprint {{ text-align:center; margin-bottom:16px; }}
        .noprint button {{ padding:7px 22px; font-size:11pt; cursor:pointer;
                           background:#1a3a6b; color:#fff; border:none; border-radius:3px; }}
        .noprint small {{ margin-left:12px; color:#555; font-size:10px; }}
        .account-block {{ margin-bottom:32px; }}
        .form-container {{ width:950px; margin:0 auto; border:1px solid #000;
                           padding:15px; background:#fff;
                           box-shadow:0 0 10px rgba(0,0,0,.1); }}
        .header {{ display:flex; justify-content:space-between; align-items:center;
                   border-bottom:2px solid #000; padding-bottom:5px; margin-bottom:10px; }}
        .grid-main {{ display:grid; grid-template-columns:repeat(5,1fr);
                      gap:0; border-top:1px solid #000; border-left:1px solid #000; }}
        .box {{ border-right:1px solid #000; border-bottom:1px solid #000;
                padding:4px; min-height:45px; position:relative; }}
        .box-num {{ position:absolute; top:2px; left:2px; border:1px solid #000;
                    font-weight:bold; padding:1px 3px; background:#eee; font-size:10px; }}
        .box-label {{ font-size:9px; line-height:1.1; font-weight:bold; margin-left:25px; }}
        .fr {{ font-style:italic; font-weight:normal; }}
        .input-area {{ width:95%; border:none; border-bottom:1px dotted #ccc;
                       margin-top:15px; outline:none; font-size:12px; background:transparent; }}
        .span-2 {{ grid-column:span 2; }}
        .span-3 {{ grid-column:span 3; }}
        .other-info-table {{ width:100%; border-collapse:collapse; margin-top:5px; }}
        .other-info-table th,
        .other-info-table td {{ border:1px solid #000; padding:2px; font-size:9px; }}
        .footer-note {{ font-size:10px; margin-top:10px;
                        display:flex; justify-content:space-between; }}
        .gen-date {{ text-align:center; color:#aaa; font-size:9px; margin-top:6px; }}
    </style>
</head>
<body>
<div class="noprint">
    <button onclick="window.print()">&#128438; Print / Save as PDF</button>
    <small>Browser Print → Save as PDF · Generated {today}</small>
</div>
{forms_html}
</body>
</html>"""

# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    # ── Hide console window when launched as GUI (frozen exe, no CLI flags) ──
    if getattr(sys, 'frozen', False) and sys.platform == 'win32' and not any(
        a in sys.argv for a in ['--step', '--all', '--export', '--gui', '--help']
    ):
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)

    config_dir = get_config_dir(script_dir)
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    args = parse_args()
    if not hasattr(args, "dry_run"):
        args.dry_run = False

    # GUI mode: frozen exe with no CLI arguments → launch the GUI
    if args.gui or (args.step is None and not args.all and getattr(sys, "frozen", False)):
        from t3_gui import main as gui_main
        gui_main()
        return

    # CLI / interactive mode
    if args.step is None and not args.all:
        check_and_run_setup(script_dir, config_dir, args)
        interactive_mode(args)
    elif args.all:
        try:
            for step_num in (1, 2, 3, 4):
                run_step(step_num, args)
            print(c(BOLD + GREEN, "\nAll steps completed."))
            if args.export:
                _do_export(args)
        except (SystemExit, Exception):
            print(c(BOLD + RED, "\nPipeline stopped due to an error."))
    else:
        run_step(args.step, args)
        if args.export and args.step == 4:
            _do_export(args)
        elif args.export and args.step != 4:
            print(c(YELLOW, "  --export only applies to --step 4, skipping."))

if __name__ == "__main__":
    main()
