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
    python run_t3.py --step 1 --funds VBAL ZCN  # override fund list
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

Config files:
    Python script (python run_t3.py / setup.bat):
        Same directory as this script — no change from previous behaviour.
    Bundled .exe (PyInstaller / Microsoft Store):
        %%LOCALAPPDATA%%\\T3Compute\\
        e.g. C:\\Users\\<you>\\AppData\\Local\\T3Compute\\

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

# ── Config directory ─────────────────────────────────────────────────────────

def get_config_dir(script_dir):
    """
    Return the directory where config files should be read/written.

    - PyInstaller bundle (.exe / MSIX): uses %LOCALAPPDATA%\\T3Compute\\
      The PyInstaller extraction folder (script_dir) is a temp path that is
      deleted on exit, so config files must live elsewhere.
    - Regular Python script: uses script_dir, preserving existing behaviour
      for setup.bat and command-line users.
    """
    if getattr(sys, "frozen", False):
        # Running as PyInstaller bundle
        local_app_data = os.environ.get(
            "LOCALAPPDATA",
            os.path.join(os.path.expanduser("~"), "AppData", "Local")
        )
        config_dir = os.path.join(local_app_data, "T3Compute")
        os.makedirs(config_dir, exist_ok=True)
        return config_dir
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
    print()
    print("  ╔══════════════════════════════════════════════════════╗")
    print("  ║           T3 Compute — First-Time Setup              ║")
    print("  ╚══════════════════════════════════════════════════════╝")
    print()
    print("  Welcome! This wizard will create your config files.")
    print("  You can edit them manually at any time after setup.")
    print()

    # ── 1. Tax year ───────────────────────────────────────────────────────────
    _separator()
    print("  STEP 1 of 5 — Tax Year")
    print()
    print("  Which tax year are you processing?")
    print("  (CDS releases T3 statements in February for the prior year)")
    print()
    while True:
        year = _prompt("Tax year (e.g. 2025)")
        if year and year.isdigit() and 2019 <= int(year) <= date.today().year:
            break
        print("  Please enter a valid 4-digit year.")

    # ── 2. Base directory ─────────────────────────────────────────────────────
    _separator()
    print("  STEP 2 of 5 — Folder for T3 source files and output")
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
    print("  STEP 3 of 5 — ACB Spreadsheet")
    print()
    print("  Full path to your ACB tracking spreadsheet (.xlsx).")
    print("  If you haven't created it yet, copy acb_worksheet_template.xlsx")
    print("  and fill in your Buy/Sell transactions first.")
    print()
    default_acb = os.path.join(config_dir, "acb_worksheet.xlsx")
    while True:
        acb_path = _prompt("ACB spreadsheet path", default=default_acb)
        if acb_path:
            break
        print("  Please enter a file path.")

    # ── 4. Funds ──────────────────────────────────────────────────────────────
    _separator()
    print("  STEP 4 of 5 — Funds")
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
    print("  STEP 5 of 5 — Brokerage Accounts")
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
    print("  ✓  Setup complete! Config files written:")
    print(f"     config.json            tax year {year}, {len(funds)} fund(s)")
    print(f"     account_periods.json   {num_accounts} account(s)")
    print(f"     funds.json             {', '.join(funds)}")
    if acb_copied:
        print(f"     acb_worksheet.xlsx     copied from template to {acb_path}")
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

    print()
    print("  ┌─────────────────────────────────────────────────────┐")
    print("  │  Welcome to T3 Compute!                             │")
    print("  │  Config files not found — starting setup wizard...  │")
    print("  └─────────────────────────────────────────────────────┘")
    print()
    print(f"  Missing: {', '.join(missing)}")

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
    # --funds is supported by steps 1, 2, and 3
    if args.funds and step_num in (1, 2, 3):
        argv += ["--funds"] + args.funds
    # --dry-run is only for step 2
    if step_num == 2 and hasattr(args, "dry_run") and args.dry_run:
        argv += ["--dry-run"]
    return argv

def run_step(step_num, args):
    """Import and execute the main() of the requested step."""
    info = STEPS[step_num]
    print()
    print("=" * 60)
    print(f"  STEP {step_num}: {info['name']}")
    print("=" * 60)

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
            print(f"\nStep {step_num} completed with warnings — see above.")
            sys.argv = original_argv
            return
        if e.code != 0:
            print(f"\nStep {step_num} exited with code {e.code}.")
            raise
    except Exception as e:
        print(f"\nERROR in step {step_num}: {e}")
        raise
    finally:
        sys.argv = original_argv

    print(f"\nStep {step_num} completed successfully.")

# ── Interactive menu ──────────────────────────────────────────────────────────

def show_menu():
    print()
    print("╔═══════════════════════════════════════════════════════════╗")
    print("║                    T3 Tax Pipeline                        ║")
    print("╠═══════════════════════════════════════════════════════════╣")
    print("║  1  Parse T3 PDFs         Reads CDS PDFs → JSONs + ACB    ║")
    print("║  2  Update ACB Spreadsheet  Inserts ROC rows from Step 1  ║")
    print("║  3  Build Assets          Reads ACB spreadsheet → JSONs   ║")
    print("║  4  Compute T3            Reads JSONs → T3 slip totals    ║")
    print("║  A  Run all steps in sequence                             ║")
    print("║  Q  Quit                                                  ║")
    print("╚═══════════════════════════════════════════════════════════╝")
    print()

def interactive_mode(args):
    """Present a menu and let the user choose which step(s) to run."""
    print()
    print("T3 Tax Pipeline — Interactive Mode")
    print("(Run with --help for CLI usage)")

    while True:
        show_menu()
        choice = input("  Enter choice [1 / 2 / 3 / 4 / A / Q]: ").strip().upper()

        if choice == "Q":
            print("Exiting.")
            break
        elif choice == "A":
            steps_to_run = [1, 2, 3, 4]
        elif choice in ("1", "2", "3", "4"):
            steps_to_run = [int(choice)]
        else:
            print("  Invalid choice, please try again.")
            continue

        # Optional year override
        year_input = input(
            "  Tax year override? (press Enter to use config default): "
        ).strip()
        if year_input:
            args.year = year_input

        # Optional fund override for steps 1, 2, and 3
        if any(s in steps_to_run for s in (1, 2, 3)):
            funds_input = input(
                "  Fund override for steps 1/2/3? (e.g. VBAL ZCN, or Enter for all): "
            ).strip()
            if funds_input:
                args.funds = funds_input.split()

        # Dry-run option for step 2
        if 2 in steps_to_run:
            dry = input(
                "  Dry-run for Step 2? Preview insertions without modifying file [Y/N]: "
            ).strip().upper()
            args.dry_run = (dry == "Y")

        # Run selected steps
        try:
            for step_num in steps_to_run:
                run_step(step_num, args)
        except (SystemExit, Exception):
            print("\nPipeline stopped due to an error.")

        print()
        again = input("  Run another step? [Y / N]: ").strip().upper()
        if again != "Y":
            print("Exiting.")
            break

# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)

    config_dir = get_config_dir(script_dir)

    # For Python/script users, sys.path already has script_dir.
    # For .exe users, also add config_dir so the pipeline scripts can be found.
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    args = parse_args()

    # Default dry_run to False if not set (older argparse won't have it)
    if not hasattr(args, "dry_run"):
        args.dry_run = False

    # First-run setup wizard — only in interactive mode, never in CLI mode
    if args.step is None and not args.all:
        check_and_run_setup(script_dir, config_dir, args)
        interactive_mode(args)
    elif args.all:
        try:
            for step_num in (1, 2, 3, 4):
                run_step(step_num, args)
            print("\nAll steps completed.")
        except (SystemExit, Exception):
            print("\nPipeline stopped due to an error.")
    else:
        run_step(args.step, args)


if __name__ == "__main__":
    main()
