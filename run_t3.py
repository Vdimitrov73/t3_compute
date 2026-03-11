"""
run_t3.py
---------
Wrapper / menu for the T3 tax pipeline.

Interactive mode (no arguments):
    python run_t3.py

CLI mode:
    python run_t3.py --step 1               # parse PDFs only
    python run_t3.py --step 2               # build assets only
    python run_t3.py --step 3               # compute T3 only
    python run_t3.py --all                  # run all three steps in sequence
    python run_t3.py --all --year 2026      # override tax year
    python run_t3.py --step 1 --funds VBAL ZCN  # override fund list
    python run_t3.py --help                 # show this help

Pipeline steps:
    Step 1 — parse_t3_pdfs.py
             Reads CDS T3 PDFs and outputs one distribution JSON per fund
             plus an Excel ACB helper file.

    Step 2 — build_assets.py
             Reads the ACB spreadsheet and the distribution JSONs, then
             outputs one JSON per brokerage account showing share balances
             on each record date.

    Step 3 — compute_t3.py
             Reads the distribution JSONs and account JSONs, computes
             T3 slip totals per account, and saves a results text file.

Config files (must be in the same directory as this script):
    config.json           — tax year, paths, fund list, tax rates
    account_periods.json  — brokerage account periods per fund per year
    funds.json            — per-fund PDF parsing hints
"""

import argparse
import sys
import os

# ── Step metadata ─────────────────────────────────────────────────────────────

STEPS = {
    1: {
        "name":   "Parse T3 PDFs",
        "module": "parse_t3_pdfs",
        "desc":   "Reads CDS T3 PDFs → distribution JSONs + ACB Excel files",
    },
    2: {
        "name":   "Build Assets",
        "module": "build_assets",
        "desc":   "Reads ACB spreadsheet + distribution JSONs → per-account share JSONs",
    },
    3: {
        "name":   "Compute T3",
        "module": "compute_t3",
        "desc":   "Reads distribution + account JSONs → T3 slip totals per account",
    },
}

# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        prog="run_t3.py",
        description="T3 tax pipeline wrapper — parse PDFs, build assets, compute T3 slips.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python run_t3.py                        Interactive menu
  python run_t3.py --all                  Run all three steps
  python run_t3.py --step 1               Run step 1 only (parse PDFs)
  python run_t3.py --step 2               Run step 2 only (build assets)
  python run_t3.py --step 3               Run step 3 only (compute T3)
  python run_t3.py --all --year 2026      Run all steps for a different year
  python run_t3.py --step 1 --funds VBAL ZCN  Run step 1 for specific funds
        """,
    )
    parser.add_argument(
        "--step",
        type=int,
        choices=[1, 2, 3],
        help="Run a single pipeline step (1, 2, or 3)",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Run all three steps in sequence",
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
    # --funds is only supported by steps 1 and 2
    if args.funds and step_num in (1, 2):
        argv += ["--funds"] + args.funds
    return argv

def run_step(step_num, args):
    """Import and execute the main() of the requested step."""
    info = STEPS[step_num]
    print()
    print("=" * 60)
    print(f"  STEP {step_num}: {info['name']}")
    print("=" * 60)

    # Temporarily replace sys.argv so the step's argparse sees our overrides
    original_argv = sys.argv
    sys.argv = build_step_argv(args, step_num)

    try:
        if step_num == 1:
            import parse_t3_pdfs
            parse_t3_pdfs.main()
        elif step_num == 2:
            import build_assets
            build_assets.main()
        elif step_num == 3:
            import compute_t3
            compute_t3.main()
    except SystemExit as e:
        # Catch sys.exit() calls inside the sub-scripts (e.g. missing files)
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
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║                    T3 Tax Pipeline                           ║")
    print("╠══════════════════════════════════════════════════════════════╣")
    print("║  1  Parse T3 PDFs    Reads CDS PDFs → distribution JSONs     ║")
    print("║  2  Build Assets     Reads ACB spreadsheet → account JSONs   ║")
    print("║  3  Compute T3       Reads JSONs → T3 slip totals            ║")
    print("║  A  Run all steps in sequence                                ║")
    print("║  Q  Quit                                                     ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print()

def interactive_mode(args):
    """Present a menu and let the user choose which step(s) to run."""
    print()
    print("T3 Tax Pipeline — Interactive Mode")
    print("(Run with --help for CLI usage)")

    while True:
        show_menu()
        choice = input("  Enter choice [1 / 2 / 3 / A / Q]: ").strip().upper()

        if choice == "Q":
            print("Exiting.")
            break
        elif choice == "A":
            steps_to_run = [1, 2, 3]
        elif choice in ("1", "2", "3"):
            steps_to_run = [int(choice)]
        else:
            print("  Invalid choice, please try again.")
            continue

        # Optional year override in interactive mode
        year_input = input(
            f"  Tax year override? (press Enter to use config default): "
        ).strip()
        if year_input:
            args.year = year_input

        # Optional fund override for steps 1 and 2
        if any(s in steps_to_run for s in (1, 2)):
            funds_input = input(
                "  Fund override for steps 1/2? (e.g. VBAL ZCN, or Enter for all): "
            ).strip()
            if funds_input:
                args.funds = funds_input.split()

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
    # Make sure the script's own directory is on the path so the three
    # pipeline scripts can always be imported regardless of where the
    # user launches from.
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)

    args = parse_args()

    if args.step is None and not args.all:
        # No CLI flags — launch interactive menu
        interactive_mode(args)
    elif args.all:
        # Run all three steps in sequence
        for step_num in (1, 2, 3):
            run_step(step_num, args)
        print("\nAll steps completed.")
    else:
        # Run the single requested step
        run_step(args.step, args)


if __name__ == "__main__":
    main()
