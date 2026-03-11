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
    # --funds is supported by steps 1, 2, and 4
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
    print("╔════════════════════════════════════════════════════════════╗")
    print("║                    T3 Tax Pipeline                         ║")
    print("╠════════════════════════════════════════════════════════════╣")
    print("║  1  Parse T3 PDFs          Reads CDS PDFs → JSONs + ACB    ║")
    print("║  2  Update ACB Spreadsheet Inserts ROC rows from Step 1    ║")
    print("║  3  Build Assets           Reads ACB spreadsheet → JSONs   ║")
    print("║  4  Compute T3             Reads JSONs → T3 slip totals    ║")
    print("║  A  Run all steps in sequence                              ║")
    print("║  Q  Quit                                                   ║")
    print("╚════════════════════════════════════════════════════════════╝")
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

        # Optional fund override for steps 1, 2, and 4
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

    args = parse_args()

    # Default dry_run to False if not set (older argparse won't have it)
    if not hasattr(args, "dry_run"):
        args.dry_run = False

    if args.step is None and not args.all:
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
