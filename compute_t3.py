"""
compute_t3.py
-------------
Reads your distribution JSONs and account shares JSONs, then computes
the T3 slip totals for each brokerage account.

Usage:
    python compute_t3.py [--config config.json] [--year 2025] [--funds VBAL ZCN]

Canadian T3 boxes computed:
    Box 21  — Capital Gains
    Box 25  — Foreign Non-Business Income
    Box 26  — Other Income
    Box 34  — Foreign Non-Business Income Tax Paid
    Box 42  — Return of Capital
    Box 49  — Actual Amount of Eligible Dividends
    Box 50  — Taxable Amount of Eligible Dividends        (Box 49 × grossup)
    Box 51  — Dividend Tax Credit for Eligible Dividends  (Box 50 × tax_credit)
    Box 23  — Actual Amount of Non-Eligible Dividends
    Box 32  — Taxable Amount of Non-Eligible Dividends    (Box 23 × grossup)
    Box 39  — Dividend Tax Credit for Non-Eligible Dividends

Phantom (non-cash) distributions:
    The "You should have received" line shows cash only.
    All income components are still fully reported for tax purposes.

Precision:
    All income box totals (21, 23, 25, 26, 34, 42, 49) are accumulated at full
    float precision and rounded to 2 decimal places in the summary.
    Derived boxes (50, 51, 32, 39) are computed from the rounded Box 49/23 values.
"""

import argparse
import json
import os
import sys
import glob
from datetime import datetime
from collections import defaultdict

# ── Config ────────────────────────────────────────────────────────────────────

KNOWN_RATES = {
    "2024": {"eligible_div_grossup": 1.38, "eligible_div_tax_credit": 0.150198},
    "2025": {"eligible_div_grossup": 1.38, "eligible_div_tax_credit": 0.150198},
}

def load_config(path="config.json"):
    with open(path) as f:
        cfg = json.load(f)
    _set_derived_paths(cfg, cfg["tax_year"])
    return cfg

def _set_derived_paths(cfg, year):
    cfg["pdf_dir"]    = os.path.join(cfg["base_dir"], year)
    cfg["dist_dir"]   = os.path.join(cfg["base_dir"], year, "distributions")
    cfg["assets_dir"] = os.path.join(cfg["base_dir"], year, "assets")
    cfg["output_txt"] = os.path.join(cfg["base_dir"], year, f"T3_results_{year}.txt")
def parse_args():
    parser = argparse.ArgumentParser(description="Compute T3 slip totals per brokerage account.")
    parser.add_argument("--config", default="config.json", help="Path to config.json")
    parser.add_argument("--year",   help="Override tax year in config")
    parser.add_argument("--funds",  nargs="+", help="Override fund list (e.g. --funds VBAL ZCN)")
    return parser.parse_args()

def check_rates(tax_year, cfg_rates):
    """Warn if the tax year is unknown or if rates differ from known values."""
    if tax_year not in KNOWN_RATES:
        print(f"WARNING: Tax rates for {tax_year} have not been verified in this script.")
        print(f"         Check https://www.canada.ca/en/revenue-agency for current gross-up rates.\n")
    else:
        known = KNOWN_RATES[tax_year]
        if cfg_rates.get("eligible_div_grossup") != known["eligible_div_grossup"]:
            print(f"WARNING: eligible_div_grossup in config ({cfg_rates.get('eligible_div_grossup')}) "
                  f"differs from known {tax_year} rate ({known['eligible_div_grossup']}).\n")
        if cfg_rates.get("eligible_div_tax_credit") != known["eligible_div_tax_credit"]:
            print(f"WARNING: eligible_div_tax_credit in config ({cfg_rates.get('eligible_div_tax_credit')}) "
                  f"differs from known {tax_year} rate ({known['eligible_div_tax_credit']}).\n")

# ── T3 box definitions ────────────────────────────────────────────────────────

# Maps JSON field name → (box number, display label)
FIELD_TO_BOX = {
    "capitalGains":                       ("21", "Capital Gains"),
    "foreignNonBusinessIncome":           ("25", "Foreign Non-Business Income"),
    "otherIncome":                        ("26", "Other Income"),
    "foreignNonBusinessIncomeTaxPaid":    ("34", "Foreign Non-Business Income Tax Paid"),
    "returnOfCapital":                    ("42", "Return of Capital"),
    "actualAmountOfEligibleDividends":    ("49", "Actual Amount of Eligible Dividends"),
    "actualAmountOfNonEligibleDividends": ("23", "Actual Amount of Non-Eligible Dividends"),
}

SUMMARY_FIELDS = [
    ("capitalGains",                             "21", "Capital Gains"),
    ("actualAmountOfEligibleDividends",          "49", "Actual Amount of Eligible Dividends"),
    ("taxableAmountOfEligibleDividends",         "50", "Taxable Amount of Eligible Dividends"),
    ("dividendTaxCreditForEligibleDividends",    "51", "Dividend Tax Credit for Eligible Dividends"),
    ("foreignNonBusinessIncome",                 "25", "Foreign Non-Business Income"),
    ("otherIncome",                              "26", "Other Income"),
    ("returnOfCapital",                          "42", "Return of Capital"),
    ("foreignNonBusinessIncomeTaxPaid",          "34", "Foreign Non-Business Income Tax Paid"),
    ("actualAmountOfNonEligibleDividends",       "23", "Actual Amount of Non-Eligible Dividends"),
    ("taxableAmountOfNonEligibleDividends",      "32", "Taxable Amount of Non-Eligible Dividends"),
    ("dividendTaxCreditForNonEligibleDividends", "39", "Dividend Tax Credit for Non-Eligible Dividends"),
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def fmt2(v):
    return f"${v:,.2f}"

def fmt4(v):
    return f"${v:,.4f}"

def parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date()

# ── Load distributions ────────────────────────────────────────────────────────

def load_distributions(dist_dir):
    """
    Returns dict: fund_name -> list of distribution dicts.
    Each dict has at minimum: recordDate (date), paymentDate (date), total (float),
    plus component fields.
    """
    funds = {}
    for path in sorted(glob.glob(os.path.join(dist_dir, "*.json"))):
        data = load_json(path)
        name = data.get("name", os.path.splitext(os.path.basename(path))[0])
        dists = []
        for d in data.get("distributions", []):
            parsed = dict(d)
            parsed["recordDate"]  = parse_date(d["recordDate"])
            parsed["paymentDate"] = parse_date(d["paymentDate"])
            dists.append(parsed)
        funds[name] = dists
    return funds

# ── Load assets (shares per account) ─────────────────────────────────────────

def load_assets(assets_dir):
    """
    Returns dict: account_name -> {fund_name -> {date -> shares}}
    """
    accounts = {}
    for path in sorted(glob.glob(os.path.join(assets_dir, "*.json"))):
        data = load_json(path)
        account = data.get("account", os.path.splitext(os.path.basename(path))[0])
        shares_map = {}
        for key, entries in data.items():
            if key == "account":
                continue
            date_shares = {}
            for entry in entries:
                d = parse_date(entry["date"])
                date_shares[d] = entry["ownedShares"]
            shares_map[key] = date_shares
        accounts[account] = shares_map
    return accounts

# ── Core computation ──────────────────────────────────────────────────────────

def compute_t3(funds, accounts, tax_rates):
    """
    Returns:
        results : dict  account -> {field -> total_dollars}
        details : dict  account -> list of strings (verbose log)
    """
    elig_grossup   = tax_rates["eligible_div_grossup"]
    elig_credit    = tax_rates["eligible_div_tax_credit"]
    ne_grossup     = tax_rates["non_eligible_div_grossup"]
    ne_credit      = tax_rates["non_eligible_div_tax_credit"]

    results = {}
    details = {}

    for account, shares_by_fund in sorted(accounts.items()):
        acc_totals  = defaultdict(float)
        acc_details = []
        acc_details.append(f"\nComputing T3 for account '{account}':")

        for fund in sorted(funds.keys()):
            dists       = funds[fund]
            date_map    = shares_by_fund.get(fund, {})
            fund_totals = defaultdict(float)
            acc_details.append(f"\n   Asset '{fund}':")

            for dist in dists:
                rec_date  = dist["recordDate"]
                pay_date  = dist["paymentDate"]
                shares    = date_map.get(rec_date, 0.0)
                total     = dist.get("total", 0.0)
                non_cash  = dist.get("nonCashCapitalGains", dist.get("capitalGains", 0.0)) if not dist.get("capitalGainsDistributedAsCash") else dist.get("nonCashCapitalGains", 0.0)
                cash_rcvd = shares * (total - non_cash)

                acc_details.append(f"\n      Distribution record={rec_date} payment={pay_date}:")

                if shares == 0:
                    acc_details.append(f"         (0 shares held — skipped)")
                    continue

                acc_details.append(f"         Shares held: {shares:,.4f}")
                acc_details.append(f"         You should have received {fmt2(cash_rcvd)} on or about {pay_date}")
                acc_details.append(f"         Breakdown:")

                for field, (box, label) in FIELD_TO_BOX.items():
                    val = dist.get(field, 0.0)
                    if val == 0.0:
                        continue
                    dollar = shares * val
                    fund_totals[field] += dollar
                    acc_details.append(
                        f"            Box {box:>2} {label:<45}: {shares:>12,.4f} × {fmt4(val)} = {fmt2(dollar)}"
                    )

            # Accumulate fund totals into account totals
            if any(v != 0 for v in fund_totals.values()):
                acc_details.append(f"\n      '{fund}' subtotals:")
                for field, total_val in fund_totals.items():
                    box, label = FIELD_TO_BOX[field]
                    acc_details.append(
                        f"         Box {box:>2} {label:<45}: {fmt2(total_val)}"
                    )
                    acc_totals[field] += total_val

        # Derived boxes — eligible dividends (computed from the rounded values)
        elig_div = acc_totals.get("actualAmountOfEligibleDividends", 0.0)
        if elig_div:
            elig_div_rounded = round(elig_div, 2)
            taxable = round(elig_div_rounded * elig_grossup, 2)
            acc_totals["taxableAmountOfEligibleDividends"]      = taxable
            acc_totals["dividendTaxCreditForEligibleDividends"] = round(taxable * elig_credit, 2)

        # Derived boxes — non-eligible dividends (computed from the rounded values)
        ne_div = acc_totals.get("actualAmountOfNonEligibleDividends", 0.0)
        if ne_div:
            ne_div_rounded = round(ne_div, 2)
            taxable_ne = round(ne_div_rounded * ne_grossup, 2)
            acc_totals["taxableAmountOfNonEligibleDividends"]      = taxable_ne
            acc_totals["dividendTaxCreditForNonEligibleDividends"] = round(taxable_ne * ne_credit, 2)

        results[account] = dict(acc_totals)
        details[account] = acc_details

    return results, details

# ── Output formatting ─────────────────────────────────────────────────────────

def print_results(results, details, output_file=None):
    lines = []

    for account in sorted(results.keys()):
        for line in details[account]:
            lines.append(line)

    lines.append("\n" + "=" * 60)
    lines.append("RESULTS SUMMARY")
    lines.append("=" * 60)

    for account, totals in sorted(results.items()):
        lines.append(f"\nT3 for '{account}':")
        for field, box, label in SUMMARY_FIELDS:
            val = totals.get(field, 0.0)
            if val == 0.0:
                continue
            lines.append(f"   Box {box:>2}  {label:<50}: {fmt2(val)}")

    output = "\n".join(lines)
    print(output)

    if output_file:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"\nResults saved to: {output_file}")

# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    cfg  = load_config(args.config)

    # Apply CLI overrides
    if args.year:
        cfg["tax_year"] = args.year
        _set_derived_paths(cfg, args.year)

    tax_year   = cfg["tax_year"]
    dist_dir   = cfg["dist_dir"]
    assets_dir = cfg["assets_dir"]
    output_txt = cfg["output_txt"]
    tax_rates  = cfg["tax_rates"]

    # Verify gross-up rates for the tax year
    check_rates(tax_year, tax_rates)

    print(f"Tax year  : {tax_year}")
    print(f"Dist dir  : {dist_dir}")
    print(f"Assets dir: {assets_dir}")
    print()

    # ── Prerequisite checks ───────────────────────────────────────────────────
    if not os.path.isdir(dist_dir):
        print(f"ERROR: Distributions folder not found: {dist_dir}")
        print("       Run Step 1 (parse_t3_pdfs.py) first.")
        sys.exit(1)
    dist_jsons = glob.glob(os.path.join(dist_dir, "*.json"))
    if not dist_jsons:
        print(f"ERROR: No distribution JSONs found in: {dist_dir}")
        print("       Run Step 1 (parse_t3_pdfs.py) first.")
        sys.exit(1)
    if not os.path.isdir(assets_dir):
        print(f"ERROR: Assets folder not found: {assets_dir}")
        print("       Run Step 2 (build_assets.py) first.")
        sys.exit(1)
    asset_jsons = glob.glob(os.path.join(assets_dir, "*.json"))
    if not asset_jsons:
        print(f"ERROR: No account JSONs found in: {assets_dir}")
        print("       Run Step 2 (build_assets.py) first.")
        sys.exit(1)

    print(f"Loading distributions from: {dist_dir}")
    funds = load_distributions(dist_dir)
    if args.funds:
        unknown = [f for f in args.funds if f not in funds]
        if unknown:
            print(f"WARNING: Fund(s) not found in distributions folder: {', '.join(unknown)}")
        funds = {k: v for k, v in funds.items() if k in args.funds}
    print(f"  Loaded {len(funds)} fund(s): {', '.join(sorted(funds.keys()))}")
    print(f"  Total distributions: {sum(len(v) for v in funds.values())}")

    print(f"\nLoading assets from: {assets_dir}")
    accounts = load_assets(assets_dir)
    print(f"  Loaded {len(accounts)} account(s): {', '.join(sorted(accounts.keys()))}")

    print("\nBeginning computation...")
    results, details = compute_t3(funds, accounts, tax_rates)

    print_results(results, details, output_txt)


if __name__ == "__main__":
    main()
