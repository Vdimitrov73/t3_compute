"""
t3_colors.py
------------
Shared ANSI colour utilities for the T3 pipeline.
No external dependencies. Import from any pipeline script.
"""

import os
import re
import sys

# Palette — standard 8 codes, renders on Windows Terminal / VS Code / any modern terminal
RESET  = "\033[0m"
BOLD   = "\033[1m"
DIM    = "\033[2m"
RED    = "\033[91m"    # errors
CYAN   = "\033[96m"    # box numbers / asset names / menu numbers
GREEN  = "\033[92m"    # dollar amounts / success messages
YELLOW = "\033[93m"    # section headers / warnings
WHITE  = "\033[97m"    # "you should have received" line
GREY   = "\033[90m"    # per-unit factors (x $0.xxxxx)


def use_color() -> bool:
    """Return True when the terminal is likely to render ANSI codes."""
    if os.environ.get("NO_COLOR"):
        return False
    if sys.stdout.isatty():
        return True
    return bool(os.environ.get("WT_SESSION") or os.environ.get("TERM_PROGRAM"))


def strip_ansi(text: str) -> str:
    return re.sub(r"\033\[[0-9;]*m", "", text)


def c(code: str, text: str) -> str:
    """Wrap text in an ANSI code + RESET, only if color is supported."""
    return f"{code}{text}{RESET}" if use_color() else text


# -- Compiled regexes used by compute_t3.print_results ------------------------
RE_DOLLARS  = re.compile(r"(\$[\d,]+\.\d{2})")
RE_BOX      = re.compile(r"(Box\s+\d+)")
RE_PERUNIT  = re.compile(r"(x \$[\d.]+)")
RE_ASSET    = re.compile(r"(Asset '.*?':)")
RE_RECEIVED = re.compile(r"(You should have received)")
RE_SUBTOTAL = re.compile(r"('.+?' subtotals:)")


def colorize_detail(line: str) -> str:
    """Apply subtle colour to a single compute_t3 detail line."""
    if RE_ASSET.search(line):
        line = RE_ASSET.sub(f"{BOLD}{CYAN}\\1{RESET}", line)
    if RE_SUBTOTAL.search(line):
        line = RE_SUBTOTAL.sub(f"{BOLD}\\1{RESET}", line)
    if RE_RECEIVED.search(line):
        line = RE_RECEIVED.sub(f"{WHITE}\\1{RESET}", line)
        line = RE_DOLLARS.sub(f"{GREEN}\\1{RESET}", line)
    elif RE_BOX.search(line):
        line = RE_BOX.sub(f"{CYAN}\\1{RESET}", line)
        line = RE_PERUNIT.sub(f"{GREY}\\1{RESET}", line)
        matches = list(RE_DOLLARS.finditer(line))
        if matches:
            m = matches[-1]
            line = line[:m.start()] + GREEN + m.group() + RESET + line[m.end():]
    return line