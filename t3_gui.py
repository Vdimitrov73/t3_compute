"""
t3_gui.py
---------
Tkinter GUI shell for the T3 Compute pipeline.

Replaces the interactive CLI menu and prompts with a Windows UI.
All pipeline logic stays in run_t3.py / the four step scripts — this
file only handles display, user input, and threading.

Launch with:
    pythonw t3_gui.py          (no console window)
    python   t3_gui.py         (with console, useful for debugging)

Requires the same directory layout as run_t3.py:
    t3_gui.py
    run_t3.py
    parse_t3_pdfs.py
    update_acb.py
    build_assets.py
    compute_t3.py
    config.json  (created by first-run wizard or manually)
    account_periods.json
    funds.json
"""

import sys
import os
import json
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import date

# ── Make pipeline scripts importable ─────────────────────────────────────────
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

import run_t3 as _run_t3

# ── Colour palette ────────────────────────────────────────────────────────────
BG          = "#1a1f2e"   # dark navy background
BG2         = "#232940"   # slightly lighter panel background
BORDER      = "#2e3650"   # panel border
ACCENT      = "#4f9cf9"   # blue accent
ACCENT2     = "#3a7bd5"   # darker blue for hover
SUCCESS     = "#4caf82"   # green
WARNING     = "#e8a838"   # amber
ERROR       = "#e05c5c"   # red
TEXT        = "#e8ecf4"   # main text
TEXT_DIM    = "#7a85a0"   # subdued text
STEP_ACTIVE = "#2a3452"   # step button active bg
STEP_HOVER  = "#2e3a5a"   # step button hover bg
LOG_BG      = "#111520"   # log area background
LOG_FG      = "#c8d0e0"   # log default text

FONT_MAIN   = ("Segoe UI", 10)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_SMALL  = ("Segoe UI", 9)
FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_STEP   = ("Segoe UI", 10)
FONT_LOG    = ("Consolas", 9)

# ── Stdout redirector ─────────────────────────────────────────────────────────

class _QueueWriter:
    """
    Redirect stdout/stderr writes into a thread-safe queue.
    Implements the full file-like interface so libraries that call
    isatty(), readable(), writable(), etc. don't crash.
    """
    def __init__(self, q):
        self._q = q

    def write(self, text):
        if text:
            self._q.put(("log", text))

    def flush(self):
        pass

    def isatty(self):
        return False

    def readable(self):
        return False

    def writable(self):
        return True

    def seekable(self):
        return False

    @property
    def encoding(self):
        return "utf-8"

    @property
    def errors(self):
        return "replace"


# ── Setup Wizard Dialog ───────────────────────────────────────────────────────

class SetupWizard(tk.Toplevel):
    """
    Full setup wizard — 4 pages:
      Page 1 — Basic config  : tax year, base folder, ACB spreadsheet
      Page 2 — Funds         : fund tickers + calc method overrides
      Page 3 — Accounts      : define brokerage account names
      Page 4 — Periods       : per-fund account assignment + mid-year transfers

    Pre-populates from existing config files so it doubles as an edit dialog.
    Writes config.json, account_periods.json, and funds.json on Finish.
    """

    def __init__(self, parent, script_dir, config_dir,
                 config_path, periods_path, funds_path, on_complete):
        super().__init__(parent)
        self.title("T3 Compute — Setup")
        self.resizable(True, True)
        self.configure(bg=BG)
        self.grab_set()   # modal

        self._script_dir   = script_dir
        self._config_dir   = config_dir
        self._config_path  = config_path
        self._periods_path = periods_path
        self._funds_path   = funds_path
        self._on_complete  = on_complete

        # ── Load existing config (pre-populate fields if files exist) ─────────
        existing_cfg     = self._load_json(config_path)
        existing_periods = self._load_json(periods_path)
        existing_funds   = self._load_json(funds_path)

        # ── Page 1 vars ───────────────────────────────────────────────────────
        default_year = existing_cfg.get("tax_year", str(date.today().year - 1))
        default_base = existing_cfg.get("base_dir",
                       os.path.join(os.path.expanduser("~"), "Documents", "Tax Documents"))
        default_acb  = existing_cfg.get("acb_spreadsheet",
                       os.path.join(default_base, "acb_worksheet.xlsx"))

        self._year = tk.StringVar(value=default_year)
        self._base = tk.StringVar(value=default_base)
        self._acb  = tk.StringVar(value=default_acb)

        # ── Page 2 vars ───────────────────────────────────────────────────────
        default_funds = " ".join(existing_cfg.get("funds", ["VBAL", "CPD", "VRE", "ZCN"]))
        self._funds_var = tk.StringVar(value=default_funds)
        # calc_method_override per fund: stored as dict {ticker: StringVar}
        self._calc_overrides = {}   # built in page 2
        self._existing_funds_meta = existing_funds  # preserve unknown funds

        # ── Page 3 vars ───────────────────────────────────────────────────────
        self._num_accounts  = tk.IntVar(value=1)
        self._account_vars  = []   # list of StringVar for account names
        self._account_frame = None

        # ── Page 4 state ──────────────────────────────────────────────────────
        # Periods: dict { fund_ticker: list of {"start": "YYYY-MM-DD", "account": "NAME"} }
        # Built fresh on page 4 from existing_periods or defaults.
        self._existing_periods = existing_periods
        self._period_widgets   = {}   # fund → list of row dicts for the UI
        self._page4_frame      = None

        self._build_page1()
        self._center()

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _load_json(path):
        if os.path.exists(path):
            try:
                with open(path) as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _center(self):
        self.update_idletasks()
        w = max(self.winfo_width(),  620)
        h = max(self.winfo_height(), 400)
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _clear(self):
        for w in self.winfo_children():
            w.destroy()

    def _lbl(self, parent, text, dim=False, bold=False, fg=None, **kw):
        if fg is None:
            fg = TEXT_DIM if dim else TEXT
        font = FONT_BOLD if bold else FONT_MAIN
        return tk.Label(parent, text=text, bg=BG, fg=fg, font=font, **kw)

    def _entry(self, parent, textvariable, width=46):
        return tk.Entry(parent, textvariable=textvariable, width=width,
                        bg=BG2, fg=TEXT, insertbackground=TEXT,
                        relief="flat", font=FONT_MAIN,
                        highlightthickness=1, highlightbackground=BORDER,
                        highlightcolor=ACCENT)

    def _btn(self, parent, text, command, accent=False, width=10, small=False):
        bg  = ACCENT  if accent else BG2
        fg  = "#ffffff" if accent else TEXT
        ab  = ACCENT2 if accent else STEP_HOVER
        fnt = FONT_SMALL if small else FONT_BOLD
        return tk.Button(parent, text=text, command=command,
                         bg=bg, fg=fg, font=fnt,
                         relief="flat", cursor="hand2",
                         padx=10, pady=5, width=width,
                         activebackground=ab, activeforeground=fg, bd=0)

    def _section(self, parent, title):
        f = tk.Frame(parent, bg=BG)
        tk.Label(f, text=title, bg=BG, fg=ACCENT, font=FONT_BOLD).pack(anchor="w")
        tk.Frame(f, bg=BORDER, height=1).pack(fill="x", pady=(2, 8))
        return f

    def _browse_folder(self, var, sync_acb=False):
        d = filedialog.askdirectory(
                parent=self,
                initialdir=var.get() or os.path.expanduser("~"))
        if d:
            var.set(d)
            if sync_acb:
                cur = self._acb.get()
                if not cur or os.path.basename(cur) == "acb_worksheet.xlsx":
                    self._acb.set(os.path.join(d, "acb_worksheet.xlsx"))

    def _browse_file(self, var):
        path = filedialog.askopenfilename(
            parent=self,
            initialdir=os.path.dirname(var.get()) or os.path.expanduser("~"),
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if path:
            var.set(path)

    def _nav_buttons(self, parent, back_cmd, next_cmd,
                     next_text="Next →", back_text="← Back", next_accent=True, show_cancel=True):
        f = tk.Frame(parent, bg=BG)
        if show_cancel:
            self._btn(f, "Cancel", self.destroy, width=8).pack(side="left")
        self._btn(f, back_text, back_cmd, width=9).pack(
            side="left", padx=((8, 0) if show_cancel else (0, 0)))
        self._btn(f, next_text, next_cmd, width=10,
                  accent=next_accent).pack(side="right")
        return f

    def _page_header(self, parent, title, subtitle, step, total=4):
        tk.Label(parent, text=f"Step {step} of {total} — {title}",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(anchor="w")
        tk.Label(parent, text=subtitle, bg=BG, fg=TEXT_DIM,
                 font=FONT_SMALL).pack(anchor="w", pady=(3, 18))

    # ═══════════════════════════════════════════════════════════════
    # PAGE 1 — Basic config
    # ═══════════════════════════════════════════════════════════════

    def _build_page1(self):
        self._clear()
        outer = tk.Frame(self, bg=BG, padx=28, pady=22)
        outer.pack(fill="both", expand=True)

        self._page_header(outer, "Basic Configuration",
                          "Tax year, T3 source folder, and ACB spreadsheet path.", 1)

        # Tax year
        s1 = self._section(outer, "Tax Year")
        s1.pack(fill="x", pady=(0, 14))
        row = tk.Frame(s1, bg=BG); row.pack(fill="x")
        self._lbl(row, "Year to process:").pack(side="left")
        self._entry(row, self._year, width=8).pack(side="left", padx=(8, 0))
        self._lbl(row, "(CDS releases T3 slips in February for the prior year)",
                  dim=True).pack(side="left", padx=(12, 0))

        # Base directory
        s2 = self._section(outer, "Base Tax Documents Folder")
        s2.pack(fill="x", pady=(0, 14))
        self._lbl(s2, "Select the base folder that contains your yearly tax folders.").pack(anchor="w")
        r2 = tk.Frame(s2, bg=BG); r2.pack(fill="x", pady=(6, 0))
        self._entry(r2, self._base, width=48).pack(side="left")
        self._btn(r2, "Browse…",
                  lambda: self._browse_folder(self._base, sync_acb=True),
                  width=8).pack(side="left", padx=(6, 0))

        # Live path preview — updates as user types year or folder
        self._path_preview = tk.StringVar()
        def _update_preview(*_):
            yr   = self._year.get().strip() or "<year>"
            base = self._base.get().strip() or "<folder>"
            self._path_preview.set(f"Pipeline expects T3 files in:  {base}\\{yr}\\")
        self._year.trace_add("write", _update_preview)
        self._base.trace_add("write", _update_preview)
        _update_preview()
        tk.Label(s2, textvariable=self._path_preview,
                 bg=BG, fg=ACCENT, font=FONT_SMALL).pack(anchor="w", pady=(6, 0))

        # ACB spreadsheet
        s3 = self._section(outer, "ACB Spreadsheet")
        s3.pack(fill="x", pady=(0, 22))
        self._lbl(s3, "Full path to your acb_worksheet.xlsx.\n"
                      "If it doesn't exist yet the template will be copied there automatically.").pack(anchor="w")
        r3 = tk.Frame(s3, bg=BG); r3.pack(fill="x", pady=(6, 0))
        self._entry(r3, self._acb, width=64).pack(side="left")
        self._btn(r3, "Browse…",
                  lambda: self._browse_file(self._acb),
                  width=8).pack(side="left", padx=(6, 0))

        self._nav_buttons(outer, self.destroy, self._p1_next,
                          back_text="Cancel", show_cancel=False).pack(anchor="e", pady=(4, 0))

    def _p1_next(self):
        yr = self._year.get().strip()
        if not yr.isdigit() or not (2019 <= int(yr) <= date.today().year):
            messagebox.showerror("Invalid year",
                f"Please enter a year between 2019 and {date.today().year}.", parent=self)
            return
        if not self._base.get().strip():
            messagebox.showerror("Missing folder",
                "Please enter a T3 files folder.", parent=self)
            return
        if not self._acb.get().strip():
            messagebox.showerror("Missing spreadsheet",
                "Please enter the ACB spreadsheet path.", parent=self)
            return
        self._build_page2()
        self._center()

    # ═══════════════════════════════════════════════════════════════
    # PAGE 2 — Funds
    # ═══════════════════════════════════════════════════════════════

    def _build_page2(self):
        self._clear()
        outer = tk.Frame(self, bg=BG, padx=28, pady=22)
        outer.pack(fill="both", expand=True)

        self._page_header(outer, "Fund Tickers",
                          "Which ETFs do you hold? Tickers must match sheet names in your ACB spreadsheet.", 2)

        s1 = self._section(outer, "Tickers (space-separated)")
        s1.pack(fill="x", pady=(0, 6))
        self._entry(s1, self._funds_var, width=50).pack(anchor="w", pady=(0, 4))
        self._lbl(s1, "Example: VBAL CPD VRE ZCN", dim=True).pack(anchor="w")

        # Calc method overrides — shown after tickers entered
        s2 = self._section(outer, "PDF Parsing Override (advanced — usually leave blank)")
        s2.pack(fill="x", pady=(14, 0))
        self._lbl(s2,
            "Most funds auto-detect correctly. Only set an override if Step 1 produces\n"
            "wrong values. RATE = per-unit dollar amounts. PERCENT = percentage of total.",
            dim=True).pack(anchor="w", pady=(0, 8))

        self._override_frame = tk.Frame(s2, bg=BG)
        self._override_frame.pack(fill="x")
        self._funds_var.trace_add("write", lambda *_: self._rebuild_overrides())
        self._rebuild_overrides()

        self._nav_buttons(outer, self._build_page1, self._p2_next).pack(
            anchor="e", pady=(16, 0))

    def _rebuild_overrides(self):
        for w in self._override_frame.winfo_children():
            w.destroy()
        tickers = [t.strip().upper() for t in self._funds_var.get().split() if t.strip()]
        # Preserve existing StringVar values when rebuilding
        for ticker in tickers:
            if ticker not in self._calc_overrides:
                existing_val = (self._existing_funds_meta.get(ticker, {}) or {}).get(
                    "calc_method_override") or "AUTO"
                self._calc_overrides[ticker] = tk.StringVar(value=existing_val)
            row = tk.Frame(self._override_frame, bg=BG)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=f"{ticker}:", bg=BG, fg=TEXT, font=FONT_SMALL,
                     width=8, anchor="w").pack(side="left")
            for opt, lbl in [("AUTO", "Auto"), ("RATE", "RATE"), ("PERCENT", "PERCENT")]:
                tk.Radiobutton(row, text=lbl, variable=self._calc_overrides[ticker],
                               value=opt, bg=BG, fg=TEXT, selectcolor=BG2,
                               activebackground=BG, font=FONT_SMALL).pack(side="left", padx=(4, 8))

    def _p2_next(self):
        funds = [t.strip().upper() for t in self._funds_var.get().split() if t.strip()]
        if not funds:
            messagebox.showerror("Missing funds",
                "Please enter at least one fund ticker.", parent=self)
            return
        self._build_page3()
        self._center()

    # ═══════════════════════════════════════════════════════════════
    # PAGE 3 — Accounts
    # ═══════════════════════════════════════════════════════════════

    def _build_page3(self):
        self._clear()
        outer = tk.Frame(self, bg=BG, padx=28, pady=22)
        outer.pack(fill="both", expand=True)

        self._page_header(outer, "Brokerage Accounts",
                          "Define the accounts that held your ETFs during the tax year.", 3)

        s = self._section(outer, "Number of brokerage accounts")
        s.pack(fill="x", pady=(0, 14))

        # Detect existing account names from periods to pre-populate
        yr = self._year.get().strip()
        existing_accts = []
        if yr in self._existing_periods:
            seen = []
            for fund_periods in self._existing_periods[yr].values():
                for p in fund_periods:
                    a = p.get("account", "")
                    if a and a not in seen:
                        seen.append(a)
            existing_accts = seen

        n_existing = max(len(existing_accts), 1)
        self._num_accounts.set(min(n_existing, 4))

        radio_row = tk.Frame(s, bg=BG); radio_row.pack(fill="x")
        for n in (1, 2, 3, 4):
            tk.Radiobutton(radio_row, text=str(n), variable=self._num_accounts, value=n,
                           bg=BG, fg=TEXT, selectcolor=BG2, activebackground=BG,
                           font=FONT_MAIN,
                           command=self._rebuild_acct_names).pack(side="left", padx=(0, 12))

        self._account_frame = tk.Frame(outer, bg=BG)
        self._account_frame.pack(fill="x", pady=(0, 14))
        self._account_name_defaults = existing_accts
        self._rebuild_acct_names()

        self._nav_buttons(outer, self._build_page2, self._p3_next).pack(
            anchor="e", pady=(4, 0))

    def _rebuild_acct_names(self):
        for w in self._account_frame.winfo_children():
            w.destroy()
        n = self._num_accounts.get()
        # Preserve existing StringVar values
        defaults = self._account_name_defaults or ["TDDI", "Wealthsimple", "DISNAT", "Questrade"]
        while len(self._account_vars) < n:
            idx = len(self._account_vars)
            val = defaults[idx] if idx < len(defaults) else f"Account{idx+1}"
            self._account_vars.append(tk.StringVar(value=val))
        s = self._section(self._account_frame, "Account Names")
        s.pack(fill="x")
        for i in range(n):
            row = tk.Frame(s, bg=BG); row.pack(fill="x", pady=3)
            tk.Label(row, text=f"Account {i+1}:", bg=BG, fg=TEXT,
                     font=FONT_MAIN, width=11, anchor="w").pack(side="left")
            self._entry(row, self._account_vars[i], width=28).pack(side="left", padx=(4, 0))

    def _p3_next(self):
        n = self._num_accounts.get()
        accounts = [self._account_vars[i].get().strip() or f"Account{i+1}"
                    for i in range(n)]
        if len(set(accounts)) != len(accounts):
            messagebox.showerror("Duplicate names",
                "Account names must be unique.", parent=self)
            return
        self._build_page4(accounts)
        self._center()

    # ═══════════════════════════════════════════════════════════════
    # PAGE 4 — Account periods per fund
    # ═══════════════════════════════════════════════════════════════

    def _build_page4(self, accounts):
        self._clear()
        self._current_accounts = accounts
        year = self._year.get().strip()
        funds = [t.strip().upper() for t in self._funds_var.get().split() if t.strip()]

        outer = tk.Frame(self, bg=BG, padx=28, pady=22)
        outer.pack(fill="both", expand=True)

        self._page_header(outer, "Account Periods",
            "For each fund, specify which account held it and when.\n"
            "Add a transfer row if a fund moved between accounts during the year.", 4)

        # Scrollable area for fund sections
        canvas_frame = tk.Frame(outer, bg=BG)
        canvas_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=BG)
        canvas_win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_resize(e):
            canvas.itemconfig(canvas_win, width=e.width)
        canvas.bind("<Configure>", _on_resize)

        def _on_frame_resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_frame_resize)

        # Mouse wheel scroll
        def _on_wheel(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        # canvas.bind_all("<MouseWheel>", _on_wheel)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self._period_widgets = {}

        for fund in funds:
            self._build_fund_section(inner, fund, year, accounts)

        self._nav_buttons(outer, self._p4_back, self._finish,
                          next_text="Finish ✓").pack(anchor="e", pady=(10, 0))

    def _build_fund_section(self, parent, fund, year, accounts):
        """Build the period-assignment UI for one fund."""
        # Load existing periods for this fund/year
        existing = []
        if year in self._existing_periods:
            existing = list(self._existing_periods[year].get(fund, []))

        if not existing:
            existing = [{"start": f"{year}-01-01", "account": accounts[0]}]

        s = self._section(parent, f"{fund}")
        s.pack(fill="x", pady=(0, 14), padx=4)

        rows_frame = tk.Frame(s, bg=BG)
        rows_frame.pack(fill="x")

        # Column headers
        hdr = tk.Frame(rows_frame, bg=BG)
        hdr.pack(fill="x", pady=(0, 2))
        tk.Label(hdr, text="From date (YYYY-MM-DD)", bg=BG, fg=TEXT_DIM,
                 font=FONT_SMALL, width=22, anchor="w").pack(side="left")
        tk.Label(hdr, text="Account", bg=BG, fg=TEXT_DIM,
                 font=FONT_SMALL, width=22, anchor="w").pack(side="left", padx=(8, 0))

        period_rows = []
        self._period_widgets[fund] = period_rows

        def add_row(start="", account=None, can_delete=True):
            row_frame = tk.Frame(rows_frame, bg=BG)
            row_frame.pack(fill="x", pady=2)

            date_var = tk.StringVar(value=start)
            date_entry = self._entry(row_frame, date_var, width=20)
            date_entry.pack(side="left")

            # Lock the first row's date to Jan 1 — can't be changed
            if not can_delete:
                date_entry.configure(state="disabled", disabledforeground=TEXT_DIM,
                                     disabledbackground=BG2)

            acct_var = tk.StringVar(value=account or accounts[0])
            acct_menu = tk.OptionMenu(row_frame, acct_var, *accounts)
            acct_menu.configure(bg=BG2, fg=TEXT, font=FONT_SMALL,
                                activebackground=STEP_HOVER, activeforeground=TEXT,
                                relief="flat", highlightthickness=0, bd=0,
                                indicatoron=True, width=18)
            acct_menu["menu"].configure(bg=BG2, fg=TEXT, font=FONT_SMALL,
                                        activebackground=ACCENT, activeforeground="#ffffff")
            acct_menu.pack(side="left", padx=(8, 0))

            row_dict = {"date_var": date_var, "acct_var": acct_var,
                        "frame": row_frame, "deletable": can_delete}
            period_rows.append(row_dict)

            if can_delete:
                def del_row(rd=row_dict):
                    rd["frame"].destroy()
                    period_rows.remove(rd)
                self._btn(row_frame, "✕", del_row, width=2, small=True).pack(
                    side="left", padx=(6, 0))

        # First row — always present, date locked to Jan 1
        first = existing[0]
        add_row(start=f"{year}-01-01",
                account=first.get("account", accounts[0]),
                can_delete=False)

        # Subsequent rows (transfers)
        for p in existing[1:]:
            add_row(start=p.get("start", ""), account=p.get("account", accounts[0]),
                    can_delete=True)

        # "Add transfer" button
        def add_transfer():
            add_row(start="", account=accounts[0], can_delete=True)
        self._btn(s, "+ Add transfer", add_transfer, width=14, small=True).pack(
            anchor="w", pady=(6, 0))

    def _p4_back(self):
        n = self._num_accounts.get()
        accounts = [self._account_vars[i].get().strip() or f"Account{i+1}" for i in range(n)]
        self._build_page3()
        self._center()

    # ═══════════════════════════════════════════════════════════════
    # FINISH — validate and write all config files
    # ═══════════════════════════════════════════════════════════════

    def _finish(self):
        year     = self._year.get().strip()
        base_dir = self._base.get().strip()
        acb_path = self._acb.get().strip()
        funds    = [t.strip().upper() for t in self._funds_var.get().split() if t.strip()]
        n        = self._num_accounts.get()
        accounts = [self._account_vars[i].get().strip() or f"Account{i+1}" for i in range(n)]

        # ── Validate all period rows ──────────────────────────────────────────
        periods_out = {}
        for fund in funds:
            rows = self._period_widgets.get(fund, [])
            fund_periods = []
            for i, rd in enumerate(rows):
                start   = rd["date_var"].get().strip()
                account = rd["acct_var"].get().strip()

                # First row date is locked — always valid
                if i == 0:
                    start = f"{year}-01-01"
                else:
                    try:
                        date.fromisoformat(start)
                    except (ValueError, TypeError):
                        messagebox.showerror("Invalid date",
                            f"{fund}: transfer date '{start}' is not valid.\n"
                            f"Please use YYYY-MM-DD format (e.g. {year}-08-11).",
                            parent=self)
                        return
                    if not start.startswith(year):
                        messagebox.showerror("Wrong year",
                            f"{fund}: transfer date '{start}' is not in {year}.",
                            parent=self)
                        return

                if account not in accounts:
                    messagebox.showerror("Invalid account",
                        f"{fund}: account '{account}' is not in your account list.",
                        parent=self)
                    return
                fund_periods.append({"start": start, "account": account})

            # Sort by date
            fund_periods.sort(key=lambda p: p["start"])
            periods_out[fund] = fund_periods

        # ── Build full periods dict (merge — preserve other years) ────────────
        merged_periods = dict(self._existing_periods)
        merged_periods[year] = periods_out

        # ── Build funds.json (preserve existing, add/update current funds) ────
        merged_funds = dict(self._existing_funds_meta)
        for fund in funds:
            override = self._calc_overrides.get(fund)
            override_val = override.get().strip() if override else "AUTO"
            merged_funds[fund] = {
                "calc_method_override": override_val if override_val not in ("", "AUTO") else None
            }
        # ── Write all three files ─────────────────────────────────────────────
        config = {
            "tax_year":        year,
            "base_dir":        base_dir,
            "acb_spreadsheet": acb_path,
            "funds":           funds,
            "tax_rates":       _run_t3._KNOWN_TAX_RATES,
            "col_indices":     {"date": 0, "share_balance": 6},
        }
        try:
            with open(self._config_path, "w") as f:
                json.dump(config, f, indent=2)
            with open(self._periods_path, "w") as f:
                json.dump(merged_periods, f, indent=2)
            with open(self._funds_path, "w") as f:
                json.dump(merged_funds, f, indent=2)
        except Exception as ex:
            messagebox.showerror("Write error",
                f"Could not save config files:\n{ex}", parent=self)
            return

        # ── Copy ACB template if not present ──────────────────────────────────
        template = os.path.join(self._script_dir, "acb_worksheet_template.xlsx")
        acb_copied = False
        if not os.path.exists(acb_path) and os.path.exists(template):
            try:
                import shutil
                os.makedirs(os.path.dirname(os.path.abspath(acb_path)), exist_ok=True)
                shutil.copy2(template, acb_path)
                acb_copied = True
            except Exception:
                pass

        # ── Summary messagebox ────────────────────────────────────────────────
        use_pdf = int(year) >= 2025
        ext = "pdf" if use_pdf else "xls"
        file_list = "\n".join(f"  {f}_T3_{year}.{ext}" for f in funds)
        acct_summary = "\n".join(
            f"  {f}: " + ", ".join(
                f"{p['account']} from {p['start']}"
                for p in periods_out.get(f, [])
            )
            for f in funds
        )

        msg = (
            f"Config saved to:\n  {self._config_dir}\n\n"
            f"Tax year : {year}\n"
            f"Funds    : {', '.join(funds)}\n"
            f"Accounts : {', '.join(accounts)}\n\n"
            f"Account assignments:\n{acct_summary}\n\n"
            f"Next: download your CDS T3 files from\n"
            f"  https://ctbsext.posttrade.cds.ca/ctbsExt/\n"
            f"and save them in:\n"
            f"  {os.path.join(base_dir, year)}\\\n"
            f"{file_list}"
        )
        if acb_copied:
            msg += f"\n\nACB template copied to:\n  {acb_path}"

        messagebox.showinfo("Setup complete", msg, parent=self)
        self.destroy()
        self._on_complete()



# ── Main application window ───────────────────────────────────────────────────

class T3App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("T3 Compute")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(860, 720)

        self._script_dir = _SCRIPT_DIR
        self._config_dir = _run_t3.get_config_dir(_SCRIPT_DIR)
        if self._config_dir not in sys.path:
            sys.path.insert(0, self._config_dir)

        self._config_path  = os.path.join(self._config_dir, "config.json")
        self._periods_path = os.path.join(self._config_dir, "account_periods.json")
        self._funds_path   = os.path.join(self._config_dir, "funds.json")

        # Runtime state
        self._running  = False
        self._ran_step4 = False
        self._log_q    = queue.Queue()
        self._args     = self._make_args()

        # Build UI
        self._build_ui()
        self._center()
        self._poll_log()

        # Check setup on first launch
        self.after(100, self._check_setup)

        # Standard keyboard shortcuts
        self.bind("<Escape>", lambda e: self._on_escape())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _make_args(self):
        """Create a minimal args namespace matching what run_step() expects."""
        import argparse
        a = argparse.Namespace(
            config  = self._config_path,
            year    = None,
            funds   = None,
            dry_run = False,
            step    = None,
        )
        setattr(a, 'all', False)
        return a

    def _center(self):
        self.update_idletasks()
        w, h = 860, 720
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _on_escape(self):
        if self._running:
            return
        if messagebox.askyesno("Exit", "Exit T3 Compute?", parent=self):
            self.destroy()

    def _on_close(self):
        if self._running:
            if not messagebox.askyesno(
                    "Step in progress",
                    "A pipeline step is still running.\nExit anyway?",
                    parent=self):
                return
        self.destroy()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ───────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=BG2, padx=20, pady=14)
        header.pack(fill="x")
        tk.Label(header, text="T3 Compute", bg=BG2, fg=TEXT,
                 font=("Segoe UI", 15, "bold")).pack(side="left")
        try:
            import re
            _vdata = open(os.path.join(self._script_dir, "version.txt")).read()
            _m = re.search(r"'FileVersion'\s*,\s*u?'([^']+)'", _vdata)
            ver_text = f"  v{_m.group(1)}" if _m else ""
        except Exception:
            ver_text = ""
        tk.Label(header, text=f"Canadian ETF T3 Tax Pipeline{ver_text}",
                 bg=BG2, fg=TEXT_DIM, font=FONT_SMALL).pack(side="left", padx=(10, 0), pady=(4, 0))
        self._config_label = tk.Label(header, text="", bg=BG2, fg=TEXT_DIM, font=FONT_SMALL)
        self._config_label.pack(side="right")
        self._refresh_config_label()

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── Body: left panel + log ────────────────────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=0, pady=0)

        # Left panel
        left = tk.Frame(body, bg=BG2, width=230)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        tk.Frame(body, bg=BORDER, width=1).pack(side="left", fill="y")

        bottom = tk.Frame(left, bg=BG2)
        bottom.pack(side="bottom", fill="x")

        tk.Frame(bottom, bg=BORDER, height=1).pack(fill="x", padx=18, pady=(8, 4))
        setup_btn = tk.Button(bottom, text="⚙  Setup Wizard",
                              command=self._open_setup,
                              bg=BG2, fg=TEXT_DIM, font=FONT_SMALL,
                              relief="flat", cursor="hand2", pady=5,
                              activebackground=STEP_HOVER, activeforeground=TEXT, bd=0)
        setup_btn.pack(fill="x", padx=18, pady=(0, 2))

        cfg_btn = tk.Button(bottom, text="📂  Open Config Folder",
                            command=self._open_config_folder,
                            bg=BG2, fg=TEXT_DIM, font=FONT_SMALL,
                            relief="flat", cursor="hand2", pady=5,
                            activebackground=STEP_HOVER, activeforeground=TEXT, bd=0)
        cfg_btn.pack(fill="x", padx=18, pady=(0, 2))

        about_btn = tk.Button(bottom, text="ℹ  About",
                              command=self._show_about,
                              bg=BG2, fg=TEXT_DIM, font=FONT_SMALL,
                              relief="flat", cursor="hand2", pady=5,
                              activebackground=STEP_HOVER, activeforeground=TEXT, bd=0)
        about_btn.pack(fill="x", padx=18, pady=(0, 2))

        tk.Frame(bottom, bg=BORDER, height=1).pack(fill="x", padx=18, pady=(4, 2))

        # Steps section
        tk.Label(left, text="PIPELINE STEPS", bg=BG2, fg=TEXT_DIM,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=18, pady=(20, 4))

        STEP_INFO = [
            (1, "Step 1", "Parse T3 PDFs",      "CDS PDFs → JSONs"),
            (2, "Step 2", "Update ACB",          "Insert ROC rows"),
            (3, "Step 3", "Build Assets",        "ACB → share JSONs"),
            (4, "Step 4", "Compute T3",          "JSONs → T3 totals"),
        ]

        self._step_btns = {}
        for num, label, name, desc in STEP_INFO:
            self._step_btns[num] = self._make_step_button(left, num, label, name, desc)

        tk.Frame(left, bg=BORDER, height=1).pack(fill="x", padx=18, pady=2)

        # Run All button
        run_all = tk.Button(left, text="▶  Run All Steps",
                            command=self._run_all,
                            bg=ACCENT, fg="#ffffff", font=FONT_BOLD,
                            relief="flat", cursor="hand2",
                            padx=14, pady=8,
                            activebackground=ACCENT2, activeforeground="#ffffff", bd=0)
        run_all.pack(fill="x", padx=18, pady=(0, 2))
        self._run_all_btn = run_all

        tk.Frame(left, bg=BORDER, height=1).pack(fill="x", padx=18, pady=2)

        # Options section
        tk.Label(left, text="OPTIONS", bg=BG2, fg=TEXT_DIM,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=18, pady=(0, 4))

        # Year override
        yr_row = tk.Frame(left, bg=BG2); yr_row.pack(fill="x", padx=18, pady=(0, 2))
        tk.Label(yr_row, text="Year override:", bg=BG2, fg=TEXT, font=FONT_SMALL).pack(anchor="w")
        self._year_var = tk.StringVar()
        yr_entry = tk.Entry(yr_row, textvariable=self._year_var, width=10,
                            bg=BG, fg=TEXT, insertbackground=TEXT,
                            relief="flat", font=FONT_SMALL,
                            highlightthickness=1, highlightbackground=BORDER,
                            highlightcolor=ACCENT)
        yr_entry.pack(anchor="w", pady=(3, 0))
        tk.Label(yr_row, text="Leave blank to use config.json",
                 bg=BG2, fg=TEXT_DIM, font=("Segoe UI", 7)).pack(anchor="w")

        # Funds override
        fd_row = tk.Frame(left, bg=BG2); fd_row.pack(fill="x", padx=18, pady=(0, 2))
        tk.Label(fd_row, text="Funds override:", bg=BG2, fg=TEXT, font=FONT_SMALL).pack(anchor="w")
        self._funds_var = tk.StringVar()
        tk.Entry(fd_row, textvariable=self._funds_var, width=18,
                 bg=BG, fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=FONT_SMALL,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT).pack(anchor="w", pady=(3, 0))
        tk.Label(fd_row, text="e.g. VBAL ZCN  (space-separated)",
                 bg=BG2, fg=TEXT_DIM, font=("Segoe UI", 7)).pack(anchor="w")

        # Dry run
        self._dry_run_var = tk.BooleanVar()
        dry_row = tk.Frame(left, bg=BG2); dry_row.pack(fill="x", padx=18, pady=(0, 2))
        dry_cb = tk.Checkbutton(dry_row, text="Dry run (Step 2 only)",
                                variable=self._dry_run_var,
                                bg=BG2, fg=TEXT, selectcolor=BG,
                                activebackground=BG2, font=FONT_SMALL)
        dry_cb.pack(anchor="w")
        self._add_tooltip(dry_cb,
            "Preview what ROC rows would be inserted\ninto your ACB spreadsheet without saving.")

        # ── Right: log area ───────────────────────────────────────────────────
        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        log_header = tk.Frame(right, bg=BG2, padx=14, pady=8)
        log_header.pack(fill="x")
        tk.Label(log_header, text="Output", bg=BG2, fg=TEXT, font=FONT_BOLD).pack(side="left")
        self._status_label = tk.Label(log_header, text="Ready", bg=BG2, fg=TEXT_DIM, font=FONT_SMALL)
        self._status_label.pack(side="left", padx=(12, 0))
        exit_btn = tk.Button(log_header, text="✕  Exit",
                             command=self._on_close,
                             bg=BG2, fg=TEXT_DIM, font=FONT_SMALL,
                             relief="flat", cursor="hand2", padx=8, pady=2,
                             activebackground=STEP_HOVER, activeforeground=ERROR, bd=0)
        exit_btn.pack(side="right", padx=(0, 4))

        self._export_btn = tk.Button(
            log_header, text="📄  Export T3",
            command=self._export_t3_html,
            bg=BG2, fg=TEXT_DIM, font=FONT_SMALL,
            relief="flat", cursor="hand2", padx=8, pady=2,
            activebackground=STEP_HOVER, activeforeground=TEXT,
            bd=0, state="disabled")
        self._export_btn.pack(side="right", padx=(0, 4))

        clear_btn = tk.Button(log_header, text="Clear",
                              command=self._clear_log,
                              bg=BG2, fg=TEXT_DIM, font=FONT_SMALL,
                              relief="flat", cursor="hand2", padx=8, pady=2,
                              activebackground=STEP_HOVER, activeforeground=TEXT, bd=0)
        clear_btn.pack(side="right")

        save_btn = tk.Button(log_header, text="Save…",
                             command=self._save_log,
                             bg=BG2, fg=TEXT_DIM, font=FONT_SMALL,
                             relief="flat", cursor="hand2", padx=8, pady=2,
                             activebackground=STEP_HOVER, activeforeground=TEXT, bd=0)
        save_btn.pack(side="right", padx=(0, 4))

        tk.Frame(right, bg=BORDER, height=1).pack(fill="x")

        self._log = scrolledtext.ScrolledText(
            right, bg=LOG_BG, fg=LOG_FG, font=FONT_LOG,
            relief="flat", wrap="word",
            insertbackground=LOG_FG,
            selectbackground=STEP_ACTIVE,
            padx=14, pady=12,
            state="disabled",
        )
        self._log.pack(fill="both", expand=True)

        # Log colour tags
        self._log.tag_config("green",   foreground=SUCCESS)
        self._log.tag_config("yellow",  foreground=WARNING)
        self._log.tag_config("red",     foreground=ERROR)
        self._log.tag_config("blue",    foreground=ACCENT)
        self._log.tag_config("dim",     foreground=TEXT_DIM)
        self._log.tag_config("heading", foreground=ACCENT, font=("Consolas", 9, "bold"))

        # Status bar
        self._statusbar = tk.Label(self, text="", bg=BG2, fg=TEXT_DIM,
                                   font=FONT_SMALL, anchor="w", padx=14, pady=4)
        self._statusbar.pack(fill="x", side="bottom")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", side="bottom")

    def _add_tooltip(self, widget, text):
        tip = None
        def show(e):
            nonlocal tip
            tip = tk.Toplevel(widget)
            tip.wm_overrideredirect(True)
            tip.wm_geometry(f"+{e.x_root+12}+{e.y_root+6}")
            tk.Label(tip, text=text, bg="#ffffcc", fg="#000000",
                     font=FONT_SMALL, relief="solid", bd=1,
                     padx=6, pady=3).pack()
        def hide(e):
            nonlocal tip
            if tip:
                tip.destroy()
                tip = None
        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)

    def _make_step_button(self, parent, num, label, name, desc):
        f = tk.Frame(parent, bg=BG2, cursor="hand2")
        f.pack(fill="x", padx=10, pady=2)
        f.configure(relief="flat")

        inner = tk.Frame(f, bg=BG2, padx=12, pady=8)
        inner.pack(fill="x")

        num_lbl  = tk.Label(inner, text=label, bg=BG2, fg=ACCENT,
                            font=("Segoe UI", 8, "bold"), width=6, anchor="w")
        num_lbl.pack(side="left")
        text_f = tk.Frame(inner, bg=BG2)
        text_f.pack(side="left", fill="x", expand=True)
        name_lbl = tk.Label(text_f, text=name, bg=BG2, fg=TEXT,
                            font=FONT_STEP, anchor="w")
        name_lbl.pack(anchor="w")
        desc_lbl = tk.Label(text_f, text=desc, bg=BG2, fg=TEXT_DIM,
                            font=("Segoe UI", 8), anchor="w")
        desc_lbl.pack(anchor="w")

        return self._finish_step_button(f, inner, num_lbl, text_f, name_lbl, desc_lbl, num)

    def _show_about(self):
        import importlib.metadata, tkinter.messagebox as mb
        try:
            import re
            _vdata = open(os.path.join(self._script_dir, "version.txt")).read()
            _m = re.search(r"'FileVersion'\s*,\s*u?'([^']+)'", _vdata)
            ver_text = f"  v{_m.group(1)}" if _m else ""
        except Exception:
            ver_text = ""
        mb.showinfo("About T3 Compute",
            f"T3 Compute  {ver_text}\n\n"
            f"Free tool for Canadian ETF investors.\n"
            f"Automates CDS T3 slip calculations\n"
            f"and ACB spreadsheet updates.\n\n"
            f"gitlab.com/vdimitrov_73/t3_compute",
            parent=self)
 
    def _save_log(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")],
            initialfile=f"T3_log_{date.today():%Y-%m-%d}.txt")
        if path:
            content = self._log.get("1.0", "end")
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)

    def _finish_step_button(self, f, inner, num_lbl, text_f, name_lbl, desc_lbl, num):
        """Wire up click/hover bindings for a step button. Called by _make_step_button."""
        def on_click(n=num):
            self._run_step(n)
        def on_enter(e, widgets=(f, inner, num_lbl, text_f, name_lbl, desc_lbl)):
            for w in widgets:
                w.configure(bg=STEP_HOVER)
        def on_leave(e, widgets=(f, inner, num_lbl, text_f, name_lbl, desc_lbl)):
            for w in widgets:
                w.configure(bg=BG2)

        for w in (f, inner, num_lbl, text_f, name_lbl, desc_lbl):
            w.bind("<Button-1>", lambda e, n=num: self._run_step(n))
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)

        return {"frame": f, "inner": inner, "text_f": text_f, "num": num_lbl, "name": name_lbl, "desc": desc_lbl}

    # ── Config helpers ────────────────────────────────────────────────────────

    def _refresh_config_label(self):
        if os.path.exists(self._config_path):
            try:
                with open(self._config_path) as f:
                    cfg = json.load(f)
                year  = cfg.get("tax_year", "?")
                funds = ", ".join(cfg.get("funds", []))
                self._config_label.configure(text=f"Tax year: {year}  ·  Funds: {funds}", fg=TEXT_DIM)
            except Exception:
                self._config_label.configure(text="config.json (unreadable)", fg=WARNING)
        else:
            self._config_label.configure(text="No config.json — run Setup Wizard", fg=WARNING)

    def _open_config_folder(self):
        import subprocess
        try:
            subprocess.Popen(["explorer", self._config_dir])
        except Exception:
            messagebox.showinfo("Config folder", self._config_dir, parent=self)

    def _open_setup(self):
        SetupWizard(
            self,
            self._script_dir, self._config_dir,
            self._config_path, self._periods_path, self._funds_path,
            on_complete=self._refresh_config_label,
        )

    def _check_setup(self):
        missing = [p for p in (self._config_path, self._periods_path, self._funds_path)
                   if not os.path.exists(p)]
        if missing:
            names = [os.path.basename(p) for p in missing]
            ans = messagebox.askyesno(
                "Welcome to T3 Compute",
                f"Config files not found:\n  {', '.join(names)}\n\nRun the setup wizard now?",
                parent=self,
            )
            if ans:
                self._open_setup()

    # ── Logging ───────────────────────────────────────────────────────────────

    def _log_write(self, text):
        self._log.configure(state="normal")

        # Colour-code based on content
        lower = text.lower()
        if any(k in lower for k in ("error", "failed", "traceback", "exited with code")):
            tag = "red"
        elif any(k in lower for k in ("warning", "warn", "skipping", "skipped")):
            tag = "yellow"
        elif any(k in lower for k in ("completed successfully", "all steps completed", " saved:", "done.")):
            tag = "green"
        elif text.startswith("=" * 10) or "STEP " in text:
            tag = "heading"
        elif any(k in lower for k in ("loading", "computing", "processing", "beginning")):
            tag = "blue"
        else:
            tag = None

        if tag:
            self._log.insert("end", text, tag)
        else:
            self._log.insert("end", text)

        self._log.configure(state="disabled")
        self._log.see("end")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _maybe_enable_export(self):
        """Enable the Export T3 button if T3_results_<year>.txt exists."""
        try:
            with open(self._config_path) as f:
                cfg = json.load(f)
            year     = self._year_var.get().strip() or cfg.get("tax_year", "")
            base_dir = cfg.get("base_dir", "")
            txt      = os.path.join(base_dir, year, f"T3_results_{year}.txt")
            if os.path.exists(txt):
                self._export_btn.configure(state="normal", fg=TEXT)
        except Exception:
            pass

    def _export_t3_html(self):
        try:
            yr_override = self._year_var.get().strip() or None
            fd = self._funds_var.get().strip()
            funds_override = [f.strip().upper() for f in fd.split() if f.strip()] or None
            out_path = _run_t3.generate_t3_html(self._config_path, yr_override, funds_override)
            self._log_write(f"✔  T3 HTML saved to: {out_path}\n")
            import webbrowser
            from pathlib import Path
            webbrowser.open(Path(out_path).as_uri())
        except Exception as ex:
            import traceback
            self._log_write(f"ERROR exporting T3 HTML: {ex}\n{traceback.format_exc()}")

    def _poll_log(self):
        """Drain the log queue — runs every 50 ms on the main thread."""
        try:
            while True:
                kind, text = self._log_q.get_nowait()
                if kind == "log":
                    self._log_write(text)
                elif kind == "done":
                    self._on_step_done(text)
        except queue.Empty:
            pass
        self.after(50, self._poll_log)

    # ── Step execution ────────────────────────────────────────────────────────

    def _sync_args(self):
        """Push current UI values into the args namespace before running."""
        yr = self._year_var.get().strip()
        self._args.year    = yr if yr else None
        self._args.config  = self._config_path

        fd = self._funds_var.get().strip()
        self._args.funds   = fd.split() if fd else None
        self._args.dry_run = self._dry_run_var.get()

    def _set_running(self, running, step_label="", step_num=None):
        self._running = running
        self._run_all_btn.configure(state="disabled" if running else "normal")
        for num, info in self._step_btns.items():
            active_bg = STEP_ACTIVE if (running and num == step_num) else BG2
            for w in info.values():
                if hasattr(w, 'configure'):
                    try:
                        w.configure(cursor="watch" if running else "hand2", bg=active_bg)
                    except Exception:
                        pass
        if running:
            self._status_label.configure(text=f"Running {step_label}…", fg=WARNING)
            self._statusbar.configure(text=f"Running {step_label}…")
        else:
            self._status_label.configure(text="Ready", fg=TEXT_DIM)
            self._statusbar.configure(text="")

    def _run_step(self, step_num):
        if self._running:
            return
        if not os.path.exists(self._config_path):
            messagebox.showwarning("No config", "config.json not found. Run the Setup Wizard first.", parent=self)
            return
        self._sync_args()
        self._ran_step4 = (step_num == 4)
        step_name = _run_t3.STEPS[step_num]["name"]
        self._set_running(True, f"Step {step_num}: {step_name}", step_num=step_num)
        self._log_write(f"\n{'='*60}\n  STEP {step_num}: {step_name}\n{'='*60}\n")

        def worker():
            self._run_worker(lambda: _run_t3.run_step(step_num, self._args))

        threading.Thread(target=worker, daemon=True).start()

    def _run_all(self):
        if self._running:
            return
        if not os.path.exists(self._config_path):
            messagebox.showwarning("No config", "config.json not found. Run the Setup Wizard first.", parent=self)
            return
        self._sync_args()
        self._ran_step4 = True
        self._set_running(True, "all steps")
        self._log_write(f"\n{'='*60}\n  RUNNING ALL STEPS\n{'='*60}\n")

        def worker():
            def run_all():
                for sn in (1, 2, 3, 4):
                    step_name = _run_t3.STEPS[sn]["name"]
                    self.after(0, lambda s=sn, n=step_name: self._set_running(True, f"Step {s}: {n}", step_num=s))
                    _run_t3.run_step(sn, self._args)
            self._run_worker(run_all)

        threading.Thread(target=worker, daemon=True).start()

    def _run_worker(self, fn):
        """
        Run fn() in the current thread with stdout/stderr redirected to the
        log queue and builtins.input patched to auto-answer 'Y'.

        Any input() call in the pipeline (e.g. the prior-year confirmation in
        update_acb.py) would crash in the GUI because sys.stdin is unavailable.
        We patch it to always return 'Y' and echo the prompt + answer to the log
        so the user can see what was auto-confirmed.
        """
        import builtins

        def _auto_input(prompt=""):
            # Echo the prompt so the user sees it, then auto-answer Y
            sys.stdout.write(f"{prompt}Y  [auto-confirmed]\n")
            return "Y"

        old_stdout, old_stderr = sys.stdout, sys.stderr
        old_input = builtins.input
        writer = _QueueWriter(self._log_q)
        sys.stdout = sys.stderr = writer
        builtins.input = _auto_input
        result = "ok"
        try:
            fn()
        except SystemExit as e:
            if e.code == 2:
                result = "warnings"
            elif e.code not in (0, None):
                result = f"error (exit {e.code})"
        except Exception as ex:
            import traceback
            sys.stdout.write(traceback.format_exc() + "\n")
            result = f"error: {ex}"
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr
            builtins.input = old_input
        self._log_q.put(("done", result))

    def _on_step_done(self, result):
        self._set_running(False)
        if result == "ok":
            self._status_label.configure(text="Completed ✓", fg=SUCCESS)
            self._statusbar.configure(text="Completed successfully.")
        elif result == "warnings" or "warning" in result.lower():
            self._status_label.configure(text="Completed with warnings ⚠", fg=WARNING)
            self._statusbar.configure(text="Completed with warnings — review output above.")
        else:
            self._status_label.configure(text=f"Failed ✗", fg=ERROR)
            self._statusbar.configure(text=f"Error: {result}")
        if self._ran_step4:
            self._maybe_enable_export()

# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    app = T3App()
    app.mainloop()


if __name__ == "__main__":
    main()
