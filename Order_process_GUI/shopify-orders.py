import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re

# ── Colour palette ──────────────────────────────────────────────────────────
BG        = "#0F1117"
SIDEBAR   = "#16181F"
CARD      = "#1C1F2B"
ACCENT    = "#4F8EF7"
ACCENT2   = "#7C5CFC"
SUCCESS   = "#22C55E"
WARN      = "#F59E0B"
DANGER    = "#EF4444"
TXT       = "#E8ECF4"
TXT_DIM   = "#6B7280"
TXT_MED   = "#9CA3AF"
BORDER    = "#2A2D3A"
ROW_ODD   = "#1C1F2B"
ROW_EVEN  = "#20232F"
ROW_SEL   = "#2A3650"

FONT_BODY   = ("Segoe UI", 10)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_TITLE  = ("Segoe UI", 20, "bold")
FONT_SUB    = ("Segoe UI", 11)
FONT_SMALL  = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)

# ── Column mapping  (CSV col → display label) ────────────────────────────────
WANTED = {
    "Created at"       : "Date",
    "Name"             : "Order Number",
    "Shipping Name"    : "Full Name",
    "Lineitem name"    : "Book Name",
    "Lineitem quantity": "Quantity",
    "Shipping Address1": "Shipping Address",
    "Shipping City"    : "Shipping City",
    "Shipping Province": "Shipping State",
    "Shipping Zip"     : "Shipping Postcode",
    "Shipping Country" : "Country Code",
    "Notes"            : "Note",
}

DISPLAY_ORDER = [
    "Date", "Order Number", "Full Name", "Book Name",
    "Quantity", "Shipping Address", "Shipping City",
    "Shipping State", "Shipping Postcode", "Country Code", "Note",
]


def load_and_process(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str)

    # keep only wanted columns that actually exist
    available = {k: v for k, v in WANTED.items() if k in df.columns}
    df = df[list(available.keys())].rename(columns=available)

    # reorder to display order (skip missing)
    ordered = [c for c in DISPLAY_ORDER if c in df.columns]
    df = df[ordered]

    # clean date  →  DD-Mon-YYYY
    if "Date" in df.columns:
        def fmt_date(val):
            if pd.isna(val) or str(val).strip() in ("", "nan"):
                return ""
            try:
                return pd.to_datetime(val).strftime("%d-%b-%Y")
            except Exception:
                return str(val)
        df["Date"] = df["Date"].apply(fmt_date)

    # strip whitespace everywhere
    df = df.fillna("").apply(lambda col: col.str.strip() if col.dtype == object else col)
    return df


class OrderExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Order Extractor")
        self.geometry("1300x780")
        self.minsize(900, 600)
        self.configure(bg=BG)

        self.df: pd.DataFrame | None = None
        self.filtered_df: pd.DataFrame | None = None
        self.current_file = tk.StringVar(value="No file loaded")
        self.search_var  = tk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        self.status_var  = tk.StringVar(value="Ready — upload a CSV to begin")
        self.sort_col    = None
        self.sort_asc    = True

        self._build_ui()
        self._apply_treeview_style()

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Left sidebar ──
        sidebar = tk.Frame(self, bg=SIDEBAR, width=220)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Logo / title
        logo_frame = tk.Frame(sidebar, bg=SIDEBAR)
        logo_frame.pack(fill="x", padx=20, pady=(28, 6))
        tk.Label(logo_frame, text="📦", font=("Segoe UI Emoji", 26),
                 bg=SIDEBAR, fg=ACCENT).pack(anchor="w")
        tk.Label(logo_frame, text="Order\nExtractor", font=("Segoe UI", 14, "bold"),
                 bg=SIDEBAR, fg=TXT, justify="left").pack(anchor="w")

        tk.Frame(sidebar, bg=BORDER, height=1).pack(fill="x", padx=20, pady=16)

        # Upload button
        self._sidebar_btn(sidebar, "⬆  Upload CSV", self._upload_file, primary=True)

        tk.Frame(sidebar, bg=BORDER, height=1).pack(fill="x", padx=20, pady=16)

        # Stats panel
        self.stat_frame = tk.Frame(sidebar, bg=SIDEBAR)
        self.stat_frame.pack(fill="x", padx=20)
        self._stat_row("Total Orders",  "—", "total_lbl")
        self._stat_row("Filtered",      "—", "filt_lbl")
        self._stat_row("Source File",   "—", "file_lbl", small=True)

        tk.Frame(sidebar, bg=BORDER, height=1).pack(fill="x", padx=20, pady=16)

        # Export buttons
        self._sidebar_btn(sidebar, "⬇  Export Filtered CSV", self._export_csv)
        self._sidebar_btn(sidebar, "⬇  Export Filtered Excel", self._export_excel)

        # Version tag at bottom
        tk.Label(sidebar, text="v1.0  •  onlinebookdeals",
                 font=("Segoe UI", 8), bg=SIDEBAR, fg=TXT_DIM).pack(side="bottom", pady=12)

        # ── Right main area ──
        main = tk.Frame(self, bg=BG)
        main.pack(side="left", fill="both", expand=True)

        # Top bar
        topbar = tk.Frame(main, bg=BG)
        topbar.pack(fill="x", padx=20, pady=(20, 0))

        tk.Label(topbar, text="Orders", font=FONT_TITLE,
                 bg=BG, fg=TXT).pack(side="left")

        # Search
        search_wrap = tk.Frame(topbar, bg=CARD, highlightbackground=BORDER,
                               highlightthickness=1)
        search_wrap.pack(side="right")
        tk.Label(search_wrap, text="🔍", font=("Segoe UI Emoji", 11),
                 bg=CARD, fg=TXT_DIM).pack(side="left", padx=(10, 4), pady=6)
        self.search_entry = tk.Entry(search_wrap, textvariable=self.search_var,
                                     font=FONT_BODY, bg=CARD, fg=TXT,
                                     insertbackground=ACCENT, bd=0,
                                     highlightthickness=0, width=28)
        self.search_entry.pack(side="left", pady=6)
        self.search_entry.insert(0, "Search orders…")
        self.search_entry.configure(fg=TXT_DIM)
        self.search_entry.bind("<FocusIn>",  self._search_focus_in)
        self.search_entry.bind("<FocusOut>", self._search_focus_out)
        tk.Label(search_wrap, text=" ", bg=CARD).pack(side="left", padx=4)

        # File path strip
        file_bar = tk.Frame(main, bg=BG)
        file_bar.pack(fill="x", padx=20, pady=(4, 12))
        tk.Label(file_bar, textvariable=self.current_file,
                 font=FONT_SMALL, bg=BG, fg=TXT_DIM).pack(side="left")

        # Table area
        table_frame = tk.Frame(main, bg=BG)
        table_frame.pack(fill="both", expand=True, padx=20)

        # Treeview + scrollbars
        tv_wrap = tk.Frame(table_frame, bg=CARD, highlightbackground=BORDER,
                           highlightthickness=1)
        tv_wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tv_wrap, style="Orders.Treeview",
                                 selectmode="extended", show="headings")
        vsb = ttk.Scrollbar(tv_wrap, orient="vertical",
                            command=self.tree.yview, style="Dark.Vertical.TScrollbar")
        hsb = ttk.Scrollbar(tv_wrap, orient="horizontal",
                            command=self.tree.xview, style="Dark.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        # Empty-state label (shown before any file is loaded)
        self.empty_label = tk.Label(
            tv_wrap,
            text="⬆  Drop or upload a Shopify orders CSV\nto extract and view your data.",
            font=("Segoe UI", 13), bg=CARD, fg=TXT_DIM, justify="center"
        )
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

        # Status bar
        status_bar = tk.Frame(main, bg=SIDEBAR, height=28)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)
        tk.Label(status_bar, textvariable=self.status_var,
                 font=FONT_SMALL, bg=SIDEBAR, fg=TXT_MED, anchor="w").pack(
                     side="left", padx=16, pady=4)

    def _sidebar_btn(self, parent, text, cmd, primary=False):
        bg_n  = ACCENT  if primary else CARD
        bg_h  = "#3A74DC" if primary else "#252839"
        fg    = "#ffffff" if primary else TXT
        relief = "flat"
        btn = tk.Button(parent, text=text, command=cmd,
                        bg=bg_n, fg=fg, font=FONT_BOLD,
                        activebackground=bg_h, activeforeground=fg,
                        bd=0, relief=relief, cursor="hand2",
                        padx=14, pady=9)
        btn.pack(fill="x", padx=20, pady=4)
        return btn

    def _stat_row(self, label, value, attr, small=False):
        row = tk.Frame(self.stat_frame, bg=SIDEBAR)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, font=FONT_SMALL if small else ("Segoe UI", 9),
                 bg=SIDEBAR, fg=TXT_DIM).pack(side="left")
        lbl = tk.Label(row, text=value,
                       font=("Segoe UI", 9, "bold") if not small else ("Segoe UI", 8),
                       bg=SIDEBAR, fg=TXT if not small else TXT_DIM)
        lbl.pack(side="right")
        setattr(self, attr, lbl)

    def _apply_treeview_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("Orders.Treeview",
                        background=ROW_ODD, foreground=TXT,
                        fieldbackground=ROW_ODD,
                        rowheight=30, font=FONT_BODY,
                        borderwidth=0, relief="flat")
        style.configure("Orders.Treeview.Heading",
                        background=CARD, foreground=TXT_MED,
                        font=("Segoe UI", 9, "bold"),
                        relief="flat", borderwidth=0,
                        padding=(8, 6))
        style.map("Orders.Treeview",
                  background=[("selected", ROW_SEL)],
                  foreground=[("selected", TXT)])
        style.map("Orders.Treeview.Heading",
                  background=[("active", "#252839")])

        style.configure("Dark.Vertical.TScrollbar",
                        troughcolor=CARD, background=BORDER,
                        arrowcolor=TXT_DIM, borderwidth=0)
        style.configure("Dark.Horizontal.TScrollbar",
                        troughcolor=CARD, background=BORDER,
                        arrowcolor=TXT_DIM, borderwidth=0)

    # ── Core logic ───────────────────────────────────────────────────────────

    def _upload_file(self):
        path = filedialog.askopenfilename(
            title="Select Shopify Orders CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            self.df = load_and_process(path)
            self.filtered_df = self.df.copy()
            self.current_file.set(f"📄  {os.path.basename(path)}")
            self._update_tree(self.df)
            self._update_stats()
            self.empty_label.place_forget()
            self.status_var.set(f"Loaded {len(self.df):,} rows from {os.path.basename(path)}")
        except Exception as exc:
            messagebox.showerror("Load Error", f"Could not process file:\n{exc}")

    def _update_tree(self, df: pd.DataFrame):
        # Clear old columns
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)

        # Column widths
        WIDTHS = {
            "Date": 100, "Order Number": 90, "Full Name": 140,
            "Book Name": 260, "Quantity": 70, "Shipping Address": 180,
            "Shipping City": 120, "Shipping State": 100,
            "Shipping Postcode": 100, "Country Code": 90, "Note": 160,
        }
        for col in df.columns:
            w = WIDTHS.get(col, 120)
            self.tree.heading(col, text=col,
                              command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=w, minwidth=60, stretch=False)

        # Insert rows with alternating colours
        for i, (_, row) in enumerate(df.iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=list(row), tags=(tag,))

        self.tree.tag_configure("odd",  background=ROW_ODD)
        self.tree.tag_configure("even", background=ROW_EVEN)

    def _update_stats(self):
        total   = len(self.df)         if self.df is not None else 0
        filt    = len(self.filtered_df) if self.filtered_df is not None else 0
        fname   = self.current_file.get().replace("📄  ", "")
        self.total_lbl.config(text=f"{total:,}")
        self.filt_lbl.config( text=f"{filt:,}")
        self.file_lbl.config( text=fname[:24] + ("…" if len(fname) > 24 else ""))

    # ── Search ───────────────────────────────────────────────────────────────

    def _search_focus_in(self, _):
        if self.search_entry.get() == "Search orders…":
            self.search_entry.delete(0, "end")
            self.search_entry.configure(fg=TXT)

    def _search_focus_out(self, _):
        if not self.search_entry.get().strip():
            self.search_entry.insert(0, "Search orders…")
            self.search_entry.configure(fg=TXT_DIM)

    def _on_search(self, *_):
        if self.df is None:
            return
        q = self.search_var.get().strip().lower()
        if not q or q == "search orders…":
            self.filtered_df = self.df.copy()
        else:
            mask = self.df.apply(
                lambda col: col.str.lower().str.contains(q, na=False)
            ).any(axis=1)
            self.filtered_df = self.df[mask].copy()
        self._update_tree(self.filtered_df)
        self._update_stats()
        self.status_var.set(f"Showing {len(self.filtered_df):,} of {len(self.df):,} rows")

    # ── Sort ─────────────────────────────────────────────────────────────────

    def _sort_by(self, col):
        if self.filtered_df is None:
            return
        if self.sort_col == col:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = col
            self.sort_asc = True
        self.filtered_df = self.filtered_df.sort_values(
            col, ascending=self.sort_asc, na_position="last"
        )
        self._update_tree(self.filtered_df)
        arrow = " ▲" if self.sort_asc else " ▼"
        self.status_var.set(f"Sorted by '{col}'{arrow}")

    # ── Export ───────────────────────────────────────────────────────────────

    def _export_csv(self):
        self._export("csv")

    def _export_excel(self):
        self._export("excel")

    def _export(self, fmt):
        if self.filtered_df is None or self.filtered_df.empty:
            messagebox.showwarning("Nothing to export", "Please load a file first.")
            return
        ext  = ".csv" if fmt == "csv" else ".xlsx"
        ft   = [("CSV", "*.csv")] if fmt == "csv" else [("Excel", "*.xlsx")]
        path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=ft,
            initialfile=f"orders_extracted{ext}"
        )
        if not path:
            return
        try:
            if fmt == "csv":
                self.filtered_df.to_csv(path, index=False)
            else:
                self.filtered_df.to_excel(path, index=False)
            self.status_var.set(f"Exported {len(self.filtered_df):,} rows → {os.path.basename(path)}")
            messagebox.showinfo("Export complete",
                                f"Saved {len(self.filtered_df):,} rows to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))


if __name__ == "__main__":
    app = OrderExtractorApp()
    app.mainloop()
