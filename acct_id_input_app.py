
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from email.message import EmailMessage
import re

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ========= External dependency you provide =========
try:
    from your_module import run_ccb_query  # <-- replace with the real import path
except Exception:
    def run_ccb_query(acct_ids: list[str]) -> pd.DataFrame:
        return pd.DataFrame({
            "ACCT_Id": acct_ids,
            "SampleCol": [f"Value for {x}" for x in acct_ids],
        })

APP_NAME = "Account Lookup App"
LOG_DIR = Path("./logs")
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH = LOG_DIR / "acct_lookup.log"

logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.INFO)
_handler = RotatingFileHandler(LOG_PATH, maxBytes=2_000_000, backupCount=3)
_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
_handler.setFormatter(_formatter)
logger.addHandler(_handler)

def parse_ids_from_text(text: str) -> list[str]:
    if not text.strip():
        return []
    raw = re.split(r"[,\s]+", text.strip())
    cleaned, seen, out = [], set(), []
    for tok in raw:
        tok = tok.strip().strip("'\"")
        if tok and tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out

def normalize_colname(name: str) -> str:
    # remove non-alnum and lower it, e.g., "ACCT_Id" -> "acctid", "ACCT ID" -> "acctid"
    return re.sub(r"[^0-9a-zA-Z]+", "", str(name)).lower()

def pick_acct_column(columns: list[str]) -> str | None:
    # direct preferred names
    preferred = {"acct_id", "acctid", "account_id", "accountid"}
    normalized = {normalize_colname(c): c for c in columns}
    # First try normalized preferred
    for want in preferred:
        if want in normalized:
            return normalized[want]
    # Next: any column whose normalized name contains both acct and id
    for c in columns:
        norm = normalize_colname(c)
        if "acct" in norm and "id" in norm:
            return c
    return None

class AccountLookupApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("900x650")
        self.minsize(800, 600)

        self.output_dir = tk.StringVar(value=str(Path.cwd() / "output"))
        self.selected_csv = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Ready.")
        self.result_df: pd.DataFrame | None = None

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        # Top: Load CSV
        top_bar = ttk.Frame(root)
        top_bar.pack(fill="x", pady=(0, 8))

        ttk.Label(top_bar, text="Load ACCT_Id from CSV:").pack(side="left")
        ttk.Button(top_bar, text="Choose CSV…", command=self.on_choose_csv).pack(side="left", padx=(8, 0))
        self.csv_label = ttk.Label(top_bar, textvariable=self.selected_csv, foreground="#555")
        self.csv_label.pack(side="left", padx=8)

        # Text area
        text_frame = ttk.LabelFrame(root, text="Account IDs (one per line)")
        text_frame.pack(fill="both", expand=True, pady=(0, 8))

        self.acct_text = tk.Text(text_frame, wrap="word", height=14)
        self.acct_text.pack(side="left", fill="both", expand=True, padx=(6,0), pady=6)

        yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.acct_text.yview)
        yscroll.pack(side="right", fill="y", padx=(0,6), pady=6)
        self.acct_text.configure(yscrollcommand=yscroll.set)

        # Output dir + Run
        out_frame = ttk.Frame(root)
        out_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(out_frame, text="Output directory:").pack(side="left")
        self.output_entry = ttk.Entry(out_frame, textvariable=self.output_dir, width=50)
        self.output_entry.pack(side="left", padx=6, fill="x", expand=True)
        ttk.Button(out_frame, text="Browse…", command=self.on_choose_output_dir).pack(side="left")
        ttk.Button(out_frame, text="Run Query", command=self.on_run_query).pack(side="left", padx=(12,0))

        # Table
        table_frame = ttk.LabelFrame(root, text="Results")
        table_frame.pack(fill="both", expand=True, pady=(0, 8))

        self.tree = ttk.Treeview(table_frame, show="headings", height=10)
        self.tree.pack(side="left", fill="both", expand=True, padx=(6,0), pady=6)

        table_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        table_scroll.pack(side="right", fill="y", padx=(0,6), pady=6)
        self.tree.configure(yscrollcommand=table_scroll.set)

        # Bottom: Export
        bottom_bar = ttk.Frame(root)
        bottom_bar.pack(fill="x")
        ttk.Button(bottom_bar, text="Export Results", command=self.on_export).pack(side="right")
        ttk.Label(bottom_bar, textvariable=self.status_var).pack(side="left")

    # ---------- Actions ----------
    def on_choose_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV with ACCT_Id column",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            # Read all values as string to avoid NaN propagation; auto-detect sep; ignore bad lines gracefully.
            df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding_errors="ignore", on_bad_lines="skip")
            logger.info(f"Loaded CSV: {path}; columns={list(df.columns)}; rows={len(df)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV:\n{e}")
            logger.exception(f"CSV read failed: {e}")
            return

        if df.empty:
            messagebox.showwarning("Empty CSV", "The selected CSV appears to be empty.")
            return

        col = pick_acct_column(list(df.columns))
        if col is None:
            first_cols = ", ".join(list(df.columns)[:8])
            messagebox.showerror(
                "Missing Column",
                "Couldn't find an account id column.\n"
                "Looking for names like ACCT_Id / acct_id / account_id.\n\n"
                f"Found columns: {first_cols}"
            )
            self.status_var.set("CSV missing ACCT_Id-like column.")
            return

        # Extract and clean
        ids_series = df[col].astype(str)
        ids = [x.strip() for x in ids_series if x and x.strip()]
        seen, unique_ids = set(), []
        for x in ids:
            if x not in seen:
                seen.add(x)
                unique_ids.append(x)

        # Ensure text widget is editable, then fill
        self.acct_text.config(state="normal")
        self.acct_text.delete("1.0", "end")
        self.acct_text.insert("1.0", "\n".join(unique_ids))
        self.acct_text.see("1.0")

        self.selected_csv.set(Path(path).name)
        self.status_var.set(f"Loaded {len(unique_ids)} unique IDs from '{Path(path).name}'.")

    def on_choose_output_dir(self):
        d = filedialog.askdirectory(title="Select Output Directory", mustexist=True)
        if not d:
            return
        self.output_dir.set(d)

    def on_run_query(self):
        ids = parse_ids_from_text(self.acct_text.get("1.0", "end"))
        if not ids:
            messagebox.showwarning("No IDs", "Please enter at least one ACCT_Id (or load a CSV).")
            return
        self.status_var.set(f"Running query for {len(ids)} account id(s)…")
        self.update_idletasks()

        try:
            df = run_ccb_query(ids)
        except Exception as e:
            logger.exception(f"Query failed: {e}")
            messagebox.showerror("Query Error", f"The query failed:\n{e}")
            self.status_var.set("Query failed.")
            self.result_df = None
            return

        if not isinstance(df, pd.DataFrame) or df.empty:
            messagebox.showinfo("No Results", "The query returned no rows.")
            self.status_var.set("No results.")
            self.result_df = None
            self.tree.delete(*self.tree.get_children())
            return

        self.result_df = df
        self.populate_table(df)
        self.status_var.set(f"Query complete. {len(df)} rows.")

    def populate_table(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="w")
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[row[c] for c in df.columns])

    def on_export(self):
        if self.result_df is None or self.result_df.empty:
            messagebox.showwarning("Nothing to Export", "Run a query first; there are no results to export.")
            return
        out_dir = Path(self.output_dir.get()).expanduser().resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("CSV", "*.csv")],
            initialdir=str(out_dir),
            initialfile="account_lookup_results.xlsx",
            title="Export Results"
        )
        if not save_path:
            return

        try:
            if save_path.lower().endswith(".csv"):
                self.result_df.to_csv(save_path, index=False)
            else:
                self.result_df.to_excel(save_path, index=False)
            self.status_var.set(f"Exported results to {save_path}")
            messagebox.showinfo("Export Complete", f"Exported to:\n{save_path}")
        except Exception as e:
            logger.exception(f"Export failed: {e}")
            messagebox.showerror("Export Error", f"Failed to export:\n{e}")

if __name__ == "__main__":
    try:
        app = AccountLookupApp()
        app.mainloop()
    except Exception as e:
        logger.exception(f"Fatal error in mainloop: {e}")
        raise
    finally:
        logger.info("=== Application exit ===")
