
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import re

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ========= External dependency you provide =========
# Expected signature: run_ccb_query(sql_query: str) -> pandas.DataFrame
try:
    from your_module import run_ccb_query  # <-- replace with the real import path
except Exception:
    # Stub for local UI testing only — replace with your real implementation.
    def run_ccb_query(sql_query: str) -> pd.DataFrame:
        # Parse back IDs from the query for a plausible dummy result
        m = re.search(r"IN\s*\((.*?)\)", sql_query, flags=re.IGNORECASE | re.S)
        ids = []
        if m:
            inside = m.group(1)
            for tok in inside.split(","):
                tok = tok.strip().strip("'\"")
                if tok:
                    ids.append(tok)
        return pd.DataFrame({"ACCT_Id": ids, "SampleCol": [f"Value for {x}" for x in ids]})

APP_NAME = "Account Lookup App"
LOG_DIR = Path("./logs")
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH = LOG_DIR / "acct_lookup.log"

# ==== Configure your query here ====
# Edit this to your actual schema/table. Column must be ACCT_Id or adjust accordingly.
QUERY_TEMPLATE = "SELECT * FROM your_schema.your_table WHERE ACCT_Id IN ({in_clause})"
CHUNK_SIZE = 1000  # Oracle-safe IN list size

# ----------------------------------
# Logging
# ----------------------------------
logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.INFO)
_handler = RotatingFileHandler(LOG_PATH, maxBytes=2_000_000, backupCount=3)
_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
_handler.setFormatter(_formatter)
logger.addHandler(_handler)

# ----------------- Helpers -----------------
def parse_ids_from_text_commas(text: str) -> list[str]:
    """Parse account IDs by splitting on commas. Trims whitespace and dedupes in order."""
    if not text:
        return []
    parts = [p.strip() for p in text.split(",")]
    seen, out = set(), []
    for p in parts:
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out

def build_in_clause(ids: list[str]) -> str:
    """Build a SQL IN clause list of quoted literals: 'id1','id2',..."""
    # Escape single quotes inside ids by doubling them; avoid f-string backslash confusion
    quoted = ["'" + i.replace("'", "''") + "'" for i in ids]
    return ",".join(quoted)

def chunk_iter(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i+size]

# ----------------- App -----------------
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
        text_frame = ttk.LabelFrame(root, text="Account IDs (comma-separated)")
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

        # Results table
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
            df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding_errors="ignore", on_bad_lines="skip")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV:\n{e}")
            logger.exception(f"CSV read failed: {e}")
            return

        if df.empty:
            messagebox.showwarning("Empty CSV", "The selected CSV appears to be empty.")
            return

        # Try to find a column named like ACCT_Id
        def norm(name: str) -> str:
            return re.sub(r"[^0-9a-zA-Z]+", "", str(name)).lower()
        preferred = {"acct_id", "acctid", "account_id", "accountid"}
        normalized = {norm(c): c for c in df.columns}
        acct_col = None
        for want in preferred:
            if want in normalized:
                acct_col = normalized[want]
                break
        if acct_col is None:
            for c in df.columns:
                n = norm(c)
                if "acct" in n and "id" in n:
                    acct_col = c
                    break

        if acct_col is None:
            first_cols = ", ".join(list(df.columns)[:8])
            messagebox.showerror("Missing Column",
                                 "Couldn't find an account id column (e.g., ACCT_Id).\n"
                                 f"Found columns: {first_cols}")
            return

        ids = [x.strip() for x in df[acct_col].astype(str) if x and x.strip()]
        # Deduplicate
        seen, unique_ids = set(), []
        for x in ids:
            if x not in seen:
                seen.add(x)
                unique_ids.append(x)

        # Insert as comma-separated values
        self.acct_text.config(state="normal")
        self.acct_text.delete("1.0", "end")
        self.acct_text.insert("1.0", ", ".join(unique_ids))
        self.acct_text.see("1.0")

        self.selected_csv.set(Path(path).name)
        self.status_var.set(f"Loaded {len(unique_ids)} unique IDs from '{Path(path).name}'.")

    def on_choose_output_dir(self):
        d = filedialog.askdirectory(title="Select Output Directory", mustexist=True)
        if not d:
            return
        self.output_dir.set(d)

    def on_run_query(self):
        raw_text = self.acct_text.get("1.0", "end")
        ids = parse_ids_from_text_commas(raw_text)
        if not ids:
            messagebox.showwarning("No IDs", "Please enter at least one ACCT_Id (comma-separated), or load a CSV.")
            return

        self.status_var.set(f"Running queries in chunks of {CHUNK_SIZE} for {len(ids)} IDs…")
        self.update_idletasks()

        all_frames = []
        try:
            for idx, chunk in enumerate(chunk_iter(ids, CHUNK_SIZE), start=1):
                in_clause = build_in_clause(chunk)
                query = QUERY_TEMPLATE.format(in_clause=in_clause)
                logger.info(f"Executing chunk {idx} with {len(chunk)} ids")
                df = run_ccb_query(query)
                if isinstance(df, pd.DataFrame) and not df.empty:
                    all_frames.append(df)
        except Exception as e:
            logger.exception(f"Query failed: {e}")
            messagebox.showerror("Query Error", f"The query failed:\n{e}")
            self.status_var.set("Query failed.")
            self.result_df = None
            return

        if not all_frames:
            messagebox.showinfo("No Results", "The query returned no rows for the provided IDs.")
            self.status_var.set("No results.")
            self.result_df = None
            self.tree.delete(*self.tree.get_children())
            return

        result = pd.concat(all_frames, ignore_index=True)
        self.result_df = result
        self.populate_table(result)
        self.status_var.set(f"Query complete. {len(result)} rows from {len(all_frames)} chunk(s).")

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

# =========================
# Run app
# =========================
if __name__ == "__main__":
    try:
        app = AccountLookupApp()
        app.mainloop()
    except Exception as e:
        logger.exception(f"Fatal error in mainloop: {e}")
        raise
    finally:
        logger.info("=== Application exit ===")
