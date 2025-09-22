import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from email.message import EmailMessage
import openpyxl
from openpyxl import load_workbook

# =========================
# Config
# =========================
DB_PATH = "sample_accounts.db"          # swap for your DB or change get_connection()
TEMPLATE_PATH = Path("account_template.xlsx")
OUTPUT_PATH = Path("account_report.xlsx")


# =========================
# Database helpers
# =========================
def get_connection():
    # Swap for Postgres/MySQL as needed.
    # Example (Postgres - psycopg2):
    # import psycopg2
    # return psycopg2.connect(host="localhost", port=5432, dbname="yourdb", user="youruser", password="yourpass")
    return sqlite3.connect(DB_PATH)

def ensure_demo_db():
    """Create a tiny demo table if using SQLite and it doesn't exist."""
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS accounts (
                acct_id TEXT PRIMARY KEY,
                name TEXT,
                segment TEXT,
                balance REAL,
                status TEXT
            )
        """)
        cur.executemany("""
            INSERT OR IGNORE INTO accounts (acct_id, name, segment, balance, status)
            VALUES (?, ?, ?, ?, ?)
        """, [
            ("A001", "Acme Corp", "Enterprise", 120000.50, "Active"),
            ("A002", "Beta LLC", "SMB", 15890.00, "Suspended"),
            ("A003", "Cygnus Inc", "Mid-Market", 5020.75, "Active"),
        ])
        conn.commit()
    except Exception as e:
        print("DB init warning:", e)
    finally:
        try: conn.close()
        except: pass

def fetch_account(acct_id: str):
    query = """
        SELECT acct_id, name, segment, balance, status
        FROM accounts
        WHERE acct_id = ?
    """  # change ? to %s if using psycopg2/mysql-connector
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(query, (acct_id,))
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description] if cur.description else []
        return query, cols, rows
    finally:
        conn.close()


# =========================
# Excel helpers
# =========================
def ensure_template_with_headers(headers):
    """Create a very simple template if none exists."""
    if not TEMPLATE_PATH.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"
        for i, col in enumerate(headers, start=1):
            ws.cell(row=1, column=i, value=col)
        wb.save(TEMPLATE_PATH)

def write_to_template(cols, rows, template_path: Path = TEMPLATE_PATH, output_path: Path = OUTPUT_PATH):
    if not cols:
        raise ValueError("No columns to write to template.")
    ensure_template_with_headers(cols)

    wb = load_workbook(template_path)
    if "Report" not in wb.sheetnames:
        ws = wb.active
        ws.title = "Report"
    ws = wb["Report"]

    # Clear existing rows except header
    if ws.max_row > 1:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.value = None

    # Write new rows
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    wb.save(output_path)
    return output_path


# =========================
# Email helpers (draft .eml)
# =========================
def create_email_draft(recipient, subject, body, attachment_path: Path, from_addr="you@example.com"):
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=attachment_path.name,
        )

    draft_path = Path("draft_email.eml")
    with open(draft_path, "wb") as f:
        f.write(msg.as_bytes())
    return draft_path


# =========================
# Tkinter UI
# =========================
class AccountLookupApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Account Lookup & Export")
        self.geometry("900x640")

        # --- Inputs ---
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="acct_id:").grid(row=0, column=0, sticky="w")
        self.acct_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.acct_var, width=30).grid(row=0, column=1, padx=8, pady=4, sticky="w")

        ttk.Button(top, text="Run Query", command=self.run_query).grid(row=0, column=2, padx=6)
        ttk.Button(top, text="Export → Excel + Draft Email", command=self.export_and_draft).grid(row=0, column=3, padx=6)

        # Email fields
        ttk.Label(top, text="To:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.to_var = tk.StringVar(value="manager@example.com")
        ttk.Entry(top, textvariable=self.to_var, width=30).grid(row=1, column=1, padx=8, pady=(6, 0), sticky="w")

        ttk.Label(top, text="Subject:").grid(row=1, column=2, sticky="e", pady=(6, 0))
        self.subj_var = tk.StringVar(value="Account Report")
        ttk.Entry(top, textvariable=self.subj_var, width=36).grid(row=1, column=3, padx=6, pady=(6, 0), sticky="w")

        # Optional: choose template location
        path_frame = ttk.Frame(self, padding=(10, 0))
        path_frame.pack(fill="x")
        self.template_label_var = tk.StringVar(value=f"Template: {TEMPLATE_PATH.resolve()}")
        self.output_label_var = tk.StringVar(value=f"Output: {OUTPUT_PATH.resolve()}")
        ttk.Label(path_frame, textvariable=self.template_label_var).grid(row=0, column=0, sticky="w", pady=(6, 0))
        ttk.Label(path_frame, textvariable=self.output_label_var).grid(row=1, column=0, sticky="w", pady=(2, 8))
        ttk.Button(path_frame, text="Change Template...", command=self.pick_template).grid(row=0, column=1, padx=8, sticky="e")

        # --- SQL shown ---
        ttk.Label(self, text="Executed SQL:").pack(anchor="w", padx=10)
        self.query_text = tk.Text(self, height=4, wrap="word", bg="#f7f7f7")
        self.query_text.pack(fill="x", padx=10, pady=5)

        # --- Table ---
        ttk.Label(self, text="Query Result (table):").pack(anchor="w", padx=10)
        self.tree = ttk.Treeview(self, columns=(), show="headings", height=8)
        self.tree.pack(fill="x", padx=10, pady=5)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.place(in_=self.tree, relx=1.0, rely=0, relheight=1.0, anchor="ne")

        # --- Formatted details ---
        ttk.Label(self, text="Formatted Result:").pack(anchor="w", padx=10)
        self.result_text = tk.Text(self, height=10, wrap="word", bg="#eef6ff")
        self.result_text.pack(fill="both", expand=True, padx=10, pady=5)

        # --- Status ---
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, anchor="w").pack(fill="x", side="bottom")

        # data cache
        self.last_cols = []
        self.last_rows = []

        ensure_demo_db()

    # ----- UI helpers -----
    def set_columns(self, columns):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns if columns else ()
        for c in columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor="w")

    def populate_rows(self, rows):
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert("", "end", values=row)

    def show_query(self, query):
        self.query_text.delete("1.0", "end")
        self.query_text.insert("1.0", query.strip())

    def show_formatted(self, cols, rows):
        self.result_text.delete("1.0", "end")
        if not rows:
            self.result_text.insert("1.0", "No results found.")
            return
        # Pretty key-value blocks
        for row in rows:
            block = "\n".join(f"{c}: {v}" for c, v in zip(cols, row))
            self.result_text.insert("end", block + "\n" + "-"*48 + "\n")

    # ----- Actions -----
    def run_query(self):
        acct_id = self.acct_var.get().strip()
        if not acct_id:
            messagebox.showinfo("Input required", "Please enter an acct_id.")
            return
        try:
            query, cols, rows = fetch_account(acct_id)
            self.last_cols, self.last_rows = cols, rows
            self.show_query(query)
            if rows:
                self.set_columns(cols)
                self.populate_rows(rows)
                self.show_formatted(cols, rows)
                self.status_var.set(f"Found {len(rows)} row(s).")
            else:
                self.set_columns(["Message"])
                self.populate_rows([("No results found.",)])
                self.show_formatted([], [])
                self.status_var.set("No results.")
        except Exception as e:
            messagebox.showerror("Error", f"Query failed:\n{e}")
            self.status_var.set("Error.")

    def export_and_draft(self):
        # Use last results if present; otherwise, run the query now
        if not self.last_rows:
            self.run_query()
            if not self.last_rows:
                return  # nothing to export

        try:
            # Excel
            out_path = write_to_template(self.last_cols, self.last_rows, TEMPLATE_PATH, OUTPUT_PATH)
            self.output_label_var.set(f"Output: {Path(out_path).resolve()}")

            # Email draft
            to = self.to_var.get().strip()
            subject = self.subj_var.get().strip() or "Account Report"
            acct_id = self.acct_var.get().strip()
            body = f"""Hello,

Please find attached the report for account {acct_id}.

Regards,
Team
"""
            draft_path = create_email_draft(
                recipient=to,
                subject=subject,
                body=body,
                attachment_path=out_path,
                from_addr="you@example.com",
            )

            messagebox.showinfo(
                "Success",
                f"Excel saved to:\n{out_path}\n\nDraft email saved to:\n{draft_path}\n\n"
                "Open the .eml in your mail client to review/send."
            )
            self.status_var.set("Exported and draft created.")
        except Exception as e:
            messagebox.showerror("Error", f"Export or draft failed:\n{e}")
            self.status_var.set("Error during export.")

    def pick_template(self):
        path = filedialog.askopenfilename(
            title="Choose Excel Template",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            global TEMPLATE_PATH
            TEMPLATE_PATH = Path(path)
            self.template_label_var.set(f"Template: {TEMPLATE_PATH.resolve()}")


# =========================
# Run app
# =========================
if __name__ == "__main__":
    app = AccountLookupApp()
    app.mainloop()
