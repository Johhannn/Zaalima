import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
from invoice_utils import (
    download_invoice_pdfs,
    extract_invoice_data,
    save_to_excel,
    save_to_sqlite,
    view_all_records,
    DB_FILE
)

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Automation Tool")
        self.root.geometry("700x500")

        ttk.Button(root, text="Download PDFs", command=self.run_download).pack(pady=5)
        ttk.Button(root, text="Extract Data", command=self.run_extract).pack(pady=5)
        ttk.Button(root, text="Save to Excel", command=self.run_excel).pack(pady=5)
        ttk.Button(root, text="Save to SQLite", command=self.run_sqlite).pack(pady=5)
        ttk.Button(root, text="View Records", command=self.run_view).pack(pady=5)

        self.log_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=20)
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.data = []

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def threaded(self, func):
        threading.Thread(target=func).start()

    def run_download(self):
        self.threaded(self._download)

    def run_extract(self):
        self.threaded(self._extract)

    def run_excel(self):
        self.threaded(self._excel)

    def run_sqlite(self):
        self.threaded(self._sqlite)

    def run_view(self):
        self.threaded(self._view)

    def _download(self):
        self.log("Downloading PDFs...")
        download_invoice_pdfs()
        self.log("PDFs downloaded.")

    def _extract(self):
        self.log("Extracting invoice data...")
        self.data = extract_invoice_data()
        for entry in self.data:
            self.log(str(entry))
        self.log("Data extraction complete.")

    def _excel(self):
        if not self.data:
            messagebox.showwarning("No Data", "Please extract data first.")
            return
        save_to_excel(self.data)
        self.log("Data saved to Excel.")

    def _sqlite(self):
        if not self.data:
            messagebox.showwarning("No Data", "Please extract data first.")
            return
        save_to_sqlite(self.data)
        self.log("Data saved to SQLite.")

    def _view(self):
        self.log("Viewing all records from database...")
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM invoices")
        rows = cursor.fetchall()
        for row in rows:
            self.log(str(row))
        conn.close()
        self.log("Records displayed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()