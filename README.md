#  Invoice Automation Tool

A Python-based desktop application that automates the process of downloading invoice PDFs from Gmail, extracting key data, and exporting it to Excel and SQLite. Built with `imaplib`, `PyPDF2`, `pandas`, and a responsive `Tkinter` GUI.

---

##  Features

-  ***Download PDFs**: Connects to Gmail via IMAP and fetches invoice attachments.
-  **Extract Data**: Parses invoice number, date, and total amount from PDF text.
-  **Export to Excel**: Saves structured data to `invoices.xlsx` using `pandas`.
-  **Save to SQLite**: Stores invoice records in a local database (`invoices.db`).
-  **GUI Interface**: User-friendly desktop app built with `Tkinter`.

---

##  Installation

1. **Clone the repository**  
   ```bash
   git clone https://github.com/yourusername/invoice-automation-tool.git
   cd Z - proj 1

## Install Dependencies

pip install -r requirements.txt

## Run the APP

python invoice_gui.py

## File Structure

Z-proj 1/
│
├── invoice_gui.py         # Tkinter GUI interface
├── invoice_utils.py       # Core logic for email, PDF, and data handling
├── requirements.txt       # Python dependencies
├── README.md              # Project documentation
├── invoices.db            # SQLite database (auto-generated)
├── invoices.xlsx          # Excel export (auto-generated)
└── attachments/           # Downloaded PDFs (auto-generated)
