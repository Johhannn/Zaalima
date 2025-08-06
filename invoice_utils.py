import imaplib
import email
import os
import PyPDF2
import re
import pandas as pd
import sqlite3


EMAIL_USER = '<Your email address>'
EMAIL_PASS = '<Your app password>'
IMAP_SERVER = 'imap.gmail.com'
DOWNLOAD_DIR = 'attachments'
EXCEL_FILE = 'invoices.xlsx'
DB_FILE = 'invoices.db'

def download_invoice_pdfs():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select('inbox')

    status, messages = mail.search(None, '(SUBJECT "Invoice")')
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    for num in messages[0].split():
        status, data = mail.fetch(num, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if filename and filename.endswith('.pdf'):
                filepath = os.path.join(DOWNLOAD_DIR, filename)
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                print(f'Downloaded: {filename}')

    mail.logout()

def extract_invoice_data():
    invoice_data = []

    invoice_no_re = r'Invoice Number:\s*(\S+)'
    date_re = r'Date:\s*(\d{4}-\d{2}-\d{2})'
    amount_re = r'Total\s*\$([0-9,]+\.\d{2})'

    for file in os.listdir(DOWNLOAD_DIR):
        if file.endswith('.pdf'):
            filepath = os.path.join(DOWNLOAD_DIR, file)
            with open(filepath, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                full_text = ''
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + '\n'

                invoice_no = re.search(invoice_no_re, full_text)
                date = re.search(date_re, full_text)
                amount = re.search(amount_re, full_text)

                invoice_data.append({
                    'File': file,
                    'Invoice Number': invoice_no.group(1) if invoice_no else 'Not found',
                    'Date': date.group(1) if date else 'Not found',
                    'Total Amount': amount.group(1) if amount else 'Not found'
                })

    return invoice_data

def save_to_excel(data):
    df = pd.DataFrame(data)
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
    print(f'\nData saved to {EXCEL_FILE}')

def save_to_sqlite(data):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file TEXT,
            invoice_number TEXT,
            date TEXT,
            total_amount TEXT
        )
    ''')

    for entry in data:
        cursor.execute('''
            INSERT INTO invoices (file, invoice_number, date, total_amount)
            VALUES (?, ?, ?, ?)
        ''', (
            entry['File'],
            entry['Invoice Number'],
            entry['Date'],
            entry['Total Amount']
        ))

    conn.commit()
    conn.close()
    print(f'Data saved to {DB_FILE}')

def view_all_records():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM invoices')
    rows = cursor.fetchall()
    for row in rows:
        print(row)
    conn.close()

if __name__ == '__main__':
    download_invoice_pdfs()
    data = extract_invoice_data()
    save_to_excel(data)
    save_to_sqlite(data)
    view_all_records()
