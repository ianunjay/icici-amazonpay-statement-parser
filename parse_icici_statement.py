"""
ICICI Amazon Pay Credit Card Statement Parser
Reads a PDF statement and outputs a clean, formatted Excel file with all transactions.
Usage: python parse_icici_statement.py <path_to_pdf> [output_excel_path]
"""

import re
import sys
from datetime import datetime
import pdfplumber
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter


# ── helpers ──────────────────────────────────────────────────────────────────

DATE_RE   = re.compile(r'\b(\d{2}/\d{2}/\d{4})\b')
AMOUNT_RE = re.compile(r'([\d,]+\.\d{2})\s*(CR)?$')

DEBIT_CATEGORIES = {
    "AMAZON PAY IN GROCERY": "Grocery",
    "AMAZON PAY IN UTILITY": "Utility",
    "AMAZON PAY IN E COMMERC": "E-Commerce",
    "AMAZON PAY IN": "Amazon",
    "MYNTRA": "Apparel",
    "ADITYA BIRLA FASHION": "Apparel",
    "BBPS": "Payment",
}

def categorise(description: str) -> str:
    desc_upper = description.upper()
    for key, label in DEBIT_CATEGORIES.items():
        if key in desc_upper:
            return label
    return "Other"


def parse_transactions(pdf_path: str) -> tuple[list[dict], dict]:
    """Extract all transactions and statement summary from the PDF."""
    transactions = []
    summary = {}

    # Patterns to detect transaction rows
    tx_line_re = re.compile(
        r'(\d{2}/\d{2}/\d{4})\s+'           # date
        r'(\d{10,})\s+'                        # serial number
        r'(.+?)\s+'                            # description (greedy → trimmed later)
        r'(\d+)\s+'                            # reward points
        r'([\d,]+\.\d{2})\s*(CR)?'            # amount + optional CR
    )

    # Summary value patterns (backtick used as ₹ in PDF extract)
    summary_patterns = {
        'statement_date':  re.compile(r'(?:STATEMENT DATE|Statement Date)[^\d]*(\w+ \d+, \d{4})'),
        'payment_due_date':re.compile(r'(?:PAYMENT DUE DATE|Payment Due Date)[^\d]*(\w+ \d+, \d{4})'),
        'total_amount_due':re.compile(r'[`₹](\d[\d,]+\.\d{2})\s*='),
        'min_amount_due':  re.compile(r'Minimum Amount due\s*[`₹]([\d,]+\.\d{2})'),
        'prev_balance':    re.compile(r'Previous Balance\s*[`₹]?([\d,]+\.\d{2})'),
        'total_purchases': re.compile(r'Purchases\s*/\s*Charges\s*[`₹]?([\d,]+\.\d{2})'),
        'total_payments':  re.compile(r'Payments\s*/\s*Credits\s*[`₹]?([\d,]+\.\d{2})'),
        'credit_limit':    re.compile(r'Credit Limit.*?[`₹]([\d,]+\.\d{2})'),
        'available_credit':re.compile(r'Available Credit.*?[`₹]([\d,]+\.\d{2})'),
    }

    full_text = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            full_text += "\n" + text

            lines = text.splitlines()
            i = 0
            while i < len(lines):
                line = lines[i].strip()

                # Try to match a complete transaction on this line
                m = tx_line_re.search(line)

                # Some transactions wrap the description across two lines
                if not m and i + 1 < len(lines):
                    combined = line + " " + lines[i + 1].strip()
                    m = tx_line_re.search(combined)
                    if m:
                        i += 1  # skip the continuation line

                if m:
                    date_str, ser_no, desc, points, amount_str, cr_flag = m.groups()
                    amount = float(amount_str.replace(',', ''))
                    tx_type = "Credit" if cr_flag else "Debit"
                    transactions.append({
                        'Date':           date_str,
                        'Serial No':      ser_no,
                        'Description':    desc.strip(),
                        'Category':       categorise(desc),
                        'Type':           tx_type,
                        'Reward Points':  int(points),
                        'Amount (₹)':     amount if tx_type == 'Debit' else -amount,
                        'Debit (₹)':      amount if tx_type == 'Debit' else None,
                        'Credit (₹)':     amount if tx_type == 'Credit' else None,
                    })
                i += 1

    # Extract summary fields from full text
    for key, pattern in summary_patterns.items():
        m = pattern.search(full_text)
        if m:
            summary[key] = m.group(1).replace(',', '')

    return transactions, summary


def style_header_cell(cell, bg_hex='1F3864', font_color='FFFFFF'):
    cell.font = Font(bold=True, color=font_color, name='Arial', size=10)
    cell.fill = PatternFill('solid', start_color=bg_hex)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    border = Border(
        bottom=Side(style='thin', color='FFFFFF'),
        right=Side(style='thin', color='FFFFFF'),
    )
    cell.border = border


def build_excel(transactions: list[dict], summary: dict, out_path: str):
    wb = openpyxl.Workbook()

    # ── Sheet 1: Transactions ─────────────────────────────────────────────
    ws = wb.active
    ws.title = "Transactions"
    ws.freeze_panes = 'A3'

    # Banner row
    ws.merge_cells('A1:I1')
    banner = ws['A1']
    banner.value = "ICICI Amazon Pay Credit Card — Transaction Statement"
    banner.font = Font(bold=True, color='FFFFFF', name='Arial', size=12)
    banner.fill = PatternFill('solid', start_color='FF6600')
    banner.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    # Column headers
    headers = ['Date', 'Serial No', 'Description', 'Category',
               'Type', 'Reward Points', 'Amount (₹)', 'Debit (₹)', 'Credit (₹)']
    ws.append(headers)
    for col_idx, _ in enumerate(headers, 1):
        style_header_cell(ws.cell(row=2, column=col_idx))
    ws.row_dimensions[2].height = 20

    # Data rows
    DEBIT_FILL  = PatternFill('solid', start_color='FFF2CC')
    CREDIT_FILL = PatternFill('solid', start_color='E2EFDA')
    ALT_FILL    = PatternFill('solid', start_color='F7F7F7')
    INR_FMT     = '#,##0.00'
    DATE_FMT    = 'DD/MM/YYYY'

    for row_idx, tx in enumerate(transactions, 3):
        ws.append([
            datetime.strptime(tx['Date'], '%d/%m/%Y'),
            tx['Serial No'],
            tx['Description'],
            tx['Category'],
            tx['Type'],
            tx['Reward Points'],
            tx['Amount (₹)'],
            tx['Debit (₹)'],
            tx['Credit (₹)'],
        ])
        row = ws[row_idx]
        fill = DEBIT_FILL if tx['Type'] == 'Debit' else CREDIT_FILL
        for cell in row:
            cell.fill = fill
            cell.font = Font(name='Arial', size=10)
            cell.alignment = Alignment(vertical='center')
            cell.border = Border(
                bottom=Side(style='hair', color='CCCCCC'),
                right=Side(style='hair', color='CCCCCC'),
            )

        # Format date & numbers
        row[0].number_format = DATE_FMT
        for col in (6, 7, 8):
            if row[col].value is not None:
                row[col].number_format = INR_FMT

    # Totals row
    n = len(transactions)
    total_row = n + 3
    ws.cell(total_row, 1, "TOTALS").font = Font(bold=True, name='Arial')
    ws.cell(total_row, 5, f'=COUNTIF(E3:E{n+2},"Debit")&" Debits / "&COUNTIF(E3:E{n+2},"Credit")&" Credits"')
    ws.cell(total_row, 6, f'=SUM(F3:F{n+2})').number_format = '#,##0'
    ws.cell(total_row, 7, f'=SUM(G3:G{n+2})').number_format = INR_FMT
    ws.cell(total_row, 8, f'=SUM(H3:H{n+2})').number_format = INR_FMT
    ws.cell(total_row, 9, f'=SUM(I3:I{n+2})').number_format = INR_FMT
    for col in range(1, 10):
        c = ws.cell(total_row, col)
        c.fill = PatternFill('solid', start_color='1F3864')
        c.font = Font(bold=True, color='FFFFFF', name='Arial', size=10)

    # Column widths
    col_widths = [13, 16, 45, 14, 10, 14, 14, 12, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Sheet 2: Summary ──────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2.column_dimensions['A'].width = 28
    ws2.column_dimensions['B'].width = 20

    ws2.merge_cells('A1:B1')
    hdr = ws2['A1']
    hdr.value = "Statement Summary"
    hdr.font = Font(bold=True, color='FFFFFF', name='Arial', size=12)
    hdr.fill = PatternFill('solid', start_color='FF6600')
    hdr.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 26

    summary_rows = [
        ("Statement Date",    summary.get('statement_date', '—')),
        ("Payment Due Date",  summary.get('payment_due_date', '—')),
        ("Total Amount Due",  f"₹{float(summary['total_amount_due']):,.2f}" if 'total_amount_due' in summary else '—'),
        ("Minimum Amount Due",f"₹{float(summary['min_amount_due']):,.2f}"  if 'min_amount_due'   in summary else '—'),
        ("Previous Balance",  f"₹{float(summary['prev_balance']):,.2f}"    if 'prev_balance'     in summary else '—'),
        ("Total Purchases",   f"₹{float(summary['total_purchases']):,.2f}" if 'total_purchases'  in summary else '—'),
        ("Total Payments",    f"₹{float(summary['total_payments']):,.2f}"  if 'total_payments'   in summary else '—'),
        ("Credit Limit",      f"₹{float(summary['credit_limit']):,.2f}"    if 'credit_limit'     in summary else '—'),
        ("Available Credit",  f"₹{float(summary['available_credit']):,.2f}"if 'available_credit' in summary else '—'),
        ("Total Transactions",str(len(transactions))),
        ("Total Reward Points",f'=Transactions!F{total_row}'),
    ]

    for r, (label, value) in enumerate(summary_rows, 2):
        ws2.cell(r, 1, label).font   = Font(bold=True, name='Arial', size=10)
        ws2.cell(r, 2, value).font   = Font(name='Arial', size=10)
        ws2.cell(r, 2).alignment     = Alignment(horizontal='right')
        if r % 2 == 0:
            for col in (1, 2):
                ws2.cell(r, col).fill = PatternFill('solid', start_color='EBF1DE')

    # ── Sheet 3: Category Summary ─────────────────────────────────────────
    ws3 = wb.create_sheet("By Category")
    ws3.column_dimensions['A'].width = 18
    ws3.column_dimensions['B'].width = 14
    ws3.column_dimensions['C'].width = 14

    ws3.merge_cells('A1:C1')
    hdr3 = ws3['A1']
    hdr3.value = "Spend by Category"
    hdr3.font = Font(bold=True, color='FFFFFF', name='Arial', size=12)
    hdr3.fill = PatternFill('solid', start_color='FF6600')
    hdr3.alignment = Alignment(horizontal='center', vertical='center')

    for col, h in enumerate(['Category', 'Total Spend (₹)', '# Transactions'], 1):
        style_header_cell(ws3.cell(2, col))
        ws3.cell(2, col, h)

    categories = sorted({tx['Category'] for tx in transactions if tx['Type'] == 'Debit'})
    for r, cat in enumerate(categories, 3):
        ws3.cell(r, 1, cat).font = Font(name='Arial', size=10)
        total = sum(tx['Debit (₹)'] for tx in transactions if tx['Category'] == cat and tx['Debit (₹)'])
        count = sum(1 for tx in transactions if tx['Category'] == cat and tx['Type'] == 'Debit')
        ws3.cell(r, 2, total).number_format = INR_FMT
        ws3.cell(r, 2).font = Font(name='Arial', size=10)
        ws3.cell(r, 2).alignment = Alignment(horizontal='right')
        ws3.cell(r, 3, count).font = Font(name='Arial', size=10)
        ws3.cell(r, 3).alignment = Alignment(horizontal='center')
        if r % 2 == 0:
            for col in range(1, 4):
                ws3.cell(r, col).fill = PatternFill('solid', start_color='EBF1DE')

    wb.save(out_path)
    print(f"✅ Excel file saved: {out_path}")
    print(f"   Transactions found : {len(transactions)}")
    print(f"   Sheets created     : Transactions, Summary, By Category")


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/icici.pdf'
    out_path  = sys.argv[2] if len(sys.argv) > 2 else 'icici_transactions.xlsx'

    print(f"Parsing: {pdf_path}")
    transactions, summary = parse_transactions(pdf_path)
    build_excel(transactions, summary, out_path)
