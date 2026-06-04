# ICICI Amazon Pay Credit Card — Statement Parser

Parses ICICI Amazon Pay Credit Card PDF statements and exports all transactions to a clean, formatted Excel file.

---

## What it does

Reads a PDF statement and produces an `.xlsx` file with three sheets:

- **Transactions** — every transaction with date, serial number, description, auto-category, type (Debit/Credit), reward points, and split debit/credit columns
- **Summary** — statement-level figures (total due, minimum due, credit limit, payment due date, etc.)
- **By Category** — spend rolled up by category (Grocery, Utility, E-Commerce, Apparel, etc.)

---

## Requirements

Python 3.8+ and two libraries:

```bash
pip install pdfplumber openpyxl
```

---

## Usage

```bash
python parse_icici_statement.py <path_to_pdf> [output_excel_path]
```

**Examples:**

```bash
# Output defaults to icici_transactions.xlsx in the current directory
python parse_icici_statement.py may_2026.pdf

# Specify a custom output path
python parse_icici_statement.py may_2026.pdf ~/Documents/may_2026_transactions.xlsx
```

---

## Compatibility

| Statement type | Works? |
|---|---|
| ICICI Amazon Pay Credit Card (text PDF) | ✅ Yes |
| Other ICICI credit card statements | ⚠️ Likely needs tweaks |
| Other banks | ❌ Not without modifications |
| Scanned / image-based PDFs | ❌ No (pdfplumber can't read images) |

The script is built around ICICI's specific column layout and formatting conventions. It should work reliably across months since ICICI keeps their statement format consistent.

---

## Output columns

| Column | Description |
|---|---|
| Date | Transaction date (DD/MM/YYYY) |
| Serial No | ICICI internal transaction reference |
| Description | Full merchant description from statement |
| Category | Auto-assigned (Grocery, Utility, E-Commerce, Apparel, Payment, Other) |
| Type | Debit or Credit |
| Reward Points | Points earned on that transaction |
| Amount (₹) | Signed amount — negative for credits |
| Debit (₹) | Populated only for debit transactions |
| Credit (₹) | Populated only for credit/payment transactions |

---

## Limitations

- Works on **text-based PDFs only** — scanned statements won't parse
- Category assignment is keyword-based and covers common merchants; uncommon merchants will fall into "Other"
- Tested on the May 2026 statement format; let me know if a future month breaks
