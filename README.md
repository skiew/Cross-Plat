# fintool – private, offline personal finance workbook

One Python script that builds and maintains an Excel budget workbook.
**Nothing leaves your computer**: no network calls, no accounts, no cloud.

## Setup
```bash
pip install openpyxl
python fintool.py init            # creates Finances.xlsx for the current year
```

## Everyday use
```bash
python fintool.py import checking.csv                       # bank export
python fintool.py import visa.csv --flip --account "Visa"   # card exports where charges are positive
python fintool.py summary                                   # quick totals in the terminal
```
Or just open `Finances.xlsx` and type transactions in by hand. Both ways work together.

* Imports are **auto-categorized** with the keyword list on the **Rules** sheet. Add your own merchants there.
* Re-importing an overlapping CSV is safe. Duplicates are skipped.
* A backup (`Finances.bak.xlsx`) is written before every change.
* **Close the file in Excel** before running `import` or `rebuild`.

## The workbook
| Sheet | What it does |
|---|---|
| Dashboard | Monthly income, spending, net, savings rate, budget vs actual, chart. Year is the yellow cell. |
| Transactions | Date, Description, Category (dropdown), Amount, Account, Notes. Type fills itself in. |
| Summary | Each category × month, year total, annual budget, remaining, % used (over-budget turns red). |
| Categories | Your categories, Type (Income / Expense / Transfer) and monthly budget. |
| Rules | Keyword → category for CSV imports. |

**Amount sign:** for expenses, positive = spent, negative = refund. For income, positive = received.
Mark credit-card payments and moves between your own accounts as **Transfers** so they don't count twice.

Blue text = your inputs · black = formulas · green = linked from another sheet.

## Privacy
`.gitignore` excludes `*.xlsx` and `*.csv`, so your data never gets committed. Keep the workbook in a
private folder (an encrypted disk or a password-protected file via Excel → File → Info → Protect Workbook).
