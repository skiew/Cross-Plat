#!/usr/bin/env python3
"""
fintool.py - a private, offline personal-finance workbook.

Everything stays on your computer: no network calls, no accounts, no cloud.
Your data lives in one Excel file (default: Finances.xlsx).

Commands
  python fintool.py init                     create a fresh Finances.xlsx
  python fintool.py import bank.csv [...]    add bank/card CSV exports (auto-categorized, de-duplicated)
  python fintool.py rebuild                  regenerate formulas/charts, keeping your data
  python fintool.py summary                  print this year's totals in the terminal

Options
  -f / --file PATH     workbook to use (default Finances.xlsx)
  --account NAME       account label for imported rows (default: CSV file name)
  --flip               flip the sign of a CSV's amounts (for card exports where charges are positive)
  --year YYYY          year the dashboard shows (init/rebuild)

Requires: Python 3.8+ and openpyxl  (pip install openpyxl)
Close the workbook in Excel before running import/rebuild.
"""
import argparse
import csv
import datetime as dt
import os
import re
import shutil
import sys

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.comments import Comment
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")

# ---------------------------------------------------------------- defaults
DEFAULT_CATEGORIES = [
    # (name, type, monthly budget)
    ("Salary", "Income", 0),
    ("Other Income", "Income", 0),
    ("Rent / Mortgage", "Expense", 1500),
    ("Utilities", "Expense", 200),
    ("Groceries", "Expense", 500),
    ("Dining Out", "Expense", 200),
    ("Transportation", "Expense", 250),
    ("Insurance", "Expense", 150),
    ("Health", "Expense", 100),
    ("Shopping", "Expense", 200),
    ("Entertainment", "Expense", 100),
    ("Subscriptions", "Expense", 50),
    ("Travel", "Expense", 150),
    ("Gifts & Donations", "Expense", 50),
    ("Fees & Interest", "Expense", 0),
    ("Uncategorized", "Expense", 0),
    ("Transfers", "Transfer", 0),
]

DEFAULT_RULES = [
    # (keyword found in description, category) - case-insensitive, first match wins
    ("payroll", "Salary"), ("direct dep", "Salary"),
    ("interest paid", "Other Income"),
    ("rent", "Rent / Mortgage"), ("mortgage", "Rent / Mortgage"),
    ("electric", "Utilities"), ("water", "Utilities"), ("comcast", "Utilities"),
    ("verizon", "Utilities"), ("at&t", "Utilities"), ("t-mobile", "Utilities"),
    ("kroger", "Groceries"), ("safeway", "Groceries"), ("trader joe", "Groceries"),
    ("whole foods", "Groceries"), ("aldi", "Groceries"), ("publix", "Groceries"),
    ("costco", "Groceries"),
    ("doordash", "Dining Out"), ("uber eats", "Dining Out"), ("grubhub", "Dining Out"),
    ("starbucks", "Dining Out"), ("mcdonald", "Dining Out"), ("chipotle", "Dining Out"),
    ("restaurant", "Dining Out"), ("cafe", "Dining Out"),
    ("shell", "Transportation"), ("chevron", "Transportation"), ("exxon", "Transportation"),
    ("uber", "Transportation"), ("lyft", "Transportation"), ("parking", "Transportation"),
    ("geico", "Insurance"), ("state farm", "Insurance"), ("progressive", "Insurance"),
    ("pharmacy", "Health"), ("cvs", "Health"), ("walgreens", "Health"),
    ("amazon", "Shopping"), ("target", "Shopping"), ("walmart", "Shopping"),
    ("netflix", "Subscriptions"), ("spotify", "Subscriptions"), ("hulu", "Subscriptions"),
    ("disney+", "Subscriptions"), ("apple.com", "Subscriptions"), ("youtube", "Subscriptions"),
    ("airline", "Travel"), ("airbnb", "Travel"), ("hotel", "Travel"), ("delta", "Travel"),
    ("steam", "Entertainment"), ("ticketmaster", "Entertainment"),
    ("overdraft", "Fees & Interest"), ("interest charge", "Fees & Interest"), ("fee", "Fees & Interest"),
    ("payment thank you", "Transfers"), ("autopay", "Transfers"), ("transfer", "Transfers"),
    ("zelle", "Transfers"), ("venmo", "Transfers"),
]

CAT_SLOTS = 60          # rows available on the Categories sheet
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# ---------------------------------------------------------------- styles
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=16, bold=True, color="1F3864")
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_TOTAL = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
B_TOP = Border(top=Side(style="thin", color="000000"))
MONEY = '$#,##0.00;($#,##0.00);"-"'
MONEY0 = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE_FMT = "yyyy-mm-dd"


def _header(ws, row, labels, widths=None):
    for i, label in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=label)
        c.font, c.fill = F_HEAD, FILL_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------- build
def build_workbook(path, year, categories, rules, transactions):
    """Write a complete workbook. transactions: list of dicts date/desc/cat/amount/account/notes."""
    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    tx = wb.create_sheet("Transactions")
    summ = wb.create_sheet("Summary")
    cats = wb.create_sheet("Categories")
    rul = wb.create_sheet("Rules")
    howto = wb.create_sheet("How To Use")

    # ---------------- Categories
    _header(cats, 1, ["Category", "Type", "Monthly Budget ($)"], [26, 12, 18])
    type_dv = DataValidation(type="list", formula1='"Income,Expense,Transfer"', allow_blank=True)
    cats.add_data_validation(type_dv)
    for r in range(2, CAT_SLOTS + 2):
        for col in (1, 2, 3):
            cats.cell(row=r, column=col).font = F_INPUT
        cats.cell(row=r, column=3).number_format = MONEY0
        type_dv.add(f"B{r}")
    for i, (name, typ, budget) in enumerate(categories[:CAT_SLOTS]):
        r = i + 2
        cats.cell(row=r, column=1, value=name)
        cats.cell(row=r, column=2, value=typ)
        cats.cell(row=r, column=3, value=budget)
    cats["E1"] = "Blue text = yours to edit. Add categories in any empty row (up to %d)." % CAT_SLOTS
    cats["E2"] = "Type: Income, Expense, or Transfer (Transfers are excluded from income/spending totals)."
    cats["E3"] = "Budgets are starting placeholders - set your own."
    for a in ("E1", "E2", "E3"):
        cats[a].font = F_NOTE
    cats.freeze_panes = "A2"

    cat_list_ref = f"Categories!$A$2:$A${CAT_SLOTS + 1}"
    cat_type_ref = f"Categories!$B$2:$B${CAT_SLOTS + 1}"

    # ---------------- Rules
    _header(rul, 1, ["Keyword (in description)", "Category"], [30, 24])
    rule_dv = DataValidation(type="list", formula1=f"={cat_list_ref}", allow_blank=True)
    rul.add_data_validation(rule_dv)
    n_rule_rows = max(len(rules) + 100, 200)
    for r in range(2, n_rule_rows + 2):
        rul.cell(row=r, column=1).font = F_INPUT
        rul.cell(row=r, column=2).font = F_INPUT
    rule_dv.add(f"B2:B{n_rule_rows + 1}")
    for i, (kw, cat) in enumerate(rules):
        rul.cell(row=i + 2, column=1, value=kw)
        rul.cell(row=i + 2, column=2, value=cat)
    rul["D1"] = "Used by 'fintool.py import' to auto-categorize. Case-insensitive; first match from the top wins."
    rul["D2"] = "Put more specific keywords above general ones (e.g. 'uber eats' above 'uber')."
    rul["D1"].font = rul["D2"].font = F_NOTE
    rul.freeze_panes = "A2"

    # ---------------- Transactions
    _header(tx, 1, ["Date", "Description", "Category", "Amount ($)", "Account", "Notes", "Type"],
            [12, 42, 22, 13, 18, 28, 10])
    tx["D1"].comment = Comment(
        "Expenses: positive number = money spent, negative = refund.\n"
        "Income: positive number = money received.", "fintool")
    tx_dv = DataValidation(type="list", formula1=f"={cat_list_ref}", allow_blank=True,
                           showErrorMessage=True, errorTitle="Unknown category",
                           error="Pick a category from the list (add new ones on the Categories sheet).")
    tx.add_data_validation(tx_dv)
    transactions = sorted(transactions, key=lambda t: (t["date"], t["desc"]))
    n_rows = max(len(transactions) + 2000, 5000)
    for r in range(2, n_rows + 2):
        tx.cell(row=r, column=1).number_format = DATE_FMT
        tx.cell(row=r, column=4).number_format = MONEY
        for col in range(1, 7):
            tx.cell(row=r, column=col).font = F_INPUT
        g = tx.cell(row=r, column=7,
                    value=f'=IF(C{r}="","",IFERROR(INDEX({cat_type_ref},MATCH(C{r},{cat_list_ref},0)),"?"))')
        g.font = F_BASE
    tx_dv.add(f"C2:C{n_rows + 1}")
    for i, t in enumerate(transactions):
        r = i + 2
        tx.cell(row=r, column=1, value=t["date"])
        tx.cell(row=r, column=2, value=t["desc"])
        tx.cell(row=r, column=3, value=t["cat"])
        tx.cell(row=r, column=4, value=t["amount"])
        tx.cell(row=r, column=5, value=t.get("account") or None)
        tx.cell(row=r, column=6, value=t.get("notes") or None)
    tx.freeze_panes = "A2"
    tx.auto_filter.ref = f"A1:G{n_rows + 1}"
    tx.conditional_formatting.add(f"G2:G{n_rows + 1}", CellIsRule(
        operator="equal", formula=['"?"'], fill=PatternFill("solid", fgColor="F8CBAD")))

    # ---------------- Dashboard (year input lives here)
    dash["A1"] = "Personal Finances"
    dash["A1"].font = F_TITLE
    dash["A3"] = "Year"
    dash["A3"].font = F_BOLD
    dash["B3"] = year
    dash["B3"].font, dash["B3"].fill = F_INPUT, FILL_INPUT
    dash["B3"].number_format = "0"
    dash["C3"] = "<- change the year to view another year"
    dash["C3"].font = F_NOTE
    YEAR = "Dashboard!$B$3"

    # ---------------- Summary (category x month)
    summ["A1"] = "Summary by Category"
    summ["A1"].font = F_TITLE
    summ["A2"] = "Year:"
    summ["B2"] = f"={YEAR}"
    summ["A2"].font = F_BOLD
    summ["B2"].font = F_LINK
    summ["B2"].number_format = "0"
    hdr = ["Category", "Type", "Monthly Budget"] + MONTHS + ["Year Total", "Annual Budget", "Remaining", "Used %"]
    _header(summ, 4, hdr, [24, 10, 12] + [11] * 12 + [13, 13, 13, 9])
    first, last = 5, 5 + CAT_SLOTS - 1
    for i in range(CAT_SLOTS):
        r = first + i
        cr = 2 + i
        summ.cell(row=r, column=1, value=f'=IF(Categories!$A{cr}="","",Categories!$A{cr})').font = F_LINK
        summ.cell(row=r, column=2, value=f'=IF(Categories!$A{cr}="","",Categories!$B{cr})').font = F_LINK
        c = summ.cell(row=r, column=3, value=f'=IF(Categories!$A{cr}="","",Categories!$C{cr})')
        c.font, c.number_format = F_LINK, MONEY0
        for m in range(1, 13):
            col = 3 + m
            f = (f'=IF($A{r}="","",SUMIFS(Transactions!$D:$D,Transactions!$C:$C,$A{r},'
                 f'Transactions!$A:$A,">="&DATE({YEAR},{m},1),'
                 f'Transactions!$A:$A,"<"&DATE({YEAR},{m + 1},1)))')
            c = summ.cell(row=r, column=col, value=f)
            c.font, c.number_format = F_BASE, MONEY0
        tot_col, ab_col, rem_col, pct_col = 16, 17, 18, 19
        c = summ.cell(row=r, column=tot_col, value=f'=IF($A{r}="","",SUM(D{r}:O{r}))')
        c.font, c.number_format = F_BOLD, MONEY0
        c = summ.cell(row=r, column=ab_col, value=f'=IF($A{r}="","",N(C{r})*12)')
        c.font, c.number_format = F_BASE, MONEY0
        c = summ.cell(row=r, column=rem_col,
                      value=f'=IF(OR($A{r}="",$B{r}<>"Expense",N(Q{r})=0),"",Q{r}-P{r})')
        c.font, c.number_format = F_BASE, MONEY0
        c = summ.cell(row=r, column=pct_col,
                      value=f'=IF(OR($A{r}="",$B{r}<>"Expense",N(Q{r})=0),"",P{r}/Q{r})')
        c.font, c.number_format = F_BASE, PCT
    summ.conditional_formatting.add(f"R{first}:R{last}", CellIsRule(
        operator="lessThan", formula=["0"], font=Font(name=FONT, color="C00000", bold=True)))
    summ.conditional_formatting.add(f"S{first}:S{last}", CellIsRule(
        operator="greaterThan", formula=["1"], fill=PatternFill("solid", fgColor="F8CBAD")))
    # totals block
    tr = last + 2
    for k, (label, typ) in enumerate([("Total Income", "Income"), ("Total Expenses", "Expense"),
                                      ("Total Transfers", "Transfer")]):
        r = tr + k
        summ.cell(row=r, column=1, value=label).font = F_BOLD
        for col in range(3, 18):
            L = get_column_letter(col)
            c = summ.cell(row=r, column=col,
                          value=f'=SUMIF($B${first}:$B${last},"{typ}",{L}${first}:{L}${last})')
            c.font, c.number_format, c.fill = F_BOLD, MONEY0, FILL_TOTAL
        summ.cell(row=r, column=1).fill = summ.cell(row=r, column=2).fill = FILL_TOTAL
    r = tr + 3
    summ.cell(row=r, column=1, value="Net (Income - Expenses)").font = F_BOLD
    for col in range(3, 18):
        L = get_column_letter(col)
        c = summ.cell(row=r, column=col, value=f"={L}{tr}-{L}{tr + 1}")
        c.font, c.number_format, c.border = F_BOLD, MONEY0, B_TOP
    summ.freeze_panes = "B5"
    sum_tot_row = tr  # income row; expenses tr+1; net tr+3

    # ---------------- Dashboard body
    _header(dash, 5, ["Month", "Income", "Expenses", "Net", "Savings Rate", "Budget", "vs Budget"],
            [14, 14, 14, 14, 13, 14, 14])
    for m in range(12):
        r = 6 + m
        L = get_column_letter(4 + m)  # Summary month column
        dash.cell(row=r, column=1, value=MONTHS[m]).font = F_BASE
        c = dash.cell(row=r, column=2, value=f"=Summary!{L}{sum_tot_row}")
        c.font, c.number_format = F_LINK, MONEY0
        c = dash.cell(row=r, column=3, value=f"=Summary!{L}{sum_tot_row + 1}")
        c.font, c.number_format = F_LINK, MONEY0
        c = dash.cell(row=r, column=4, value=f"=B{r}-C{r}")
        c.font, c.number_format = F_BASE, MONEY0
        c = dash.cell(row=r, column=5, value=f'=IF(B{r}=0,"",D{r}/B{r})')
        c.font, c.number_format = F_BASE, PCT
        c = dash.cell(row=r, column=6, value=f"=Summary!$C${sum_tot_row + 1}")
        c.font, c.number_format = F_LINK, MONEY0
        c = dash.cell(row=r, column=7, value=f'=IF(C{r}=0,"",F{r}-C{r})')
        c.font, c.number_format = F_BASE, MONEY0
    r = 18
    dash.cell(row=r, column=1, value="Year").font = F_BOLD
    for col, f, fmt in [(2, "=SUM(B6:B17)", MONEY0), (3, "=SUM(C6:C17)", MONEY0),
                        (4, "=B18-C18", MONEY0), (5, '=IF(B18=0,"",D18/B18)', PCT),
                        (6, "=SUM(F6:F17)", MONEY0)]:
        c = dash.cell(row=r, column=col, value=f)
        c.font, c.number_format, c.border, c.fill = F_BOLD, fmt, B_TOP, FILL_TOTAL
    dash.cell(row=r, column=1).fill = FILL_TOTAL
    dash.cell(row=r, column=7).fill = FILL_TOTAL
    dash.conditional_formatting.add("D6:D18", CellIsRule(
        operator="lessThan", formula=["0"], font=Font(name=FONT, color="C00000", bold=True)))
    dash.conditional_formatting.add("G6:G17", CellIsRule(
        operator="lessThan", formula=["0"], font=Font(name=FONT, color="C00000", bold=True)))

    # Key stats
    dash["I5"] = "At a glance"
    dash["I5"].font = F_BOLD
    stats = [
        ("Avg monthly income", '=IFERROR(AVERAGEIF(B6:B17,"<>0"),0)', MONEY0),
        ("Avg monthly spending", '=IFERROR(AVERAGEIF(C6:C17,"<>0"),0)', MONEY0),
        ("Savings rate (year)", "=E18", PCT),
        ("Uncategorized rows", '=COUNTIF(Transactions!C:C,"Uncategorized")', "0"),
        ("Transactions this year",
         f'=COUNTIFS(Transactions!A:A,">="&DATE({YEAR},1,1),Transactions!A:A,"<"&DATE({YEAR}+1,1,1))', "0"),
    ]
    for k, (label, f, fmt) in enumerate(stats):
        r = 6 + k
        dash.cell(row=r, column=9, value=label).font = F_BASE
        c = dash.cell(row=r, column=10, value=f)
        c.font, c.number_format = F_BOLD, fmt
    dash.column_dimensions["H"].width = 3
    dash.column_dimensions["I"].width = 24
    dash.column_dimensions["J"].width = 14
    dash["A20"] = ("Savings rate = Net / Income. Budget = sum of monthly expense budgets on the Categories sheet. "
                   "Transfers (e.g. credit-card payments) are excluded so they aren't double-counted.")
    dash["A20"].font = F_NOTE

    chart = BarChart()
    chart.type = "col"
    chart.title = "Income vs Expenses"
    chart.y_axis.title = "$"
    chart.add_data(Reference(dash, min_col=2, max_col=3, min_row=5, max_row=17), titles_from_data=True)
    chart.set_categories(Reference(dash, min_col=1, min_row=6, max_row=17))
    chart.height, chart.width = 8, 18
    chart.roundedCorners = False
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.numFmt = "$#,##0"
    dash.add_chart(chart, "A22")

    # ---------------- How To
    howto.column_dimensions["A"].width = 110
    lines = [
        ("How to use this workbook", F_TITLE),
        ("", F_BASE),
        ("Everything is local. This file never connects to the internet; keep it somewhere private.", F_BOLD),
        ("", F_BASE),
        ("Colors: blue text = your inputs, black = formulas, green = linked from another sheet, "
         "yellow = key setting (the year on Dashboard).", F_BASE),
        ("", F_BASE),
        ("1. Categories sheet - set your categories, their Type (Income / Expense / Transfer) and monthly budgets.", F_BASE),
        ("2. Transactions sheet - type rows in by hand, OR import bank CSVs:  python fintool.py import mybank.csv", F_BASE),
        ("   Amount: for expenses a positive number is money spent (negative = refund); for income, positive = received.", F_BASE),
        ("   The Type column fills itself in; a '?' means the category isn't on the Categories sheet.", F_BASE),
        ("3. Rules sheet - keywords used to auto-categorize imported rows. Add your own merchants.", F_BASE),
        ("4. Summary - spending per category per month, vs your annual budget.", F_BASE),
        ("5. Dashboard - monthly income, spending, savings rate and a chart. Change the year in the yellow cell.", F_BASE),
        ("", F_BASE),
        ("Tips", F_BOLD),
        ("- Mark credit-card payments and moves between your own accounts as 'Transfers' so they don't count twice.", F_BASE),
        ("- Re-importing the same CSV is safe: duplicate rows (same date, description, amount, account) are skipped.", F_BASE),
        ("- Close this file in Excel before running import or rebuild. A backup (.bak.xlsx) is made each time.", F_BASE),
        ("- After adding categories you can just keep going - Summary and Dashboard pick them up automatically.", F_BASE),
    ]
    for i, (text, font) in enumerate(lines, 1):
        c = howto.cell(row=i, column=1, value=text)
        c.font = font

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ---------------------------------------------------------------- read existing
def read_workbook(path):
    wb = load_workbook(path)  # formulas as strings; we only read input columns
    year = wb["Dashboard"]["B3"].value or dt.date.today().year
    categories = []
    for row in wb["Categories"].iter_rows(min_row=2, max_col=3, values_only=True):
        if row[0] not in (None, ""):
            categories.append((str(row[0]).strip(), row[1] or "Expense", row[2] or 0))
    rules = []
    for row in wb["Rules"].iter_rows(min_row=2, max_col=2, values_only=True):
        if row[0] not in (None, "") and row[1] not in (None, ""):
            rules.append((str(row[0]).strip(), str(row[1]).strip()))
    transactions = []
    for row in wb["Transactions"].iter_rows(min_row=2, max_col=6, values_only=True):
        if row[0] in (None, "") and row[1] in (None, ""):
            continue
        d = row[0]
        if isinstance(d, dt.datetime):
            d = d.date()
        elif isinstance(d, str):
            d = parse_date(d) or d
        transactions.append({"date": d, "desc": row[1] or "", "cat": row[2] or "Uncategorized",
                             "amount": row[3] or 0, "account": row[4], "notes": row[5]})
    return int(year), categories, rules, transactions


# ---------------------------------------------------------------- CSV import
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%Y/%m/%d",
                "%m-%d-%Y", "%d-%b-%Y", "%b %d, %Y", "%d %b %Y", "%Y%m%d"]


def parse_date(s):
    s = str(s).strip().split(" ")[0] if re.match(r"^\d{4}-\d{2}-\d{2}[ T]", str(s)) else str(s).strip()
    s = s.split("T")[0]
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def parse_amount(s):
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")") or s.startswith("-") or s.endswith("-")
    s = re.sub(r"[^\d.]", "", s)
    if not s:
        return None
    v = float(s)
    return -v if neg else v


def _find(headers, *candidates):
    low = [h.lower().strip() for h in headers]
    for cand in candidates:
        for i, h in enumerate(low):
            if h == cand:
                return i
    for cand in candidates:
        for i, h in enumerate(low):
            if cand in h:
                return i
    return None


def read_bank_csv(path, flip=False):
    """Return list of (date, description, signed_amount) where negative = money out."""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        sample = fh.read(4096)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        rows = list(csv.reader(fh, dialect))
    # find header row (first row that has a date-ish and amount-ish column)
    hi = None
    for i, row in enumerate(rows[:20]):
        if _find(row, "date", "transaction date", "posted date", "posting date") is not None and \
           _find(row, "amount", "debit", "credit", "withdrawal", "deposit") is not None:
            hi = i
            break
    if hi is None:
        raise SystemExit(f"{path}: couldn't find a header row with Date and Amount columns.")
    h = rows[hi]
    di = _find(h, "transaction date", "date", "posted date", "posting date", "trans. date")
    ni = _find(h, "description", "payee", "merchant", "name", "memo", "details", "narrative")
    ai = _find(h, "amount")
    debit_i = _find(h, "debit", "withdrawal", "money out")
    credit_i = _find(h, "credit", "deposit", "money in")
    if ai is not None and h[ai].lower().strip() not in ("amount",) and debit_i is not None:
        ai = None  # e.g. "Debit Amount" matched "amount" - use debit/credit pair
    out = []
    for row in rows[hi + 1:]:
        if not row or all(not c.strip() for c in row):
            continue
        d = parse_date(row[di]) if di is not None and di < len(row) else None
        if d is None:
            continue
        desc = " ".join(row[ni].split()) if ni is not None and ni < len(row) else ""
        if ai is not None:
            amt = parse_amount(row[ai] if ai < len(row) else None)
        else:
            deb = parse_amount(row[debit_i]) if debit_i is not None and debit_i < len(row) else None
            cre = parse_amount(row[credit_i]) if credit_i is not None and credit_i < len(row) else None
            amt = (cre or 0) - abs(deb or 0) if (deb or cre) else None
        if amt is None:
            continue
        out.append((d, desc, -amt if flip else amt))
    return out


def categorize(desc, rules, signed_amount):
    low = desc.lower()
    for kw, cat in rules:
        if kw.lower() in low:
            return cat
    return "Other Income" if signed_amount > 0 else "Uncategorized"


def _key(t):
    d = t["date"].isoformat() if hasattr(t["date"], "isoformat") else str(t["date"])
    return (d, str(t["desc"]).strip().lower(), round(float(t["amount"] or 0), 2), (t.get("account") or "").lower())


def backup(path):
    if os.path.exists(path):
        bak = os.path.splitext(path)[0] + ".bak.xlsx"
        shutil.copy2(path, bak)
        return bak


# ---------------------------------------------------------------- commands
def cmd_init(args):
    if os.path.exists(args.file) and not args.force:
        sys.exit(f"{args.file} already exists. Use 'rebuild' to refresh it, or --force to start over.")
    year = args.year or dt.date.today().year
    example = [{"date": dt.date(year, 1, 1), "desc": "EXAMPLE - Coffee shop (delete this row)",
                "cat": "Dining Out", "amount": 4.75, "account": "Checking", "notes": "example row"}]
    build_workbook(args.file, year, DEFAULT_CATEGORIES, DEFAULT_RULES, example)
    print(f"Created {args.file} for {year}. Open it in Excel, or import a CSV:\n"
          f"  python fintool.py import your_bank_export.csv")


def cmd_rebuild(args):
    year, cats, rules, txs = read_workbook(args.file)
    bak = backup(args.file)
    build_workbook(args.file, args.year or year, cats, rules, txs)
    print(f"Rebuilt {args.file} ({len(txs)} transactions). Backup: {bak}")


def cmd_import(args):
    if not os.path.exists(args.file):
        build_workbook(args.file, args.year or dt.date.today().year, DEFAULT_CATEGORIES, DEFAULT_RULES, [])
        print(f"Created new workbook {args.file}")
    year, cats, rules, txs = read_workbook(args.file)
    cat_type = {c[0]: c[1] for c in cats}
    seen = {_key(t) for t in txs}
    total_new = 0
    for csv_path in args.csv:
        account = args.account or os.path.splitext(os.path.basename(csv_path))[0]
        added = dup = 0
        for d, desc, signed in read_bank_csv(csv_path, flip=args.flip):
            cat = categorize(desc, rules, signed)
            typ = cat_type.get(cat, "Expense")
            # Expense: positive = spent. Income/Transfer: keep bank sign (positive = money in).
            amount = round(-signed if typ == "Expense" else signed, 2)
            t = {"date": d, "desc": desc, "cat": cat, "amount": amount, "account": account, "notes": None}
            k = _key(t)
            if k in seen:
                dup += 1
                continue
            seen.add(k)
            txs.append(t)
            added += 1
        total_new += added
        print(f"{csv_path}: {added} added, {dup} duplicates skipped (account '{account}')")
    if total_new:
        bak = backup(args.file)
        build_workbook(args.file, year, cats, rules, txs)
        unc = sum(1 for t in txs if t["cat"] == "Uncategorized")
        print(f"Saved {args.file}. Backup: {bak}")
        if unc:
            print(f"{unc} rows are 'Uncategorized' - filter the Category column in Excel to fix them,\n"
                  f"and add keywords on the Rules sheet so future imports catch them.")
    else:
        print("Nothing new to add.")


def cmd_summary(args):
    year, cats, rules, txs = read_workbook(args.file)
    year = args.year or year
    cat_type = {c[0]: c[1] for c in cats}
    budget = {c[0]: float(c[2] or 0) for c in cats}
    by_cat, income, spend = {}, [0.0] * 12, [0.0] * 12
    for t in txs:
        d = t["date"]
        if not hasattr(d, "year") or d.year != year:
            continue
        typ = cat_type.get(t["cat"], "Expense")
        amt = float(t["amount"] or 0)
        if typ == "Income":
            income[d.month - 1] += amt
        elif typ == "Expense":
            spend[d.month - 1] += amt
            by_cat[t["cat"]] = by_cat.get(t["cat"], 0) + amt
    print(f"\n  {year}        Income     Spending          Net")
    for m in range(12):
        if income[m] or spend[m]:
            print(f"  {MONTHS[m]:<6}{income[m]:>12,.2f} {spend[m]:>12,.2f} {income[m] - spend[m]:>12,.2f}")
    ti, ts = sum(income), sum(spend)
    print(f"  {'Total':<6}{ti:>12,.2f} {ts:>12,.2f} {ti - ts:>12,.2f}")
    if ti:
        print(f"  Savings rate: {(ti - ts) / ti:.1%}")
    months_active = max(1, sum(1 for m in range(12) if income[m] or spend[m]))
    print(f"\n  Spending by category        Spent   Budget({months_active}mo)")
    for cat, amt in sorted(by_cat.items(), key=lambda kv: -kv[1]):
        b = budget.get(cat, 0) * months_active
        flag = "  OVER" if b and amt > b else ""
        print(f"  {cat:<22}{amt:>12,.2f} {b:>12,.2f}{flag}")
    print()


def main():
    p = argparse.ArgumentParser(description="Private offline finance workbook tool.")
    p.add_argument("-f", "--file", default="Finances.xlsx", help="workbook path (default Finances.xlsx)")
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("init", help="create a new workbook")
    s.add_argument("--year", type=int)
    s.add_argument("--force", action="store_true", help="overwrite an existing file")
    s.set_defaults(func=cmd_init)
    s = sub.add_parser("import", help="import bank/card CSV exports")
    s.add_argument("csv", nargs="+")
    s.add_argument("--account", help="account label (default: CSV file name)")
    s.add_argument("--flip", action="store_true", help="flip amount signs (card exports where charges are positive)")
    s.add_argument("--year", type=int, help="dashboard year if creating a new workbook")
    s.set_defaults(func=cmd_import)
    s = sub.add_parser("rebuild", help="regenerate formulas/charts, keeping your data")
    s.add_argument("--year", type=int)
    s.set_defaults(func=cmd_rebuild)
    s = sub.add_parser("summary", help="print totals for the year")
    s.add_argument("--year", type=int)
    s.set_defaults(func=cmd_summary)
    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
