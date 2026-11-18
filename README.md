# Expense Dashboard

Static dashboard (HTML + Chart.js) built from the `EXPENSES.xlsx` pivot. No login, no build step.

## Pages
- **Dashboard**: a card per category (total + month-by-month trend line), Monthly trend bars with a dropdown to pick any expense, Expense by category bars (click a bar for the breakup)
- **Table view**: category totals by month, click a row to see its items

Filters: period, category, breakup item, search.

## Update the data
```
pip install openpyxl
python scripts/build_data.py path/to/EXPENSES.xlsx
```
Reads the pivot (bold rows = category totals, indented rows = breakup), checks every category equals the sum of its items and that all categories equal the Grand Total, then rewrites `data/expenses.json` and the data inside `index.html`. Commit and push.

Owner's Injection, Loan and Opening Deposit are treated as funding (money in); everything else is expense.

## Deploy
Vercel: import the repo, no build step. Or open `index.html` locally.
