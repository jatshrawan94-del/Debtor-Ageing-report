# Debtor & Creditor Ageing Report (Tally)

Turns a raw **Tally "Group Outstandings" age-wise export** into a monthly ageing report:

- **Excel workbook**: Dashboard (live formulas + charts), party-wise Ageing Detail, Advances, Movement vs prior month, Checks (reconciliation to Tally Grand Total + exceptions), Source.
- **HTML dashboard**: KPIs, ageing chart, top-10 exposures, searchable/sortable party table, month-on-month movement, reconciliation.
- **Browser tester** (`tester/index.html`): drop any export and see the dashboard instantly. Files are read in the browser and are not uploaded.

Works for **Sundry Debtors and Sundry Creditors**, any company, and any F6 ageing slabs (detected from the file).

## 1. Export from Tally

Display More Reports → Statements of Accounts → Outstandings → Group → *Sundry Debtors* (or *Sundry Creditors*) → **F6 Age-wise** (set your periods) → **Alt+E** → Excel (.xlsx).

If Tally gives `.xls`, open it in Excel and save as `.xlsx`.

## 2. Run it

**Windows, easiest:** drag one or more Tally exports onto `Run_Ageing.bat`. It installs `openpyxl` if needed, picks the latest month of each type as the report (an earlier month becomes the comparison), saves to `Reports\` and opens the dashboard. Double-click it with no files to run the demo.

**Command line:**

```bash
pip install -r requirements.txt

# Debtors only
python tally_ageing.py "Debtors Sep-26.xlsx" --out Reports

# Debtors + creditors in one dashboard
python tally_ageing.py "Debtors Sep-26.xlsx" "Creditors Sep-26.xlsx" --out Reports

# With month-on-month movement (prior files matched by type)
python tally_ageing.py "Debtors Sep-26.xlsx" "Creditors Sep-26.xlsx" \
  --prev "Debtors Aug-26.xlsx" "Creditors Aug-26.xlsx" --out Reports

# Try it with the demo file
python tally_ageing.py sample/Sample_Debtors_Sep-26.xlsx --out Reports
```

Output: `Debtor_Ageing_<Mon-YY>.xlsx`, `Creditor_Ageing_<Mon-YY>.xlsx`, `Ageing_Dashboard_<Mon-YY>.html`.

Option: `--immaterial 5` hides balances below 5 (default 1) from the party list. They are still counted in the reconciliation.

## 3. Or use the browser tester

Open `tester/index.html` (or the GitHub Pages URL if Pages is enabled for this repo). It loads the demo file. Then drag in your own exports. Add last month's file of the same type and it becomes the comparison automatically. **Download Excel** saves a summary workbook.

## Conventions

| Item | Rule |
|---|---|
| Net balance | Debtors = Debit − Credit; Creditors = Credit − Debit (per bucket) |
| Gross receivables / payables | Parties with a balance due. Advances (opposite balances) are shown separately so overdue KPIs aren't understated by netting |
| Overdue > 90 | Every aged bucket after the first |
| Older than 180 | Buckets whose label mentions 180 or 365 |
| Sub-groups | Tally sub-group subtotal rows are detected and excluded so totals aren't doubled |
| Checks | Ledger total and every bucket must agree to Tally's Grand Total (tolerance 1) |

Exception flags: advance/opposite balance · >50% older than 180 days · Dr & Cr in different buckets (needs bill-wise knock-off) · provision/ECL ledger under debtors · foreign-currency ledger · immaterial balance.

## Troubleshooting

- **`checks DIFFERENCES`**: open the *Checks* sheet. Usually the export is missing the Grand Total row or was edited by hand.
- **`Could not find bucket header`**: the report wasn't exported age-wise (F6).
- **Permission error**: close the file in Excel and run again.

## Data privacy

Don't commit real Tally exports or generated reports. `.gitignore` excludes `*.xlsx`/`*.html` outside `sample/` and `tester/`. `sample/` contains fictional demo data only.

## Files

| File | Purpose |
|---|---|
| `Run_Ageing.bat` | Windows drag-and-drop launcher |
| `run_ageing.py` | Sorts dropped files into current / prior month for the launcher |
| `tally_ageing.py` | Engine: parse → analyse → Excel + HTML |
| `dashboard_template.html` | Template for the generated HTML dashboard (keep next to the script) |
| `tester/index.html` | Self-contained browser tester |
| `sample/Sample_Debtors_Sep-26.xlsx` | Fictional demo export in Tally's layout |
| `docs/claude-skill.md` | Instructions for running this monthly with Claude |
