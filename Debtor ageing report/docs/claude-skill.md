
# Tally debtor & creditor ageing report (monthly)

Use when the user uploads a raw Tally ageing export (Sundry Debtors and/or Sundry Creditors, "Group Outstandings", age-wise) and wants the monthly ageing report / dashboard.

## Expected input (Tally export)
Gateway of Tally > Display More Reports > Statements of Accounts > Outstandings > Group > Sundry Debtors (or Sundry Creditors) > Alt+F1 detailed if needed > F6 Age-wise (e.g. <90, 90-120, 120-180, 180-365, >365) > Alt+E export to Excel.
Layout the engine relies on:
- Header block in column A (company, address, "Sundry Debtors"/"Sundry Creditors", "Group Outstandings", period "1-Jan-26 to 30-Sep-26").
- A row whose column A = "Particulars"; next row = bucket labels (first pair "Pending Bills" = total, then buckets, last usually "On Account"); next row = Debit / Credit.
- Party rows until "Grand Total". Bold, non-indented rows are sub-group subtotals (skipped); indented/italic rows are ledgers inside that sub-group.
Bucket labels are read from the file, so any F6 ageing slabs work.

## Procedure
1. Get `tally_ageing.py` and `dashboard_template.html` from this repository (same folder). Needs Python 3 + openpyxl.
2. Copy each upload into its own input folder, then run:
   `python3 tally_ageing.py DEBTORS.xlsx [CREDITORS.xlsx] [--prev PRIOR_DEBTORS.xlsx [PRIOR_CREDITORS.xlsx]] --out /mnt/user-data/outputs`
   - Type (Debtors/Creditors) and month are auto-detected from the header; prior-month files are matched by type.
   - Ask for last month's raw export(s) if the user wants month-on-month movement; never fabricate a prior month.
   - `--immaterial 1` (default) hides rounding balances under 1 from the party list (still counted in reconciliation).
3. Verify before delivering:
   - Console line must say `checks OK` (ledger total and every bucket agree to Tally Grand Total). If DIFFERENCES, inspect the Checks sheet and the Source sheet layout before delivering.
   - Recalculate with LibreOffice (`soffice --headless --convert-to xlsx --outdir recalc FILE.xlsx`) and confirm no `#` errors.
   - Screenshot the HTML (Playwright) in light and dark mode and look at it.
4. Deliver `<Debtor|Creditor>_Ageing_<Mon-YY>.xlsx` (one per type) and `Ageing_Dashboard_<Mon-YY>.html` (one dashboard, tabs per type). Summarise in 3-5 lines: gross balance, % over 90 days, advances, key exceptions, reconciliation status.

## Conventions 
- Net per bucket: Debtors = Debit - Credit; Creditors = Credit - Debit. Party-wise sheet shows net buckets.
- Dashboard splits parties into: balance due (gross receivables/payables), advances (opposite balance) and nil-net parties with unadjusted bills, so overdue KPIs are not understated by netting. Net column always agrees to Tally.
- Overdue = every aged bucket after the first; ">180" = buckets whose label mentions 180 or 365.
- Exception flags: advance/opposite balance; >50% older than 180 days; Dr & Cr sitting in different buckets (needs bill-wise knock-off in Tally); provision/ECL ledgers under debtors (classification review); foreign-currency ledgers (e.g. "- USD"); immaterial balances.
- Workbook sheets: Dashboard (live SUMIF formulas to Ageing Detail + 2 charts), Ageing Detail (formulas for totals, %), Advances / Debit balances, Movement (only with --prev), Checks, Source (raw copy).
- Amounts in base currency, format #,##0 with negatives in brackets.

