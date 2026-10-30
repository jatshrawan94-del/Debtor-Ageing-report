# Debtor & Creditor Ageing Report (Tally)

Turns raw **Tally "Group Outstandings" age-wise exports** into the house-format ageing workbook:

| Sheet | Contents |
|---|---|
| **Debtors Dashboard** | KPI tiles (total receivables, customer advances, net receivable, >90 days), commentary, *Receivables (net debit balances)* and *Customer advances (net credit balances)* with ageing buckets and **Amount in USD** |
| **Creditor Dashboard** | Same for payables / advances paid, grouped by Tally sub-group (Fish, Saudi, Service, Trade suppliers…) |
| DR_Row Data / CR_Row Data | The Tally export as received (hidden, like the house file) |
| **FX Rates** | AED per unit for each currency, with source and date. Yellow input cells drive every converted amount |
| **Checks** | Ledger rows vs Tally Grand Total for every bucket, plus exceptions (advances, unadjusted bills, provision ledgers, foreign-currency ledgers) |

Also produced: an interactive HTML dashboard (KPIs, ageing chart, top exposures, searchable party table, month-on-month movement).

Both Tally export layouts are read automatically:
- **Single column per bucket**, Dr/Cr shown in the cell format (`0.00 Dr` / `0.00 Cr`)
- **Debit and Credit columns per bucket**

## Foreign exchange (Central Bank of the UAE)

For each currency in `config.json` → `"currencies"` (default `["USD"]`), the rate is taken in this order:

1. Manual override: `"fx_override": {"EUR": 4.2851}` in `config.json`, or `--fx EUR=4.2851`
2. **CBUAE published rate for the report date**, downloaded from centralbank.ae (best effort, see below)
3. `fx_rates.csv`: your own table (`Date,Currency,AED_per_unit,Source`). Successful CBUAE downloads are saved here too.
4. USD only: the CBUAE peg, **3.6725 AED**

The CBUAE site loads its rate table from an internal page endpoint, not a documented API. If it changes, step 2 fails quietly and steps 3–4 are used. The FX Rates sheet always shows which source and date were used. Converted amounts are formulas (`Pending Bills ÷ rate`), so you can correct a rate in Excel and everything updates.

## Run it

### Windows launcher (easiest)

1. Install Python from python.org and tick **Add Python to PATH**.
2. Export from Tally: Display More Reports → Statements of Accounts → Outstandings → Group → *Sundry Debtors* / *Sundry Creditors* → **F6 Age-wise** → **Alt+E** → Excel (.xlsx).
3. **Drag the export(s) onto `Run_Ageing.bat`.** The report and dashboard are saved in `Reports\` and opened.

Double-click `Run_Ageing.bat` with no files to run the demo in `sample\`.

### Command line

```bash
pip install -r requirements.txt
python tally_ageing.py "Debtors.xlsx" "Creditors.xlsx"            # or one workbook with both sheets
python tally_ageing.py "Sep.xlsx" --prev "Aug.xlsx"               # month-on-month in the dashboard
python tally_ageing.py "Sep.xlsx" --currencies USD,EUR --fx EUR=4.2851
python tally_ageing.py "Sep.xlsx" --offline --detail --out "D:\Reports"
```

Output name: `Creditor_Debtors ageing report_DD.MM.YYYY.xlsx` (or `Debtors…` / `Creditors…` when only one type), dated from the Tally period end.

### Browser (no install)

Open `tester/index.html`, or the GitHub Pages link `…/tester/`. Drop the exports, set FX rates in the yellow boxes (USD defaults to 3.6725), and click **Download Excel report**. It produces the same workbook as the launcher. Files are read in the browser and never uploaded. The browser can't download CBUAE rates itself, so enter other currencies' rates by hand.

## config.json

```json
{
  "currencies": ["USD"],          // add "EUR", "GBP", "INR"… for more "Amount in …" columns
  "fx_override": {},              // e.g. {"EUR": 4.2851} to force a rate
  "offline": false,               // true = don't try CBUAE online
  "immaterial": 1,                // balances below this are treated as rounding (hidden)
  "output_folder": "Reports"
}
```
(Comments shown for explanation only; the real file is plain JSON.)

## Rules used

| Item | Rule |
|---|---|
| Section | Party net balance: debtors Dr → *Receivables*, Cr → *Customer advances*; creditors Cr → *Payables*, Dr → *Advances paid*. Advances are shown as positive amounts |
| >90 days | Sum of every bucket after "< 90 days", excluding On Account |
| Groups | Tally sub-groups (bold rows) become section headers; ledgers directly under Sundry Creditors appear under **OTHER SUPPLIERS** |
| Rounding | Bucket residues below 1 (e.g. 0.04) are shown as "-" |
| Totals | `SUM` formulas over the listed rows; KPI tiles link to the totals |
| Checks | Every bucket must agree to Tally's Grand Total (tolerance 1). A difference means rows were removed or edited after export |

## Files

| File | Purpose |
|---|---|
| `Run_Ageing.bat` | Windows drag-and-drop launcher |
| `run_ageing.py` | Launcher helper (runs the report, opens results) |
| `tally_ageing.py` | Reader for both Tally layouts, analysis, HTML dashboard, CLI |
| `report_format.py` | House-format Excel workbook |
| `fx.py` | CBUAE / fx_rates.csv / peg exchange rates |
| `config.json`, `fx_rates.csv` | Settings and your rate table |
| `dashboard_template.html` | Template for the HTML dashboard |
| `tester/index.html` | Browser version (same Excel output) |
| `sample/` | Fictional demo exports (both layouts) |

## Data privacy

Don't commit real Tally exports or generated reports. `.gitignore` excludes `*.xlsx` and `*.html` outside `sample/` and `tester/`, but only when you use Git or GitHub Desktop. It doesn't apply to files dragged onto GitHub's web upload page.
