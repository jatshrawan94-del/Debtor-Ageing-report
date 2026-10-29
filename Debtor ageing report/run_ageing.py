"""Helper for Run_Ageing.bat: sorts dropped files into current and prior month by type."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tally_ageing as ta

files = [f for f in sys.argv[1:] if f.lower().endswith((".xlsx", ".xlsm"))]
if not files:
    sys.exit("No .xlsx files given. If Tally exported .xls, open it in Excel and save as .xlsx.")
by_kind = {}
for f in files:
    d = ta.parse(f)
    by_kind.setdefault(d["kind"], []).append((d["as_at"] or ta.dt.date.min, f))
current, prev = [], []
for kind, items in by_kind.items():
    items.sort()
    current.append(items[-1][1])
    if len(items) > 1:
        prev.append(items[-2][1])
sys.argv = [sys.argv[0], *current] + (["--prev", *prev] if prev else []) + ["--out", os.path.join(os.path.dirname(os.path.abspath(__file__)), "Reports")]
ta.main()
