"""Helper for Run_Ageing.bat: runs the report on the dropped files and opens the results."""
import os, sys, traceback
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tally_ageing as ta

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    files = [f for f in sys.argv[1:] if f.lower().endswith((".xlsx", ".xlsm"))]
    skipped = [f for f in sys.argv[1:] if f not in files]
    for f in skipped:
        print(f"Skipped {os.path.basename(f)}: only .xlsx is supported. If Tally exported .xls, open it in Excel and save as .xlsx.")
    if not files:
        print("No Tally files given - running the demo files in the sample folder.")
        files = [os.path.join(here, "sample", f) for f in sorted(os.listdir(os.path.join(here, "sample"))) if f.endswith(".xlsx")]
    xl, hp = ta.main(files)
    if os.name == "nt":
        os.startfile(xl); os.startfile(hp)

if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0): print(f"\nStopped: {e.code}"); sys.exit(1)
    except Exception:
        traceback.print_exc(); sys.exit(1)
