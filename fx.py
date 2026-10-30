"""
Foreign-exchange rates for the ageing report (AED per 1 unit of foreign currency).

Order of precedence for each currency:
  1. --fx CODE=RATE on the command line / "fx_override" in config.json
  2. Central Bank of the UAE (CBUAE) published rate for the report date (online; best effort)
  3. fx_rates.csv next to this file (your own maintained table; also a cache of CBUAE downloads)
  4. Built-in: USD 3.6725 - the AED/USD peg published by CBUAE

The CBUAE website loads its rate table through an internal page endpoint that is not a
documented API, so step 2 may stop working if the site changes. When it fails the report
still runs on steps 3-4 and the FX Rates sheet says which source was used. Rates on the
FX Rates sheet are input cells: change one there and every converted amount updates.
"""
import csv, datetime as dt, html, os, re, ssl, urllib.request

PEG = {"USD": 3.6725}
HERE = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(HERE, "fx_rates.csv")

CBUAE_PAGE = "https://www.centralbank.ae/en/forex-eibor/exchange-rates/"
CBUAE_ENDPOINTS = [
    "https://www.centralbank.ae/umbraco/Surface/Exchange/GetExchangeRateAllCurrencyDate?dateTime={iso}",
    "https://www.centralbank.ae/umbraco/Surface/Exchange/GetExchangeRateAllCurrency",
]

NAMES = {  # CBUAE currency names -> ISO codes
    "us dollar": "USD", "euro": "EUR", "gb pound": "GBP", "british pound": "GBP", "pound sterling": "GBP",
    "indian rupee": "INR", "saudi riyal": "SAR", "chinese yuan": "CNY", "chinese renminbi": "CNY",
    "japanese yen": "JPY", "swiss franc": "CHF", "canadian dollar": "CAD", "australian dollar": "AUD",
    "kuwaiti dinar": "KWD", "qatari riyal": "QAR", "omani riyal": "OMR", "bahraini dinar": "BHD",
    "pakistani rupee": "PKR", "south african rand": "ZAR", "nigerian naira": "NGN", "ghanaian cedi": "GHS",
    "cfa franc": "XOF", "west african cfa franc": "XOF", "singapore dollar": "SGD", "hong kong dollar": "HKD",
    "turkish lira": "TRY", "egyptian pound": "EGP", "kenyan shilling": "KES", "korean won": "KRW",
    "south korean won": "KRW", "malaysian ringgit": "MYR", "thai baht": "THB", "russian ruble": "RUB",
    "bangladesh taka": "BDT", "sri lankan rupee": "LKR", "new zealand dollar": "NZD",
}

def _read_csv():
    rows = []
    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                try:
                    rows.append({"date": r.get("Date", "").strip(), "code": r["Currency"].strip().upper(),
                                 "rate": float(r["AED_per_unit"]), "source": r.get("Source", "fx_rates.csv").strip()})
                except (KeyError, ValueError):
                    continue
    return rows

def _append_csv(date, rates, source):
    new = not os.path.exists(CSV_PATH)
    try:
        with open(CSV_PATH, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new: w.writerow(["Date", "Currency", "AED_per_unit", "Source"])
            for k, v in sorted(rates.items()): w.writerow([date, k, v, source])
    except OSError:
        pass

def _parse_cbuae(text):
    """Pull (currency name, rate) pairs out of the CBUAE rate table HTML."""
    out = {}
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", text, re.S | re.I):
        cells = [html.unescape(re.sub(r"<[^>]+>", " ", c)).strip()
                 for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S | re.I)]
        cells = [re.sub(r"\s+", " ", c) for c in cells if c]
        if len(cells) < 2: continue
        name = cells[0].lower()
        code = NAMES.get(name) or (cells[0].upper() if re.fullmatch(r"[A-Z]{3}", cells[0]) else None)
        if not code: continue
        nums = [float(x) for c in cells[1:] for x in re.findall(r"\d+\.\d+|\d+", c.replace(",", ""))]
        if nums: out[code] = nums[0]
    return out

def fetch_cbuae(date, timeout=12):
    ctx = ssl.create_default_context()
    hdrs = {"User-Agent": "Mozilla/5.0 (ageing-report)", "Referer": CBUAE_PAGE, "X-Requested-With": "XMLHttpRequest"}
    for url in CBUAE_ENDPOINTS:
        try:
            req = urllib.request.Request(url.format(iso=date.isoformat()), headers=hdrs)
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                rates = _parse_cbuae(r.read().decode("utf-8", "replace"))
            if "USD" in rates and abs(rates["USD"] - PEG["USD"]) < 0.01:   # sanity check against the peg
                return rates
        except Exception:
            continue
    return {}

def get_rates(currencies, date, overrides=None, online=True, log=print):
    """Return {code: {"rate": float, "source": str, "date": str}} for the requested currencies."""
    date = date or dt.date.today()
    overrides = {k.upper(): float(v) for k, v in (overrides or {}).items()}
    want = [c.upper() for c in currencies]
    res = {}
    for c in want:
        if c in overrides: res[c] = {"rate": overrides[c], "source": "Manual override", "date": date.isoformat()}
    todo = [c for c in want if c not in res]
    if todo and online:
        live = fetch_cbuae(date)
        if live:
            _append_csv(date.isoformat(), {k: v for k, v in live.items()}, "CBUAE")
            for c in todo:
                if c in live: res[c] = {"rate": live[c], "source": "Central Bank of the UAE", "date": date.isoformat()}
        else:
            log("  FX: could not download CBUAE rates (offline or site changed) - using fx_rates.csv / USD peg.")
    todo = [c for c in want if c not in res]
    if todo:
        table = _read_csv()
        for c in todo:
            hits = [r for r in table if r["code"] == c]
            if hits:
                same = [r for r in hits if r["date"] == date.isoformat()]
                best = same[-1] if same else sorted(hits, key=lambda r: r["date"])[-1]
                res[c] = {"rate": best["rate"], "source": best["source"] + " (fx_rates.csv)", "date": best["date"]}
    for c in [c for c in want if c not in res]:
        if c in PEG:
            res[c] = {"rate": PEG[c], "source": "CBUAE official peg", "date": date.isoformat()}
        else:
            log(f"  FX: no rate for {c}. Add it to fx_rates.csv or pass --fx {c}=RATE. Column skipped.")
    return {c: res[c] for c in want if c in res}
