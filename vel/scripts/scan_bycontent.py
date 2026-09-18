import openpyxl, os, re, hashlib, sys, warnings
warnings.filterwarnings("ignore")
import pandas as pd

BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
REF_FIRST = ["GSTIN", "Invoice Reference No", "Acknowledgment No"]  # anchor of the CT raw layout
months = sorted(d for d in os.listdir(BASE) if re.match(r"^\d\d ", d))
start, end = int(sys.argv[1]), int(sys.argv[2])

def sig(hdr):
    return hashlib.sha1("|".join(str(h) for h in hdr).encode()).hexdigest()[:12]

def looks_raw(hdr):
    s = [str(h).strip() for h in hdr[:3]]
    return s == REF_FIRST

for m in months[start:end]:
    d = os.path.join(BASE, m)
    cands = [f for f in os.listdir(d)
             if re.search(r"clear ?tax", f, re.I) and re.search(r"sale", f, re.I)
             and not f.startswith("~$")]
    print(f"\n### {m}")
    for f in cands:
        p = os.path.join(d, f)
        print(f"  FILE {f}")
        try:
            sheets = pd.read_excel(p, sheet_name=None, header=None, nrows=1)
        except Exception as e:
            print(f"    LOAD FAILED {type(e).__name__}: {e}"); continue
        full = None
        for name, df in sheets.items():
            if df.empty: 
                print(f"    - {name!r}: empty first row"); continue
            hdr = [h for h in df.iloc[0].tolist() if pd.notna(h)]
            mark = "  <== RAW CT LAYOUT" if looks_raw(hdr) else ""
            print(f"    - {name!r}: {len(hdr)} headers sig={sig(hdr)}{mark}")
            if looks_raw(hdr): full = name
        if full:
            df = pd.read_excel(p, sheet_name=full, header=0)
            print(f"    >>> RAW SHEET {full!r}: data_rows={len(df)} cols={df.shape[1]}")
