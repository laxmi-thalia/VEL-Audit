import os, re, sys, warnings, hashlib
warnings.filterwarnings("ignore")
import pandas as pd

BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
MAR_OVERRIDE = {"12 Mar 2026": "Cleartax sales register Mar 2026.xlsx"}
months = sorted(d for d in os.listdir(BASE) if re.match(r"^\d\d ", d))
start, end = int(sys.argv[1]), int(sys.argv[2])

def looks_header(vals):
    """A header row: mostly non-null strings, few numbers."""
    nn = [v for v in vals if pd.notna(v)]
    if len(nn) < 5: return False
    strs = sum(1 for v in nn if isinstance(v, str))
    return strs / len(nn) > 0.8

for m in months[start:end]:
    d = os.path.join(BASE, m)
    if m in MAR_OVERRIDE:
        cands = [MAR_OVERRIDE[m]]
    else:
        cands = [f for f in os.listdir(d)
                 if re.search(r"clear ?tax", f, re.I) and re.search(r"sale", f, re.I)
                 and not f.startswith("~$")]
    print(f"\n### {m}")
    for f in cands:
        p = os.path.join(d, f)
        if f.lower().endswith(".xls"):
            print(f"  {f}: LEGACY .xls -- still blocked"); continue
        try:
            xl = pd.ExcelFile(p)
        except Exception as e:
            print(f"  {f}: LOAD FAILED {type(e).__name__}: {e}"); continue
        wsheets = [s for s in xl.sheet_names if "working" in s.lower()]
        print(f"  {f}\n    all sheets: {xl.sheet_names}\n    working-like: {wsheets}")
        for s in wsheets:
            probe = xl.parse(s, header=None, nrows=12)
            hdr_row = None
            for i in range(len(probe)):
                if looks_header(probe.iloc[i].tolist()):
                    hdr_row = i; break
            print(f"    - {s!r}: probe_shape={probe.shape} header_row_index={hdr_row}")
            if hdr_row is None:
                for i in range(min(6, len(probe))):
                    nn = probe.iloc[i].notna().sum()
                    print(f"        row{i}: {nn} non-empty")
                continue
            full = xl.parse(s, header=hdr_row)
            hdrs = [str(c) for c in full.columns]
            sig = hashlib.sha1("|".join(hdrs).encode()).hexdigest()[:12]
            print(f"        data_rows={len(full)} cols={full.shape[1]} sig={sig}")
