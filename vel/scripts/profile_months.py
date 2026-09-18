import os, re, sys, hashlib, warnings
warnings.filterwarnings("ignore")
import pandas as pd

BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
OVERRIDE = {"12 Mar 2026": "Cleartax sales register Mar 2026.xlsx"}
months = sorted(d for d in os.listdir(BASE) if re.match(r"^\d\d ", d))
start, end = int(sys.argv[1]), int(sys.argv[2])

def norm(c): return str(c).strip().lower()
def col(df, name):
    for c in df.columns:
        if norm(c) == name.lower(): return c
def looks_header(vals):
    nn = [v for v in vals if pd.notna(v)]
    if len(nn) < 5: return False
    return sum(1 for v in nn if isinstance(v, str)) / len(nn) > 0.8

for m in months[start:end]:
    d = os.path.join(BASE, m)
    cands = [OVERRIDE[m]] if m in OVERRIDE else [
        f for f in os.listdir(d)
        if re.search(r"clear ?tax", f, re.I) and re.search(r"sale", f, re.I)
        and not f.startswith("~$")]
    print(f"\n================ {m} ================")
    for f in cands:
        p = os.path.join(d, f)
        if f.lower().endswith(".xls"):
            print(f"  {f}: LEGACY .xls -- blocked"); continue
        xl = pd.ExcelFile(p)
        ws = [s for s in xl.sheet_names if "working" in s.lower()]
        print(f"  {f}\n    sheets={xl.sheet_names}\n    working-like={ws}")
        for s in ws:
            probe = xl.parse(s, header=None, nrows=12)
            hr = next((i for i in range(len(probe)) if looks_header(probe.iloc[i].tolist())), None)
            if hr is None:
                print(f"    - {s!r}: NO HEADER ROW FOUND in first 12"); continue
            df = xl.parse(s, header=hr)
            hdrs = [norm(c) for c in df.columns]
            sig = hashlib.sha1("|".join(hdrs).encode()).hexdigest()[:12]
            dt, st, irn = col(df,"Document Type"), col(df,"Supply Type"), col(df,"Invoice Reference No")
            print(f"    - {s!r}: hdr_row={hr} rows={len(df)} cols={df.shape[1]} normsig={sig}")
            if dt is None or st is None:
                print(f"        !! missing key column(s): DocType={dt} SupplyType={st}"); continue
            print(f"        DocType : {df[dt].value_counts(dropna=False).to_dict()}")
            print(f"        Supply  : {df[st].value_counts(dropna=False).to_dict()}")
            if irn is not None:
                b = df[irn].isna() | (df[irn].astype(str).str.strip()=="")
                print(f"        blank IRN: {int(b.sum())}")
