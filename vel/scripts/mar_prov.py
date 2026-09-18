import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
D="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data/12 Mar 2026/"
def n(c): return str(c).strip().lower()
def looks_header(v):
    nn=[x for x in v if pd.notna(x)]
    return len(nn)>=5 and sum(1 for x in nn if isinstance(x,str))/len(nn)>0.8
for f in ["Cleartax sales register Mar 2026.xlsx","Cleartax Sales prov.xlsx"]:
    try:
        xl=pd.ExcelFile(D+f)
    except Exception as e:
        print(f"{f}: LOAD FAILED {e}"); continue
    ws=[s for s in xl.sheet_names if "working" in s.lower()]
    print(f"\n{f}\n   sheets={xl.sheet_names}")
    for s in ws:
        pr=xl.parse(s,header=None,nrows=12)
        hr=next((i for i in range(len(pr)) if looks_header(pr.iloc[i].tolist())),None)
        if hr is None: print(f"   {s}: no header found"); continue
        d=xl.parse(s,header=hr); cm={n(c):c for c in d.columns}
        tv=pd.to_numeric(d[cm["item assessable amount"]],errors="coerce").fillna(0).sum() if "item assessable amount" in cm else None
        print(f"   {s!r}: rows={len(d)} cols={d.shape[1]} taxable={tv:,.2f}" if tv is not None else f"   {s!r}: rows={len(d)}")
