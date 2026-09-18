import os, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
FILES = {
 "April": ("01 April 2025", "Cleartax Sales April 2025.xlsx", "Working", 1),
 "May":   ("02 May 2025",   "Cleartax sales may 2025.xlsx",   "working", 2),
 "June":  ("03 June 2025",  "Cleartax Sales Register June 2025.xlsx", "Working", 1),
 "Aug":   ("05 Aug 2025",   "Cleartax Sales Report August 2025.xlsx", "WORKING", 2),
}
def norm(c): return str(c).strip().lower()
def col(df, name):
    for c in df.columns:
        if norm(c) == name.lower(): return c
    return None

for k,(d,f,s,h) in FILES.items():
    df = pd.read_excel(os.path.join(BASE,d,f), sheet_name=s, header=h)
    irn = col(df,"Invoice Reference No"); dt = col(df,"Document Type")
    st  = col(df,"Supply Type");          dn = col(df,"Document Number")
    print(f"\n================ {k}  ({len(df)} rows) ================")
    print(f"  Document Type  : {df[dt].value_counts(dropna=False).to_dict()}")
    print(f"  Supply Type    : {df[st].value_counts(dropna=False).to_dict()}")
    blank_irn = df[irn].isna() | (df[irn].astype(str).str.strip()=="")
    print(f"  rows with BLANK IRN: {int(blank_irn.sum())}  / populated: {int((~blank_irn).sum())}")
    if blank_irn.any():
        sub = df[blank_irn]
        print(f"     their Document Type: {sub[dt].value_counts(dropna=False).to_dict()}")
        print(f"     their Supply Type  : {sub[st].value_counts(dropna=False).to_dict()}")
        print(f"     blank Document Number too: {int(sub[dn].isna().sum())}")
