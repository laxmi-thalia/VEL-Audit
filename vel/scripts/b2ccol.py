import os, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
T = [("01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),
     ("09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx","Working",2),
     ("12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)]
for m,f,s,h in T:
    df = pd.read_excel(os.path.join(BASE,m,f), sheet_name=s, header=h)
    print(f"\n=== {m}  rows={len(df)}")
    for c in df.columns:
        col = df[c].astype(str).str.strip().str.upper()
        n = int(col.str.contains("B2C", na=False).sum())
        if n:
            vals = col[col.str.contains("B2C", na=False)].value_counts().to_dict()
            print(f"    column {str(c)!r}: {n} rows  values={vals}")
