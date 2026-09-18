import warnings; warnings.filterwarnings("ignore")
import pandas as pd
def S(v): return "" if pd.isna(v) else str(v).strip()
p="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/Financials/Statewise PnL FY25-26.xlsx"
xl=pd.ExcelFile(p); print("sheets:", xl.sheet_names)
for s in xl.sheet_names[:3]:
    d=xl.parse(s, header=None)
    print("\n=== %r shape=%s ===" % (s, str(d.shape)))
    for i in range(min(25, len(d))):
        row=[(j, str(v)[:24]) for j, v in enumerate(d.iloc[i].tolist()) if pd.notna(v)]
        if row: print("r%02d: %s" % (i, " | ".join("c%d:%s" % x for x in row[:12])))
