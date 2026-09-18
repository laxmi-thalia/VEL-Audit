import warnings; warnings.filterwarnings("ignore")
import pandas as pd
p="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-3B/AnnualReport-VIKRAN ENGINEERING LIMITED-Maharashtra-2025-26 (1).xlsx"
xl=pd.ExcelFile(p); print("sheets:", xl.sheet_names)
for s in xl.sheet_names[:4]:
    df=xl.parse(s, header=None)
    print(f"\n=== {s!r} shape={df.shape}")
    for i in range(min(22,len(df))):
        row=[(j,str(v)[:26]) for j,v in enumerate(df.iloc[i].tolist()) if pd.notna(v)]
        if row: print(f"  r{i:2d}: " + " | ".join(f"c{j}:{v}" for j,v in row[:8]))
