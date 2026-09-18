import warnings; warnings.filterwarnings("ignore")
import pandas as pd
p="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/Tax Comparison Reports/2025-26_19AAECR0503Q1Z5_Tax liability and ITC comparison.xlsx"
xl=pd.ExcelFile(p)
print("sheets:", xl.sheet_names)
for s in xl.sheet_names:
    df=xl.parse(s, header=None)
    print(f"\n=== {s!r} shape={df.shape}")
    for i in range(min(18,len(df))):
        row=[str(v) for v in df.iloc[i].tolist() if pd.notna(v)]
        if row: print(f"  r{i}: {row[:9]}")
