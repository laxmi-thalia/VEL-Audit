import warnings; warnings.filterwarnings("ignore")
import pandas as pd
p="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/Tax Comparison Reports/2025-26_19AAECR0503Q1Z5_Tax liability and ITC comparison.xlsx"
for s in ["Tax Liability Summary","Tax liability"]:
    df=pd.read_excel(p, sheet_name=s, header=None)
    print(f"\n=== {s!r} (rows {len(df)}) — Jan/Feb/Mar + total")
    for i in range(len(df)):
        cell=str(df.iloc[i,0])
        if any(k in cell for k in ["Jan-26","Feb-26","Mar-26","Total","TOTAL"]):
            row=[str(v) for v in df.iloc[i].tolist() if pd.notna(v)]
            print(f"  r{i}: {row[:10]}")
