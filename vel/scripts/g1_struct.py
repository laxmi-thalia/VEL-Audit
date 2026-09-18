import warnings; warnings.filterwarnings("ignore")
import pandas as pd
B="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-1/"
f=B+"GSTR1-Portal-Net-VIKRAN ENGINEERING LIMITED-Maharashtra-Apr 2025-Mar 2026.xlsx"
xl=pd.ExcelFile(f); print("sheets:", xl.sheet_names)
for s in ["Overview","Sales-Net"]:
    df=xl.parse(s, header=None, nrows=14)
    print(f"\n=== {s!r} shape(first14)={df.shape}")
    for i in range(len(df)):
        row=[str(v) for v in df.iloc[i].tolist() if pd.notna(v)]
        if row: print(f"  r{i}: {row[:12]}")
