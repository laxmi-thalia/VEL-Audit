import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
B="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-1/"
f=B+"GSTR1-Portal-Net-VIKRAN ENGINEERING LIMITED-Maharashtra-Apr 2025-Mar 2026.xlsx"
df=pd.read_excel(f, sheet_name="Sales-Net", header=0)
print(f"Sales-Net: {len(df)} rows x {df.shape[1]} cols")
for i,c in enumerate(df.columns,1): print(f"  {i:2d}. {c!r}")
print("\nDoc Type values:", df["Doc Type"].value_counts(dropna=False).to_dict())
print("Sale Type values:", df["Sale Type"].value_counts(dropna=False).to_dict())
