import warnings; warnings.filterwarnings("ignore")
import pandas as pd
p="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-1/GSTR1-Portal-Net-VIKRAN ENGINEERING LIMITED-Maharashtra-Apr 2025-Mar 2026.xlsx"
df=pd.read_excel(p, sheet_name="SalesSummary-Net", header=None)
print("SalesSummary-Net shape:", df.shape)
for i in range(min(40,len(df))):
    row=[str(v) for v in df.iloc[i].tolist() if pd.notna(v)]
    if row: print(f"  r{i}: {row[:6]}")
