import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
B="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-1/"
tot=0
for f in sorted(os.listdir(B)):
    if not f.endswith(".xlsx"): continue
    st=f.split("LIMITED-")[1].rsplit("-Apr",1)[0]
    try:
        df=pd.read_excel(B+f, sheet_name="SalesSummary-Net")
    except Exception as e:
        print(f"{st}: FAILED {e}"); continue
    n=len(df)
    tot+=n
    types = df["Summary Type"].astype(str).str.strip().value_counts().to_dict() if n and "Summary Type" in df.columns else {}
    print(f"{st:22s} rows={n:4d}  SummaryType={types}")
print(f"\nTOTAL summary-level rows across states: {tot}")
