import warnings; warnings.filterwarnings("ignore")
import pandas as pd
p="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-3B/AnnualReport-VIKRAN ENGINEERING LIMITED-Maharashtra-2025-26 (1).xlsx"
df=pd.read_excel(p, sheet_name="GSTR-3B", header=None)
print("header row:", [str(v) for v in df.iloc[0].tolist() if pd.notna(v)])
print()
print("all Section / Type rows present:")
for i in range(1,len(df)):
    sec=df.iloc[i,1]; typ=df.iloc[i,2]
    if pd.isna(sec): continue
    filled=sum(1 for j in range(3,15) if pd.notna(df.iloc[i,j]))
    print(f"  r{i:2d}  {str(sec)[:56]:58s} | {str(typ)[:16]:16s} | months with a value: {filled}/12")
