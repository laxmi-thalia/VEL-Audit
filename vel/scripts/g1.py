import warnings; warnings.filterwarnings("ignore")
import pandas as pd
B="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-1/"
for st in ["Maharashtra","Andhra Pradesh"]:
    p=B+f"GSTR1-Portal-Net-VIKRAN ENGINEERING LIMITED-{st}-Apr 2025-Mar 2026.xlsx"
    xl=pd.ExcelFile(p)
    print(f"\n=== {st}: {xl.sheet_names}")
    for s in xl.sheet_names:
        if "b2c" in s.lower().replace(" ",""):
            df=xl.parse(s, header=None)
            print(f"    {s!r}: shape={df.shape}")
