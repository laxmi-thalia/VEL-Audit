import os, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
base = [str(c).strip().lower() for c in pd.read_excel(
    os.path.join(BASE,"09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx"),
    sheet_name="Working", header=2, nrows=0).columns]
mar = [str(c).strip().lower() for c in pd.read_excel(
    os.path.join(BASE,"12 Mar 2026","Cleartax sales register Mar 2026.xlsx"),
    sheet_name="Working", header=1, nrows=0).columns]
print(f"Dec baseline={len(base)}  Mar={len(mar)}")
print(f"  MISSING in Mar: {[c for c in base if c not in set(mar)]}")
print(f"  EXTRA   in Mar: {[c for c in mar if c not in set(base)]}")
