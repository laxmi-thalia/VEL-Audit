import os, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
F = {
 "April":("01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),
 "May":  ("02 May 2025","Cleartax sales may 2025.xlsx","working",2),
 "June": ("03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),
 "Aug":  ("05 Aug 2025","Cleartax Sales Report August 2025.xlsx","WORKING",2),
}
def norm(c): return str(c).strip().lower()
base=[norm(c) for c in pd.read_excel(BASE+"/09 Dec 2025/Cleartax Sales Register Dec 2025.xlsx",
      sheet_name="Working", header=2, nrows=0).columns]
print(f"Dec baseline: {len(base)} cols")
for k,(d,f,s,h) in F.items():
    cols=[norm(c) for c in pd.read_excel(os.path.join(BASE,d,f),sheet_name=s,header=h,nrows=0).columns]
    print(f"\n{k}: {len(cols)} cols")
    print(f"  MISSING vs Dec: {[c for c in base if c not in set(cols)]}")
    print(f"  EXTRA   vs Dec: {[c for c in cols if c not in set(base)]}")
