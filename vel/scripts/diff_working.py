import os, re, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
FILES = {
 "April": ("01 April 2025", "Cleartax Sales April 2025.xlsx", "Working", 1),
 "May":   ("02 May 2025",   "Cleartax sales may 2025.xlsx",   "working", 2),
 "June":  ("03 June 2025",  "Cleartax Sales Register June 2025.xlsx", "Working", 1),
 "Aug":   ("05 Aug 2025",   "Cleartax Sales Report August 2025.xlsx", "WORKING", 2),
}
hdrs = {}
for k,(d,f,s,h) in FILES.items():
    df = pd.read_excel(os.path.join(BASE,d,f), sheet_name=s, header=h, nrows=0)
    hdrs[k] = [str(c).strip() for c in df.columns]
    print(f"{k}: {len(hdrs[k])} cols")

# raw CT layout from April's raw sheet
raw = pd.read_excel(os.path.join(BASE,"01 April 2025","Cleartax Sales April 2025.xlsx"),
                    sheet_name="CT_e-Invoices_detailed_report_b", header=0, nrows=0)
raw = [str(c).strip() for c in raw.columns]
print(f"RAW: {len(raw)} cols")

print("\n--- do the first 146 Working cols equal the RAW layout? ---")
for k,v in hdrs.items():
    same = v[:146] == raw
    print(f"  {k}: {'IDENTICAL' if same else 'DIFFERS'}")
    if not same:
        for i,(a,b) in enumerate(zip(v[:146], raw)):
            if a != b:
                print(f"      col {i+1}: working={a!r}  raw={b!r}")

print("\n--- the 4 extra (147-150) per month ---")
for k,v in hdrs.items():
    print(f"  {k}: {v[146:]}")
