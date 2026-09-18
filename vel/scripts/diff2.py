import os, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
FILES = {
 "April": ("01 April 2025", "Cleartax Sales April 2025.xlsx", "Working", 1),
 "May":   ("02 May 2025",   "Cleartax sales may 2025.xlsx",   "working", 2),
 "June":  ("03 June 2025",  "Cleartax Sales Register June 2025.xlsx", "Working", 1),
 "Aug":   ("05 Aug 2025",   "Cleartax Sales Report August 2025.xlsx", "WORKING", 2),
}
raw = [str(c).strip() for c in pd.read_excel(
    os.path.join(BASE,"01 April 2025","Cleartax Sales April 2025.xlsx"),
    sheet_name="CT_e-Invoices_detailed_report_b", header=0, nrows=0).columns]
rawset = set(raw)
hdrs = {}
for k,(d,f,s,h) in FILES.items():
    hdrs[k] = [str(c).strip() for c in pd.read_excel(os.path.join(BASE,d,f), sheet_name=s, header=h, nrows=0).columns]

for k,v in hdrs.items():
    extra   = [ (i+1,c) for i,c in enumerate(v) if c not in rawset ]
    missing = [ c for c in raw if c not in set(v) ]
    firstdiff = next((i+1 for i,(a,b) in enumerate(zip(v,raw)) if a!=b), None)
    print(f"\n=== {k}: {len(v)} cols | first positional divergence at col {firstdiff}")
    print(f"    ADDED (not in raw CT layout): {extra}")
    print(f"    MISSING from working:        {missing}")

print("\n=== pairwise: are the working header LISTS equal across months? ===")
ks = list(hdrs)
for i in range(len(ks)):
    for j in range(i+1,len(ks)):
        a,b = hdrs[ks[i]], hdrs[ks[j]]
        if a==b: print(f"    {ks[i]} == {ks[j]}")
        else:
            fd = next((n+1 for n,(x,y) in enumerate(zip(a,b)) if x!=y), None)
            print(f"    {ks[i]} != {ks[j]}  first diff col {fd}: {a[fd-1] if fd else None!r} vs {b[fd-1] if fd else None!r}")
