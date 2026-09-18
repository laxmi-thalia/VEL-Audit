import os, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
F = {
 "Dec(baseline)": ("09 Dec 2025", "Cleartax Sales Register Dec 2025.xlsx", "Working", 2),
 "Nov":           ("08 Nov 2025", "Cleartax Sales Reg Nov 2025.xlsx",      "Working", 2),
 "Jan":           ("10 Jan 2026", "Cleartax sales register Jan 2026.xlsx", "Working", 2),
}
h = {}
for k,(d,f,s,hr) in F.items():
    cols = pd.read_excel(os.path.join(BASE,d,f), sheet_name=s, header=hr, nrows=0).columns
    h[k] = [str(c).strip().lower() for c in cols]
    print(f"{k}: {len(h[k])} cols")
base = h["Dec(baseline)"]
for k in ("Nov","Jan"):
    miss = [c for c in base if c not in set(h[k])]
    extra= [c for c in h[k] if c not in set(base)]
    print(f"\n=== {k} vs Dec baseline ===")
    print(f"  MISSING in {k}: {miss}")
    print(f"  EXTRA   in {k}: {extra}")
