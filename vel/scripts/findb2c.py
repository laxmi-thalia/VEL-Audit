import os, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
TARGETS = [("01 April 2025","Cleartax Sales April 2025.xlsx"),
           ("09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx"),
           ("12 Mar 2026","Cleartax sales register Mar 2026.xlsx")]
for m,f in TARGETS:
    p = os.path.join(BASE,m,f)
    print(f"\n================ {m} / {f}")
    xl = pd.ExcelFile(p)
    for s in xl.sheet_names:
        try:
            df = xl.parse(s, header=None)
        except Exception as e:
            print(f"  {s!r}: parse failed {e}"); continue
        hits = {}
        for c in df.columns:
            col = df[c].astype(str).str.strip().str.upper()
            n = int(col.str.contains("B2C", na=False).sum())
            if n: hits[c] = n
        print(f"  sheet {s!r} ({df.shape[0]}x{df.shape[1]}): B2C cells by col-index -> {hits if hits else 'NONE'}")
