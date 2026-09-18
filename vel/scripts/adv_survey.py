import warnings; warnings.filterwarnings("ignore")
import pandas as pd
def S(v): return "" if pd.isna(v) else str(v).strip()
# last year's advance reco (format + opening balances)
p = "//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit Data of FY 2024-25/Advance from Customer.xlsx"
xl = pd.ExcelFile(p)
print("LAST YEAR FILE sheets:", xl.sheet_names)
for s in xl.sheet_names[:4]:
    d = xl.parse(s, header=None, nrows=10)
    print("\n=== %r shape=%s" % (s, str(d.shape)))
    for i in range(min(8, len(d))):
        row = [str(v)[:24] for v in d.iloc[i].tolist() if pd.notna(v)]
        if row: print("   r%d: %s" % (i, " | ".join(row[:10])))
