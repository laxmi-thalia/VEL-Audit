import warnings; warnings.filterwarnings("ignore")
import pandas as pd
def S(v): return "" if pd.isna(v) else str(v).strip()
p = "//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit Data of FY 2024-25/Advance from Customer.xlsx"
xl = pd.ExcelFile(p)
fin = xl.parse("Final", header=None)
print("=== 'Final' full (%s) ===" % str(fin.shape))
for i in range(len(fin)):
    row = [(j, str(v)[:20]) for j, v in enumerate(fin.iloc[i].tolist()) if pd.notna(v)]
    if row: print("r%02d: %s" % (i, " | ".join("c%d:%s" % x for x in row[:14])))
pv = xl.parse("Advance pivot", header=None)
print("\n=== 'Advance pivot' full (%s) ===" % str(pv.shape))
for i in range(len(pv)):
    row = [(j, str(v)[:22]) for j, v in enumerate(pv.iloc[i].tolist()) if pd.notna(v)]
    if row: print("r%02d: %s" % (i, " | ".join("c%d:%s" % x for x in row[:12])))
