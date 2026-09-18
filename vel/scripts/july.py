import warnings, hashlib
warnings.filterwarnings("ignore")
import pandas as pd
P = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\July2025_converted.xlsx"
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
def norm(c): return str(c).strip().lower()
def looks_header(vals):
    nn=[v for v in vals if pd.notna(v)]
    return len(nn)>=5 and sum(1 for v in nn if isinstance(v,str))/len(nn)>0.8
xl = pd.ExcelFile(P)
probe = xl.parse("Working", header=None, nrows=12)
hr = next(i for i in range(len(probe)) if looks_header(probe.iloc[i].tolist()))
df = xl.parse("Working", header=hr)
cols=[norm(c) for c in df.columns]
print(f"July Working: hdr_row={hr} rows={len(df)} cols={df.shape[1]} sig={hashlib.sha1('|'.join(cols).encode()).hexdigest()[:12]}")
def col(n):
    for c in df.columns:
        if norm(c)==n: return c
dt,st,irn = col("document type"), col("supply type"), col("invoice reference no")
print("  DocType:", df[dt].value_counts(dropna=False).to_dict())
print("  Supply :", df[st].value_counts(dropna=False).to_dict())
b = df[irn].isna() | (df[irn].astype(str).str.strip()=="")
print("  blank IRN:", int(b.sum()))
base=[norm(c) for c in pd.read_excel(BASE+"/09 Dec 2025/Cleartax Sales Register Dec 2025.xlsx",
      sheet_name="Working", header=2, nrows=0).columns]
print("  MISSING vs Dec:", [c for c in base if c not in set(cols)])
print("  EXTRA   vs Dec:", [c for c in cols if c not in set(base)])
