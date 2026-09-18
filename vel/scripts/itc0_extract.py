"""ITC step 0a: extract the Working sheet to itc_working.pkl + field diagnostics."""
import openpyxl, pandas as pd
from collections import Counter
def S(v): return "" if v is None else str(v).strip()
wb = openpyxl.load_workbook("itc_allstate.xlsx", read_only=True, data_only=True)
ws = wb["Working"]
hdr=None; rows=[]; empty=0
for r in ws.iter_rows(values_only=True):
    if hdr is None:
        vals=[S(v) for v in r]
        if sum(1 for v in vals if v)>8: hdr=vals
        continue
    if r[hdr.index("Document Number")] is None or S(r[hdr.index("Document Number")])=="":
        empty+=1
        if empty>500: break
        continue
    empty=0
    rows.append(r[:len(hdr)])
wb.close()
df = pd.DataFrame(rows, columns=hdr)
df = df[df["Company"].map(S)=="VEL"]
df.to_pickle("itc_working.pkl")
print("rows:", len(df))
for c in ("Taxable Value","Reference","GSTR 2B/6A PERIOD","GST CREDIT YES/NO","Type","Type for GSTR9","GSTR3B Month","Tax Rate"):
    s = df[c]
    nn = s.map(lambda v: v is not None and S(v)!="").sum()
    print("%-20s non-empty %6d | sample:" % (c, nn), [S(v)[:18] for v in s.dropna().head(3)])
print("GSTR3B Month values:", sorted({S(v) for v in df["GSTR3B Month"].dropna()})[:15])
print("credit yes/no:", dict(Counter(S(v) for v in df["GST CREDIT YES/NO"])))
