import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd
SP = os.path.dirname(os.path.abspath(__file__))
B = "//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/GLs/Outward Tax/"
def S(v): return "" if pd.isna(v) else str(v).strip()
def num(v):
    v = S(v)
    try: return str(int(float(v)))
    except Exception: return v
d = pd.read_excel(B + "GST Output.xlsx", sheet_name="Data", header=5)
d.columns = [S(c) for c in d.columns]
print("rows:", len(d))
nod = d[d["Document Number"].isna()]
print("rows WITHOUT a Document Number (trailer/total?):", len(nod), "| their Amount sum:", pd.to_numeric(nod["Amount in Local Currency"], errors="coerce").sum())
d = d[d["Document Number"].notna()].copy()
d["amt"] = pd.to_numeric(d["Amount in Local Currency"], errors="coerce").fillna(0)
d["gl"] = d["G/L Account"].map(num)
print("\nG/L Accounts:")
for gl, x in d.groupby("gl"):
    print("   %s  rows %5d  amount %18,.2f" % (gl, len(x), x["amt"].sum()) if False else "   %s  rows %5d  amount %18s" % (gl, len(x), format(round(x["amt"].sum(), 2), ",.2f")))
print("\nDocument Type:", d["Document Type"].map(S).value_counts().to_dict())
print("Business place:", dict(sorted(collections.Counter(d["Business place"].map(S)).items())))
d["fy"] = d["Fiscal Year"].map(num)
print("Fiscal Year:", d["fy"].value_counts().to_dict())
print("Year/Month range:", S(d["Year/Month"].min()), "->", S(d["Year/Month"].max()))
pd_ = pd.to_datetime(d["Posting Date"], errors="coerce")
print("Posting Date range:", pd_.min(), "->", pd_.max())
ref_blank = d["Reference"].isna() | (d["Reference"].map(S) == "")
print("\nrows without Reference (invoice no):", int(ref_blank.sum()), "| their doc types:", d[ref_blank]["Document Type"].map(S).value_counts().to_dict())
print("their amount:", format(round(d[ref_blank]["amt"].sum(), 2), ",.2f"))
# join coverage vs the register documents
R = pd.read_pickle(os.path.join(SP, "register.pkl"))
regdocs = {S(x).upper() for x in R["docno"]}
refs = {S(x).upper() for x in d[~ref_blank]["Reference"]}
print("\ndistinct References:", len(refs), "| found in register:", len(refs & regdocs), "| not in register:", len(refs - regdocs))
print("sample refs not in register:", sorted(refs - regdocs)[:6])
print("\ntotal amount (docs only):", format(round(d['amt'].sum(), 2), ",.2f"), " (register total GST was 1,22,34,58,748.34)")
