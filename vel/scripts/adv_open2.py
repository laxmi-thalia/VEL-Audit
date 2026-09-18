import warnings; warnings.filterwarnings("ignore")
import pandas as pd
def S(v): return "" if pd.isna(v) else str(v).strip()
p = "//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit Data of FY 2024-25/Advance from Customer.xlsx"
fin = pd.ExcelFile(p).parse("Final", header=None)
print("state header row r1:", [(j, S(v)) for j, v in enumerate(fin.iloc[1]) if S(v)])
print("labels row r2   :", [(j, S(v)) for j, v in enumerate(fin.iloc[2]) if S(v)][:20])
for lbl, i in (("Closing 31.03.25", 26), ("Net GSTR-1", 28)):
    print(lbl, ":", [(j, round(float(v), 2)) for j, v in enumerate(fin.iloc[i]) if pd.notna(v) and not isinstance(v, str)])
# GST Advance ledger FY 25-26
q = "//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/GLs/Outward Tax/GST Advance.xlsx"
d = pd.read_excel(q, sheet_name="Data", header=5)
d.columns = [S(c) for c in d.columns]
d = d[d["Document Number"].notna()].copy()
d["amt"] = -pd.to_numeric(d["Amount in Local Currency"], errors="coerce").fillna(0)
d["gl"] = d["G/L Account"].map(lambda v: S(int(v)) if pd.notna(v) else "")
d["ym"] = d["Year/Month"].map(S)
print("\nGST Advance ledger: rows", len(d), "| G/L accounts:", d["gl"].value_counts().to_dict())
print("Year/Month range:", d["ym"].min(), "->", d["ym"].max())
fy = d[d["ym"].isin({"2025/%02d" % m for m in range(1, 13)})]
print("FY 25-26 rows:", len(fy), "| net by business place (tax amount, + = liability):")
print(fy.groupby(fy["Business place"].map(S))["amt"].sum().round(2).to_string())
