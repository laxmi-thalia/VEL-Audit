import warnings; warnings.filterwarnings("ignore")
import pandas as pd
def S(v): return "" if pd.isna(v) else str(v).strip()
p="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/Financials/Statewise PnL FY25-26.xlsx"
d=pd.read_excel(p, sheet_name="Data", header=0)
d.columns=[S(c) for c in d.columns]
print("all columns:", list(d.columns))
num=[c for c in d.columns if pd.api.types.is_numeric_dtype(d[c])]
print("numeric columns:", num)
# find the amount column: try each numeric col summed over revenue-ish accounts
acc=[c for c in d.columns if "Account Number" in c][2]
d["acct"]=d[acc].map(S)
rev=d[d["acct"].str.startswith("3")]
for c in num:
    t=rev[c].sum()
    if abs(t)>1e8: print("candidate amount col %r: revenue-accounts sum %.2f" % (c, t))
