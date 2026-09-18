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
d = d[d["Document Number"].notna()].copy()
d["amt"] = pd.to_numeric(d["Amount in Local Currency"], errors="coerce").fillna(0)
d["gl"] = d["G/L Account"].map(num); d["dt"] = d["Document Type"].map(S)
d["ym"] = d["Year/Month"].map(S)
FY = {"2025/%02d" % m for m in range(1, 13)}
f = d[d["ym"].isin(FY)].copy()
print("FY 25-26 rows (Year/Month 2025/01-12):", len(f), "of", len(d))
LG = {"2610080100": "CGST?", "2610080101": "SGST?", "2610080102": "IGST?"}
SALES = ["RV", "DA", "DG", "DR", "XD"]
print("\nby ledger x doc type (FY only), amounts sign-flipped (credit = liability):")
for gl, x in f.groupby("gl"):
    row = {t: round(-x[x["dt"] == t]["amt"].sum(), 2) for t in SALES}
    other = round(-x[~x["dt"].isin(SALES)]["amt"].sum(), 2)
    print("  %s (%s): %s | other types %s | SALES TOTAL %s"
          % (gl, LG.get(gl), {k: format(v, ',.0f') for k, v in row.items()}, format(other, ",.0f"),
             format(round(sum(row.values()), 2), ",.2f")))
R = pd.read_pickle(os.path.join(SP, "register.pkl"))
adv = R["dtc"].astype(str).str.startswith("MOB")
print("\nregister totals: CGST %s | SGST %s | IGST %s"
      % (format(round(R['cgst'].sum(), 2), ',.2f'), format(round(R['sgst'].sum(), 2), ',.2f'), format(round(R['igst'].sum(), 2), ',.2f')))
print("register ex-advances: CGST %s | IGST %s"
      % (format(round(R[~adv]['cgst'].sum(), 2), ',.2f'), format(round(R[~adv]['igst'].sum(), 2), ',.2f')))
adv_c = R[adv]["cgst"].sum()
print("register advances only: CGST", format(round(adv_c, 2), ",.2f"))
# per business place, CGST ledger sales-total vs register CGST
BP = {"AP01": "Andhra Pradesh", "AR01": "Arunachal Pradesh", "AS01": "Assam", "BR01": "Bihar", "CG01": "Chhattisgarh",
      "GU01": "Gujarat", "JH01": "Jharkhand", "JK01": "Jammu & Kashmir", "KA01": "Karnataka", "MH01": "Maharashtra",
      "MP01": "Madhya Pradesh", "PB01": "Punjab", "RJ01": "Rajasthan", "TG01": "Telangana", "TN01": "TamilNadu",
      "UP01": "Uttar Pradesh", "WB01": "West Bengal"}
reg = R.groupby(R["state"].map(S))["cgst"].sum()
print("\nCGST by state: GL(2610080100, sales types) vs register:")
g = f[(f["gl"] == "2610080100") & (f["dt"].isin(SALES))].groupby(f["Business place"].map(S))["amt"].sum() * -1
for bp, st in sorted(BP.items()):
    a = round(g.get(bp, 0.0), 2); b = round(reg.get(st, 0.0), 2)
    flag = "" if abs(a - b) < 1 else "   <<< diff %s" % format(round(a - b, 2), ",.2f")
    print("  %-6s %-18s GL %16s | register %16s%s" % (bp, st, format(a, ",.2f"), format(b, ",.2f"), flag))
