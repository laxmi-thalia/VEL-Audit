import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
T=pd.read_pickle(os.path.join(SP,"b3.pkl"))
G=pd.read_pickle(os.path.join(SP,"gstr1.pkl")); S=pd.read_pickle(os.path.join(SP,"g1_summary.pkl"))
def s(x): return "" if pd.isna(x) else str(x).strip().upper()
for df in (G,S):
    for c in ["Taxable Value (Net)","CGST (Net)"]: df[c]=pd.to_numeric(df[c],errors="coerce").fillna(0)
G["month"]=pd.to_datetime(G["Tax Period"],errors="coerce").dt.strftime("%b %Y")
S["month"]=pd.to_datetime(S["Tax Period"],errors="coerce").dt.strftime("%b %Y")
MONTHS=sorted({c.split("|")[0] for c in T.columns if "|taxable" in c}, key=lambda m: pd.to_datetime(m,format="%b %Y"))
for st,g in [("Bihar","10AAECR0503Q1ZN"),("Telangana","36AAECR0503Q1Z9"),("Punjab","03AAECR0503Q1ZI")]:
    r=T[T["gstin"].map(s)==g].iloc[0]
    print(f"\n===== {st} — month-wise taxable (GSTR-1 vs 3B) and CGST")
    for m in MONTHS:
        g1t=G[(G["Company GSTIN"].map(s)==g)&(G["month"]==m)]["Taxable Value (Net)"].sum() \
           +S[(S["Company GSTIN"].map(s)==g)&(S["month"]==m)]["Taxable Value (Net)"].sum()
        g1c=G[(G["Company GSTIN"].map(s)==g)&(G["month"]==m)]["CGST (Net)"].sum() \
           +S[(S["Company GSTIN"].map(s)==g)&(S["month"]==m)]["CGST (Net)"].sum()
        b3t=r[f"{m}|taxable"]; b3c=r[f"{m}|cgst"]
        dt=round(g1t-b3t,2); dc=round(g1c-b3c,2)
        if abs(dt)>=1 or abs(dc)>=1:
            print(f"   {m}: taxable G1={g1t:>16,.2f} 3B={b3t:>16,.2f} diff={dt:>15,.2f} | CGST diff={dc:>13,.2f}")
