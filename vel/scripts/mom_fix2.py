import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
R=pd.read_pickle(os.path.join(SP,"register.pkl")); G=pd.read_pickle(os.path.join(SP,"gstr1.pkl"))
TY={"Invoice":"INV","Credit Note":"CRN","Debit Note":"DBN"}
def s(x): return "" if pd.isna(x) else str(x).strip().upper()
R["g"]=R["gstin"].map(s); G["g"]=G["Company GSTIN"].map(s)
lab=R.groupby("g")["state"].first().to_dict()
for k,v in G.groupby("g")["__state"].first().items(): lab.setdefault(k,v)
R["adv"]=R["dtc"].astype(str).str.startswith("MOB")
rb=R[~R["adv"]].copy()
# month = document date; FALLBACK to the source-file month when the date is missing
m1=pd.to_datetime(rb["ddate"],errors="coerce").dt.strftime("%Y-%m")
m2=pd.to_datetime(rb["src"],format="%b-%y",errors="coerce").dt.strftime("%Y-%m")
rb["month"]=m1.fillna(m2).fillna("(no date)")
fb=rb.groupby(["g","month","dtc"]).agg(taxable=("tax","sum"),igst=("igst","sum"),
    cgst=("cgst","sum"),sgst=("sgst","sum")).reset_index(); fb["source"]="Books"
G["dtc"]=G["Doc Type"].map(TY); G["month"]=pd.to_datetime(G["Tax Period"],errors="coerce").dt.strftime("%Y-%m")
for c in ["Taxable Value (Net)","IGST (Net)","CGST (Net)","SGST (Net)"]: G[c]=pd.to_numeric(G[c],errors="coerce").fillna(0)
fg=G.groupby(["g","month","dtc"]).agg(taxable=("Taxable Value (Net)","sum"),igst=("IGST (Net)","sum"),
    cgst=("CGST (Net)","sum"),sgst=("SGST (Net)","sum")).reset_index(); fg["source"]="GSTR-1"
piv=pd.concat([fb,fg],ignore_index=True).pivot_table(index=["g","month","dtc"],columns="source",
     values=["taxable","igst","cgst","sgst"],aggfunc="sum",fill_value=0).reset_index()
piv.columns=[a if not b else f"{a}_{b}" for a,b in piv.columns]
for f in ["taxable","igst","cgst","sgst"]:
    for src in ["Books","GSTR-1"]:
        if f"{f}_{src}" not in piv: piv[f"{f}_{src}"]=0.0
    piv[f"{f}_DIFF"]=(piv[f"{f}_Books"]-piv[f"{f}_GSTR-1"]).round(2)
piv.insert(0,"state",piv["g"].map(lab))
WB_NO_REPORT={"West Bengal"}
piv["note"]=piv.apply(lambda r: "GSTR-1 report not downloaded - not a real difference"
                      if r["state"] in WB_NO_REPORT and abs(r["taxable_DIFF"])>=1 else "", axis=1)
piv=piv.sort_values(["state","month","dtc"]); piv.to_pickle(os.path.join(SP,"mom2.pkl"))
d=piv[piv["taxable_DIFF"].abs()>=1]
print(f"rows: {len(piv)} | with taxable diff >= Rs 1: {len(d)}")
print(d[["state","month","dtc","taxable_Books","taxable_GSTR-1","taxable_DIFF","note"]].to_string(index=False))
