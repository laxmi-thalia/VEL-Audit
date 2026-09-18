import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),"sr_compare2.py")).read().split("a,b=int")[0])
S=lambda v: "" if pd.isna(v) else str(v).strip()
onlysap=[]; dtcount=collections.Counter(); nog_by_dt=collections.Counter(); tot=0
for lbl,d,f in SAPF:
    got=pick(os.path.join(BASE,d,f))
    if not got: continue
    sh,hr,sap=got
    cd,cf,cs,ch=CTF[lbl]; ct=pd.read_excel(cf if cd is None else os.path.join(BASE,cd,cf),sheet_name=cs,header=ch)
    odn=col(sap,"ODN"); gstn=col(sap,"Bill to Customer GSTN"); base=col(sap,"Base Value")
    dtyp=col(sap,"Invoice Doc Type"); irn=col(sap,"Irn No"); glt=col(sap,"G/L Text")
    cdset={S(x) for x in ct[col(ct,"Document Number")] if S(x)}
    sap["_odn"]=sap[odn].map(S)
    for dt,x in sap.groupby(sap[dtyp].map(S)):
        dtcount[dt]+=len(x)
        nog_by_dt[dt]+=int((x[gstn].map(S)=="").sum())
    sub=sap[(sap["_odn"]!="") & (~sap["_odn"].isin(cdset))]
    tot+=len(sub)
    for o,x in sub.groupby("_odn"):
        onlysap.append((lbl,o,S(x[dtyp].iloc[0]),S(x[gstn].iloc[0]),
                        pd.to_numeric(x[base],errors="coerce").sum(),
                        S(x[irn].iloc[0]) if irn else "", S(x[glt].iloc[0]) if glt else ""))
O=pd.DataFrame(onlysap,columns=["month","odn","doctype","cust_gstin","base","irn","gl_text"])
print(f"SAP documents NOT in ClearTax, full year: {len(O)}  (rows {tot})")
print(f"  by doc type : {O['doctype'].value_counts().to_dict()}")
print(f"  with customer GSTIN : {(O['cust_gstin']!='').sum()}   without : {(O['cust_gstin']=='').sum()}")
print(f"  with IRN            : {(O['irn']!='').sum()}   without : {(O['irn']=='').sum()}")
print(f"  total base value    : {O['base'].sum():,.2f}")
print(f"\n  top G/L Text on these documents:")
for k,v in O['gl_text'].value_counts().head(10).items(): print(f"     {v:4d}  {k}")
print(f"\nALL SAP rows by doc type, and how many lack a customer GSTIN:")
for dt,c in dtcount.most_common():
    print(f"   {dt:12s} rows={c:6d}  no-GSTIN={nog_by_dt[dt]:6d}  ({100*nog_by_dt[dt]/c:.0f}%)")
