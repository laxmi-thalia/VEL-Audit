import os, warnings, collections, pickle; warnings.filterwarnings("ignore")
import pandas as pd
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),"sr_compare2.py")).read().split("a,b=int")[0])
S=lambda v: "" if pd.isna(v) else str(v).strip()
by_irn={}; by_odn={}; sto=[]; uncert=collections.defaultdict(lambda:[0,0.0])
multi_gl=0; docs=0
for lbl,d,f in SAPF:
    got=pick(os.path.join(BASE,d,f))
    if not got: print(f"{lbl}: NO DATA SHEET"); continue
    sh,hr,sap=got
    C={k:col(sap,k) for k in ["ODN","Irn No","G/L Account Number","G/L Text","Profit Center",
       "WBS element/Cost Center","Business Place","Base Value","Remark","Invoice Doc Type","Billing Type"]}
    missing=[k for k,v in C.items() if v is None]
    if missing: print(f"{lbl}: MISSING COLUMNS {missing}")
    sap["_odn"]=sap[C["ODN"]].map(S); sap["_irn"]=sap[C["Irn No"]].map(lambda v:S(v).lower()) if C["Irn No"] else ""
    sap["_base"]=pd.to_numeric(sap[C["Base Value"]],errors="coerce").fillna(0)
    # STO occurrences: Remark / Billing Type / GL Text containing STO as a word
    for cname in ["Remark","Billing Type","G/L Text"]:
        cc=C[cname]
        if cc is None: continue
        hit=sap[sap[cc].map(lambda v:"STO" in S(v).upper().split() or S(v).upper()=="STO")]
        for _,r in hit.iterrows():
            sto.append((lbl,cname,S(r[cc]),S(r["_odn"]),r["_base"]))
    # uncertified
    u=sap[sap[C["G/L Text"]].map(lambda v:"uncertified" in S(v).lower())]
    uncert[lbl][0]+=len(u); uncert[lbl][1]+=u["_base"].sum()
    # per-document aggregation
    for o,x in sap[sap["_odn"]!=""].groupby("_odn"):
        docs+=1
        gls=sorted({S(v) for v in x[C["G/L Text"]] if S(v)})
        if len(gls)>1: multi_gl+=1
        dom=x.loc[x["_base"].abs().idxmax()]
        rec={"gl_no":S(dom[C["G/L Account Number"]]),"gl":S(dom[C["G/L Text"]]),
             "gl_all":" | ".join(gls),
             "pc":S(dom[C["Profit Center"]]),"cc":S(dom[C["WBS element/Cost Center"]]),
             "bp":S(dom[C["Business Place"]]),"month":lbl}
        by_odn[o]=rec
        for irn in {v for v in x["_irn"] if v}:
            by_irn[irn]=rec
    print(f"{lbl}: indexed")
print(f"\ndocuments indexed: {docs} | with >1 distinct GL text: {multi_gl}")
print(f"IRN keys: {len(by_irn)} | ODN keys: {len(by_odn)}")
print(f"STO occurrence rows: {len(sto)}")
print("uncertified by month:")
tot=[0,0.0]
for k,(n_,v_) in uncert.items():
    tot[0]+=n_; tot[1]+=v_
    print(f"   {k}: rows={n_:5d}  base={v_:>18,.2f}")
print(f"   TOTAL rows={tot[0]}  base={tot[1]:,.2f}")
pickle.dump({"by_irn":by_irn,"by_odn":by_odn,"sto":sto,
             "uncert":{k:tuple(v) for k,v in uncert.items()}},
            open(os.path.join(SP,"sap_lookup.pkl"),"wb"))
print("saved sap_lookup.pkl")
