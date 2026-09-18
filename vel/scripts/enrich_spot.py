import os, random, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),"sr_compare2.py")).read().split("a,b=int")[0])
S=lambda v: "" if pd.isna(v) or v is None else str(v).strip()
wb=openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx",read_only=True,data_only=True)
ws=wb["SR_2025-26"]
rows=[]
for i,r in enumerate(ws.iter_rows(min_row=5,max_row=27006,min_col=1,max_col=52,values_only=True),start=5):
    if r[51]:  # matched
        rows.append((i,S(r[5]),S(r[23]).lower(),S(r[36]),S(r[37]),S(r[38]),S(r[50]),S(r[49]),S(r[51])))
wb.close()
random.seed(23)
sample=random.sample(rows,15)
bymonth={}
for s_ in sample: bymonth.setdefault(s_[7],[]).append(s_)
checked=0; bad=[]
for lbl,items in bymonth.items():
    d,f=next((d,f) for L,d,f in SAPF if L==lbl)
    sh,hr,sap=pick(os.path.join(BASE,d,f))
    C={k:col(sap,k) for k in ["ODN","Irn No","G/L Text","Profit Center","WBS element/Cost Center","Business Place","Base Value"]}
    sap["_odn"]=sap[C["ODN"]].map(S); sap["_base"]=pd.to_numeric(sap[C["Base Value"]],errors="coerce").fillna(0)
    for (i,odn,irn,gl,pc,cc,bp,mon,how) in items:
        x=sap[sap["_odn"]==odn]
        checked+=1
        if len(x)==0: bad.append((i,odn,"ODN NOT IN SAP FILE")); continue
        dom=x.loc[x["_base"].abs().idxmax()]
        exp_gl=S(dom[C["G/L Text"]]); exp_pc=S(dom[C["Profit Center"]]); exp_cc=S(dom[C["WBS element/Cost Center"]]); exp_bp=S(dom[C["Business Place"]])
        irn_ok = (how!="IRN") or any(S(v).lower()==irn for v in x[C["Irn No"]])
        for name,got,exp in [("GL",gl,exp_gl),("PC",pc,exp_pc),("CC",cc,exp_cc),("BP",bp,exp_bp)]:
            # PC may be float-formatted; normalise numerics
            g2,e2=got.rstrip("0").rstrip("."),exp.rstrip("0").rstrip(".")
            if g2!=e2: bad.append((i,odn,f"{name}: register={got!r} sap={exp!r}"))
        if not irn_ok: bad.append((i,odn,"IRN claimed but not found in SAP doc"))
print(f"spot-checked {checked} matched rows across {len(bymonth)} months, 4 fields + join key each")
print(f"mismatches: {len(bad)}")
for b in bad[:10]: print("  ", b)
print("VERDICT:", "ENRICHMENT VALUES MATCH SAP SOURCE" if not bad else "*** DISCREPANCIES ***")
