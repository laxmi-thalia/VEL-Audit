import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
D=pd.read_pickle(os.path.join(SP,"reg_docs.pkl"))
G=pd.read_pickle(os.path.join(SP,"gstr1.pkl"))
TY={"Invoice":"INV","Credit Note":"CRN","Debit Note":"DBN"}
def s(x): return "" if pd.isna(x) else str(x).strip()
D["k_irn"]=D["irn"].map(lambda v: s(v).lower())
D["k_doc"]=D["gstin"].map(s).str.upper()+"|"+D["docno"].map(s).str.upper()
G["dtc"]=G["Doc Type"].map(TY)
G["k_irn"]=G["IRN"].map(lambda v: s(v).lower())
G["k_doc"]=G["Company GSTIN"].map(s).str.upper()+"|"+G["Doc No"].map(s).str.upper()
for c in ["Taxable Value (Net)","IGST (Net)","CGST (Net)","SGST (Net)"]:
    G[c]=pd.to_numeric(G[c],errors="coerce").fillna(0)

gi={r.k_irn:i for i,r in G.iterrows() if r.k_irn}
gd={}
for i,r in G.iterrows(): gd.setdefault(r.k_doc,[]).append(i)
usedG=set(); rowsout=[]
for i,r in D.iterrows():
    j=None; how=None
    if r.k_irn and r.k_irn in gi and gi[r.k_irn] not in usedG:
        j=gi[r.k_irn]; how="IRN"
    else:
        for cand in gd.get(r.k_doc,[]):
            if cand not in usedG: j=cand; how="GSTIN+DocNo"; break
    if j is None:
        rowsout.append((r.state,r.gstin,r.docno,r.dtc,r.tax,None,None,"IN BOOKS ONLY",None)); continue
    usedG.add(j); g=G.loc[j]
    dt=round(r.tax-g["Taxable Value (Net)"],2)
    status = "MATCHED" if abs(dt)<1 else "MATCHED - VALUE DIFF"
    if r.dtc!=g["dtc"]: status="MATCHED - DOC TYPE DIFF"
    rowsout.append((r.state,r.gstin,r.docno,r.dtc,r.tax,g["Taxable Value (Net)"],g["dtc"],status,how))
for i,r in G.iterrows():
    if i not in usedG:
        rowsout.append((r["__state"],r["Company GSTIN"],r["Doc No"],None,None,r["Taxable Value (Net)"],r["dtc"],"IN GSTR-1 ONLY",None))
M=pd.DataFrame(rowsout, columns=["state","gstin","docno","reg_type","reg_taxable","g1_taxable","g1_type","status","matched_by"])
M.to_pickle(os.path.join(SP,"step2_match.pkl"))
print("STEP 2 - INVOICE-LEVEL MATCH (books vs GSTR-1)\n")
sm=M.groupby("status").agg(docs=("status","size"),
      reg_taxable=("reg_taxable","sum"), g1_taxable=("g1_taxable","sum")).reset_index()
for _,x in sm.iterrows():
    print(f"  {x['status']:26s} {x['docs']:5d} docs | books {round(x['reg_taxable'],2):>18} | GSTR-1 {round(x['g1_taxable'],2):>18}")
print(f"\n  matched by: {M[M.status.str.startswith('MATCHED')]['matched_by'].value_counts().to_dict()}")
print(f"  TOTAL books taxable  : {round(M['reg_taxable'].sum(),2)}")
print(f"  TOTAL GSTR-1 taxable : {round(M['g1_taxable'].sum(),2)}")
print(f"  DIFFERENCE           : {round(M['reg_taxable'].sum()-M['g1_taxable'].sum(),2)}")
