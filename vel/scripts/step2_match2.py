import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
D=pd.read_pickle(os.path.join(SP,"reg_docs.pkl")); G=pd.read_pickle(os.path.join(SP,"gstr1.pkl"))
R=pd.read_pickle(os.path.join(SP,"register.pkl"))
TY={"Invoice":"INV","Credit Note":"CRN","Debit Note":"DBN"}
def s(x): return "" if pd.isna(x) else str(x).strip()
# counterparty GSTIN per register document (first non-blank buyer gstin on its lines)
R["adv"]=R["dtc"].astype(str).str.startswith("MOB")
cp=(R[~R["adv"]].assign(k=lambda d: d["gstin"].map(s).str.upper()+"|"+d["docno"].map(s).str.upper()+"|"+d["dtc"].map(s))
      .groupby("k")["bgstin"].apply(lambda x: next((s(v) for v in x if s(v)), "")).to_dict())
D["k3"]=D["gstin"].map(s).str.upper()+"|"+D["docno"].map(s).str.upper()+"|"+D["dtc"].map(s)
D["cust"]=D["k3"].map(cp).fillna("")
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
# tier 3 index: own GSTIN + customer GSTIN + doc type + taxable rounded to paisa
g3={}
for i,r in G.iterrows():
    key=(s(r["Company GSTIN"]).upper(), s(r["Customer GSTIN"]).upper(), r["dtc"], round(float(r["Taxable Value (Net)"]),2))
    g3.setdefault(key,[]).append(i)
usedG=set(); out=[]
for _,r in D.iterrows():
    j=None; how=None
    if r.k_irn and r.k_irn in gi and gi[r.k_irn] not in usedG: j,how=gi[r.k_irn],"IRN"
    if j is None:
        for c in gd.get(r.k_doc,[]):
            if c not in usedG: j,how=c,"GSTIN+DocNo"; break
    if j is None:
        key=(s(r["gstin"]).upper(), s(r["cust"]).upper(), r["dtc"], round(float(r["tax"]),2))
        for c in g3.get(key,[]):
            if c not in usedG: j,how=c,"Customer+Amount"; break
    if j is None:
        out.append((r["state"],r["gstin"],r["docno"],r["dtc"],r["tax"],None,None,None,"IN BOOKS ONLY",None)); continue
    usedG.add(j); g=G.loc[j]
    dt=round(r["tax"]-g["Taxable Value (Net)"],2)
    st="MATCHED" if abs(dt)<1 else "MATCHED - VALUE DIFF"
    out.append((r["state"],r["gstin"],r["docno"],r["dtc"],r["tax"],g["Doc No"],g["Taxable Value (Net)"],g["dtc"],st,how))
for i,r in G.iterrows():
    if i not in usedG:
        out.append((r["__state"],r["Company GSTIN"],None,None,None,r["Doc No"],r["Taxable Value (Net)"],r["dtc"],"IN GSTR-1 ONLY",None))
M=pd.DataFrame(out,columns=["state","gstin","docno","reg_type","reg_taxable","g1_docno","g1_taxable","g1_type","status","matched_by"])
M.to_pickle(os.path.join(SP,"step2_match2.pkl"))
print("counts:", M["status"].value_counts().to_dict())
print("matched by:", M[M.status.str.startswith("MATCHED")]["matched_by"].value_counts().to_dict())
print("tier-3 pairs found:")
print(M[M.matched_by=="Customer+Amount"][["state","docno","g1_docno","reg_type","reg_taxable","g1_taxable"]].to_string(index=False))
