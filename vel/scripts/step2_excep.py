import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
M=pd.read_pickle(os.path.join(SP,"step2_match.pkl"))
G=pd.read_pickle(os.path.join(SP,"gstr1.pkl"))
def s(x): return "" if pd.isna(x) else str(x).strip()
G["k_doc"]=G["Company GSTIN"].map(s).str.upper()+"|"+G["Doc No"].map(s).str.upper()
blank=set(G[G["IRN"].isna() | (G["IRN"].astype(str).str.strip().isin(["","nan"]))]["k_doc"])

for st in ["IN BOOKS ONLY","IN GSTR-1 ONLY"]:
    X=M[M.status==st]
    print(f"\n===== {st}: {len(X)} documents =====")
    print("  by state :", X.groupby("state").size().to_dict())
    tcol="reg_type" if st=="IN BOOKS ONLY" else "g1_type"
    print("  by type  :", X[tcol].value_counts(dropna=False).to_dict())
    val = X["reg_taxable"].sum() if st=="IN BOOKS ONLY" else X["g1_taxable"].sum()
    print(f"  taxable  : {round(val,2)}")
    if st=="IN GSTR-1 ONLY":
        k=(X["gstin"].map(s).str.upper()+"|"+X["docno"].map(s).str.upper())
        nb=int(k.isin(blank).sum())
        print(f"  of these, IRN was BLANK on the GSTR-1 side: {nb} of {len(X)}")
