import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
SP=os.path.dirname(os.path.abspath(__file__))
R=pd.read_pickle(os.path.join(SP,"register.pkl"))
G=pd.read_pickle(os.path.join(SP,"gstr1.pkl"))
M=pd.read_pickle(os.path.join(SP,"step2_match.pkl"))
TY={"Invoice":"INV","Credit Note":"CRN","Debit Note":"DBN"}
def s(x): return "" if pd.isna(x) else str(x).strip()
# ---- month-on-month fact table (long form: one row per GSTIN x month x doctype x source)
R["adv"]=R["dtc"].astype(str).str.startswith("MOB")
rb=R[~R["adv"]].copy()
rb["month"]=pd.to_datetime(rb["ddate"],errors="coerce").dt.strftime("%Y-%m").fillna("(no date)")
fb=rb.groupby(["state","gstin","month","dtc"]).agg(taxable=("tax","sum"),igst=("igst","sum"),
     cgst=("cgst","sum"),sgst=("sgst","sum")).reset_index(); fb["source"]="Books"
G["dtc"]=G["Doc Type"].map(TY)
G["month"]=pd.to_datetime(G["Tax Period"],errors="coerce").dt.strftime("%Y-%m")
for c in ["Taxable Value (Net)","IGST (Net)","CGST (Net)","SGST (Net)"]:
    G[c]=pd.to_numeric(G[c],errors="coerce").fillna(0)
fg=G.groupby(["__state","Company GSTIN","month","dtc"]).agg(taxable=("Taxable Value (Net)","sum"),
     igst=("IGST (Net)","sum"),cgst=("CGST (Net)","sum"),sgst=("SGST (Net)","sum")).reset_index()
fg.columns=["state","gstin","month","dtc","taxable","igst","cgst","sgst"]; fg["source"]="GSTR-1"
FACT=pd.concat([fb,fg],ignore_index=True)
piv=FACT.pivot_table(index=["state","gstin","month","dtc"],columns="source",
     values=["taxable","igst","cgst","sgst"],aggfunc="sum",fill_value=0).reset_index()
piv.columns=[a if not b else f"{a}_{b}" for a,b in piv.columns]
for f in ["taxable","igst","cgst","sgst"]:
    for src in ["Books","GSTR-1"]:
        if f"{f}_{src}" not in piv: piv[f"{f}_{src}"]=0.0
    piv[f"{f}_DIFF"]=(piv[f"{f}_Books"]-piv[f"{f}_GSTR-1"]).round(2)
piv=piv.sort_values(["state","month","dtc"])

wb=Workbook(); H=Font(bold=True,color="FFFFFF"); F=PatternFill("solid",fgColor="1F4E79")
BAD=PatternFill("solid",fgColor="FFC7CE"); OKF=PatternFill("solid",fgColor="C6EFCE")
ws=wb.active; ws.title="Reconciliation Summary"
tb=round(M["reg_taxable"].sum(),2); tg=round(M["g1_taxable"].sum(),2)
rows=[["STEP 2 — Sales Register vs GSTR-1 (FY 2025-26)","",""],["","",""],
 ["Books — taxable value (excluding advances)",tb,""],
 ["GSTR-1 — taxable value (net of amendments)",tg,""],
 ["DIFFERENCE",round(tb-tg,2),""],["","",""],
 ["RECONCILED BY:","",""]]
for st_ in ["MATCHED","IN BOOKS ONLY","IN GSTR-1 ONLY"]:
    X=M[M.status==st_]
    v = X["reg_taxable"].sum() if st_=="IN BOOKS ONLY" else (X["g1_taxable"].sum() if st_=="IN GSTR-1 ONLY" else 0)
    rows.append([st_, len(X), round(v,2)])
rows += [["","",""],
 ["Matched documents agree on taxable value to the rupee","",""],
 ["Matched by IRN", int((M["matched_by"]=="IRN").sum()),""],
 ["Matched by GSTIN + Doc No", int((M["matched_by"]=="GSTIN+DocNo").sum()),""],
 ["","",""],["NOTE: advances are excluded here — they are summary-level in GSTR-1","",""],
 ["NOTE: West Bengal has no GSTR-1 file, so its books rows cannot match","",""]]
for r in rows: ws.append(r)
for c in ws[1]: c.font=H; c.fill=F
ws.column_dimensions["A"].width=58; ws.column_dimensions["B"].width=20; ws.column_dimensions["C"].width=20

e=wb.create_sheet("Exceptions")
e.append(["Status","State","My GSTIN","Doc No","Books type","Books taxable","GSTR-1 type","GSTR-1 taxable","Likely reason"])
for _,r in M[M.status!="MATCHED"].sort_values(["status","state"]).iterrows():
    reason=("West Bengal — GSTR-1 report not downloaded" if r["state"]=="West Bengal"
            else "Credit note declared in GSTR-1 with no IRN and no entry in books" if r["status"]=="IN GSTR-1 ONLY"
            else "In books, not found in GSTR-1")
    e.append([r["status"],r["state"],r["gstin"],r["docno"],r["reg_type"],r["reg_taxable"],r["g1_type"],r["g1_taxable"],reason])
for c in e[1]: c.font=H; c.fill=F
for col,w in zip("ABCDEFGHI",[20,18,20,22,11,18,11,18,52]): e.column_dimensions[col].width=w
e.freeze_panes="A2"

m=wb.create_sheet("Month-on-Month")
m.append(["State","My GSTIN","GSTR-1 Month","Doc Type","Taxable Books","Taxable GSTR-1","Taxable DIFF",
          "IGST DIFF","CGST DIFF","SGST DIFF"])
for _,r in piv.iterrows():
    m.append([r["state"],r["gstin"],r["month"],r["dtc"],round(r["taxable_Books"],2),round(r["taxable_GSTR-1"],2),
              r["taxable_DIFF"],r["igst_DIFF"],r["cgst_DIFF"],r["sgst_DIFF"]])
for c in m[1]: c.font=H; c.fill=F
for row in m.iter_rows(min_row=2):
    row[6].fill = OKF if abs(row[6].value or 0)<1 else BAD
for col,w in zip("ABCDEFGHIJ",[20,20,14,10,18,18,16,14,14,14]): m.column_dimensions[col].width=w
m.freeze_panes="A2"

mt=wb.create_sheet("Matched")
mt.append(["State","My GSTIN","Doc No","Type","Books taxable","GSTR-1 taxable","Matched by"])
for _,r in M[M.status=="MATCHED"].iterrows():
    mt.append([r["state"],r["gstin"],r["docno"],r["reg_type"],r["reg_taxable"],r["g1_taxable"],r["matched_by"]])
for c in mt[1]: c.font=H; c.fill=F
for col,w in zip("ABCDEFG",[20,20,22,8,18,18,16]): mt.column_dimensions[col].width=w
mt.freeze_panes="A2"
OUT=r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
wb.save(OUT)
nb=int((piv["taxable_DIFF"].abs()>=1).sum())
print("WROTE",OUT)
print(f"month-on-month rows: {len(piv)} | rows with a taxable difference >= Rs 1: {nb}")
