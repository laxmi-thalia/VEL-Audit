import openpyxl, warnings, datetime as dt
warnings.filterwarnings("ignore")
from openpyxl.styles import Font, PatternFill
P=r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
wb=openpyxl.load_workbook(P); ws=wb["SR_2025-26"]
R0,R1=5,27006
rows=[]
for i,r in enumerate(ws.iter_rows(min_row=R0,max_row=R1,min_col=1,max_col=50,values_only=True),start=R0):
    rows.append((i,r))
def g(r,idx): return r[idx]
F=[]  # (check, register row, detail)
FY0,FY1=dt.datetime(2025,4,1),dt.datetime(2026,3,31)
seen={}
for i,r in rows:
    state,gstin,ddate,dnum = r[1],r[2],r[4],r[5]
    dtc,sup,bg = r[7],r[8],r[9]
    tax,igst,cgst,sgst = r[16],r[17],r[18],r[19]
    rate,irn,hsn,qty,uom = r[14],r[23],r[32],r[34],r[35]
    if sup in (None,""): F.append(("Blank Supply Type",i,f"doc type {dtc}"))
    if ddate in (None,""): F.append(("Blank Document Date",i,f"doc type {dtc}, no date to assign a GSTR-1 month"))
    elif isinstance(ddate,dt.datetime) and not (FY0<=ddate<=FY1):
        F.append(("Document Date outside FY 25-26",i,str(ddate.date())))
    if irn not in (None,"") and len(str(irn).strip())!=64:
        F.append(("IRN length not 64",i,f"len={len(str(irn).strip())}"))
    oc = str(gstin)[:2] if gstin else None
    cc = str(bg)[:2] if bg not in (None,"") else None
    if oc and cc:
        intra = oc==cc
        I,C,S = (igst or 0),(cgst or 0),(sgst or 0)
        if intra and abs(I)>0.01: F.append(("Intra-state row carrying IGST",i,f"{oc}->{cc}"))
        if intra and abs(C-S)>0.01: F.append(("Intra-state CGST <> SGST",i,f"diff {round(C-S,2)}"))
        if not intra and abs(C)+abs(S)>0.01: F.append(("Inter-state row carrying CGST/SGST",i,f"{oc}->{cc}"))
    if hsn in (None,""): F.append(("Blank HSN/SAC",i,""))
    if qty in (None,""): F.append(("Blank Quantity",i,""))
    if uom in (None,""): F.append(("Blank Unit of Measurement",i,""))
    if tax not in (None,"") and rate not in (None,"") and abs(tax)>0.01:
        eff=round(((igst or 0)+(cgst or 0)+(sgst or 0))/tax*100,2)
        if abs(eff-float(rate))>0.05: F.append(("Effective rate <> stated GST Rate",i,f"computed {eff} vs stated {rate}"))
    k=(str(gstin),str(dnum),str(dtc))
    if dnum not in (None,""):
        if k in seen: F.append(("Duplicate GSTIN+DocNo+DocType",i,f"first seen row {seen[k]}"))
        else: seen[k]=i
import collections
cnt=collections.Counter(f[0] for f in F)
print("FINDINGS SUMMARY")
for k,v in cnt.most_common(): print(f"  {v:6d}  {k}")
print(f"  {len(F):6d}  TOTAL exception rows flagged")
sh=wb.create_sheet("Step 1 Findings")
sh.append(["Check","Register row","Detail"])
for c in sh[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="C00000")
sh.append(["SUMMARY — counts by check","",""])
for k,v in cnt.most_common(): sh.append([k,v,"see detail rows below"])
sh.append(["","",""]); sh.append(["DETAIL","",""])
for chk,i,d in F: sh.append([chk,i,d])
for col,w in zip("ABC",[38,14,60]): sh.column_dimensions[col].width=w
sh.freeze_panes="A2"
wb.save(P); print("appended 'Step 1 Findings' to", P)
