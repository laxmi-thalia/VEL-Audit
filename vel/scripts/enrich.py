import os, pickle, collections, warnings; warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.styles import Font, PatternFill
SP=os.path.dirname(os.path.abspath(__file__))
L=pickle.load(open(os.path.join(SP,"sap_lookup.pkl"),"rb"))
by_irn,by_odn=L["by_irn"],L["by_odn"]
P=r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
wb=openpyxl.load_workbook(P); ws=wb["SR_2025-26"]
S=lambda v: "" if v is None else str(v).strip()
AK,AL,AM,AY,AZ=37,38,39,51,52
ws.cell(4,AY).value="Business Place"; ws.cell(4,AZ).value="SAP Matched By"
for c in (AY,AZ):
    ws.cell(4,c).font=Font(bold=True,color="FFFFFF"); ws.cell(4,c).fill=PatternFill("solid",fgColor="1F4E79")
stats=collections.defaultdict(lambda:[0,0,0,0])  # month -> [rows, irn, odn, none]
unmatched=collections.Counter(); unmatched_docs={}
for i in range(5,27007):
    m=S(ws.cell(i,50).value)
    irn=S(ws.cell(i,24).value).lower(); odn=S(ws.cell(i,6).value)
    rec=None; how=""
    if irn and irn in by_irn: rec,how=by_irn[irn],"IRN"
    elif odn and odn in by_odn: rec,how=by_odn[odn],"ODN"
    st=stats[m]; st[0]+=1
    if rec:
        st[1 if how=="IRN" else 2]+=1
        ws.cell(i,AK).value=rec["gl"]; ws.cell(i,AL).value=rec["pc"]; ws.cell(i,AM).value=rec["cc"]
        ws.cell(i,AY).value=rec["bp"]; ws.cell(i,AZ).value=how
    else:
        st[3]+=1
        dt=S(ws.cell(i,8).value)
        unmatched[(m,dt)]+=1
        unmatched_docs.setdefault((m,dt,odn), S(ws.cell(i,2).value))
# ---- report sheet
if "SAP Enrichment" in wb.sheetnames: del wb["SAP Enrichment"]
sh=wb.create_sheet("SAP Enrichment")
H=Font(bold=True,color="FFFFFF"); F=PatternFill("solid",fgColor="1F4E79")
sh.append(["SAP enrichment of the Master Sales Register — GL Name / Profit Centre / Cost Centre / Business Place"])
sh.append(["Source: monthly SAP 'Sales Register <Month>' files. Join: IRN first, ODN fallback. NO figures changed; NO SAP rows added."])
sh.append([])
sh.append(["Month","Register rows","Matched by IRN","Matched by ODN","Unmatched"])
for c in sh[4]: c.font=H; c.fill=F
tot=[0,0,0,0]
for m in sorted(stats, key=lambda x:("Apr May Jun Jul Aug Sep Oct Nov Dec Jan Feb Mar".split().index(x[:3]))):
    st=stats[m]; sh.append([m]+st)
    for j in range(4): tot[j]+=st[j]
sh.append(["TOTAL"]+tot)
sh.append([])
sh.append(["Unmatched by month and document type (advances have no IRN and often no SAP ODN):"])
sh.append(["Month","Doc type","Rows"])
for c in sh[sh.max_row]: c.font=H; c.fill=F
for (m,dt),n_ in sorted(unmatched.items()):
    sh.append([m,dt,n_])
sh.append([])
sh.append(["STO occurrences found in SAP registers — NOT classified; for CA confirmation:"])
sh.append(["Month","Found in column","Value","ODN","Base value"])
for c in sh[sh.max_row]: c.font=H; c.fill=F
for row in L["sto"]: sh.append(list(row))
sh.append([])
sh.append(["Uncertified / accrued revenue in SAP (G/L Text contains 'Uncertified') — kept OUTSIDE Step 1;"])
sh.append(["flagged for the financial-statements reconciliation (GSTR-9C stage):"])
sh.append(["Month","SAP rows","Base value"])
for c in sh[sh.max_row]: c.font=H; c.fill=F
ut=[0,0.0]
for m,(n_,v_) in L["uncert"].items():
    sh.append([m,n_,round(v_,2)]); ut[0]+=n_; ut[1]+=v_
sh.append(["TOTAL",ut[0],round(ut[1],2)])
sh.append([])
sh.append(["Separately: 267 SAP documents (Rs 8,79,72,580.91 base) exist in SAP but not in ClearTax —"])
sh.append(["mostly 'Uncertified' accruals with no invoice/IRN. Deliberately excluded from the register."])
for col,w in zip("ABCDE",[16,22,44,22,18]): sh.column_dimensions[col].width=w
wb.save(P)
print(f"rows: {tot[0]} | IRN {tot[1]} | ODN {tot[2]} | unmatched {tot[3]}  ({100*(tot[1]+tot[2])/tot[0]:.2f}% matched)")
print("unmatched by doc type:", dict(collections.Counter({dt:0 for (_,dt) in unmatched}) + collections.Counter({dt:n for (m,dt),n in unmatched.items()})))
