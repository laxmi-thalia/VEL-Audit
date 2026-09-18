import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
SP=os.path.dirname(os.path.abspath(__file__))
wb=openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx", read_only=True, data_only=True)
ws=wb["SR_2025-26"]; rec=[]
for r in ws.iter_rows(min_row=5,max_row=27006,min_col=1,max_col=50,values_only=True):
    rec.append(dict(month=r[0],state=r[1],gstin=r[2],ddate=r[4],docno=r[5],dtc=r[7],sup=r[8],
                    bgstin=r[9],tax=r[16] or 0,igst=r[17] or 0,cgst=r[18] or 0,sgst=r[19] or 0,
                    irn=r[23],src=r[49]))
wb.close()
R=pd.DataFrame(rec)
R.to_pickle(os.path.join(SP,"register.pkl"))
print(f"register item rows: {len(R)}")
adv=R["dtc"].astype(str).str.startswith("MOB")
print(f"  advance rows (excluded from invoice match): {int(adv.sum())}")
B=R[~adv].copy()
print(f"  invoice/CN/DN item rows: {len(B)}")
D=B.groupby(["gstin","docno","dtc"], dropna=False).agg(
    tax=("tax","sum"),igst=("igst","sum"),cgst=("cgst","sum"),sgst=("sgst","sum"),
    lines=("tax","size"), irn=("irn","first"), state=("state","first"), ddate=("ddate","first")).reset_index()
print(f"  -> distinct DOCUMENTS in register: {len(D)}")
print(f"     by doc type: {D['dtc'].value_counts().to_dict()}")
print(f"     register taxable (ex-advances): {round(B['tax'].sum(),2)}")
G=pd.read_pickle(os.path.join(SP,"gstr1.pkl"))
print(f"\nGSTR-1 document rows: {len(G)}")
print(f"     by doc type: {G['Doc Type'].value_counts().to_dict()}")
print(f"     GSTR-1 taxable (net): {round(pd.to_numeric(G['Taxable Value (Net)'],errors='coerce').sum(),2)}")
D.to_pickle(os.path.join(SP,"reg_docs.pkl"))
wbstate = R[R['state']=='West Bengal']
print(f"\nnote: West Bengal in register = {len(wbstate)} item rows (no GSTR-1 file), taxable {round(wbstate['tax'].sum(),2)}")
