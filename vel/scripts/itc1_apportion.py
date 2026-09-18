"""Fix 2B_ stamps: apportion each matched 2B doc's taxes across its register lines
(weights = books total tax per line; doc sums tie; single-line docs get full amounts)."""
import pandas as pd
import win32com.client as win32, pythoncom
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
reg = pd.read_pickle("itc_working.pkl").reset_index(drop=True)
mt = pd.read_pickle("itc_reg_match.pkl").reset_index(drop=True)
b2 = pd.read_pickle("b2itc.pkl").reset_index(drop=True)
N = len(reg)
for c in ("IGST","CGST","SGST"):
    reg[c+"_n"] = pd.to_numeric(reg[c], errors="coerce").fillna(0)
reg["tot"] = reg["IGST_n"]+reg["CGST_n"]+reg["SGST_n"]
groups = {}
for ri, bi in enumerate(mt["b2i"]):
    if bi is None or (isinstance(bi,float) and pd.isna(bi)): continue
    groups.setdefault(int(bi), []).append(ri)
arr = {k: [0.0]*N for k in ("I","C","S")}
docs_matched = 0; tax_matched = 0.0
for bi, lines in groups.items():
    docs_matched += 1
    dI = float(b2.at[bi,"Integrated Tax(₹)"] or 0); dC = float(b2.at[bi,"Central Tax(₹)"] or 0); dS = float(b2.at[bi,"State/UT Tax(₹)"] or 0)
    tax_matched += dI+dC+dS
    wts = [max(float(reg.at[ri,"tot"]), 0.0) for ri in lines]
    tw = sum(wts)
    if tw <= 0: wts = [1.0]*len(lines); tw = float(len(lines))
    accI=accC=accS=0.0
    for j, ri in enumerate(lines):
        if j == len(lines)-1:
            i_, c_, s_ = dI-accI, dC-accC, dS-accS
        else:
            f_ = wts[j]/tw
            i_, c_, s_ = round(dI*f_,2), round(dC*f_,2), round(dS*f_,2)
            accI+=i_; accC+=c_; accS+=s_
        arr["I"][ri], arr["C"][ri], arr["S"][ri] = i_, c_, s_
print("matched 2B docs:", docs_matched, "| matched 2B tax total: %.2f" % tax_matched)
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P); xl.Calculation=-4135
    ws=wb.Worksheets("ITC Register 2025-26")
    H={}
    for c in range(1,90):
        h=str(ws.Cells(5,c).Value or "").strip()
        if h and h not in H: H[h]=c
    R0=6; RN=R0+N-1
    for nm,k in (("2B_IGST","I"),("2B_CGST","C"),("2B_SGST","S")):
        ci=H[nm]
        ws.Range(ws.Cells(R0,ci), ws.Cells(RN,ci)).Value = [[v] for v in arr[k]]
    xl.Calculation=-4105; xl.CalculateFullRebuild()
    b=ws.Cells(4,H["B_Total GST"]).Value; t2=ws.Cells(4,H["2B_Total GST"]).Value; d=ws.Cells(4,H["D_Total GST"]).Value
    print("B_Total %.2f | 2B_Total %.2f | D_Total %.2f" % (b,t2,d))
    ok = abs(t2-tax_matched)<=2 and abs(d-(b-t2))<=2
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
    print("formula ERROR cells:",tot)
    wb.Close(SaveChanges=True)
    print("VERDICT:","PASS" if ok and tot==0 else "FAIL")
finally:
    xl.Quit()
