import openpyxl
from collections import Counter
def S(v): return "" if v is None else str(v).strip()
wb = openpyxl.load_workbook("itc_allstate.xlsx", read_only=True, data_only=True)
ws = wb["Working"]
hdr=None; hr=0; n=0; empty_run=0
comp=Counter(); typ=Counter(); state=Counter()
sums={"tax":0.0,"igst":0.0,"cgst":0.0,"sgst":0.0}
idx={}
for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
    if hdr is None:
        vals=[S(v) for v in r]
        if sum(1 for v in vals if v)>8:
            hdr=vals; hr=i
            idx={v:j for j,v in enumerate(vals) if v}
            print("header row",i,"cols:",[v for v in vals if v])
        continue
    key = r[idx.get("Document Number", 0)] if "Document Number" in idx else r[0]
    if key is None or S(key)=="":
        empty_run+=1
        if empty_run>500: break
        continue
    empty_run=0; n+=1
    if "COMPANY" in idx: comp[S(r[idx["COMPANY"]])[:24]]+=1
    elif "Company" in idx: comp[S(r[idx["Company"]])[:24]]+=1
    if "TYPE" in idx: typ[S(r[idx["TYPE"]])[:24]]+=1
    if "STATE NAME" in idx: state[S(r[idx["STATE NAME"]])[:20]]+=1
    for nm,k in (("TAXABLE VALUE","tax"),("IGST","igst"),("CGST","cgst"),("SGST","sgst")):
        j=idx.get(nm)
        if j is not None and isinstance(r[j],(int,float)): sums[k]+=r[j]
wb.close()
print("real data rows:",n)
print("company values:",dict(comp.most_common(10)))
print("type values:",dict(typ.most_common(12)))
print("states:",len(state),dict(state.most_common(5)))
print("sums:",{k:round(v,2) for k,v in sums.items()})
