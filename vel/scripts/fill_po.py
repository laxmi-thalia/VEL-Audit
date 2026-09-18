"""Fill 'Sample POs' for selected documents from the SAP monthly sales registers:
Contract No (the customer contract = PO), Sales Doc Number, RA Bill Client No, keyed by ODN."""
import os, collections, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
SP = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(SP, "sr_compare2.py"), encoding="utf-8").read().split("a,b=int")[0]
src = src.replace("//server/", "//192.168.1.69/")
exec(src)   # gives BASE, SAPF, pick(), col()
def S(v): return "" if pd.isna(v) or v is None else str(v).strip()
def num(v):
    v = S(v)
    try: return str(int(float(v)))
    except Exception: return v
P = r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
wb = openpyxl.load_workbook(P); ws = wb["SR_2025-26"]
hdr = {S(ws.cell(4, c).value): c for c in range(1, ws.max_column + 1)}
cSI, cSP, cF = hdr["Sample Invoices"], hdr["Sample POs"], hdr["Document Number"]
need = {S(ws.cell(i, cSI).value) for i in range(5, 27007) if ws.cell(i, cSI).value}
print("selected documents to resolve:", len(need))
ref = collections.defaultdict(lambda: {"contract": set(), "so": set(), "ra": set()})
for lbl, d, f in SAPF:
    got = pick(os.path.join(BASE, d, f))
    if not got: continue
    sh, hr, sap = got
    odn = col(sap, "ODN"); cn = col(sap, "Contract No"); so = col(sap, "Sales Doc Number"); ra = col(sap, "RA Bill Client No")
    sub = sap[sap[odn].map(S).isin(need)]
    for _, r in sub.iterrows():
        k = S(r[odn])
        if cn is not None and S(r[cn]): ref[k]["contract"].add(num(r[cn]))
        if so is not None and S(r[so]): ref[k]["so"].add(num(r[so]))
        if ra is not None and S(r[ra]): ref[k]["ra"].add(S(r[ra]))
filled = unresolved = 0
for i in range(5, 27007):
    d = S(ws.cell(i, cSI).value)
    if not d: continue
    r = ref.get(d)
    if not r or not (r["contract"] or r["so"] or r["ra"]):
        ws.cell(i, cSP).value = "PO ref not in SAP register"; unresolved += 1; continue
    parts = []
    if r["contract"]: parts.append("Contract " + "/".join(sorted(r["contract"])))
    if r["so"]: parts.append("SO " + "/".join(sorted(r["so"])))
    if r["ra"]: parts.append("/".join(sorted(r["ra"])))
    ws.cell(i, cSP).value = " | ".join(parts); filled += 1
wb.save(P)
docs_resolved = sum(1 for d in need if d in ref and (ref[d]["contract"] or ref[d]["so"] or ref[d]["ra"]))
print("documents with a PO/contract ref:", docs_resolved, "of", len(need), "| rows filled:", filled, "| rows unresolved:", unresolved)
