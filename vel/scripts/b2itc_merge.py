"""Merge the 'GSTR 2B' sheets from all state July-2026 working files -> b2itc.pkl.
Header row found per file by 'GSTIN of supplier'; invoice numbers unquoted; numerics coerced."""
import os, re, pandas as pd, openpyxl, warnings; warnings.filterwarnings("ignore")
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
B = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit data of FY 2026-27/04 July 2026/Final"
KEEP = ["F.Y","2B Return Period","GSTR 3B Month","Curent previous","GSTIN of supplier","Trade/Legal name",
"Taxpayer Type","Invoice number","Invoice type","Invoice Date","Invoice Value(₹)","Place of supply",
"Supply Attract Reverse Charge","Rate","Taxable Value (₹)","Integrated Tax(₹)","Central Tax(₹)",
"State/UT Tax(₹)","Cess(₹)","GSTR-1/5 Period","GSTR-1/5 Filing Date","ITC Availability","Reason",
"Source","IRN","Remarks"]
frames=[]
for st in sorted(os.listdir(B)):
    d = os.path.join(B, st)
    if not os.path.isdir(d): continue
    cands = [f for f in os.listdir(d) if re.match(r"gstr[\s-]*3b", f.lower()) and f.lower().endswith(".xlsx")]
    if not cands:
        print("  %-20s NO 3B WORKING FILE" % st); continue
    fn = os.path.join(d, cands[0])
    try:
        wb = openpyxl.load_workbook(fn, read_only=True, data_only=True)
    except Exception as e:
        print("  %-20s OPEN FAIL %s" % (st, e)); continue
    tgt = next((n for n in wb.sheetnames if n.strip().upper() in ("GSTR 2B","GSTR-2B","GSTR2B")), None)
    if tgt is None:
        print("  %-20s no GSTR 2B sheet (has: %s)" % (st, wb.sheetnames[:6])); wb.close(); continue
    ws = wb[tgt]
    hdr=None; rows=[]; empty=0
    for r in ws.iter_rows(values_only=True):
        if hdr is None:
            vals=[S(v) for v in r]
            if "GSTIN of supplier" in vals: hdr=vals
            continue
        gi = hdr.index("GSTIN of supplier")
        if len(r)<=gi or r[gi] is None or S(r[gi])=="":
            empty+=1
            if empty>300: break
            continue
        empty=0
        rows.append(r[:len(hdr)])
    wb.close()
    if hdr is None:
        print("  %-20s header NOT FOUND" % st); continue
    df = pd.DataFrame(rows, columns=[h if h else "col%d"%i for i,h in enumerate(hdr)])
    # dedupe duplicated header names (two 'Total Tax' etc.) - keep first occurrence of KEEP names
    df = df.loc[:, ~df.columns.duplicated()]
    for k in KEEP:
        if k not in df.columns: df[k] = None
    sub = df[KEEP].copy()
    sub["State folder"] = st
    frames.append(sub)
    print("  %-20s rows %6d" % (st, len(sub)))
all_ = pd.concat(frames, ignore_index=True)
all_["Invoice number"] = all_["Invoice number"].map(lambda v: S(v).strip("'").strip())
for c in ("Invoice Value(₹)","Rate","Taxable Value (₹)","Integrated Tax(₹)","Central Tax(₹)","State/UT Tax(₹)","Cess(₹)"):
    all_[c] = pd.to_numeric(all_[c], errors="coerce")
all_.to_pickle("b2itc.pkl")
print("\nTOTAL rows:", len(all_))
print("FY spread:", dict(all_["F.Y"].map(S).value_counts()))
mk = all_["GSTR 3B Month"].map(S)
lab = mk.map(lambda s: "Lapsed" if s.lower().startswith("lapsed") else ("blank" if s=="" else ("Claimed-in style" if s.lower().startswith("claimed") else "month")))
print("claim marking classes:", dict(lab.value_counts()))
fy26 = all_[all_["F.Y"].map(S)=="2025-26"]
print("FY 25-26 rows:", len(fy26), "| taxable %.2f | IGST %.2f | CGST %.2f | SGST %.2f" % (
    fy26["Taxable Value (₹)"].sum(), fy26["Integrated Tax(₹)"].sum(), fy26["Central Tax(₹)"].sum(), fy26["State/UT Tax(₹)"].sum()))
