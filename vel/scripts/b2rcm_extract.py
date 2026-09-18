"""Extract Reverse Charge = Yes line items from all GSTR-2B files -> b2rcm.pkl + stats."""
import os, pandas as pd, warnings; warnings.filterwarnings("ignore")
def S(v): return "" if v is None or (isinstance(v,float) and pd.isna(v)) else str(v).strip()
B = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-2B/"
KEEP = ["Company GSTIN","Tax Period","Doc Type","Doc No","Doc Date","Supplier GSTIN","Supplier Name",
        "Supplier State","Place of Supply","Reverse Charge","Taxable Value (Net)","IGST (Net)","CGST (Net)",
        "SGST (Net)","ITC Eligible","Is Amendment","Source","IRN"]
frames=[]
for fn in sorted(os.listdir(B)):
    if not fn.lower().endswith(".xlsx"): continue
    xf = pd.ExcelFile(B+fn)
    sheet = "Purchase-Net" if "Purchase-Net" in xf.sheet_names else "Purchase"
    df = xf.parse(sheet)
    df.columns = [S(c) for c in df.columns]
    if "Taxable Value (Net)" not in df.columns and "Taxable Value" in df.columns:
        for a,b in (("Taxable Value (Net)","Taxable Value"),("IGST (Net)","IGST"),("CGST (Net)","CGST"),("SGST (Net)","SGST")):
            df[a] = df.get(b, 0)
    rc = df[df["Reverse Charge"].map(S).str.lower()=="yes"].copy() if "Reverse Charge" in df.columns else df.iloc[0:0]
    for k in KEEP:
        if k not in rc.columns: rc[k] = ""
    frames.append(rc[KEEP])
    st = fn.split("LIMITED-")[1].split("-Apr")[0]
    print("  %-22s total rows %5d | RCM=Yes %4d" % (st, len(df), len(rc)))
all_ = pd.concat(frames, ignore_index=True)
for c in ("Taxable Value (Net)","IGST (Net)","CGST (Net)","SGST (Net)"):
    all_[c] = pd.to_numeric(all_[c], errors="coerce").fillna(0)
all_.to_pickle("b2rcm.pkl")
print("\n2B RCM=Yes line items:", len(all_))
print("taxable %,.2f | IGST %,.2f | CGST %,.2f | SGST %,.2f".replace("%,",'%')%tuple())  if False else None
print("taxable {:,.2f} | IGST {:,.2f} | CGST {:,.2f} | SGST {:,.2f}".format(
    all_["Taxable Value (Net)"].sum(), all_["IGST (Net)"].sum(), all_["CGST (Net)"].sum(), all_["SGST (Net)"].sum()))
