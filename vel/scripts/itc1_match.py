"""ITC step 1 engine: invoice-level match register <-> merged 2B.
Cascade: (vendor GSTIN, normalized inv no) -> (vendor GSTIN, inv date, total tax ~1) ->
(vendor GSTIN, total tax ~1 unique). Each 2B row consumable once per register doc group.
Outputs: itc_match.pkl (per-register-row + per-2B-row verdicts) + stats."""
import pandas as pd, re
from collections import defaultdict
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
def ninv(v):
    return re.sub(r"[^A-Z0-9]", "", S(v).upper().lstrip("'"))
reg = pd.read_pickle("itc_working.pkl")
reg = reg.reset_index(drop=True)
reg["vg"] = reg["Vendor GSTIN"].map(lambda v: S(v).upper())
reg["inv"] = reg["Reference"].map(ninv)
for c in ("IGST","CGST","SGST","Taxable Value"):
    reg[c+"_n"] = pd.to_numeric(reg[c], errors="coerce").fillna(0)
reg["tot"] = (reg["IGST_n"]+reg["CGST_n"]+reg["SGST_n"]).round(2)
reg["ddate"] = pd.to_datetime(reg["Document Date"], errors="coerce")
b2 = pd.read_pickle("b2itc.pkl").reset_index(drop=True)
b2["vg"] = b2["GSTIN of supplier"].map(lambda v: S(v).upper())
b2["inv"] = b2["Invoice number"].map(ninv)
for c in ("Integrated Tax(₹)","Central Tax(₹)","State/UT Tax(₹)","Taxable Value (₹)"):
    b2[c] = b2[c].fillna(0)
b2["tot"] = (b2["Integrated Tax(₹)"]+b2["Central Tax(₹)"]+b2["State/UT Tax(₹)"]).round(2)
b2["idate"] = pd.to_datetime(b2["Invoice Date"], errors="coerce")
# ITC rows only (register RCM/ISD rows have their own flows)
regI = reg[reg["TYPE"].map(S)=="ITC"] if "TYPE" in reg.columns else reg
regI = reg[reg["TYPE"].map(S).isin(["ITC"])] if False else reg  # keep all; tag later
# --- index 2B by keys
by_inv = defaultdict(list); by_dt = defaultdict(list); by_amt = defaultdict(list)
for i, r in b2.iterrows():
    if r["vg"]:
        if r["inv"]: by_inv[(r["vg"], r["inv"])].append(i)
        if pd.notna(r["idate"]): by_dt[(r["vg"], r["idate"].date(), r["tot"])].append(i)
        by_amt[(r["vg"], r["tot"])].append(i)
used = {}          # b2 idx -> reg idx (first claim wins)
reg_verdict = [""] * len(reg)
reg_b2idx = [None] * len(reg)
def take(cands, ri):
    for bi in cands:
        if bi not in used:
            used[bi] = ri
            return bi
    return cands[0] if cands else None   # duplicate claim allowed but marked
for ri, r in reg.iterrows():
    if not r["vg"]:
        reg_verdict[ri] = "No vendor GSTIN (URD/ISD/RCM-self)"
        continue
    c1 = by_inv.get((r["vg"], r["inv"])) if r["inv"] else None
    if c1:
        bi = take(c1, ri); reg_b2idx[ri] = bi
        reg_verdict[ri] = "Matched (invoice no)"
        continue
    if pd.notna(r["ddate"]):
        c2 = by_dt.get((r["vg"], r["ddate"].date(), r["tot"]))
        if c2:
            bi = take(c2, ri); reg_b2idx[ri] = bi
            reg_verdict[ri] = "Matched (date+amount)"
            continue
    c3 = by_amt.get((r["vg"], r["tot"]))
    if c3 and len(c3) <= 3 and r["tot"] != 0:
        bi = take(c3, ri); reg_b2idx[ri] = bi
        reg_verdict[ri] = "Matched (vendor+amount)"
        continue
    reg_verdict[ri] = "NOT FOUND IN 2B"
reg["match"] = reg_verdict
reg["b2i"] = reg_b2idx
print("register verdicts:", dict(reg["match"].value_counts()))
# 2B side: claimed month from register
b2["claimed_by_reg"] = ""
b2["reg_claim_month"] = ""
cm = reg["GSTR3B Month"].map(S)
for ri, bi in enumerate(reg_b2idx):
    if bi is not None:
        b2.at[bi, "claimed_by_reg"] = "Claimed in register"
        b2.at[bi, "reg_claim_month"] = cm.iloc[ri]
unm = b2[(b2["claimed_by_reg"]=="") & (b2["vg"]!="")]
print("2B rows not matched by any register claim:", len(unm), "of", len(b2))
for fy in ("2024-25","2025-26","2026-27"):
    s = unm[unm["FY (derived)"]==fy]
    print("  unclaimed-in-register %s: rows %5d | taxable %.2f | total tax %.2f" % (fy, len(s), s["Taxable Value (₹)"].sum(), s["tot"].sum()))
# cross-check vs client marking where present (Bihar/TN)
mk = b2["GSTR 3B Month"].map(S)
have_mk = b2[(mk!="") & (~mk.str.lower().str.startswith("lapsed"))]
agree = (have_mk["claimed_by_reg"]=="Claimed in register").sum()
print("client-marked-claimed rows also matched by register:", agree, "of", len(have_mk))
reg[["match","b2i"]].to_pickle("itc_reg_match.pkl")
b2.to_pickle("b2itc.pkl")
print("saved")
