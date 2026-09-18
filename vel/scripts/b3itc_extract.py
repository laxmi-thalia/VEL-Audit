"""Extract 3B Table-4 figures per GSTIN-month from Octa files: 4A total, 4B total,
Net ITC (A-B), plus 4A(3) RCM and 4A(4) ISD separately. -> b3itc.pkl"""
import os, re, warnings; warnings.filterwarnings("ignore")
import pandas as pd
B = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-3B/"
MEAS = {"Integrated Tax": "igst", "Central Tax": "cgst", "State/UT Tax": "sgst"}
rows = []
for f in sorted(os.listdir(B)):
    if not f.lower().endswith((".xlsx", ".xls")): continue
    st = re.sub(r"-2025-26.*$", "", f.split("LIMITED-")[1])
    ov = pd.read_excel(B + f, sheet_name="Overview", header=None)
    gstin = None
    for i in range(len(ov)):
        if str(ov.iloc[i, 1]).strip().lower() == "gstin": gstin = str(ov.iloc[i, 2]).split("(")[0].strip()
    d = pd.read_excel(B + f, sheet_name="GSTR-3B", header=None)
    months = [str(d.iloc[0, j]).strip() for j in range(3, 15)]
    rec = {"state": st, "gstin": gstin}
    for m in months:
        for grp in ("a", "b", "a3", "a4"):
            for v in MEAS.values(): rec["%s|%s|%s" % (m, grp, v)] = 0.0
    for i in range(1, len(d)):
        sec = str(d.iloc[i, 1]).strip(); typ = str(d.iloc[i, 2]).strip()
        if typ not in MEAS: continue
        grp = None
        if re.match(r"^4\.A\.", sec): grp = "a"
        elif re.match(r"^4\.B\.", sec): grp = "b"
        if grp is None: continue
        for j, m in enumerate(months, start=3):
            v = d.iloc[i, j]
            v = float(v) if pd.notna(v) else 0.0
            rec["%s|%s|%s" % (m, grp, MEAS[typ])] += v
            if sec.startswith("4.A.3"): rec["%s|a3|%s" % (m, MEAS[typ])] += v
            if sec.startswith("4.A.4"): rec["%s|a4|%s" % (m, MEAS[typ])] += v
    rows.append(rec)
    print("  %-22s %s" % (st, gstin))
T = pd.DataFrame(rows)
T.to_pickle("b3itc.pkl")
tots = {}
for v in ("igst", "cgst", "sgst"):
    a = sum(T[c].sum() for c in T.columns if c.endswith("|a|" + v))
    b_ = sum(T[c].sum() for c in T.columns if c.endswith("|b|" + v))
    tots[v] = (a, b_, a - b_)
print("\nFY totals: 4A / 4B / Net")
for v, (a, b_, n) in tots.items():
    print("  %s: %15.2f %15.2f %15.2f" % (v, a, b_, n))
print("Net ITC total: %.2f" % sum(n for _, _, n in tots.values()))
