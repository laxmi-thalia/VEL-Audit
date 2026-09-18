"""ITC step 1b: (a) fallback 2B for J&K / Arunachal / Kerala from the FY 25-26 Portal
Reports 2B files (their July workings lack the GSTR 2B sheet); (b) merge the ISD sheets
(all states) and match the register's ISD rows. Updates register stamps for previously
NOT-FOUND rows, appends the new 2B rows (source-tagged) + a '2B ISD Data' sheet."""
import os, re, pandas as pd, openpyxl, datetime, warnings; warnings.filterwarnings("ignore")
import win32com.client as win32, pythoncom
from collections import defaultdict
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
def ninv(v): return re.sub(r"[^A-Z0-9]", "", S(v).upper().lstrip("'"))
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
B = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-2B/"
FALL = {"Jammu and Kashmir": "Jammu Kashmir", "Arunachal Pradesh": "Arunachal Pradesh", "Kerala": "Kerala"}
supp_frames = []; isd_frames = []
for fn in sorted(os.listdir(B)):
    if not fn.lower().endswith(".xlsx"): continue
    st = fn.split("LIMITED-")[1].split("-Apr")[0]
    xf = pd.ExcelFile(B + fn)
    if st in FALL:
        sheet = "Purchase-Net" if "Purchase-Net" in xf.sheet_names else "Purchase"
        df = xf.parse(sheet)
        df.columns = [S(c) for c in df.columns]
        if len(df):
            if "Taxable Value (Net)" not in df.columns and "Taxable Value" in df.columns:
                for a, b_ in (("Taxable Value (Net)","Taxable Value"),("IGST (Net)","IGST"),("CGST (Net)","CGST"),("SGST (Net)","SGST")):
                    df[a] = df.get(b_, 0)
            df["State folder"] = FALL[st]
            supp_frames.append(df)
        print("  fallback %-20s rows %d" % (st, len(df)))
    if "ISD" in xf.sheet_names:
        di = xf.parse("ISD")
        di.columns = [S(c) for c in di.columns]
        if len(di):
            di["State folder"] = st
            isd_frames.append(di)
supp = pd.concat(supp_frames, ignore_index=True) if supp_frames else pd.DataFrame()
isd = pd.concat(isd_frames, ignore_index=True) if isd_frames else pd.DataFrame()
print("supplemental B2B rows:", len(supp), "| ISD rows:", len(isd))
reg = pd.read_pickle("itc_working.pkl").reset_index(drop=True)
mt = pd.read_pickle("itc_reg_match.pkl").reset_index(drop=True)
b2 = pd.read_pickle("b2itc.pkl").reset_index(drop=True)
for c in ("IGST","CGST","SGST"):
    reg[c+"_n"] = pd.to_numeric(reg[c], errors="coerce").fillna(0)
reg["tot"] = (reg["IGST_n"]+reg["CGST_n"]+reg["SGST_n"]).round(2)
reg["vg"] = reg["Vendor GSTIN"].map(lambda v: S(v).upper())
reg["inv"] = reg["Reference"].map(ninv)
reg["match"] = mt["match"]; reg["b2i"] = mt["b2i"]

# ---- (a) fallback matching for the 3 states
NEW = []
if len(supp):
    for c in ("Taxable Value (Net)","IGST (Net)","CGST (Net)","SGST (Net)"):
        supp[c] = pd.to_numeric(supp.get(c), errors="coerce").fillna(0)
    supp["vg"] = supp["Supplier GSTIN"].map(lambda v: S(v).upper())
    supp["inv"] = supp["Doc No"].map(ninv)
    supp["tot"] = (supp["IGST (Net)"]+supp["CGST (Net)"]+supp["SGST (Net)"]).round(2)
    by_inv = defaultdict(list); by_amt = defaultdict(list)
    for i, r in supp.iterrows():
        if r["vg"]:
            if r["inv"]: by_inv[(r["vg"], r["inv"])].append(i)
            by_amt[(r["vg"], r["tot"])].append(i)
    STF = {"Jammu & Kashmir":"Jammu Kashmir","Arunachal Pradesh":"Arunachal Pradesh","Kerala":"Kerala"}
    fixed = 0
    b2n = len(b2)
    supp_used = {}
    for ri, r in reg.iterrows():
        if r["match"] != "NOT FOUND IN 2B": continue
        if S(r["STATE NAME"]) not in STF: continue
        c1 = by_inv.get((r["vg"], r["inv"])) if r["inv"] else None
        c2 = None if c1 else (by_amt.get((r["vg"], r["tot"])) if r["tot"] else None)
        cands = c1 or (c2 if c2 and len(c2) <= 3 else None)
        if not cands: continue
        si = cands[0]
        reg.at[ri, "match"] = "Matched (portal 2B FY25-26)" + ("" if c1 else " by amount")
        supp_used.setdefault(si, []).append(ri)
        fixed += 1
    print("fallback-matched register rows:", fixed)
    # append matched supp rows into b2 frame (for the sheet)
    rows2 = []
    for si, ris in supp_used.items():
        r = supp.loc[si]
        rows2.append({"State folder": r["State folder"], "FY (derived)": "2025-26", "F.Y": "2025-26",
            "2B Return Period": S(r.get("Tax Period")), "GSTR 3B Month": "", "Curent previous": "",
            "GSTIN of supplier": r["vg"], "Trade/Legal name": S(r.get("Supplier Name")),
            "Invoice number": S(r.get("Doc No")), "Invoice type": S(r.get("Doc Type")),
            "Invoice Date": r.get("Doc Date"), "Invoice Value(₹)": r.get("Doc Value"),
            "Place of supply": S(r.get("Place of Supply")), "Supply Attract Reverse Charge": S(r.get("Reverse Charge")),
            "Rate": None, "Taxable Value (₹)": r["Taxable Value (Net)"], "Integrated Tax(₹)": r["IGST (Net)"],
            "Central Tax(₹)": r["CGST (Net)"], "State/UT Tax(₹)": r["SGST (Net)"], "Cess(₹)": 0,
            "GSTR-1/5 Period": "", "GSTR-1/5 Filing Date": None, "ITC Availability": S(r.get("ITC Eligible")),
            "Reason": "", "Source": "Portal 2B FY25-26 (fallback)", "IRN": S(r.get("IRN")),
            "Remarks": "", "claimed_by_reg": "Claimed in register", "reg_claim_month": S(reg.at[supp_used[si][0], "GSTR3B Month"]),
            "vg": r["vg"], "inv": r["inv"], "tot": r["tot"], "idate": pd.NaT})
    if rows2:
        b2 = pd.concat([b2, pd.DataFrame(rows2)], ignore_index=True)
        # map register rows to the appended indices + apportion
        k = b2n
        for si, ris in supp_used.items():
            dI = float(supp.at[si,"IGST (Net)"]); dC = float(supp.at[si,"CGST (Net)"]); dS = float(supp.at[si,"SGST (Net)"])
            wts = [max(float(reg.at[ri,"tot"]),0.0) for ri in ris]; tw = sum(wts) or len(ris)
            accI=accC=accS=0.0
            for j, ri in enumerate(ris):
                reg.at[ri,"b2i"] = k
                if j == len(ris)-1: i_,c_,s_ = dI-accI, dC-accC, dS-accS
                else:
                    f_ = (wts[j] if sum(wts) else 1.0)/tw
                    i_,c_,s_ = round(dI*f_,2), round(dC*f_,2), round(dS*f_,2); accI+=i_;accC+=c_;accS+=s_
                NEW.append((ri, i_, c_, s_, S(supp.at[si,"Doc No"]), supp.at[si,"Doc Date"], supp.at[si,"vg"], S(supp.at[si,"Tax Period"])))
            k += 1

# ---- (b) ISD matching
isd_matched = 0
if len(isd):
    for c in ("IGST","CGST","SGST"):
        isd[c] = pd.to_numeric(isd.get(c), errors="coerce").fillna(0)
    isd["tot"] = (isd["IGST"]+isd["CGST"]+isd["SGST"]).round(2)
    isd["inv"] = isd["Doc No"].map(ninv)
    by_amt_isd = defaultdict(list)
    by_inv_isd = defaultdict(list)
    ST2G = {"Andhra Pradesh":"37AAECR0503Q1Z7"}
    for i, r in isd.iterrows():
        cg = S(r.get("Company GSTIN")).upper()
        by_inv_isd[(cg, r["inv"])].append(i)
        by_amt_isd[(cg, r["tot"])].append(i)
    G = {"Jammu & Kashmir":"01AAECR0503Q1ZM","Punjab":"03AAECR0503Q1ZI","Haryana":"06AAECR0503Q1ZC",
    "Rajasthan":"08AAECR0503Q1Z8","Uttar Pradesh":"09AAECR0503Q1Z6","Bihar":"10AAECR0503Q1ZN",
    "Arunachal Pradesh":"12AAECR0503Q1ZJ","Assam":"18AAECR0503Q1Z7","West Bengal":"19AAECR0503Q1Z5",
    "Jharkhand":"20AAECR0503Q1ZM","Chhattisgarh":"22AAECR0503Q1ZI","Madhya Pradesh":"23AAECR0503Q1ZG",
    "Gujarat":"24AAECR0503Q1ZE","Maharashtra":"27AAECR0503Q1Z8","Karnataka":"29AAECR0503Q1Z4",
    "Kerala":"32AAECR0503Q1ZH","Tamil Nadu":"33AAECR0503Q1ZF","Telangana":"36AAECR0503Q1Z9",
    "Andhra Pradesh":"37AAECR0503Q1Z7"}
    for ri, r in reg.iterrows():
        if S(r["TYPE"]) != "ISD" or r["match"] not in ("NOT FOUND IN 2B","No vendor GSTIN (URD/ISD/RCM-self)"): continue
        my = G.get(S(r["STATE NAME"]), "")
        c1 = by_inv_isd.get((my, r["inv"])) if r["inv"] else None
        c2 = None if c1 else by_amt_isd.get((my, r["tot"]))
        cands = c1 or (c2 if c2 and len(c2) <= 3 else None)
        if cands:
            reg.at[ri, "match"] = "Matched (2B ISD section)"
            si = cands[0]
            NEW.append((ri, float(isd.at[si,"IGST"]), float(isd.at[si,"CGST"]), float(isd.at[si,"SGST"]),
                        S(isd.at[si,"Doc No"]), isd.at[si,"Doc Date"], S(isd.at[si,"GSTIN"]), S(isd.at[si,"Tax Period"])))
            isd_matched += 1
    print("ISD-matched register rows:", isd_matched)
print("updated register verdicts:", dict(reg["match"].value_counts()))
reg[["match","b2i"]].to_pickle("itc_reg_match.pkl")
b2.to_pickle("b2itc.pkl")

# ---- write into master
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    ws = wb.Worksheets("ITC Register 2025-26")
    H = {}
    for c in range(1, 90):
        h = str(ws.Cells(5, c).Value or "").strip()
        if h and h not in H: H[h] = c
    R0 = 6
    # update only changed rows
    for (ri, i_, c_, s_, dno, ddt, vgs, per) in NEW:
        r = R0 + ri
        ws.Cells(r, H["2B_IGST"]).Value = i_
        ws.Cells(r, H["2B_CGST"]).Value = c_
        ws.Cells(r, H["2B_SGST"]).Value = s_
        ws.Cells(r, H["Correct Invoice no."]).Value = dno
        ws.Cells(r, H["Invoice as per 2B"]).Value = dno
        if ddt is not None and not (isinstance(ddt, float) and pd.isna(ddt)) and ddt is not pd.NaT:
            ws.Cells(r, H["Correct Invoice date"]).Value = ddt.to_pydatetime() if isinstance(ddt, pd.Timestamp) else ddt
        ws.Cells(r, H["Correct GSTIN"]).Value = vgs
        ws.Cells(r, H["Supplier GSTN as per 2B"]).Value = vgs
        ws.Cells(r, H["2B Period"]).Value = per
    # rewrite whole match column (verdicts changed)
    ws.Range(ws.Cells(R0, H["Invoice Level Match"]), ws.Cells(R0 + len(reg) - 1, H["Invoice Level Match"])).Value = [[v] for v in reg["match"]]
    # append supplemental rows to 2B sheet
    ws2 = wb.Worksheets("GSTR-2B ITC Data")
    COLS = ["State folder","FY (derived)","F.Y","2B Return Period","GSTR 3B Month","Curent previous",
            "GSTIN of supplier","Trade/Legal name","Invoice number","Invoice type","Invoice Date",
            "Invoice Value(₹)","Place of supply","Supply Attract Reverse Charge","Rate","Taxable Value (₹)",
            "Integrated Tax(₹)","Central Tax(₹)","State/UT Tax(₹)","Cess(₹)","GSTR-1/5 Period",
            "ITC Availability","Reason","Source","IRN","claimed_by_reg","reg_claim_month"]
    startb = 48124
    extra = b2.iloc[startb:]
    if len(extra):
        grid = []
        for row in extra[COLS].itertuples(index=False):
            rr = []
            for v in row:
                if v is None or v is pd.NaT: rr.append(None)
                elif isinstance(v, pd.Timestamp): rr.append(v.to_pydatetime())
                elif not isinstance(v, (str, datetime.datetime, datetime.date)) and pd.isna(v): rr.append(None)
                else: rr.append(v)
            grid.append(rr)
        r0 = 6 + startb
        ws2.Range(ws2.Cells(r0, 1), ws2.Cells(r0 + len(grid) - 1, len(COLS))).Value = grid
        print("appended", len(grid), "fallback rows to GSTR-2B ITC Data")
    # ISD data sheet
    for w in list(wb.Worksheets):
        if w.Name == "2B ISD Data": w.Delete()
    ws3 = wb.Worksheets.Add(None, ws2); ws3.Name = "2B ISD Data"; ws3.Activate()
    ws3.Rows(1).RowHeight = 21.0
    ws3.Cells(2, 1).Value = "GSTR-2B ISD section - merged from the Portal Reports 2B files (FY 25-26)."
    ws3.Cells(2, 1).Font.Bold = True
    ICOLS = ["State folder","Company GSTIN","Tax Period","Doc Type","Doc No","Doc Date","GSTIN","Trade Name","IGST","CGST","SGST","Cess","ITC Eligible"]
    hr2 = ws3.Range(ws3.Cells(4, 1), ws3.Cells(4, len(ICOLS))); hr2.Value = [ICOLS]
    hr2.Font.Bold = True; hr2.Font.Color = 0xFFFFFF; hr2.Interior.Color = 0x794E1F
    if len(isd):
        grid = []
        for row in isd.reindex(columns=ICOLS).itertuples(index=False):
            rr = []
            for v in row:
                if v is None or v is pd.NaT: rr.append(None)
                elif isinstance(v, pd.Timestamp): rr.append(v.to_pydatetime())
                elif not isinstance(v, (str, datetime.datetime, datetime.date)) and pd.isna(v): rr.append(None)
                else: rr.append(v)
            grid.append(rr)
        ws3.Range(ws3.Cells(5, 1), ws3.Cells(4 + len(grid), len(ICOLS))).Value = grid
        for nm in ("IGST","CGST","SGST","Cess"):
            ci = ICOLS.index(nm) + 1
            ws3.Range(ws3.Cells(5, ci), ws3.Cells(4 + len(grid), ci)).NumberFormat = "#,##0.00"
    for ci, w_ in ((1,14),(2,18),(5,16),(7,18),(8,24)): ws3.Columns(ci).ColumnWidth = w_
    shp = ws3.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); shp.Name = "btnIndex"
    shp.Fill.ForeColor.RGB = 0x794E1F; shp.Line.Visible = False; shp.Placement = 2
    t = shp.TextFrame2.TextRange; t.Text = "<< INDEX"; t.Font.Bold = True; t.Font.Size = 9; t.Font.Fill.ForeColor.RGB = 0xFFFFFF
    ws3.Hyperlinks.Add(Anchor=shp, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
    ix = wb.Worksheets("INDEX")
    last = ix.Cells(ix.Rows.Count, 2).End(-4162).Row
    have = {str(ix.Cells(r, 2).Value or "").strip() for r in range(5, last + 1)}
    if "GSTR-2B ISD section" not in have:
        r = last + 1
        ix.Cells(r, 1).Value = "ITC"; ix.Cells(r, 2).Value = "GSTR-2B ISD section"
        ix.Hyperlinks.Add(Anchor=ix.Cells(r, 7), Address="", SubAddress="'2B ISD Data'!A1", TextToDisplay="2B ISD Data")
        for c in range(1, 11): ix.Cells(r, c).Borders.LineStyle = 1
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    b = ws.Cells(4, H["B_Total GST"]).Value; t2 = ws.Cells(4, H["2B_Total GST"]).Value; d = ws.Cells(4, H["D_Total GST"]).Value
    print("B_Total %.2f | 2B_Total %.2f | D_Total %.2f" % (b, t2, d))
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot)
    print("goldens: T3", wb.Worksheets("SR_2025-26").Range("T3").Value, "| RCM", wb.Worksheets("RCM Register").Cells(4, 44).Value)
    wb.Close(SaveChanges=True)
    print("VERDICT:", "PASS" if tot == 0 else "FAIL")
finally:
    xl.Quit()
