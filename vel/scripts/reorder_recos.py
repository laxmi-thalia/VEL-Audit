"""Standard column sequence for all recos:
[Side] Taxable, IGST, CGST, SGST, Total Tax  x  (SR/GSTR-1/3B/GL side -> comparator -> Difference)
Sheets: S2 Month-on-Month, S3 1 vs 3B, S3 Month-on-Month, S3 SR vs 3B, S3 SR vs 3B MoM, S4 GL vs SR.
Dependents repointed: S2 Reco (CA format) AI remarks, S3 1 vs 3B auto remarks. Pivot rebuilt separately."""
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
wb = openpyxl.load_workbook(P)
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); TOT = Font(bold=True)
NF = "#,##0.00"
ws0 = wb["SR_2025-26"]
H = {S(ws0.cell(4, c).value): L(c) for c in range(1, ws0.max_column + 1)}
SR = lambda n: "'SR_2025-26'!$%s$5:$%s$27006" % (H[n], H[n])
ng = wb["GSTR-1 Data"].max_row
G1 = lambda c: "'GSTR-1 Data'!$%s$2:$%s$%d" % (c, c, ng)
n3 = wb["3B Data"].max_row
B3 = lambda c: "'3B Data'!$%s$2:$%s$%d" % (c, c, n3)
ngl = wb["GL Data"].max_row
GLD = lambda c: "'GL Data'!$%s$2:$%s$%d" % (c, c, ngl)

def hdr_write(ws, row, heads, start=1):
    for i, h in enumerate(heads):
        x = ws.cell(row, start + i, h); x.font = HF; x.fill = HB

def clear(ws, r0, r1, c0, c1):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            ws.cell(r, c).value = None

# ---------- 1. S2 Month-on-Month ----------
mm = wb["S2 Month-on-Month"]
R0, R1 = 4, 1371
old = {}
for r in range(R0, R1 + 1):
    old[r] = dict(
        state=mm.cell(r, 1).value, gstin=mm.cell(r, 2).value, month=mm.cell(r, 3).value, typ=mm.cell(r, 4).value,
        bt=mm.cell(r, 5).value, gt=mm.cell(r, 6).value, bc=mm.cell(r, 8).value, gc=mm.cell(r, 9).value,
        a1=mm.cell(r, 11).value, a2=mm.cell(r, 12).value, rem=mm.cell(r, 14).value, key=mm.cell(r, 15).value,
        bi=mm.cell(r, 16).value, gi=mm.cell(r, 17).value, bs=mm.cell(r, 19).value, gs=mm.cell(r, 20).value)
clear(mm, 3, R1, 1, 21)
hdr_write(mm, 3, ["State", "GSTIN", "Month", "GSTR-1 Type",
 "SR Taxable", "SR IGST", "SR CGST", "SR SGST", "SR Total Tax",
 "GSTR-1 Taxable", "GSTR-1 IGST", "GSTR-1 CGST", "GSTR-1 SGST", "GSTR-1 Total Tax",
 "Diff Taxable", "Diff IGST", "Diff CGST", "Diff SGST", "Diff Total Tax",
 "Auto: books docs not in GSTR-1", "Auto: GSTR-1 docs not in books", "UNEXPLAINED residual (taxable)", "Remark (type here)", "key"])
for r in range(R0, R1 + 1):
    o = old[r]
    if not S(o["state"]):
        continue
    vals = [o["state"], o["gstin"], o["month"], o["typ"],
            o["bt"], o["bi"], o["bc"], o["bs"], "=F%d+G%d+H%d" % (r, r, r),
            o["gt"], o["gi"], o["gc"], o["gs"], "=K%d+L%d+M%d" % (r, r, r),
            "=E%d-J%d" % (r, r), "=F%d-K%d" % (r, r), "=G%d-L%d" % (r, r), "=H%d-M%d" % (r, r), "=I%d-N%d" % (r, r),
            o["a1"], o["a2"], "=O%d-T%d-U%d" % (r, r, r), o["rem"], o["key"]]
    for i, v in enumerate(vals):
        c = mm.cell(r, 1 + i); c.value = v
        if 5 <= i + 1 <= 22:
            c.number_format = NF
for c_, w in zip(range(1, 25), [18, 17, 8, 14] + [14] * 15 + [22, 22, 16, 32, 18]):
    mm.column_dimensions[L(c_)].width = w
mm.freeze_panes = "E4"
mm.auto_filter.ref = "A3:X%d" % R1
print("S2 MoM reordered (24 cols)")

# ---------- 2. S2 Reco (CA format): repoint AI remarks N->W, O->X ----------
rc = wb["S2 Reco (CA format)"]
n = 0
for r in range(54, 75):
    v = rc.cell(r, 35).value
    if isinstance(v, str) and "S2 Month-on-Month" in v:
        rc.cell(r, 35).value = (v
            .replace("'S2 Month-on-Month'!$N$4:$N$1371", "'S2 Month-on-Month'!$W$4:$W$1371")
            .replace("'S2 Month-on-Month'!$O$4:$O$1371", "'S2 Month-on-Month'!$X$4:$X$1371"))
        n += 1
print("S2 Reco remarks repointed:", n)

# ---------- 3. S3 Month-on-Month (GSTR-1 vs 3B) ----------
m3 = wb["S3 Month-on-Month"]
rows3 = []
for r in range(4, m3.max_row + 1):
    if S(m3.cell(r, 1).value) and S(m3.cell(r, 1).value) != "Total":
        rows3.append((r, m3.cell(r, 1).value, m3.cell(r, 2).value, m3.cell(r, 3).value, m3.cell(r, 10).value, m3.cell(r, 11).value))
clear(m3, 3, m3.max_row, 1, 11)
hdr_write(m3, 3, ["State", "GSTIN", "Month",
 "GSTR-1 Taxable", "GSTR-1 IGST", "GSTR-1 CGST", "GSTR-1 SGST", "GSTR-1 Total Tax",
 "3B Taxable", "3B IGST", "3B CGST", "3B SGST", "3B Total Tax",
 "Diff Taxable", "Diff IGST", "Diff CGST", "Diff SGST", "Diff Total Tax", "Remark (type here)", "key"])
maxr3 = 0
for (r, st, g_, mo, rem, key) in rows3:
    maxr3 = max(maxr3, r)
    def g1f(c, rr=r): return '=SUMIFS(%s,%s,$B%d,%s,$C%d)' % (G1(c), G1("A"), rr, G1("K"), rr)
    def b3f(c, rr=r): return '=SUMIFS(%s,%s,$B%d,%s,$C%d)' % (B3(c), B3("B"), rr, B3("C"), rr)
    vals = [st, g_, mo, g1f("F"), g1f("G"), g1f("H"), g1f("I"), "=E%d+F%d+G%d" % (r, r, r),
            b3f("D"), b3f("E"), b3f("F"), b3f("G"), "=J%d+K%d+L%d" % (r, r, r),
            "=D%d-I%d" % (r, r), "=E%d-J%d" % (r, r), "=F%d-K%d" % (r, r), "=G%d-L%d" % (r, r), "=H%d-M%d" % (r, r),
            rem, key if key is not None else '=B%d&"|"&C%d' % (r, r)]
    for i, v in enumerate(vals):
        c = m3.cell(r, 1 + i); c.value = v
        if 4 <= i + 1 <= 18:
            c.number_format = NF
tr = maxr3 + 1
m3.cell(tr, 1, "Total").font = TOT
for c in range(4, 19):
    m3.cell(tr, c).value = "=SUM(%s4:%s%d)" % (L(c), L(c), maxr3); m3.cell(tr, c).number_format = NF; m3.cell(tr, c).font = TOT
for c_, w in zip(range(1, 21), [18, 17, 8] + [14] * 15 + [32, 18]):
    m3.column_dimensions[L(c_)].width = w
m3.freeze_panes = "D4"
m3.auto_filter.ref = "A3:T%d" % maxr3
print("S3 MoM reordered; %d data rows; total row %d" % (len(rows3), tr))

# ---------- 4. S3 1 vs 3B ----------
s31 = wb["S3 1 vs 3B"]
pairs31 = []
for r in range(4, s31.max_row + 1):
    a = S(s31.cell(r, 1).value)
    if a and a != "Total" and len(a) == 15:
        pairs31.append((r, s31.cell(r, 1).value, s31.cell(r, 2).value))
last31 = max(r for r, _, _ in pairs31) + 1
clear(s31, 3, s31.max_row, 1, 12)
hdr_write(s31, 3, ["GSTIN", "State",
 "GSTR-1 Taxable", "GSTR-1 IGST", "GSTR-1 CGST", "GSTR-1 SGST", "GSTR-1 Total Tax",
 "3B Taxable", "3B IGST", "3B CGST", "3B SGST", "3B Total Tax",
 "Diff Taxable", "Diff IGST", "Diff CGST", "Diff SGST", "Diff Total Tax", "Remarks (auto from S3 Month-on-Month)"])
for (r, g_, st) in pairs31:
    def g1c(c, rr=r): return '=SUMIFS(%s,%s,$A%d)' % (G1(c), G1("A"), rr)
    def b3c(c, rr=r): return '=SUMIFS(%s,%s,$A%d)' % (B3(c), B3("B"), rr)
    rem = ('=TRIM(IFERROR(IF(LEN(INDEX(\'S3 Month-on-Month\'!$S$4:$S$%d,MATCH($A%d&"|"&"Apr-25",\'S3 Month-on-Month\'!$T$4:$T$%d,0)))>0,'
           '"see month rows",""),""))' % (maxr3, r, maxr3))
    rem = ('=IFERROR(TEXTJOIN(" | ",TRUE,IF(LEN(\'S3 Month-on-Month\'!$S$4:$S$%d)*(\'S3 Month-on-Month\'!$B$4:$B$%d=$A%d),'
           '\'S3 Month-on-Month\'!$C$4:$C$%d&": "&\'S3 Month-on-Month\'!$S$4:$S$%d,"")),"")' % (maxr3, maxr3, r, maxr3, maxr3))
    vals = [g_, st, g1c("F"), g1c("G"), g1c("H"), g1c("I"), "=C%d+D%d+E%d" % (r, r, r),
            b3c("D"), b3c("E"), b3c("F"), b3c("G"), "=H%d+I%d+J%d" % (r, r, r),
            "=C%d-H%d" % (r, r), "=D%d-I%d" % (r, r), "=E%d-J%d" % (r, r), "=F%d-K%d" % (r, r), "=G%d-L%d" % (r, r),
            rem]
    for i, v in enumerate(vals):
        c = s31.cell(r, 1 + i); c.value = v
        if 3 <= i + 1 <= 17:
            c.number_format = NF
s31.cell(last31, 2, "Total").font = TOT
for c in range(3, 18):
    s31.cell(last31, c).value = "=SUM(%s4:%s%d)" % (L(c), L(c), last31 - 1); s31.cell(last31, c).number_format = NF; s31.cell(last31, c).font = TOT
for c_, w in zip(range(1, 19), [18, 20] + [14] * 15 + [50]):
    s31.column_dimensions[L(c_)].width = w
print("S3 1 vs 3B reordered; %d states; total row %d" % (len(pairs31), last31))

# ---------- 5. S3 SR vs 3B ----------
sv = wb["S3 SR vs 3B"]
pairs = []
for r in range(4, sv.max_row + 1):
    a = S(sv.cell(r, 1).value)
    if a and len(a) == 15:
        pairs.append((r, sv.cell(r, 1).value, sv.cell(r, 2).value))
lastv = max(r for r, _, _ in pairs) + 1
clear(sv, 3, sv.max_row, 1, 8)
hdr_write(sv, 3, ["GSTIN", "State",
 "SR Taxable (incl adv)", "SR IGST", "SR CGST", "SR SGST", "SR Total Tax",
 "3B Taxable", "3B IGST", "3B CGST", "3B SGST", "3B Total Tax",
 "Diff Taxable", "Diff IGST", "Diff CGST", "Diff SGST", "Diff Total Tax"])
for (r, g_, st) in pairs:
    def srf(nm, rr=r): return '=SUMIFS(%s,%s,$A%d)' % (SR(nm), SR("My GSTIN"), rr)
    def b3c(c, rr=r): return '=SUMIFS(%s,%s,$A%d)' % (B3(c), B3("B"), rr)
    vals = [g_, st, srf("Taxable Value"), srf("IGST Amount"), srf("CGST Amount"), srf("SGST Amount"), "=C%d+D%d+E%d" % (r, r, r),
            b3c("D"), b3c("E"), b3c("F"), b3c("G"), "=H%d+I%d+J%d" % (r, r, r),
            "=C%d-H%d" % (r, r), "=D%d-I%d" % (r, r), "=E%d-J%d" % (r, r), "=F%d-K%d" % (r, r), "=G%d-L%d" % (r, r)]
    vals[6] = "=D%d+E%d+F%d" % (r, r, r)
    for i, v in enumerate(vals):
        c = sv.cell(r, 1 + i); c.value = v
        if 3 <= i + 1 <= 17:
            c.number_format = NF
sv.cell(lastv, 2, "Total").font = TOT
for c in range(3, 18):
    sv.cell(lastv, c).value = "=SUM(%s4:%s%d)" % (L(c), L(c), lastv - 1); sv.cell(lastv, c).number_format = NF; sv.cell(lastv, c).font = TOT
for c_, w in zip(range(1, 18), [18, 20] + [15] * 15):
    sv.column_dimensions[L(c_)].width = w
print("S3 SR vs 3B reordered; total row %d" % lastv)

# ---------- 6. S3 SR vs 3B MoM ----------
sm = wb["S3 SR vs 3B MoM"]
rowsm = []
for r in range(4, sm.max_row + 1):
    a = S(sm.cell(r, 1).value)
    if a and a != "Total":
        rowsm.append((r, sm.cell(r, 1).value, sm.cell(r, 2).value, sm.cell(r, 3).value, sm.cell(r, 10).value))
lastm = max(r for r, _, _, _, _ in rowsm) + 1
clear(sm, 3, sm.max_row, 1, 10)
hdr_write(sm, 3, ["State", "GSTIN", "Month",
 "SR Taxable (incl adv)", "SR IGST", "SR CGST", "SR SGST", "SR Total Tax",
 "3B Taxable", "3B IGST", "3B CGST", "3B SGST", "3B Total Tax",
 "Diff Taxable", "Diff IGST", "Diff CGST", "Diff SGST", "Diff Total Tax", "Remark (type here)"])
for (r, st, g_, mo, rem) in rowsm:
    def srf(nm, rr=r): return '=SUMIFS(%s,%s,$B%d,%s,$C%d)' % (SR(nm), SR("My GSTIN"), rr, SR("Month"), rr)
    def b3f(c, rr=r): return '=SUMIFS(%s,%s,$B%d,%s,$C%d)' % (B3(c), B3("B"), rr, B3("C"), rr)
    vals = [st, g_, mo, srf("Taxable Value"), srf("IGST Amount"), srf("CGST Amount"), srf("SGST Amount"), "=E%d+F%d+G%d" % (r, r, r),
            b3f("D"), b3f("E"), b3f("F"), b3f("G"), "=J%d+K%d+L%d" % (r, r, r),
            "=D%d-I%d" % (r, r), "=E%d-J%d" % (r, r), "=F%d-K%d" % (r, r), "=G%d-L%d" % (r, r), "=H%d-M%d" % (r, r), rem]
    for i, v in enumerate(vals):
        c = sm.cell(r, 1 + i); c.value = v
        if 4 <= i + 1 <= 18:
            c.number_format = NF
sm.cell(lastm, 1, "Total").font = TOT
for c in range(4, 19):
    sm.cell(lastm, c).value = "=SUM(%s4:%s%d)" % (L(c), L(c), lastm - 1); sm.cell(lastm, c).number_format = NF; sm.cell(lastm, c).font = TOT
for c_, w in zip(range(1, 20), [18, 17, 8] + [15] * 15 + [32]):
    sm.column_dimensions[L(c_)].width = w
sm.auto_filter.ref = "A3:S%d" % (lastm - 1)
print("S3 SR vs 3B MoM reordered; %d rows" % len(rowsm))

# ---------- 7. S4 GL vs SR (GL carries no taxable: IGST, CGST, SGST, Total) ----------
s4 = wb["S4 GL vs SR"]
pairs4 = []
for r in range(4, s4.max_row + 1):
    b = S(s4.cell(r, 2).value)
    if b and b != "Total":
        pairs4.append((r, s4.cell(r, 1).value, s4.cell(r, 2).value, s4.cell(r, 9).value))
last4 = max(r for r, _, _, _ in pairs4) + 1
clear(s4, 3, s4.max_row, 1, 10)
hdr_write(s4, 3, ["Business place", "State",
 "GL IGST", "GL CGST", "GL SGST", "GL Total Tax",
 "SR IGST", "SR CGST", "SR SGST", "SR Total Tax",
 "Diff IGST", "Diff CGST", "Diff SGST", "Diff Total Tax",
 "Explained (S4 Exceptions)", "Residual (CGST diff - explained)"])
for (r, bp, st, exf) in pairs4:
    def glf(hd, rr=r): return '=SUMIFS(%s,%s,$B%d,%s,"%s")' % (GLD("J"), GLD("A"), rr, GLD("I"), hd)
    def srf(nm, rr=r): return '=SUMIFS(%s,%s,$B%d)' % (SR(nm), SR("My State"), rr)
    vals = [bp, st, glf("IGST"), glf("CGST"), glf("SGST"), "=C%d+D%d+E%d" % (r, r, r),
            srf("IGST Amount"), srf("CGST Amount"), srf("SGST Amount"), "=G%d+H%d+I%d" % (r, r, r),
            "=C%d-G%d" % (r, r), "=D%d-H%d" % (r, r), "=E%d-I%d" % (r, r), "=F%d-J%d" % (r, r),
            exf, "=L%d-O%d" % (r, r)]
    for i, v in enumerate(vals):
        c = s4.cell(r, 1 + i); c.value = v
        if 3 <= i + 1 <= 16:
            c.number_format = NF
s4.cell(last4, 2, "Total").font = TOT
for c in range(3, 17):
    s4.cell(last4, c).value = "=SUM(%s4:%s%d)" % (L(c), L(c), last4 - 1); s4.cell(last4, c).number_format = NF; s4.cell(last4, c).font = TOT
for c_, w in zip(range(1, 17), [13, 20] + [14] * 12 + [22, 24]):
    s4.column_dimensions[L(c_)].width = w
print("S4 reordered; %d states; total row %d" % (len(pairs4), last4))
wb.save(P)
print("SAVED")
