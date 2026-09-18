"""Tax comp report - deterministic reasons engine (no master writes here; tcr_reasons.py imports this).
Facts established from the data (18-09):
  * D-F (3B ITC claimed)  == filed 4A(4) + 4A(5) - 4D(1)   for all 204 rows (portal definition; the sheet header text says -4B(2))
  * H-J (2B auto-drafted) is the portal figure; the CA's working carries its own '2B current month ITC as per portal' (A_2B) + ISD (A_4A4)
CA method (Computation New, PART A): 4A5 filed = A_2B (2B net of CN) + reclaim + permanent reversal; 4D1 filed = reclaim + permanent
reversal; 4A4 = ISD; 4B2 = carry-forward (unmatched 2B); 4B1 = permanent reversal.  When the method is followed exactly the portal
comparison closes to zero, so every rupee of difference is one of:
  c_4D1  = -(filed 4D1 - (reclaim + permrev))            -> S-U  'Difference to be added to 4D1'
  c_2B   = (A_2B + A_4A4) - portal 2B                     -> V-X  'Difference in GSTR-2B amounts' (or AB-AD when it equals the CN netted in 4A5)
  c_isd  = filed 4A4 - A_4A4                              -> Y-AA
  c_4A5  = filed 4A5 - (A_2B + reclaim + permrev)         -> Y-AA (sub-explained by the month's accounting correction / amendment rows when they tie)
  residual = diff - sum(components)                       -> stays in AE-AG (sheet formula), flagged for CA review
"""
import pickle, collections
HEADS = ["IGST", "CGST", "SGST"]
def rs(x):
    """Indian grouping, 2dp trimmed: 1505238.5 -> 15,05,238.50 ; 35640 -> 35,640"""
    neg = x < 0; s = "%.2f" % abs(x); ip, fp = s.split(".")
    if len(ip) > 3:
        head, last3 = ip[:-3], ip[-3:]; groups = []
        while len(head) > 2: groups.insert(0, head[-2:]); head = head[:-2]
        if head: groups.insert(0, head)
        ip = ",".join(groups) + "," + last3
    out = ip + ("" if fp == "00" else "." + fp)
    return ("-" if neg else "") + out
def head_str(vec, tol=1.0):
    """'IGST Rs. X / C-SGST Rs. Y' for the non-zero heads of a 3-vector (C/SGST merged when equal)."""
    parts = []
    if abs(vec[0]) > tol: parts.append("IGST Rs. %s" % rs(vec[0]))
    if abs(vec[1]) > tol or abs(vec[2]) > tol:
        if abs(vec[1] - vec[2]) <= tol: parts.append("C/SGST Rs. %s each" % rs(vec[1]))
        else: parts.append("CGST Rs. %s, SGST Rs. %s" % (rs(vec[1]), rs(vec[2])))
    return " & ".join(parts)
def V(rec, k): return rec.get(k) or [0.0, 0.0, 0.0]
def nz(v, tol=1.0): return any(abs(x) > tol for x in v)
def sub(a, b): return [a[j] - b[j] for j in range(3)]
def add(a, b): return [a[j] + b[j] for j in range(3)]
def decompose(row, filed, comp):
    """row: tcr_base row dict (b3, b2, short); filed: dict 4A3.. -> [I,C,S] (filed 3B); comp: extractor-2 record or None.
    Returns dict with components, reason lines, impact, buckets {S,V,Y,AB} (3-vectors) and flags."""
    diff = row["short"][:3]; out = {"diff": diff, "lines": [], "buckets": {"S": [0, 0, 0], "V": [0, 0, 0], "Y": [0, 0, 0], "AB": [0, 0, 0]}, "flags": []}
    tol = 1.0
    if not nz(diff, tol):
        out["reason"] = "Matched"; out["impact"] = "NA"; out["note"] = "NA - Minor amount/ positive diff in 8D"; return out
    if filed is None:
        out["reason"] = "Filed GSTR-3B figures not available in 3B Data for this month - difference Rs. %s not analysed" % rs(sum(diff)); out["impact"] = "Pending"; out["flags"].append("no filed"); return out
    b3_chk = sub(add(filed["4A4"], filed["4A5"]), filed["4D1"])
    if nz(sub(b3_chk, row["b3"][:3]), 1.5): out["flags"].append("3B col != filed 4A4+4A5-4D1")
    if comp is None or not comp["vals"].get("A_2B"):
        out["reason"] = ("GSTR-3B working file (Computation sheet) not available for this month - difference %s not explained; CA review" % head_str(diff)); out["impact"] = "Pending"; out["flags"].append("no computation"); return out
    A = comp["vals"]; a2b, aisd, arec, aperm, acn = V(A, "A_2B"), V(A, "A_4A4"), V(A, "A_reclaim"), V(A, "A_permrev"), V(A, "A_4B2")
    exp_4A5 = add(add(a2b, arec), aperm); exp_4D1 = add(arec, aperm)
    c_4D1 = [-(filed["4D1"][j] - exp_4D1[j]) for j in range(3)]
    c_2B = sub(add(a2b, aisd), row["b2"][:3])
    c_isd = sub(filed["4A4"], aisd)
    c_4A5 = sub(filed["4A5"], exp_4A5)
    resid = [diff[j] - (c_4D1[j] + c_2B[j] + c_isd[j] + c_4A5[j]) for j in range(3)]
    out.update(c_4D1=c_4D1, c_2B=c_2B, c_isd=c_isd, c_4A5=c_4A5, resid=resid)
    lines = []; impacts = []
    # (0) 2B net of CN negative for a head (CN exceeds ITC) and 4A(5) filed as nil: the floored amount is its own component
    floored = [(-a2b[j] if (a2b[j] < -tol and abs(filed["4A5"][j]) <= tol) else 0.0) for j in range(3)]
    c_4A5 = sub(c_4A5, floored)
    if nz(floored, tol):
        lines.append("%s - Not reported in 4A(5)-Negative [2B current month net of CN is negative %s (CN exceeds ITC); 4A(5) filed as nil]" % (head_str(floored), head_str([a2b[j] if floored[j] else 0 for j in range(3)])))
        out["buckets"]["Y"] = add(out["buckets"]["Y"], floored); impacts.append("Pending")
    # (a) amounts that sit in BOTH 4A(5) and 4D(1) beyond the Computation (extra reclaim, +) or in NEITHER (permanent reversal /
    #     reclaim shown in the Computation but not reported, -): they net off in the comparison - report once, no bucket.
    overlap = [0.0, 0.0, 0.0]
    for j in range(3):
        if c_4D1[j] * c_4A5[j] < 0 and abs(c_4D1[j]) > tol and abs(c_4A5[j]) > tol:
            overlap[j] = (1 if c_4A5[j] > 0 else -1) * min(abs(c_4D1[j]), abs(c_4A5[j]))
    c_4A5 = sub(c_4A5, overlap); c_4D1 = add(c_4D1, overlap)
    pos = [x if x > 0 else 0 for x in overlap]; neg = [x if x < 0 else 0 for x in overlap]
    if nz(pos, tol):
        lines.append("%s - reclaim reported in both 4A(5) and 4D(1) over and above the Computation's reclaim row - nets off in the comparison, no impact" % head_str(pos))
    if nz(neg, tol):
        what = "Permanent reversal" if nz(aperm, tol) and not nz(sub([-x for x in neg], aperm), 1.5) else "Permanent reversal / reclaim"
        lines.append("%s - %s as per Computation reported in neither 4A(5) nor 4D(1) - nets off in the comparison, no impact" % (head_str([-x for x in neg]), what))
    # (b) 4D(1) deviations -> S-U 'Difference to be added to 4D1'
    if nz(c_4D1, tol):
        neg_exp = [j for j in range(3) if abs(c_4D1[j]) > tol and exp_4D1[j] < -tol and abs(filed["4D1"][j]) <= tol]
        if neg_exp and all(abs(c_4D1[j]) <= tol or j in neg_exp for j in range(3)):
            lines.append("%s - Not reported in 4D(1)-Negative [reclaim %s + permanent reversal %s as per Computation is negative; filed 4D1 %s]" % (
                head_str(c_4D1), head_str(arec) or "nil", head_str(aperm) or "nil", head_str(filed["4D1"]) or "nil"))
        else:
            under = sum(c_4D1) > 0   # filed 4D1 smaller than reclaim + permrev -> 3B claimed overstated
            lines.append("%s - %s in 4D(1) [filed 4D1 %s vs reclaim %s + permanent reversal %s as per Computation]" % (
                head_str(c_4D1), "Not reported" if under else "Excess reported", head_str(filed["4D1"]) or "nil", head_str(arec) or "nil", head_str(aperm) or "nil"))
        out["buckets"]["S"] = c_4D1; impacts.append("Add to 4D1")
    # (c) portal 2B vs the CA's 2B working -> AB-AD when it is the CN netted in 4A(5), else V-X
    if nz(c_2B, tol):
        racc, ramd = V(A, "R_acc"), V(A, "R_amend")
        if not nz(add(c_2B, acn), 1.5) and nz(acn, tol):
            lines.append("%s - CN netted in 4A(5) instead of reversal in 4B(2) [2B gross %s less CN %s = %s reported]" % (head_str(c_2B), head_str(V(A, "A_4A5")), head_str(acn), head_str(a2b)))
            out["buckets"]["AB"] = c_2B; impacts.append("Ignore")
        elif nz(ramd, tol) and not nz(add(c_2B, ramd), 1.5):
            lines.append("%s - 2B amount expensed out / amended as per Computation reconciliation (amendment-expense out %s)" % (head_str(c_2B), head_str(ramd)))
            out["buckets"]["V"] = c_2B; impacts.append("Ignore")
        else:
            lines.append("%s - Mismatch in 2B amounts [portal auto-drafted 2B %s vs 2B as per Computation %s (2B current month %s + ISD %s)] - amendment / 2B glitch to be checked" % (
                head_str(c_2B), head_str(row["b2"][:3]) or "nil", head_str(add(a2b, aisd)) or "nil", head_str(a2b) or "nil", head_str(aisd) or "nil"))
            out["buckets"]["V"] = c_2B; impacts.append("Pending")
    # (d) ISD -> Y-AA
    if nz(c_isd, tol):
        lines.append("%s - ISD credit reported in 4A(4) %s vs %s as per Computation" % (head_str(c_isd), head_str(filed["4A4"]) or "nil", head_str(aisd) or "nil"))
        out["buckets"]["Y"] = add(out["buckets"]["Y"], c_isd); impacts.append("Pending")
    # (e) remaining 4A(5) deviation -> Y-AA
    if nz(c_4A5, tol):
        racc, ramd = V(A, "R_acc"), V(A, "R_amend"); adj = sub(racc, ramd)
        floored = [j for j in range(3) if abs(c_4A5[j]) > tol and exp_4A5[j] < -tol and abs(filed["4A5"][j]) <= tol]
        if nz(adj, tol) and not nz(sub(c_4A5, adj), 1.5):
            lines.append("%s - Accounting correction %s / amendment-expense out %s adjusted in 4A(5) as per Computation reconciliation" % (head_str(c_4A5), head_str(racc) or "nil", head_str(ramd) or "nil"))
        elif floored and all(abs(c_4A5[j]) <= tol or j in floored for j in range(3)):
            lines.append("%s - Not reported in 4A(5)-Negative [2B current month net of CN is negative %s (CN exceeds ITC); 4A(5) filed as nil]" % (head_str(c_4A5), head_str(exp_4A5)))
        else:
            lines.append("%s - %s reported in 4A(5) other than CN [filed 4A5 %s vs 2B current month %s + reclaim %s + permanent reversal %s] - not explained by Computation sheet" % (
                head_str(c_4A5), "Excess" if sum(c_4A5) > 0 else "Short", head_str(filed["4A5"]) or "nil", head_str(a2b) or "nil", head_str(arec) or "nil", head_str(aperm) or "nil"))
        out["buckets"]["Y"] = add(out["buckets"]["Y"], c_4A5); impacts.append("Pending")
    if nz(resid, 1.5):
        lines.append("%s - residual not reconciled with Computation sheet - CA review" % head_str(resid)); impacts.append("Pending"); out["flags"].append("residual")
    if not lines:   # components individually within tolerance but total > 1: rounding across heads
        out["reason"] = "Matched"; out["impact"] = "NA"; out["note"] = "NA - Minor amount/ positive diff in 8D"; return out
    out["lines"] = lines; out["reason"] = "\n".join(lines)
    out["impact"] = "Add to 4D1" if "Add to 4D1" in impacts else ("Pending" if "Pending" in impacts else "Ignore")
    out["note"] = "NA - Minor amount/ positive diff in 8D" if out["impact"] != "Pending" else ""
    return out
def load():
    b = pickle.load(open("tcr_base.pkl", "rb")); c2 = pickle.load(open("tcr_comp2.pkl", "rb"))
    return b["tcr"], b["filed"], c2["data"], c2["missing"]
if __name__ == "__main__":
    tcr, filed, comp, missing = load()
    print("rows", len(tcr), "| comp", len(comp), "| missing", len(missing))
    # identity survey
    ids = collections.Counter()
    for r in tcr:
        k = (r["gstin"], r["month"]); f = filed.get(k); c = comp.get(k)
        if not f or not c: continue
        A = c["vals"]
        if "A_2B" in A: ids["b2 == A_2B + A_4A4"] += not nz(sub(add(V(A, "A_2B"), V(A, "A_4A4")), r["b2"][:3]), 1.5); ids["b2 == A_4A5 + A_4A4 (gross of CN)"] += not nz(sub(add(V(A, "A_4A5"), V(A, "A_4A4")), r["b2"][:3]), 1.5)
        ids["filed 4A5 == A_2B+reclaim+permrev"] += not nz(sub(f["4A5"], add(add(V(A, "A_2B"), V(A, "A_reclaim")), V(A, "A_permrev"))), 1.5)
        ids["filed 4D1 == reclaim+permrev"] += not nz(sub(f["4D1"], add(V(A, "A_reclaim"), V(A, "A_permrev"))), 1.5)
        ids["filed 4A4 == A_4A4"] += not nz(sub(f["4A4"], V(A, "A_4A4")), 1.5)
        ids["filed 4B2 == A_cf"] += not nz(sub([abs(x) for x in f["4B2"]], V(A, "A_cf")), 1.5)
        ids["filed 4B1 == permrev"] += not nz(sub([abs(x) for x in f["4B1"]], V(A, "A_permrev")), 1.5)
        for p, q in (("P_4A5", "4A5"), ("P_4D1", "4D1"), ("P_4A4", "4A4")): ids["working portal-block %s == filed" % q] += not nz(sub(V(A, p), f[q]), 1.5)
        ids["n"] += 1
    print("identities:", dict(ids))
    res = collections.Counter(); flags = collections.Counter(); imp = collections.Counter()
    for r in tcr:
        k = (r["gstin"], r["month"]); d = decompose(r, filed.get(k), comp.get(k)); r["_d"] = d
        res[d["reason"].split("\n")[0][:40] if d["reason"] in ("Matched",) else ("explained" if not d["flags"] else ",".join(d["flags"]))] += 1; imp[d["impact"]] += 1
    print("outcome:", dict(res)); print("impact:", dict(imp))
    for r in tcr:
        d = r["_d"]
        if d["reason"] != "Matched": print("\n%s %s | diff %s\n   %s | %s" % (r["state"], r["month"], head_str(d["diff"]), d["impact"], d["reason"].replace("\n", "\n   ")))
