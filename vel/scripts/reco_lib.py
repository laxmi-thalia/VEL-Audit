"""Matching rules shared by cascade_fix.py (FY 25-26) and fy2627_reco.py (FY 26-27). Pure Python, no Excel.
Layers (Pawan 18-09/21-09, conservative):
  1  exact vendor GSTIN + zero-insensitive invoice (recipient checked; flagged when the 2B document sits under another VEL GSTIN)
  1b malformed vendor GSTIN (not 15 chars) rescued only by PAN + exact invoice, flagged; otherwise 'Not matched - vendor GSTIN invalid'
  2  GSTIN + document amount (+/-100, same recipient, single candidate) / 3 similar invoice + amount (same gates)
  2b line-level date+amount / amount, single candidate, same recipient + FY
  4  FY 24-25 2B (prior-year keys/dates) -> 'Matched with 2B of FY 24-25 - Table 6A1'
Vocabulary is the 21-09 standard (see gst-audit-vel-itc skill changes-log)."""
import re, collections, datetime as dt
TOL_ABS = 100.0
GST = re.compile(r"^\d{2}[A-Z0-9]{13}$")
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v or 0)
    except Exception: return 0.0
def norm(s): return re.sub(r"[ \-/.'_]", "", S(s).upper())
def zkey(s): return norm(re.sub(r"(?<!\d)0+(?=\d)", "", S(s)))   # zeros stripped per digit run BEFORE separators go (007 == 7)
def fy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
def fy_of_label(v, d):
    m = re.search(r"(20)?(\d{2})-(\d{2})", S(v))
    return ("20%s-%s" % (m.group(2), m.group(3))) if m else fy(d)
def inv_core(s):
    s = S(s).upper(); s = re.sub(r"\bINVOICE\b|\bINV\b", "", s); s = re.sub(r"[/\-\s]?\d{2}[-/]\d{2}\s*$", "", s)
    return re.sub(r"[^A-Z0-9]", "", s)
def _lev(a, b):
    prev = list(range(len(b) + 1))
    for x in range(1, len(a) + 1):
        cur = [x]
        for y in range(1, len(b) + 1): cur.append(min(prev[y] + 1, cur[y - 1] + 1, prev[y - 1] + (a[x - 1] != b[y - 1])))
        prev = cur
    return prev[-1]
def inv_gate(a, b):
    A, B = S(a).upper(), S(b).upper()
    if not A or not B or abs(len(A) - len(B)) > 3: return False
    def split(s):
        for sep in ("/", "-", " "):
            if sep in s: i = s.index(sep); return s[:i], s[i + 1:]
        return s, ""
    pa, ta = split(A); pb, tb = split(B)
    return pa == pb and ta.isdigit() and tb.isdigit() and _lev(ta, tb) <= 2
def inv_similar(a, b):
    ca, cb = inv_core(a), inv_core(b)
    return (len(ca) >= 3 and len(cb) >= 3 and (ca in cb or cb in ca)) or inv_gate(a, b)
def classify_vendor_gstin(g):
    s = S(g).upper()
    if not s or len(s) < 5 or s in ("MISSING", "NA", "NONE", "NIL", "URD", "UNREGISTERED", "NOT AVAILABLE"): return ("missing", "")   # '0', '-', 'NA' = no GSTIN
    pan = s[2:12] if len(s) >= 12 else ""
    return ("valid", pan) if GST.match(s) else ("malformed", pan)
def _tot(r): return round(num(r["igst"]) + num(r["cgst"]) + num(r["sgst"]), 2)
# Numbered vocabulary (CA Priyesh 21-09: same numbered remark on the register and the 2B sheet so filter n-n ties on both sides;
# only 10 'Not in 2B' (books only) and 11 'Not in books' (2B only) differ). Non-exact matches carry '– review' except 4 (amount & date tie).
V = {"exact": "1 – Matched with 2B – invoice no",
     "exact_rcpt": "2 – Matched with 2B – invoice no – recipient GSTIN differs (2B under %s) – review",
     "pan_fix": "3 – Matched with 2B – invoice no, vendor GSTIN corrected from 2B (%s) – review",
     "amt_date": "4 – Matched with 2B – amount & date tie, invoice no differs",
     "similar": "5 – Matched with 2B – similar invoice no + amount (±100) – review",
     "gstin_amt": "6 – Matched with 2B – GSTIN + amount (±100), invoice no differs – review",
     "date_amt": "7 – Matched with 2B – date + amount, invoice no differs – review",
     "amt_only": "8 – Matched with 2B – amount only, invoice no differs – review",
     "py": "9 – Matched with 2B of FY 24-25 – Table 6A1",
     "not_in_2b": "10 – Not in 2B – Apr-25 to Aug-26",
     "not_in_books": "11 – Not in books – FY 25-26 claims",
     "rcm": "12 – Not applicable – RCM self-invoice", "isd": "13 – Not applicable – ISD", "urd": "14 – Not applicable – no vendor GSTIN (URD)",
     "invalid": "15 – Not matched – vendor GSTIN invalid (%d chars) – review",
     "rcm_in2b": "12 – Not applicable – RCM line (in 2B: %s)"}
# RCM-category lines are outside the Countif / B_ / 2B_ frame; when their vendor GSTIN + invoice sit in 2B they keep number 12 on both
# sheets (Pawan 22-09) so a '1 –' filter holds ITC lines only. The basis of the match is kept in brackets.
RCM_BASIS = {"1": "invoice no", "2": "invoice no", "3": "invoice no", "4": "amount & date", "5": "similar invoice + amount",
             "6": "GSTIN + amount", "7": "date + amount", "8": "amount", "9": "FY 24-25 2B"}
def match_register(reg, b2, py_keys, py_dates):
    """reg rows: vendor_gstin, invoice, invoice_date, invoice_year, category, vel_gstin, igst, cgst, sgst.
    b2 rows: supplier_gstin, doc_no, doc_date, company_gstin, key, igst, cgst, sgst.
    py_keys: {vendorGSTIN + zkey(invoice)} of the FY 24-25 2B; py_dates: {(vendorGSTIN, date, total)}.
    Returns [{'verdict': str, 'key2': str}] aligned to reg."""
    N = len(reg); verdict = [""] * N; key2 = [""] * N
    b2key = [S(r["key"]) for r in b2]; by_z = collections.defaultdict(list); by_dt = collections.defaultdict(list); by_amt = collections.defaultdict(list); by_pan = collections.defaultdict(list)
    for j, r in enumerate(b2):
        vg = S(r["supplier_gstin"]).upper()
        if not vg: continue
        zk = zkey(r["doc_no"]); by_z[vg + zk].append(j); by_pan[(vg[2:12], zk)].append(j)
        if isinstance(r["doc_date"], dt.datetime): by_dt[(vg, r["doc_date"].date(), _tot(r))].append(j)
        by_amt[(vg, _tot(r))].append(j)
    exact_owned = set()
    def pick_rcpt(i, cands):
        """A8 (M1 23-09): the same supplier can bill the same invoice number to two of the client's states. Among the 2B
        rows with that key, take the one reported under THIS line's recipient GSTIN; only when none is, fall back to the
        first candidate and let other_rcpt() flag it (remark 2)."""
        vel = S(reg[i]["vel_gstin"]).upper()
        for j in cands:
            if j not in exact_owned and vel and S(b2[j]["company_gstin"]).upper() == vel: return j
        free = [j for j in cands if j not in exact_owned]
        return free[0] if free else cands[0]
    def other_rcpt(i, j):
        cg = S(b2[j]["company_gstin"]).upper(); vel = S(reg[i]["vel_gstin"]).upper()
        return cg if cg and vel and cg != vel else ""
    # ---- pass 1: exact (+ malformed-GSTIN PAN rescue)
    for i, r in enumerate(reg):
        kind, pan = classify_vendor_gstin(r["vendor_gstin"]); vg = S(r["vendor_gstin"]).upper(); cat = S(r["category"])
        if kind == "missing":
            verdict[i] = V["rcm"] if cat == "RCM" else (V["isd"] if cat == "ISD" else V["urd"]); continue
        if kind == "malformed":
            c = by_pan.get((pan, zkey(r["invoice"]))) if pan else None
            if c: j = pick_rcpt(i, c); key2[i] = b2key[j]; exact_owned.add(j); verdict[i] = V["pan_fix"] % S(b2[j]["supplier_gstin"]).upper()
            else: verdict[i] = V["invalid"] % len(vg)
            continue
        c = by_z.get(vg + zkey(r["invoice"]))
        if c: j = pick_rcpt(i, c); key2[i] = b2key[j]; exact_owned.add(j); o = other_rcpt(i, j); verdict[i] = (V["exact_rcpt"] % o) if o else V["exact"]
    # ---- pass 2: document-level (same recipient, +/-100, single candidate)
    fb_used = set(); b2doc = {}
    for j, r in enumerate(b2):
        vg = S(r["supplier_gstin"]).upper()
        if not vg or j in exact_owned: continue
        dk = (vg, zkey(r["doc_no"])); d = b2doc.setdefault(dk, {"tot": 0.0, "fy": fy(r["doc_date"]), "rows": [], "rcpt": S(r["company_gstin"]).upper(), "raw": S(r["doc_no"]), "date": r["doc_date"]})
        d["tot"] += _tot(r); d["rows"].append(j)
    by_vendor = collections.defaultdict(list)
    for dk in b2doc: by_vendor[dk[0]].append(dk)
    regdoc = collections.OrderedDict()
    for i, r in enumerate(reg):
        if verdict[i] or S(r["category"]) != "ITC": continue
        vg = S(r["vendor_gstin"]).upper(); dk = (vg, zkey(r["invoice"]))
        d = regdoc.setdefault(dk, {"tot": 0.0, "fy": fy_of_label(r["invoice_year"], r["invoice_date"]), "lines": [], "rcpt": S(r["vel_gstin"]).upper(), "raw": S(r["invoice"]), "date": r["invoice_date"]})
        d["tot"] += _tot(r); d["lines"].append(i)
    used = set()
    def same_day(a, b): return isinstance(a, dt.datetime) and isinstance(b, dt.datetime) and a.date() == b.date()
    for (vg, inv), d in regdoc.items():
        if not d["tot"]: continue
        cands = [dk for dk in by_vendor.get(vg, []) if dk not in used and (not d["fy"] or b2doc[dk]["fy"] == d["fy"]) and b2doc[dk]["rcpt"] == d["rcpt"] and abs(b2doc[dk]["tot"] - d["tot"]) <= TOL_ABS]
        if not cands: continue
        sim = [dk for dk in cands if inv_similar(d["raw"], b2doc[dk]["raw"])]
        if len(sim) == 1: pick = sim[0]
        elif len(cands) == 1: pick = cands[0]
        else: continue
        # CA 21-09: amount ties to the rupee AND the dates agree -> nothing to review, only the invoice number is written differently
        if abs(b2doc[pick]["tot"] - d["tot"]) < 1 and same_day(b2doc[pick]["date"], d["date"]): why = V["amt_date"]
        else: why = V["similar"] if pick in sim else V["gstin_amt"]
        used.add(pick); j0 = b2doc[pick]["rows"][0]; fb_used.update(b2doc[pick]["rows"])
        for i in d["lines"]: key2[i] = b2key[j0]; verdict[i] = why
    # ---- pass 2b: line-level, single candidate; pass 3: prior-year 2B
    for i, r in enumerate(reg):
        if verdict[i]: continue
        vg = S(r["vendor_gstin"]).upper(); tot = _tot(r); d = r["invoice_date"]; ifY = fy_of_label(r["invoice_year"], d)
        ok = lambda j: j not in exact_owned and j not in fb_used and (not ifY or fy(b2[j]["doc_date"]) == ifY) and S(b2[j]["company_gstin"]).upper() == S(r["vel_gstin"]).upper()
        c = [j for j in (by_dt.get((vg, d.date(), tot), []) if isinstance(d, dt.datetime) else []) if ok(j)]
        if len(c) == 1: key2[i] = b2key[c[0]]; fb_used.add(c[0]); verdict[i] = V["date_amt"]; continue
        c = [j for j in by_amt.get((vg, tot), []) if ok(j)]
        if tot and len(c) == 1: key2[i] = b2key[c[0]]; fb_used.add(c[0]); verdict[i] = V["amt_only"]; continue
        if (vg + zkey(r["invoice"])) in py_keys or (isinstance(d, dt.datetime) and (vg, d.date(), tot) in py_dates):
            verdict[i] = V["py"]; key2[i] = "PY:" + vg + norm(r["invoice"]); continue
        verdict[i] = V["not_in_2b"]
    for i, r in enumerate(reg):
        if S(r["category"]) == "RCM" and key2[i]:
            b = RCM_BASIS.get(verdict[i].split(" – ")[0])
            if b: verdict[i] = V["rcm_in2b"] % b
    return [{"verdict": verdict[i], "key2": key2[i]} for i in range(N)]
