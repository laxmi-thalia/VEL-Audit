# VEL ITC changes — batch 3 (21-09-2026) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Apply Pawan's five ITC-side changes to `VEL_GST_Audit_FY2025-26_MASTER (2).xlsx`: precise remarks for malformed vendor GSTINs and recipient mismatches, rebuild the FY 24-25 2B base from the Octa PAN-level export, bring last year's permanent-reversal flags into the 2B sheets, and give `ITC Register 2026-27` the same 2B reconciliation format as `ITC Register 2025-26`.

**Architecture:** The matching rules move out of `cascade_fix.py` into a pure-Python module `vel/scripts/reco_lib.py` (unit-tested with synthetic rows) so the FY 25-26 cascade and the new FY 26-27 reco share one implementation. Every workbook write is a single Excel COM session (open → write → recalc → verify → save → SaveAs xlsb) that refuses to save when a golden moves unexpectedly. Sheet inputs are read by header name, never by remembered column letters; new columns are appended at the end of a sheet, never inserted in the middle.

**Tech Stack:** Python 3.12 (`C:\PROJECTS\accountic\backend\.venv`), openpyxl (read-only reads), pywin32 Excel COM (all writes), pyxlsb (last year's `.xlsb`), pytest.

**Spec:** this conversation (Pawan, 21-09-2026) + `vel/HANDOFF.md` + `.claude/skills/gst-audit-vel-itc/SKILL.md` (changes-logs 18-09 → 21-09 hold every ruling referenced below).

## Global Constraints

- Only write to `C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx`; re-export `…MASTER (2).xlsb` (COM `SaveAs FileFormat=50`) at the end of every task that saves.
- Before any write: `open(path,'r+b')` on both files; if `PermissionError`, stop and ask Pawan to close Excel without saving. Snapshot the master to the scratchpad (`master2_snapshot_before_<task>.xlsx`) before each write.
- No LLM calls. No client data in chat (counts, totals, document numbers are fine).
- Everything formula-based where a value can be derived in-sheet; values only for stamped verdicts.
- Goldens that must not move unless the task says so: `ITC Register 2025-26` row-4 `Total GST` = **1,069,969,542.15**; `ITC Summary` row 25 Net-ITC check `[0,0,0]`; 0 error cells workbook-wide (`SpecialCells(-4123,16)` on every sheet).
- Remark vocabulary (21-09): `Matched with 2B – <basis>` (+ ` – review` on non-exact), `Not in 2B – Apr-25 to Aug-26`, `Not applicable – RCM self-invoice / – ISD / – no vendor GSTIN (URD)`. Strings tested by other formulas and NOT to be changed: `Table 6A1 of GSTR-9 - Unclaimed`, `Table 6A1 of GSTR-9 - Claimed`, Table 13 `Claimed`/`Unclaimed`, `Correction Entries- ITC dated 24-25 reversed in 25-26`, `Permanent Reversals`, Countif `Consider`/`Not consider`.
- Never insert rows at the first data row of a sheet (ranges starting at `$6` shift); insert inside the range and delete the old rows. Never insert columns before existing ones on `ITC Register 2026-27` (ITC Summary references its columns by letter).
- Scripts live in `C:\PROJECTS\gst-audit-engine\vel\scripts\` and are copied to the session scratchpad to run; commit after every task; Pawan pushes (or `git push` if the remote accepts it).

---

## File Structure

- Create `vel/scripts/reco_lib.py` — pure functions: normalisation (`norm`, `zkey`, `inv_core`, `inv_gate`, `fy`, `fy_of_label`), `classify_vendor_gstin(gstin) -> ("valid"|"malformed"|"missing", pan)`, `match_register(reg_rows, b2_rows, py_index) -> list[Verdict]` implementing the four layers + malformed-GSTIN PAN rescue + recipient check. No Excel imports.
- Create `tests/vel/test_reco_lib.py` — pytest on synthetic rows.
- Modify `vel/scripts/cascade_fix.py` — replace passes 1/2a/2b/3 with `reco_lib.match_register`; keep the COM stamping / extract rebuild as is.
- Create `vel/scripts/rebuild_2b_itc_data.py` — `GSTR-2B ITC Data` from `PAN GSTR2B 2024-25.xlsx`.
- Create `vel/scripts/perm_reversals_ly.py` — stamp last year's `Permanent Reversals` flags into both 2B sheets.
- Create `vel/scripts/fy2627_reco.py` — the 25-26 reco format on `ITC Register 2026-27`.
- Modify `.claude/skills/gst-audit-vel-itc/SKILL.md` (both copies) and `vel/HANDOFF.md` — changes-log + state.

---

### Task 1: `reco_lib.py` — matching rules as a tested module

**Files:**
- Create: `vel/scripts/reco_lib.py`
- Create: `tests/vel/test_reco_lib.py`, `tests/vel/__init__.py` (empty)

**Interfaces:**
- Consumes: nothing (pure Python).
- Produces:
  - `norm(s: str) -> str` — upper, strip ` -/.'_` (identical to the sheets' KEY formula)
  - `zkey(s: str) -> str` — `norm` + leading zeros of digit runs dropped
  - `inv_core(s) -> str`, `inv_gate(a, b) -> bool`, `fy(d: datetime) -> str`, `fy_of_label(label, d) -> str`
  - `classify_vendor_gstin(g: str) -> tuple[str, str]` → `("valid", pan)` for `^\d{2}[A-Z0-9]{13}$`, `("malformed", pan_or_"")` for any other non-empty string (pan = chars 3-12 when len ≥ 12 else ""), `("missing", "")` for blank/`Missing`/`NA`
  - `Verdict = dict(verdict: str, key2: str, consider_key: str)`
  - `match_register(reg: list[dict], b2: list[dict], py_keys: set[str], py_dates: set[tuple]) -> list[Verdict]` — `reg` rows have keys `vendor_gstin, invoice, invoice_date, invoice_year, category, vel_gstin, igst, cgst, sgst`; `b2` rows have `supplier_gstin, doc_no, doc_date, company_gstin, key, igst, cgst, sgst`. Returns one Verdict per register row, in order.

- [ ] **Step 1: Write the failing tests**

```python
# tests/vel/test_reco_lib.py
import datetime as dt
from vel.scripts import reco_lib as R

def reg(**kw):
    base = dict(vendor_gstin="27AABCI4971Q1ZW", invoice="GZ/04", invoice_date=dt.datetime(2025, 5, 30), invoice_year="2025-26",
                category="ITC", vel_gstin="09AAECR0503Q1Z6", igst=1000.0, cgst=0.0, sgst=0.0)
    base.update(kw); return base

def b2(**kw):
    base = dict(supplier_gstin="27AABCI4971Q1ZW", doc_no="GZ/04", doc_date=dt.datetime(2025, 5, 30), company_gstin="09AAECR0503Q1Z6",
                igst=1000.0, cgst=0.0, sgst=0.0)
    base.update(kw); base["key"] = base["supplier_gstin"] + R.norm(base["doc_no"]); return base

def test_norm_and_zkey():
    assert R.norm("gz/04-24.25") == "GZ042425"
    assert R.zkey("SDIP/25-26/007") == R.zkey("SDIP/25-26/7")

def test_classify_vendor_gstin():
    assert R.classify_vendor_gstin("27AABCI4971Q1ZW") == ("valid", "AABCI4971Q")
    assert R.classify_vendor_gstin("27AABCI4971QZW") == ("malformed", "AABCI4971Q")      # 14 chars
    assert R.classify_vendor_gstin("27AAECM2933K1ZNB") == ("malformed", "AAECM2933K")    # 16 chars
    assert R.classify_vendor_gstin("Missing") == ("missing", "")
    assert R.classify_vendor_gstin("") == ("missing", "")

def test_exact_match_same_recipient():
    v = R.match_register([reg()], [b2()], set(), set())[0]
    assert v["verdict"] == "Matched with 2B – invoice no" and v["key2"] == "27AABCI4971Q1ZWGZ04"

def test_exact_match_recipient_differs():
    v = R.match_register([reg()], [b2(company_gstin="27AAECR0503Q1Z8")], set(), set())[0]
    assert v["verdict"] == "Matched with 2B – invoice no – recipient GSTIN differs (2B under 27AAECR0503Q1Z8) – review"
    assert v["key2"] == "27AABCI4971Q1ZWGZ04"

def test_malformed_gstin_rescued_by_pan_and_invoice():
    v = R.match_register([reg(vendor_gstin="27AABCI4971QZW")], [b2()], set(), set())[0]
    assert v["verdict"] == "Matched with 2B – invoice no, vendor GSTIN corrected from 2B (27AABCI4971Q1ZW) – review"
    assert v["key2"] == "27AABCI4971Q1ZWGZ04"

def test_malformed_gstin_not_rescued():
    v = R.match_register([reg(vendor_gstin="27AABCI4971QZW", invoice="XX/99")], [b2()], set(), set())[0]
    assert v["verdict"] == "Not matched – vendor GSTIN invalid (14 chars) – review" and v["key2"] == ""

def test_missing_gstin_by_category():
    assert R.match_register([reg(vendor_gstin="Missing", category="RCM")], [], set(), set())[0]["verdict"] == "Not applicable – RCM self-invoice"
    assert R.match_register([reg(vendor_gstin="Missing", category="ISD")], [], set(), set())[0]["verdict"] == "Not applicable – ISD"
    assert R.match_register([reg(vendor_gstin="", category="ITC")], [], set(), set())[0]["verdict"] == "Not applicable – no vendor GSTIN (URD)"

def test_document_amount_layer_single_candidate_only():
    rows = [reg(invoice="GZ/04/24-25", igst=600.0), reg(invoice="GZ/04/24-25", igst=400.0)]
    out = R.match_register(rows, [b2(doc_no="ABC-1")], set(), set())
    assert all(o["verdict"] == "Matched with 2B – GSTIN + amount (±100), invoice no differs – review" for o in out)
    out2 = R.match_register(rows, [b2(doc_no="ABC-1"), b2(doc_no="ABC-2")], set(), set())
    assert all(o["verdict"] == "Not in 2B – Apr-25 to Aug-26" for o in out2)

def test_invoice_similar_layer_prefers_similar_candidate():
    rows = [reg(invoice="3/GZ/03", igst=1000.0)]
    out = R.match_register(rows, [b2(doc_no="GZ/03"), b2(doc_no="ZZ/77")], set(), set())
    assert out[0]["verdict"] == "Matched with 2B – similar invoice no + amount (±100) – review" and out[0]["key2"].endswith("GZ03")

def test_prior_year_2b_layer():
    r = reg(invoice="PY/1", invoice_date=dt.datetime(2024, 11, 1), invoice_year="2024-25")
    v = R.match_register([r], [], {"27AABCI4971Q1ZW" + R.zkey("PY/1")}, set())[0]
    assert v["verdict"] == "Matched with 2B of FY 24-25 – Table 6A1" and v["key2"].startswith("PY:")
```

- [ ] **Step 2: Run tests to verify they fail**

Run (from `C:\PROJECTS\gst-audit-engine`): `C:\PROJECTS\accountic\backend\.venv\Scripts\python.exe -m pytest tests/vel/test_reco_lib.py -q`
Expected: `ImportError`/`ModuleNotFoundError: vel.scripts.reco_lib` (create empty `vel/__init__.py` and `vel/scripts/__init__.py` if the import path itself fails).

- [ ] **Step 3: Implement `reco_lib.py`**

```python
# vel/scripts/reco_lib.py
"""Matching rules shared by cascade_fix.py (FY 25-26) and fy2627_reco.py (FY 26-27). Pure Python, no Excel.
Layers (Pawan 18-09/21-09, conservative): 1 exact GSTIN + zero-insensitive invoice (recipient checked, flagged if it differs);
1b malformed vendor GSTIN rescued only by PAN + exact invoice (flagged); 2 GSTIN + document amount (±100, same recipient,
single candidate); 3 similar invoice + amount (±100, same recipient, single candidate); 2b line date+amount / amount single
candidate; 4 FY 24-25 2B (prior-year keys). Vocabulary is the 21-09 standard."""
import re, collections, datetime as dt
TOL_ABS = 100.0
GST = re.compile(r"^\d{2}[A-Z0-9]{13}$")
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v or 0)
    except Exception: return 0.0
def norm(s): return re.sub(r"[ \-/.'_]", "", S(s).upper())
def zkey(s): return re.sub(r"(?<!\d)0+(?=\d)", "", norm(s))
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
    if not s or s in ("MISSING", "NA", "NONE"): return ("missing", "")
    pan = s[2:12] if len(s) >= 12 else ""
    return ("valid", pan) if GST.match(s) else ("malformed", pan)
def _tot(r): return round(num(r["igst"]) + num(r["cgst"]) + num(r["sgst"]), 2)
def match_register(reg, b2, py_keys, py_dates):
    N = len(reg); verdict = [""] * N; key2 = [""] * N
    # 2B indices
    b2key = [S(r["key"]) for r in b2]; by_z = collections.defaultdict(list); by_dt = collections.defaultdict(list); by_amt = collections.defaultdict(list); by_pan = collections.defaultdict(list)
    for j, r in enumerate(b2):
        vg = S(r["supplier_gstin"]).upper()
        if not vg: continue
        zk = zkey(r["doc_no"]); by_z[vg + zk].append(j); by_pan[(vg[2:12], zk)].append(j)
        if isinstance(r["doc_date"], dt.datetime): by_dt[(vg, r["doc_date"].date(), _tot(r))].append(j)
        by_amt[(vg, _tot(r))].append(j)
    exact_owned = set()
    def rcpt_suffix(i, j):
        cg = S(b2[j]["company_gstin"]).upper(); vel = S(reg[i]["vel_gstin"]).upper()
        return (" – recipient GSTIN differs (2B under %s) – review" % cg) if cg and vel and cg != vel else ""
    # ---- pass 1: exact (+ malformed-GSTIN PAN rescue)
    for i, r in enumerate(reg):
        kind, pan = classify_vendor_gstin(r["vendor_gstin"]); vg = S(r["vendor_gstin"]).upper(); cat = S(r["category"])
        if kind == "missing":
            verdict[i] = "Not applicable – RCM self-invoice" if cat == "RCM" else ("Not applicable – ISD" if cat == "ISD" else "Not applicable – no vendor GSTIN (URD)"); continue
        if kind == "malformed":
            c = by_pan.get((pan, zkey(r["invoice"]))) if pan else None
            if c: j = c[0]; key2[i] = b2key[j]; exact_owned.add(j); verdict[i] = "Matched with 2B – invoice no, vendor GSTIN corrected from 2B (%s) – review" % S(b2[j]["supplier_gstin"]).upper()
            else: verdict[i] = "Not matched – vendor GSTIN invalid (%d chars) – review" % len(vg)
            continue
        c = by_z.get(vg + zkey(r["invoice"]))
        if c: j = c[0]; key2[i] = b2key[j]; exact_owned.add(j); verdict[i] = "Matched with 2B – invoice no" + rcpt_suffix(i, j)
    # ---- pass 2: document-level (same recipient, ±100, single candidate)
    fb_used = set(); b2doc = {}
    for j, r in enumerate(b2):
        vg = S(r["supplier_gstin"]).upper()
        if not vg or j in exact_owned: continue
        dk = (vg, zkey(r["doc_no"])); d = b2doc.setdefault(dk, {"tot": 0.0, "fy": fy(r["doc_date"]), "rows": [], "rcpt": S(r["company_gstin"]).upper(), "raw": S(r["doc_no"])}); d["tot"] += _tot(r); d["rows"].append(j)
    by_vendor = collections.defaultdict(list)
    for dk in b2doc: by_vendor[dk[0]].append(dk)
    regdoc = collections.OrderedDict()
    for i, r in enumerate(reg):
        if verdict[i] or S(r["category"]) != "ITC": continue
        vg = S(r["vendor_gstin"]).upper(); dk = (vg, zkey(r["invoice"])); d = regdoc.setdefault(dk, {"tot": 0.0, "fy": fy_of_label(r["invoice_year"], r["invoice_date"]), "lines": [], "rcpt": S(r["vel_gstin"]).upper(), "raw": S(r["invoice"])})
        d["tot"] += _tot(r); d["lines"].append(i)
    used = set()
    for (vg, inv), d in regdoc.items():
        if not d["tot"]: continue
        cands = [dk for dk in by_vendor.get(vg, []) if dk not in used and (not d["fy"] or b2doc[dk]["fy"] == d["fy"]) and b2doc[dk]["rcpt"] == d["rcpt"] and abs(b2doc[dk]["tot"] - d["tot"]) <= TOL_ABS]
        if not cands: continue
        sim = [dk for dk in cands if inv_similar(d["raw"], b2doc[dk]["raw"])]
        if len(sim) == 1: pick, why = sim[0], "Matched with 2B – similar invoice no + amount (±100) – review"
        elif len(cands) == 1: pick, why = cands[0], "Matched with 2B – GSTIN + amount (±100), invoice no differs – review"
        else: continue
        used.add(pick); j0 = b2doc[pick]["rows"][0]; fb_used.update(b2doc[pick]["rows"])
        for i in d["lines"]: key2[i] = b2key[j0]; verdict[i] = why
    # ---- pass 2b: line-level, single candidate, same FY; pass 3: prior-year 2B
    for i, r in enumerate(reg):
        if verdict[i]: continue
        vg = S(r["vendor_gstin"]).upper(); tot = _tot(r); d = r["invoice_date"]; ifY = fy_of_label(r["invoice_year"], d)
        ok = lambda j: j not in exact_owned and j not in fb_used and (not ifY or fy(b2[j]["doc_date"]) == ifY) and S(b2[j]["company_gstin"]).upper() == S(r["vel_gstin"]).upper()
        c = [j for j in (by_dt.get((vg, d.date(), tot), []) if isinstance(d, dt.datetime) else []) if ok(j)]
        if len(c) == 1: key2[i] = b2key[c[0]]; fb_used.add(c[0]); verdict[i] = "Matched with 2B – date + amount, invoice no differs – review"; continue
        c = [j for j in by_amt.get((vg, tot), []) if ok(j)]
        if tot and len(c) == 1: key2[i] = b2key[c[0]]; fb_used.add(c[0]); verdict[i] = "Matched with 2B – amount only, invoice no differs – review"; continue
        if (vg + zkey(r["invoice"])) in py_keys or (isinstance(d, dt.datetime) and (vg, d.date(), tot) in py_dates):
            verdict[i] = "Matched with 2B of FY 24-25 – Table 6A1"; key2[i] = "PY:" + vg + norm(r["invoice"]); continue
        verdict[i] = "Not in 2B – Apr-25 to Aug-26"
    return [{"verdict": verdict[i], "key2": key2[i]} for i in range(N)]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\PROJECTS\accountic\backend\.venv\Scripts\python.exe -m pytest tests/vel/test_reco_lib.py -q`
Expected: `10 passed`

- [ ] **Step 5: Commit**

```bash
cd C:\PROJECTS\gst-audit-engine && git add vel/__init__.py vel/scripts/__init__.py vel/scripts/reco_lib.py tests/vel && git commit -m "feat(vel-itc): reco_lib - shared, tested matching rules (malformed GSTIN rescue, recipient check)"
```

---

### Task 2: `cascade_fix.py` uses `reco_lib`

**Files:**
- Modify: `vel/scripts/cascade_fix.py` (passes 1 → 3, lines "# ---------------- pass 1" through the line before "# ---------------- Consider + labels")

**Interfaces:**
- Consumes: `reco_lib.match_register`, `reco_lib.zkey/norm/fy/fy_of_label/S/num`.
- Produces: `verdict[i]`, `key2[i]` lists exactly as before (downstream stamping, Countif labels, six/sixrem, extract rebuild unchanged). Verdict prefixes the downstream code tests: `Matched with 2B of FY 24-25`, `Not in 2B`, `Not applicable` — keep them.

- [ ] **Step 1: Replace the matching passes with a call into reco_lib**

Delete everything from `# ---------------- 2B indices` up to (not including) `# ---------------- Consider + labels` and put:

```python
import sys; sys.path.insert(0, r"C:\PROJECTS\gst-audit-engine")
from vel.scripts.reco_lib import match_register, zkey as _zk
reg_rows = [{"vendor_gstin": g(r, "Vendor GSTIN"), "invoice": g(r, "Invoice No."), "invoice_date": g(r, "Invoice Date"), "invoice_year": g(r, "Invoice Year"),
             "category": g(r, "Category"), "vel_gstin": g(r, "VEL GSTIN"), "igst": g(r, "IGST"), "cgst": g(r, "CGST"), "sgst": g(r, "SGST")} for r in rows]
b2_rows = [{"supplier_gstin": bg(r, "Supplier GSTIN"), "doc_no": bg(r, "Doc No"), "doc_date": bg(r, "Doc Date"), "company_gstin": bg(r, "Company GSTIN"),
            "key": S(bg(r, "Supplier GSTIN")).upper() + norm(bg(r, "Doc No")), "igst": bg(r, "IGST (Net)"), "cgst": bg(r, "CGST (Net)"), "sgst": bg(r, "SGST (Net)")} for r in B]
out = match_register(reg_rows, b2_rows, set(o_z), set(o_dt))
verdict = [o["verdict"] for o in out]; key2 = [o["key2"] for o in out]
b2key = [r["key"] for r in b2_rows]
print("verdicts:", dict(collections.Counter(v.split(" – ")[0] for v in verdict)))
```

(`o_z`/`o_dt` are the FY 24-25 2B dicts already built above from `GSTR-2B ITC Data`; `set(o_z)` gives the zero-insensitive keys, `set(o_dt)` the `(gstin, date, total)` tuples. `b2key` is still used by the extract rebuild.)

- [ ] **Step 2: Fix the label tests below to the new prefixes**

In the `# ---------------- Consider + labels` block keep `v.startswith("Matched with 2B of FY 24-25")`, `v.startswith("Not in 2B")`, `v.startswith("Not applicable")`; add: rows whose verdict starts with `"Not matched – vendor GSTIN invalid"` get label `"Not consider - vendor GSTIN invalid"` and `six = None`.

- [ ] **Step 3: Dry-run the matching only (no COM)**

Run: `python -X utf8 -c "exec(open('cascade_fix.py',encoding='utf-8').read().split('# ---------------- T6A1 extract rows')[0])"` from the scratchpad (master closed).
Expected: `verdicts:` line shows `Matched with 2B ≈ 35,700`, `Not in 2B ≈ 5,081`, `Not applicable ≈ 2,527 + 0 URD`, `Not matched – vendor GSTIN invalid ≤ 26` (some of the 26 rescued by PAN), and 5 verdicts containing `recipient GSTIN differs`. Anything else → stop and inspect.

- [ ] **Step 4: Commit**

```bash
git add vel/scripts/cascade_fix.py && git commit -m "refactor(vel-itc): cascade_fix matches through reco_lib (malformed GSTIN + recipient remarks)"
```

---

### Task 3: Rebuild `GSTR-2B ITC Data` from the Octa PAN-level FY 24-25 export

**Files:**
- Create: `vel/scripts/rebuild_2b_itc_data.py`

**Interfaces:**
- Consumes: `\\192.168.1.69\gst folder\…\VEL\Audit Data of FY 2024-25\PAN GSTR2B 2024-25.xlsx` — sheet `Inv+CDN(Document level)` (header row 3: `My GSTIN, 2B Return Period (MMYYYY text), Supplier GSTIN, Supplier Legal Name, Document Type, Section Name, Supply Type, Document Date (dd/mm/yyyy text), Document Number, Total Taxable Value, Total Tax Value, IGST Amount, CGST Amount, SGST Amount, CESS Amount, Total Document Value, State Place of Supply, Is Reverse Charge Appl…, GSTR-1 filing period, GSTR-1 filing date, Itc Availability, Reason for Non Account…`; row 1 holds column totals IGST/CGST/SGST/Cess) and sheet `ISD + ISDA` (header rows 1–2, 187 rows).
- Produces: `GSTR-2B ITC Data` with the SAME header row 5 as today (27 columns A..AA, KEY in AB — `cascade_fix.py` reads `FY (derived)`, `GSTIN of supplier`, `Invoice number`, `Invoice Date`, `Integrated Tax(₹)`, `Central Tax(₹)`, `State/UT Tax(₹)` by name; `t6a1_2b_period.py` asserts `2B Return Period`=D, `GSTIN of supplier`=G, `Invoice number`=I and uses the KEY column found after AA). Rows 6.. replaced; sheet stays hidden.

- [ ] **Step 1: Write the script**

```python
"""GSTR-2B ITC Data (FY 24-25 2B base) rebuilt from Octa's PAN-level export 'PAN GSTR2B 2024-25.xlsx' (Pawan 21-09) - replaces
the working-file tabs (48,802 rows across FY 20-21..26-27, mixed period formats). Same header row 5, rows 6.. replaced, KEY col
kept; CDN amounts negative; ties to the file's row-1 column totals before saving."""
import os, re, time, shutil, collections, datetime as dt, warnings, openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
SRC = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit Data of FY 2024-25/PAN GSTR2B 2024-25.xlsx"
for p in (P, B):
    try: open(p, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + p)
shutil.copy(P, "master2_snapshot_before_2bitcdata.xlsx")
E0 = dt.datetime(1899, 12, 30)
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(str(v).replace(",", "") or 0)
    except Exception: return 0.0
def ddate(s):
    s = S(s)
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try: return dt.datetime.strptime(s[:10], f)
        except Exception: pass
    return s if isinstance(s, dt.datetime) else None
def period(s):   # '032025' -> 2025-03-01
    s = S(s); return dt.datetime(int(s[2:6]), int(s[:2]), 1) if re.fullmatch(r"\d{6}", s) else None
def fy(d): return ("%d-%02d" % (d.year if d.month >= 4 else d.year - 1, (d.year + (1 if d.month >= 4 else 0)) % 100)) if isinstance(d, dt.datetime) else None
ST = {"01": "Jammu & Kashmir", "03": "Punjab", "06": "Haryana", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "12": "Arunachal Pradesh", "18": "Assam", "19": "West Bengal", "20": "Jharkhand", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat", "27": "Maharashtra", "29": "Karnataka", "32": "Kerala", "33": "Tamil Nadu", "36": "Telangana", "37": "Andhra Pradesh"}
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True); ws = wb["Inv+CDN(Document level)"]; rows = list(ws.iter_rows(values_only=True))
totals_row1 = [num(v) for v in rows[0] if v not in (None, "")]           # IGST, CGST, SGST, Cess column totals as exported
h = {S(v): j for j, v in enumerate(rows[2]) if S(v)}; body = [r for r in rows[3:] if r and S(r[h["My GSTIN"]])]
out = []
for r in body:
    g = lambda k: r[h[k]] if k in h else None
    typ = S(g("Document Type")).upper(); sign = -1.0 if "CREDIT" in typ else 1.0
    d = ddate(g("Document Date")); p = period(g("2B Return Period")); my = S(g("My GSTIN"))
    out.append([ST.get(my[:2], my[:2]), fy(d), "2024-25", p, None, "Current", S(g("Supplier GSTIN")).upper(), S(g("Supplier Legal Name")), S(g("Document Number")), typ.title(),
                d, sign * num(g("Total Document Value")), S(g("State Place of Supply")), S(g("Is Reverse Charge Applicable")) if "Is Reverse Charge Applicable" in h else S(next((g(k) for k in h if k.startswith("Is Reverse")), "")),
                None, sign * num(g("Total Taxable Value")), sign * num(g("IGST Amount")), sign * num(g("CGST Amount")), sign * num(g("SGST Amount")), sign * num(g("CESS Amount")),
                S(next((g(k) for k in h if k.startswith("GSTR-1/IFF/GSTR-5 Filing Period")), "")), S(g("Itc Availability")), S(next((g(k) for k in h if k.startswith("Reason")), "")), "PAN GSTR2B 2024-25 (Octa)", None, None, None])
isd = wb["ISD + ISDA"]; irows = list(isd.iter_rows(values_only=True)); ih = {S(v): j for j, v in enumerate(irows[0]) if S(v)}
for r in irows[2:]:
    if not r or not S(r[ih["My GSTIN"]]): continue
    my = S(r[ih["My GSTIN"]]); d = ddate(r[ih["ISD Document date"]]); p = period(S(r[ih["ISD GSTR-6 Period"]]))
    out.append([ST.get(my[:2], my[:2]), fy(d), "2024-25", p, None, "Current", S(r[ih["GSTIN of ISD"]]).upper(), S(r[ih["Legal Name of ISD"]]), S(r[ih["ISD Document number"]]), "ISD", d, None, "", "N", None, 0.0,
                num(r[ih["Tax amount"]]), num(r[ih["Tax amount"] + 1]), num(r[ih["Tax amount"] + 2]), num(r[ih["Tax amount"] + 3]), "", S(r[ih["Eligibility of ITC"]]), "", "PAN GSTR2B 2024-25 (Octa) - ISD", None, None, None])
wb.close()
igst, cgst, sgst = (sum(x[16] for x in out if x[9] != "ISD"), sum(x[17] for x in out if x[9] != "ISD"), sum(x[18] for x in out if x[9] != "ISD"))
print("documents %d (+ISD %d) | IGST %.2f CGST %.2f SGST %.2f | file row-1 totals %s" % (sum(1 for x in out if x[9] != "ISD"), sum(1 for x in out if x[9] == "ISD"), igst, cgst, sgst, totals_row1[:3]))
assert all(abs(a - b) < 1 for a, b in zip((igst, cgst, sgst), totals_row1[:3])), "column totals do not tie to the export - stop"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wbx = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wbx.Worksheets("GSTR-2B ITC Data")
    hdr = [sh.Cells(5, c).Value for c in range(1, 30)]; assert hdr[:27] == ['State folder', 'FY (derived)', 'F.Y', '2B Return Period', 'GSTR 3B Month', 'Curent previous', 'GSTIN of supplier', 'Trade/Legal name', 'Invoice number', 'Invoice type', 'Invoice Date', 'Invoice Value(₹)', 'Place of supply', 'Supply Attract Reverse Charge', 'Rate', 'Taxable Value (₹)', 'Integrated Tax(₹)', 'Central Tax(₹)', 'State/UT Tax(₹)', 'Cess(₹)', 'GSTR-1/5 Period', 'ITC Availability', 'Reason', 'Source', 'IRN', 'Claimed in FY 25-26 register?', 'Claim month (derived)'], hdr[:27]
    keycol = hdr.index("KEY (supplier GSTIN + invoice, normalised)") + 1
    old_n = sh.Cells(sh.Rows.Count, 1).End(-4162).Row - 5; K = len(out)
    sh.Rows("7:%d" % (7 + K - 2)).Insert(-4121)                                      # inside the range -> dependents follow
    vals = [[(v - E0).days if isinstance(v, dt.datetime) else v for v in row] for row in out]
    sh.Range(sh.Cells(6, 1), sh.Cells(5 + K, 27)).Value = vals
    if old_n > 1: sh.Rows("%d:%d" % (6 + K, 6 + K + old_n - 2)).Delete(-4162)
    for c in (4, 11): sh.Range(sh.Cells(6, c), sh.Cells(5 + K, c)).NumberFormat = "dd-mm-yyyy"
    sh.Range(sh.Cells(6, keycol), sh.Cells(5 + K, keycol)).Formula = '=UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($G6&$I6," ",""),"-",""),"/",""),".",""),"\'",""),"_",""))'
    sh.Cells(1, 2).Value = "FY 24-25 GSTR-2B: Octa PAN-level export 'PAN GSTR2B 2024-25.xlsx' (Inv+CDN document level %d rows + ISD %d), CDN negative. Rebuilt 21-09-2026." % (sum(1 for x in out if x[9] != "ISD"), sum(1 for x in out if x[9] == "ISD"))
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    RN = sh.Cells(sh.Rows.Count, 1).End(-4162).Row; print("rows 6..%d (was %d rows) | error cells %d | golden %.2f" % (RN, old_n, e, wbx.Worksheets("ITC Register 2025-26").Range("V4").Value))
    assert e == 0 and RN == 5 + K
    wbx.Save(); wbx.SaveAs(B, FileFormat=50); wbx.Close(False); print("saved + xlsb (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
```

(If the `Is Reverse Charge` / `GSTR-1 … Filing Period` / `Reason` header spellings differ, print `h` and adjust the three `startswith` prefixes — do not guess column letters.)

- [ ] **Step 2: Run it**

Run from the scratchpad: `python -X utf8 rebuild_2b_itc_data.py`
Expected: `documents 10323 (+ISD 187)`, the IGST/CGST/SGST line equal to the file's row-1 totals (49,19,627.35 / 11,22,193.14 / 11,22,193.14 — the assertion enforces it), `error cells 0`, golden `1069969542.15`, `saved + xlsb`.

- [ ] **Step 3: Commit**

```bash
git add vel/scripts/rebuild_2b_itc_data.py && git commit -m "feat(vel-itc): GSTR-2B ITC Data rebuilt from the Octa PAN-level FY 24-25 export"
```

---

### Task 4: Permanent-reversal flags from last year's 9C

**Files:**
- Create: `vel/scripts/perm_reversals_ly.py`

**Interfaces:**
- Consumes: `VEL_GSTR 9_9C FY 24-25.xlsb` sheet `GSTR-2B Apr 24-Oct 25` (header row 4: `My GSTIN`, `Supplier GSTIN`, `Document Number`, `Permanent Reversals`, `Reclaim - Table 6H`), read with pyxlsb from the local copy `scratchpad\VEL_2425.xlsb` (copy the file if absent; COM/pyxlsb cannot read UNC reliably).
- Produces: `GSTR-2B Apr25-Aug26` column `Permanent Reversals` (BG) = `"Permanent Reversals"` on rows whose (Company GSTIN, Supplier GSTIN, zkey(Doc No)) is flagged last year; new column `Permanent Reversals (LY 9C)` appended after `KEY` on `GSTR-2B ITC Data` with the same flag. ITC Summary CE:CG (Table 8C block) sums BG — the delta is reported, and it is the expected effect of this task.

- [ ] **Step 1: Write the script**

```python
"""Stamp last year's 'Permanent Reversals' flags (VEL_GSTR 9_9C FY 24-25.xlsb, sheet 'GSTR-2B Apr 24-Oct 25') into
'GSTR-2B Apr25-Aug26' col 'Permanent Reversals' and into 'GSTR-2B ITC Data' (new col after KEY). Key = my GSTIN + supplier
GSTIN + zero-insensitive doc no. Values only (the flag is a CA decision from last year); ITC Summary 8C block delta reported."""
import os, re, time, shutil, collections, win32com.client as win32, pythoncom
from pyxlsb import open_workbook
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for p in (P, B):
    try: open(p, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + p)
shutil.copy(P, "master2_snapshot_before_permrev.xlsx")
def S(v): return "" if v is None else str(v).strip()
def zkey(s): return re.sub(r"(?<!\d)0+(?=\d)", "", re.sub(r"[ \-/.'_]", "", S(s).upper()))
flags = set(); recl = set()
with open_workbook("VEL_2425.xlsb") as w:
    with w.get_sheet("GSTR-2B Apr 24-Oct 25") as sh:
        rows = [[c.v for c in r] for r in sh.rows()]
h = {S(x): j for j, x in enumerate(rows[3]) if S(x)}
for r in rows[4:]:
    if not r or not S(r[h["My GSTIN"]]): continue
    k = (S(r[h["My GSTIN"]]).upper(), S(r[h["Supplier GSTIN"]]).upper(), zkey(r[h["Document Number"]]))
    if S(r[h["Permanent Reversals"]]): flags.add(k)
    if S(r[h["Reclaim - Table 6H"]]): recl.add(k)
print("LY flagged documents: permanent reversals %d | reclaim 6H %d" % (len(flags), len(recl)))
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    isum = wb.Worksheets("ITC Summary"); before = [round(isum.Cells(25, c).Value or 0, 2) for c in (83, 84, 85)]
    b2 = wb.Worksheets("GSTR-2B Apr25-Aug26"); BH = {b2.Cells(2, c).Value: c for c in range(1, 70) if b2.Cells(2, c).Value}; NB = b2.Cells(b2.Rows.Count, 1).End(-4162).Row
    col = lambda h_: [v[0] for v in b2.Range(b2.Cells(3, BH[h_]), b2.Cells(NB, BH[h_])).Value]
    my, sup, dn = col("Company GSTIN"), col("Supplier GSTIN"), col("Doc No"); n = 0; out = []
    for a, b, c in zip(my, sup, dn):
        k = (S(a).upper(), S(b).upper(), zkey(c)); hit = k in flags; n += hit; out.append(["Permanent Reversals" if hit else None])
    b2.Range(b2.Cells(3, BH["Permanent Reversals"]), b2.Cells(NB, BH["Permanent Reversals"])).Value = out
    o = wb.Worksheets("GSTR-2B ITC Data"); OH = {o.Cells(5, c).Value: c for c in range(1, 40) if o.Cells(5, c).Value}; ON = o.Cells(o.Rows.Count, 1).End(-4162).Row
    kc = max(OH.values()) + 1
    if o.Cells(5, kc - 1).Value != "Permanent Reversals (LY 9C)":
        x = o.Cells(5, kc); x.Value = "Permanent Reversals (LY 9C)"; x.Font.Bold = True; x.Font.Color = 0xFFFFFF; x.Interior.Color = 0xB09784
    else: kc -= 1
    # GSTR-2B ITC Data has no My GSTIN column: match on (supplier GSTIN, zkey(invoice)) against the LY flags projected the same way
    flags2 = {(k[1], k[2]) for k in flags}
    sg = [v[0] for v in o.Range(o.Cells(6, OH["GSTIN of supplier"]), o.Cells(ON, OH["GSTIN of supplier"])).Value]; iv = [v[0] for v in o.Range(o.Cells(6, OH["Invoice number"]), o.Cells(ON, OH["Invoice number"])).Value]
    out2 = [["Permanent Reversals" if (S(a).upper(), zkey(b)) in flags2 else None] for a, b in zip(sg, iv)]; n2 = sum(1 for x in out2 if x[0])
    o.Range(o.Cells(6, kc), o.Cells(ON, kc)).Value = out2
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    after = [round(isum.Cells(25, c).Value or 0, 2) for c in (83, 84, 85)]; e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("stamped: 2B Apr25-Aug26 %d rows | GSTR-2B ITC Data %d rows | ITC Summary CE:CG before %s after %s | error cells %d" % (n, n2, before, after, e))
    assert e == 0
    wb.Save(); wb.SaveAs(B, FileFormat=50); wb.Close(False); print("saved + xlsb (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
```

- [ ] **Step 2: Run it** — Expected: `LY flagged documents: permanent reversals ≥ 468`, stamped counts > 0 on both sheets, `error cells 0`, CE:CG delta printed (report it to Pawan — it is the intended change).

- [ ] **Step 3: Commit**

```bash
git add vel/scripts/perm_reversals_ly.py && git commit -m "feat(vel-itc): last year's permanent-reversal flags stamped into both 2B sheets"
```

---

### Task 5: Re-run the FY 25-26 chain and verify

**Files:** none new — run `cascade_fix.py` → `final_countif_rule.py` → `t6a1_2b_period.py` from the scratchpad (master closed).

- [ ] **Step 1: Snapshot and run the chain, capturing the log**

```bash
cp "/c/Users/pawar/Downloads/VEL_GST_Audit_FY2025-26_MASTER (2).xlsx" master2_snapshot_before_chain2.xlsx
PY=/c/PROJECTS/accountic/backend/.venv/Scripts/python.exe
{ $PY -X utf8 cascade_fix.py; $PY -X utf8 final_countif_rule.py; $PY -X utf8 t6a1_2b_period.py; } > chain2.log 2>&1; grep -v SyntaxWarning chain2.log
```

- [ ] **Step 2: Check the log against these expectations**

- `verdicts:` — `Not applicable` = RCM 2,527 + ISD 171 only (URD 0); `Not matched – vendor GSTIN invalid` ≤ 26; `Matched with 2B of FY 24-25` may change from 932 (new FY 24-25 base) — report the new number.
- `error cells: 0`, `Total GST golden 1069969542.15`, `Net-ITC diff [-0.0, 0.0, 0.0]`, `B_Total (all ITC docs) 937528722.19`.
- `t6a1`: `2B Return Period` distribution printed; `Not in 2B` should drop versus 1,773 (report).
- Any assertion or `Traceback` → restore `master2_snapshot_before_chain2.xlsx` and stop.

- [ ] **Step 3: Read the five recipient-mismatch rows and the ≤26 malformed rows back (COM read-only) and paste their verdicts into the report (document numbers are fine, no vendor names).**

- [ ] **Step 4: Commit the scratchpad scripts copied to the repo**

```bash
cp cascade_fix.py final_countif_rule.py t6a1_2b_period.py /c/PROJECTS/gst-audit-engine/vel/scripts/ && cd /c/PROJECTS/gst-audit-engine && git add -A && git commit -m "chore(vel-itc): chain re-run after 2B base rebuild"
```

---

### Task 6: `ITC Register 2026-27` — 2B reconciliation in the 25-26 format

**Files:**
- Create: `vel/scripts/fy2627_reco.py`

**Interfaces:**
- Consumes: `reco_lib.match_register` (b2 = `GSTR-2B Apr25-Aug26` rows, all periods; `py_keys`/`py_dates` empty — FY 25-26 invoices never sit in the FY 24-25 2B); register = `ITC Register 2026-27` rows 5.. (header row 4; columns `Vendor GSTIN`, `Invoice No.`, `Invoice Date`, `Invoice Year`, `Category`, `VEL GSTN`, `IGST`, `CGST`, `SGST`, `Vendor Name/RCM Category`).
- Produces: columns APPENDED after the last header (`Source`) in this order and with these exact headers: `KEY`, `Countif`, `B_IGST`, `B_CGST`, `B_SGST`, `B_Total GST`, `KEY2 (matched 2B key)`, `2B_IGST`, `2B_CGST`, `2B_SGST`, `2B_Total GST`, `D_IGST`, `D_CGST`, `D_SGST`, `D_Total GST`, `Reco Remarks` — same semantics as the 25-26 register: `KEY` formula (vendor GSTIN & normalised invoice), `Countif` value `Consider` on the first line of each (vendor GSTIN or vendor name + normalised invoice) / `Not consider` others / blank for RCM-ISD, `B_` = SUMIFS by own KEY + vendor name on Consider lines, `KEY2` value, `2B_` = SUMIFS of `GSTR-2B Apr25-Aug26` V/W/X by KEY2 (0 when blank), `D_ = B_ − 2B_`, `Reco Remarks` value (21-09 vocabulary). Existing `Available in 2B` / `2B Inv` / `2B Period` / `Final Remarks` are left in place.

- [ ] **Step 1: Write the script**

```python
"""ITC Register 2026-27: 2B reconciliation in the FY 25-26 format (Pawan 21-09) - KEY / Countif / B_ / KEY2 / 2B_ / D_ / Reco Remarks
appended after 'Source' (ITC Summary references this sheet by column letter - never insert in the middle). Matching via reco_lib
against 'GSTR-2B Apr25-Aug26' (all periods). Single COM session; B_ must equal the sheet's ITC tax before saving."""
import re, sys, os, time, shutil, collections, datetime as dt, warnings, openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
sys.path.insert(0, r"C:\PROJECTS\gst-audit-engine"); from vel.scripts.reco_lib import match_register, norm, S, num
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for p in (P, B):
    try: open(p, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + p)
shutil.copy(P, "master2_snapshot_before_fy2627reco.xlsx")
wb = openpyxl.load_workbook(P, read_only=True)
n = wb["ITC Register 2026-27"]; hd = [c.value for c in next(n.iter_rows(min_row=4, max_row=4))]; H = {h: i for i, h in enumerate(hd) if h}
rows = [r for r in n.iter_rows(min_row=5, values_only=True) if r[3]]; g = lambda r, h: r[H[h]]
b2 = wb["GSTR-2B Apr25-Aug26"]; BH = {c.value: i for i, c in enumerate(next(b2.iter_rows(min_row=2, max_row=2))) if c.value}
B2 = [r for r in b2.iter_rows(min_row=3, values_only=True) if r[0]]; NB = 2 + len(B2); wb.close()
reg = [{"vendor_gstin": g(r, "Vendor GSTIN"), "invoice": g(r, "Invoice No."), "invoice_date": g(r, "Invoice Date"), "invoice_year": g(r, "Invoice Year"), "category": g(r, "Category"),
        "vel_gstin": g(r, "VEL GSTN"), "igst": g(r, "IGST"), "cgst": g(r, "CGST"), "sgst": g(r, "SGST")} for r in rows]
b2r = [{"supplier_gstin": r[BH["Supplier GSTIN"]], "doc_no": r[BH["Doc No"]], "doc_date": r[BH["Doc Date"]], "company_gstin": r[BH["Company GSTIN"]], "key": S(r[BH["Supplier GSTIN"]]).upper() + norm(r[BH["Doc No"]]),
        "igst": r[BH["IGST (Net)"]], "cgst": r[BH["CGST (Net)"]], "sgst": r[BH["SGST (Net)"]]} for r in B2]
out = match_register(reg, b2r, set(), set())
GST = re.compile(r"^\d{2}[A-Z0-9]{13}$"); seen = set(); labels = []
for r in rows:
    if g(r, "Category") != "ITC": labels.append(None); continue
    vg = S(g(r, "Vendor GSTIN")).upper(); k = (vg if GST.match(vg) else "NM:" + S(g(r, "Vendor Name/RCM Category")).upper()) + "|" + norm(g(r, "Invoice No."))
    labels.append("Not consider" if k in seen else "Consider"); seen.add(k)
print("rows %d | verdicts %s | Countif %s" % (len(rows), dict(collections.Counter(o["verdict"].split(" – ")[0] for o in out)), dict(collections.Counter(labels))))
NEW = ["KEY", "Countif", "B_IGST", "B_CGST", "B_SGST", "B_Total GST", "KEY2 (matched 2B key)", "2B_IGST", "2B_CGST", "2B_SGST", "2B_Total GST", "D_IGST", "D_CGST", "D_SGST", "D_Total GST", "Reco Remarks"]
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wbx = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wbx.Worksheets("ITC Register 2026-27")
    HX = {sh.Cells(4, c).Value: c for c in range(1, 90) if sh.Cells(4, c).Value}; RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row; assert RN - 4 == len(rows)
    start = max(HX.values()) + 1
    for i, h in enumerate(NEW):
        if h in HX: raise SystemExit("column already exists: " + h)
        x = sh.Cells(4, start + i); x.Value = h; x.Font.Bold = True; x.Font.Color = 0xFFFFFF; x.Interior.Color = 0x4F3F33 if not h.startswith(("B_", "2B_", "D_")) else 0xB09784
    C = {h: L(start + i) for i, h in enumerate(NEW)}; c = lambda h: L(HX[h])
    rg = lambda h: sh.Range("%s5:%s%d" % (C[h], C[h], RN))
    rg("KEY").Formula = '=UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($%s5&$%s5," ",""),"-",""),"/",""),".",""),"\'",""),"_",""))' % (c("Vendor GSTIN"), c("Invoice No."))
    rg("Countif").Value = [[v] for v in labels]; rg("KEY2 (matched 2B key)").Value = [[o["key2"] or None] for o in out]; rg("Reco Remarks").Value = [[o["verdict"]] for o in out]
    KEY, VN, K2, CF = C["KEY"], c("Vendor Name/RCM Category"), C["KEY2 (matched 2B key)"], C["Countif"]
    for b, src in (("B_IGST", "IGST"), ("B_CGST", "CGST"), ("B_SGST", "SGST")):
        rg(b).Formula = '=IF($%s5="Consider",SUMIFS($%s$5:$%s$%d,$%s$5:$%s$%d,$%s5,$%s$5:$%s$%d,$%s5),"NA")' % (CF, c(src), c(src), RN, KEY, KEY, RN, KEY, VN, VN, RN, VN)
    rg("B_Total GST").Formula = '=IF($%s5="Consider",%s5+%s5+%s5,"NA")' % (CF, C["B_IGST"], C["B_CGST"], C["B_SGST"])
    for t, col in (("2B_IGST", "V"), ("2B_CGST", "W"), ("2B_SGST", "X")):
        rg(t).Formula = '=IF($%s5="Consider",IF($%s5="",0,SUMIFS(\'GSTR-2B Apr25-Aug26\'!$%s$3:$%s$%d,\'GSTR-2B Apr25-Aug26\'!$AW$3:$AW$%d,$%s5)),"NA")' % (CF, K2, col, col, NB, NB, K2)
    rg("2B_Total GST").Formula = '=IF($%s5="Consider",%s5+%s5+%s5,"NA")' % (CF, C["2B_IGST"], C["2B_CGST"], C["2B_SGST"])
    for d, b, t in (("D_IGST", "B_IGST", "2B_IGST"), ("D_CGST", "B_CGST", "2B_CGST"), ("D_SGST", "B_SGST", "2B_SGST"), ("D_Total GST", "B_Total GST", "2B_Total GST")):
        rg(d).Formula = '=IF($%s5="Consider",%s5-%s5,"NA")' % (CF, C[b], C[t])
    for h in ("B_IGST", "B_CGST", "B_SGST", "B_Total GST", "2B_IGST", "2B_CGST", "2B_SGST", "2B_Total GST", "D_IGST", "D_CGST", "D_SGST", "D_Total GST"): rg(h).NumberFormat = "#,##0.00"
    sh.Columns(C["Reco Remarks"]).ColumnWidth = 60; sh.Range(sh.Cells(4, 1), sh.Cells(RN, start + len(NEW) - 1)).AutoFilter()
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    cs = lambda h: sum(v[0] for v in rg(h).Value if isinstance(v[0], (int, float)))
    itc_tax = sum(num(g(r, "IGST")) + num(g(r, "CGST")) + num(g(r, "SGST")) for r in rows if g(r, "Category") == "ITC")
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("B_Total %.2f vs ITC tax %.2f | 2B_Total %.2f | D_Total %.2f | error cells %d | FY 25-26 golden %.2f" % (cs("B_Total GST"), itc_tax, cs("2B_Total GST"), cs("D_Total GST"), e, wbx.Worksheets("ITC Register 2025-26").Range("V4").Value))
    assert e == 0 and abs(cs("B_Total GST") - itc_tax) < 1
    wbx.Save(); wbx.SaveAs(B, FileFormat=50); wbx.Close(False); print("saved + xlsb (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
```

- [ ] **Step 2: Run it** — Expected: `rows 2109`, verdict counts, `Countif {Consider ≈ 2000, Not consider ≈ 100}`, `B_Total == ITC tax 24,280,458.10` (assertion), `error cells 0`, FY 25-26 golden unchanged, `saved + xlsb`.

- [ ] **Step 3: Commit**

```bash
git add vel/scripts/fy2627_reco.py && git commit -m "feat(vel-itc): ITC Register 2026-27 reconciled to 2B in the FY 25-26 format"
```

---

### Task 7: Documentation, handoff, push

**Files:**
- Modify: `.claude/skills/gst-audit-vel-itc/SKILL.md` (repo copy and `C:\PROJECTS\accountic\.claude\skills\gst-audit-vel-itc\SKILL.md`), `vel/HANDOFF.md`

- [ ] **Step 1: Append a dated changes-log entry (both skill copies)** covering: reco_lib as the single matching implementation; malformed-GSTIN vocabulary (`Not matched – vendor GSTIN invalid (n chars) – review`, PAN rescue text); recipient-differs suffix; `GSTR-2B ITC Data` = Octa PAN-level FY 24-25 export (row-1 totals as the tie-out); permanent-reversal flags from the LY 9C (with the CE:CG delta); FY 26-27 reco columns appended after `Source`; the numbers printed by Tasks 5 and 6.

- [ ] **Step 2: Update `vel/HANDOFF.md` "Verified state" and "Open items"** (remove the resolved items, keep ZFI06 / IMS columns / missing Octa states).

- [ ] **Step 3: Commit and push**

```bash
git add -A && git commit -m "docs(vel): batch-3 rulings and handoff state" && GIT_TERMINAL_PROMPT=0 git push
```

---

### Task 8: T6A1 `2B Return Period` lookup keyed on GSTIN + invoice + invoice FY (Pawan's `VEL 21.9.26.docx`)

**Files:**
- Modify: `vel/scripts/t6a1_2b_period.py`

**Problem (from the docx screenshots):** supplier `09DCEPK6815A2ZS` invoice `3` exists twice in the FY 24-25 2B base — once dated 20.07.2023 (2B Oct-23) and once dated 03.06.2025 (2B Apr-25). The extract row (a FY 24-25 document) looked up `GSTIN & invoice` and took the first hit, Oct-23. Invoice numbers restart every year, so the key must carry the invoice FY.

**Interfaces:**
- Consumes: `GSTR-2B ITC Data` col `FY (derived)` (B) and its helper KEY column; `GSTR-2B Apr25-Aug26` col `Doc FY (doc date)` (AV) and `KEY` (AW); extract cols `GSTIN of supplier` (J), `Invoice Date` (L), `Invoice number` (M).
- Produces: helper `KEY+FY` columns on both 2B sheets (`GSTR-2B ITC Data`: rewrite the existing helper to `=<KEY>&"|"&$B6`; `GSTR-2B Apr25-Aug26`: new column after the last header `=$AW3&"|"&$AV3`) and the extract formula in C using key `=<norm(J&M)>&"|"&<FY of L>` where FY of L = `IF(MONTH($L5)>=4,YEAR($L5)&"-"&RIGHT(YEAR($L5)+1,2),YEAR($L5)-1&"-"&RIGHT(YEAR($L5),2))`.

- [ ] **Step 1: Change the helper KEY formula on `GSTR-2B ITC Data`** (in `t6a1_2b_period.py`, the line `o.Range(... KL ...).Formula = "=" + norm("$G6&$I6")`) to `"=" + norm("$G6&$I6") + '&"|"&$B6'`; rename the header to `KEY (supplier GSTIN + invoice + FY, normalised)`.

- [ ] **Step 2: Add the CY helper on `GSTR-2B Apr25-Aug26`**: find the first empty header cell on row 2 after the last DPS column, write `KEY+FY`, fill rows 3..NB with `=$AW3&"|"&$AV3`; remember its letter as `CYK`.

- [ ] **Step 3: Rewrite the extract lookup formula** `f` so that `key` = `norm("$J5&$M5") & '&"|"&IF(MONTH($L5)>=4,YEAR($L5)&"-"&RIGHT(YEAR($L5)+1,2),YEAR($L5)-1&"-"&RIGHT(YEAR($L5),2))'` and the CY `MATCH` runs against `'GSTR-2B Apr25-Aug26'!$<CYK>$3:$<CYK>$NB` instead of `$AW`.

- [ ] **Step 4: Run the script** (after Task 3, master closed). Expected: the docx row (Bihar, supplier `09DCEPK6815A2ZS`, invoice `3`, FY 24-25) now shows a FY 24-25 2B period or `Not in 2B (Apr-24 to Aug-26)` — never Oct-23; distribution printed; `error cells 0`.

- [ ] **Step 5: Commit**

```bash
git add vel/scripts/t6a1_2b_period.py && git commit -m "fix(vel-itc): T6A1 2B Return Period lookup keyed on GSTIN + invoice + invoice FY"
```

---

## Self-review

- Spec coverage: item 1 (URD confusion) → Tasks 1, 2, 5; item 2 (ITCR 26-27 2B reco) → Task 6; item 3 (recipient mismatch remarks) → Tasks 1, 2, 5; item 4 (permanent reversals from LY 9C) → Task 4; item 5 (GSTR-2B ITC Data) → Task 3 (+5 for the knock-on). Pawan's "2B path of prev year" folder is the source of Task 3.
- Names used across tasks: `match_register`, `norm`, `zkey`, `classify_vendor_gstin`, `S`, `num` defined in Task 1 and imported in Tasks 2 and 6; new 26-27 headers listed once in Task 6 `NEW`; `Permanent Reversals` column header on the 2B sheet is the existing BG header.
- Order of execution matters: 3 → 4 → (1, 2) → 5 → 6 → 7. Tasks 1–2 can be built before 3–4 but the chain (5) must run after 3–4.
