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
