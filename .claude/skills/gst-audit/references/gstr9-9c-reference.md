# GSTR 9 / 9C — distilled working reference

Source: Bharat's *Analysis of GST Returns* (OCR), ch. 7 (GSTR 9) & ch. 8
(GSTR 9C), read 2026-07-27. Engine-mapping notes in [brackets].

## GSTR 9 — table map & cross-checks

### Part II — outward (Tables 4, 5)
- **4A-4E**: B2C/B2B/exports-with-pay/SEZ-with-pay/deemed, net-of CDN &
  amendments rows 4I-4L. **4F** advances, **4G** inward RCM liability,
  **4G1** (FY 24-25+) ECO 9(5) supplies where ECO pays. **4N** total.
- **Instruction 2A**: rows of Table 4 must contain ONLY current-FY liability —
  prior-FY liability reported in current-year 3B/GSTR-1 is REMOVED from Table
  4 (it belongs to last year's Tables 10/11). Reduce Table 9 tax-payable
  correspondingly; the payable-vs-paid gap that creates should equal LAST
  year's Table 10, else it is short-payment → DRC-03 ("Annual Return",
  **cash only** per instruction 9).
- **Table 5**: exempt (5D)/nil/non-GST/no-supply + LUT exports (5A? no — 5A
  is export w/o pay), SEZ w/o pay, RCM-outward (5C), net of amendments.
  5D+5E+5F drives the Rule 42/43 exempt-turnover cross-check:
  `Table 6I ITC × exempt(5C+5D+5E) / total turnover (5N+10-11)` ≈ annual
  Rule 42/43 reversal in 7C+7D (tentative only). [engine: FINDING-level
  check candidate]
- **5N + 10 − 11 = Total turnover** (F206 in the working ✓).

### Part III — ITC (Tables 6, 7, 8)
- **6A** auto = Σ Table 4A of 3B for the FY.
- **6A1 (FY 24-25+)**: ITC of the PRECEDING FY availed in current-FY 3Bs
  (Apr–Oct, filed by 30 Nov) that sits inside 6A — carved out here.
  **6A2 = 6A − 6A1** feeds the current-year splits. Routing rules:
  - reclaim of a **Rule 37/37A** reversal from a prior FY → **6H**, not 6A1;
  - reclaim of a prior-FY reversal **other than** 37/37A → **6A1**, not 6H.
  [engine: side-working M/N..R cols; DPS spill-in sheet + tags drive it]
- **6B** inward supplies (excl imports/RCM, incl SEZ services) — FIRST-TIME
  availment only (FY 24-25+ tri-split: availed→6B, reversed→7, reclaimed→6H).
  ↔ 3B 4A5. **6C** RCM-unregistered / **6D** RCM-registered (both s.9(3);
  9(4) in 6C) ↔ 3B 4A3. **6E** import goods (incl SEZ goods) ↔ 4A1.
  **6F** import services ↔ 4A2. **6G** ISD ↔ 4A4. **6H** reclaims ↔ 4A5/4D1.
  **6M** ITC-01/02/02A. 6I subtotal, 6J = 6I − 6A2 (ideally 0).
- **Table 7**: 7A Rule 37 (180-day, temporary, reclaim exempt from 16(4)),
  **7A1 Rule 37A** (supplier's 3B unfiled by 30 Sep → recipient reverses by
  30 Nov, interest-free if by then; reclaim when supplier pays), 7A2 Rule 38
  (banking 50%), 7B Rule 39 (ISD excess), 7C Rule 42 / 7D Rule 43 (common
  credit), 7E s.17(5) blocked, 7H other. 7A-7E ↔ 3B 4B1/4B2 (post-Circular
  170/2022: 17(5) goes through 4B1; 4D1 is disclosure-only). **7J = 6O − 7**
  is what flows to 9C 12E.
- **Table 8**: 8A auto from **2B Table 3(I)** (2A until FY 22-23); 8B =
  6B+6H auto; **8C** = current-FY invoices availed NEXT FY (forward charge)
  — and per the book, **8C should also equal Table 13 − Table 12 (to the
  forward-charge extent)**. 8D = A−(B+C): if **negative**, check reversals
  (8D doesn't see Table 7); if positive, split 8E (not availed) / 8F
  (ineligible). **8K lapse figure creates no payment obligation.**

### Part IV/V (Tables 9-13) & others
- **Table 9** payable vs paid; paid auto from 3B 6.1 (from Sep-24: negative-
  liability netting columns — only positive net flows to payable). Payable
  should equal Table 4 after the 2A adjustment.
- **Tables 10/11**: current-FY supplies amended/declared (10, +) or reduced
  (11, −) in next-FY returns Apr–Nov.
- **Tables 12/13**: current-FY ITC reversed (12) / availed (13) in next-FY
  returns to 30 Nov. Full-year ITC = 6 − 7 + 13 − 12, but 9C receives only
  7J — 12/13 therefore surface inside 9C **Table 12B/12C**, not 12E.
  A Rule 37/37A reversal of THIS year reclaimed NEXT year belongs in NEXT
  year's annual return (6H), not this year's 13.
- **Table 15**: refunds claimed/sanctioned/rejected/pending + demands —
  check RFD-01s, **shipping-bill IGST scroll**, DRC-07s, and Part II of the
  Electronic Liability Register.
- **Table 16**: 16A purchases from composition dealers (3B Table 5 source);
  16B s.143 job-work deemed supply; 16C goods sent on approval >180 days.
  Disclosure-only per author (liability, if any, routes via Table 4/10).
- **Table 17 HSN outward**: 6-digit mandatory (turnover > 5 Cr), 4-digit B2B
  otherwise; qty net of returns; portal excel from GSTR-1 Table 12 available
  FY 24-25+. **Table 18** inward HSN: only items ≥10% of inward value,
  optional. **Table 19** late fee.

## GSTR 9C — table map

- **Table 5**: 5A audited-FS turnover (per-GSTIN, not PAN) + adjustments
  5B-5O (unbilled revenue, unadjusted advances, deemed supply, post-FY CNs,
  s.15 valuation e.g. excise-in-value, **5N forex** rule-34 vs books, 5O
  residual) → **5P**; **5Q = GSTR-9 (5N + 10 − 11)**; 5R = Q − P; Table 6
  reasons. 5B-5N optional FY 19-20 through 24-25 (5O carries the net).
- **Table 7**: 7A = 5P; 7B exempt/nil/non-GST/no-supply; 7C zero-rated
  WITHOUT pay (LUT); 7D outward-RCM supplies; **7D1** (FY 24-25+) ECO 9(5);
  7E = A−B−C−D taxable; **7F auto = 4N − 4G + 10 − 11**; 7G ideally NIL.
- **Table 9**: rate-wise (9A..9K1 forward; 9B/9D/9F/9H/9H2 RCM-inward;
  **9K2** ECO 9(5), no taxable value) + 9L-9O interest/late-fee/penalty/
  others. 9P = ΣA..O; **9Q auto = Table 9 payable + 10 − 11 of GSTR 9,
  consolidated** — since 9 of GSTR-9 lumps interest with tax but 9Q wants
  head-wise, an interest-sized 9R difference is EXPECTED and explained in
  Table 10. [standard Table-10 reason — don't chase it as an error]
- **Table 11**: additional liability from 6/8/10 → DRC-03 ("Reconciliation
  Statement"); FY 24-25+ payable through cash **or credit** ledger.
- **Table 12**: 12A books ITC (GSTIN-wise); **12B** ITC booked earlier FY,
  claimed this FY (+); **12C** ITC booked this FY, claimed next FY (−) —
  equals GSTR-9 Table 13 − Table 12; 12D = A+B−C; 12E auto = 7J; 12F
  unreconciled; Table 13 reasons.
  - **Author's caveat (instr. 2A era)**: 12E and 12A both exclude prior-FY
    ITC, so a populated 12B can double-count — the book argues 12B should be
    ~nil post-2A. **DPS practice** (golden FY 24-25): 12B = prior-FY-booked
    ITC availed this year, with the format-driven difference explained as a
    Table 13 reason. Keep DPS practice, keep the reason narrative — but know
    the argument if a reviewer pushes back.
- **Table 14/15/16**: expense-head-wise eligible-ITC reco (optional through
  FY 24-25; ≈ 3CD clause 44). Part V auditor recommendation. **Table 17
  late fee** (Circular 246/03/2025): GSTR-9C filed late ⇒ s.47(2) late fee
  runs from GSTR-9 filing (or due date) to 9C filing — self-calculated by
  portal.

## Standing audit checks this reference adds (query fodder)

1. **Rule 37A sweep**: any supplier whose FY GSTR-3B was unfiled by 30 Sep of
   the next year (2B shows supplier filing status) → recipient must reverse
   by 30 Nov; reclaim only when supplier pays.
2. **Rule 37 180-day** payment ageing on creditors (interest u/s 50 on late
   reversal; reclaim not hit by 16(4)).
3. **Table 15 tie-out**: refunds (esp. IGST-paid exports via shipping-bill
   scroll), DRC-07 demands, liability-register Part II.
4. **DRC-03s**: any filed during/for the FY (Annual Return = cash-only;
   9C Reconciliation Statement = cash or credit from FY 24-25).
5. **Table 16 disclosures**: composition-dealer purchases, s.143 job-work
   (ITC-04), goods on approval >180 days.
6. **8D sign logic** and the 5-series exempt-turnover Rule 42/43 sanity
   formula (above).
7. **HSN digits**: 6-digit outward mandatory >5 Cr; inward Table 18 only for
   ≥10% items.
