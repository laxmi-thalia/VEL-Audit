# VEL GSTR-9/9C FY 2025-26 — handoff (as of 22-09-2026)

Start a Claude Code session in this repo and say: "Read vel/HANDOFF.md and the four gst-audit-vel-* skills, then continue."
The skills (`.claude/skills/gst-audit-vel-{sales,rcm,itc}` + `gst-audit`) hold the full method and the dated changes-logs
with every ruling the CAs / Pawan gave. This file is only the current state and the open items.

## Deliverable (NOT in the repo — client data)
- Working copy: `C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx` (~65 MB, 72 sheets).
- Sent to the CAs: the `.xlsb` next to it (re-export after every change: COM `SaveAs FileFormat=50`).
- Only ever write to that path. Before any write: `open(path,'r+b')` — if locked, ask the user to close Excel without saving.
- Snapshots of every step live in the session scratchpad (`master2_snapshot_before_*.xlsx`); take one before each write.
- Server data: `\\192.168.1.69\gst folder\GST Returns\GST Audit & Annual Return\FY 2025-26\1. Corporate Clients\VEL\`
  (COM cannot open UNC paths — copy locally; Git-Bash `/c/...` paths do not work inside Python `open()`).

## Verified state (last full check 21-09-2026 19:42 after chain4, 0 error cells)
- ITC Register 2025-26 golden Total GST **1,06,99,69,542.15**; B_ = ITC-category tax 93,75,28,722.19; 2B_ 87,03,42,198.29;
  D_ 6,71,86,523.90; Net-ITC check 0.00; ITCR vs 3B net 2,52,024.32; ITC Summary 6A1 blocks G/J/M/P/S
  2,67,48,108.62 / 1,45,33,807.83 / 3,75,19,734.83 / 78,67,313.94 / 26,09,545.00.
- Matching = `vel/scripts/reco_lib.py` (tests in `tests/vel/`), numbered remarks 1–15 identical on the register and the 2B sheet
  (2B side is a live lookup, remarks_mirror.py). Register: 1 exact 32,075 | 2 recipient differs 5 | 3 GSTIN corrected 17 |
  4 amount & date tie 2,714 | 5 similar 311 | 6 GSTIN+amount 457 | 7 date+amount 82 | 8 amount 1 | 9 FY 24-25 2B 1,164 |
  10 not in 2B 4,948 | 12 RCM 2,527 | 14 URD 7 | 15 invalid GSTIN 2. 2B side: 11 not in books 7,239, matched 6,628.
- `GSTR-2B ITC Data` = Octa PAN-level FY 24-25 export (10,322 docs + 188 ISD) with LY-9C permanent-reversal flags; T6A1 Extract
  3,395 rows keyed on GSTIN + invoice + FY (1,354 dated / 2,041 not in 2B).
- ITC Register 2026-27: 2,109 lines = 181 documents, reconciled in the 25-26 format (B_ 2,42,80,458.10 == ITC tax, 2B_ 2,36,63,716.09).
- RCM Register 6..3636; RCM GL restricted to FY 25-26 + Mar-25 output + open items (22-09; Found in Output GL 3,630/1);
  POS block NA on 2,819 rows without vendor GSTIN; Statewise RCM vs 3B +5.98L (HOIS 5.50L/99k + Gujarat /
  TN / Telangana taxable-only). Tax comp Reasons: 166 Matched / 38 explained / 0 residual.

## Open items
1. **ZFI06**: client export (1,717 docs) contains none of the register's 6,018 FY 25-26 documents → Expense GL Element /
   PO Number / Expense Description filled on 144 rows only; `ZFI06 status` column explains it. Needs a fresh ZFI06 run
   (company code 1000, 01.04.2025–31.03.2026, all BPs, no selection). On arrival: reload `ZFI06 Data` sheet (header row 2,
   doc number text) — formulas recompute.
2. FY 25-26 Octa 2B exports missing for TN, TG, WB, HR (Kerala file empty) → those states' register lines show "not in 2B".
3. HOIS rows (2, no GSTIN) carrying 5.50L RCM — CA to decide where they belong.
4. Haryana Apr–Dec working files are nil-ITC stubs; MP July RCM Statewise now ties.
5. Eligibility column deferred by the CAs; user to confirm the Tax comp header text (says −4B(2), data is −4D(1)).
6. Optional: raw-data sheet header font sizes (Calibri 9/10) not forced to 11.

## How to re-run the ITC chain after a data change
`b2_remerge_2627.py` → `strip_orphan_pivots.py` → `readd_buttons2.py` → then `vel/scripts/chain4.sh` (= `cascade_fix.py` →
`final_countif_rule.py` → `t6a1_2b_period.py` → `fy2627_reco.py` → `remarks_mirror.py` → `rcm_gl_rebuild.py`, ~11 min, one COM
session per script, log to chain4.log) → verify (0 errors, golden, Net-ITC 0.00, 6A1 blocks) → export xlsb. Snapshot the master
first. Never insert rows at the first data row (row 6) — ranges that start at $6 shift; insert inside the range and delete the old
rows. Never insert columns before existing ones on ITC Register 2026-27 (ITC Summary reads it by letter) — append after the last header.
