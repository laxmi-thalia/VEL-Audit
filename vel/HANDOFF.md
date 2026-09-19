# VEL GSTR-9/9C FY 2025-26 — handoff (as of 19-09-2026)

Start a Claude Code session in this repo and say: "Read vel/HANDOFF.md and the four gst-audit-vel-* skills, then continue."
The skills (`.claude/skills/gst-audit-vel-{sales,rcm,itc}` + `gst-audit`) hold the full method and the dated changes-logs
with every ruling the CAs / Pawan gave. This file is only the current state and the open items.

## Deliverable (NOT in the repo — client data)
- Working copy: `C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx` (~65 MB, 72 sheets).
- Sent to the CAs: the `.xlsb` next to it (re-export after every change: COM `SaveAs FileFormat=50`).
- Only ever write to that path. Before any write: `open(path,'r+b')` — if locked, ask the user to close Excel without saving.
- Snapshots of every step live in the session scratchpad (`master2_snapshot_before_*.xlsx`); take one before each write.
- Server data: `\192.168.1.69\gst folder\GST Returns\GST Audit & Annual Return\FY 2025-26\1. Corporate Clients\VEL\`
  (COM cannot open UNC paths — copy locally; Git-Bash `/c/...` paths do not work inside Python `open()`).

## Verified state (last full check 19-09-2026, 0 error cells)
- ITC Register 2025-26 golden Total GST **1,06,99,69,542.15**; B_ = ITC-category tax 93,75,28,722.19; Net-ITC check 0.00.
- Matching: exact 32,035 lines; GSTIN+amount 2,229; invoice-similar+amount 1,273; date+amount 205; FY 24-25 2B 932;
  not found 5,081; RCM/ISD/URD 2,553. Fallbacks are single-candidate, ±100, same recipient, flagged "review".
- RCM Register 6..3636 (MP Jul-25 + Jan-26 from monthly files); Statewise RCM vs 3B +5.98L = HOIS 5.50L/99k + Gujarat
  +94,649 / TN +7,500 / Telangana −54,372 taxable-only. RCM GL keyed on Posting Date | Document Number.
- Tax comp report Reasons: 166 Matched / 38 explained / 0 residual (Computation sheets vs filed 3B).
- T6A1 Extract: 3,256 rows, 2B Return Period live; ITC Register 2026-27: 2,109 document rows from the client's Inputs sheets.

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
`b2_remerge_2627.py` → `strip_orphan_pivots.py` → `readd_buttons2.py` → `cascade_fix.py` → `final_countif_rule.py` →
`t6a1_2b_period.py` → verify (0 errors, golden, Net-ITC 0.00) → export xlsb. Do everything in ONE COM session where possible
(each open/recalc/save of the master costs ~2 min). Never insert rows at the first data row (row 6) — ranges that start at $6
shift; insert inside the range and delete the old rows.
