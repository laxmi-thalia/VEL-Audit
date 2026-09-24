#!/bin/bash
cd "/c/Users/pawar/AppData/Local/Temp/claude/c--PROJECTS-accountic/ed6b1fc9-75d0-4eb0-9100-e931702d9fa7/scratchpad"
PY="/c/PROJECTS/accountic/backend/.venv/Scripts/python.exe"
for s in a6_rcm_paid_vs_claimed a3_t13_diff_remarks; do
  echo "=== $s START $(date +%H:%M:%S)"
  "$PY" -X utf8 "$s.py" || { echo "=== $s FAILED rc=$?"; echo "=== POST13 ABORTED"; exit 1; }
  echo "=== $s END $(date +%H:%M:%S)"
done
cp "/c/Users/pawar/Downloads/VEL_GST_Audit_FY2025-26_MASTER (2).xlsx" master2_snapshot_before_chain7.xlsx
echo "=== snapshot before_chain7 taken $(date +%H:%M:%S)"
bash chain5.sh || { echo "=== POST13 ABORTED (chain)"; exit 1; }
echo "=== a7_reco_sheet START $(date +%H:%M:%S)"
"$PY" -X utf8 a7_reco_sheet.py || { echo "=== a7_reco_sheet FAILED"; echo "=== POST13 ABORTED (a7)"; exit 1; }
echo "=== POST13 DONE $(date +%H:%M:%S)"
