#!/bin/bash
cd "/c/Users/pawar/AppData/Local/Temp/claude/c--PROJECTS-accountic/ed6b1fc9-75d0-4eb0-9100-e931702d9fa7/scratchpad"
PY="/c/PROJECTS/accountic/backend/.venv/Scripts/python.exe"
for s in a7_reco_sheet a6_rcm_paid_vs_claimed a3_t13_diff_remarks; do
  echo "=== $s START $(date +%H:%M:%S)"
  "$PY" -X utf8 "$s.py" || { echo "=== $s FAILED rc=$?"; echo "=== POST11 ABORTED"; exit 1; }
  echo "=== $s END $(date +%H:%M:%S)"
done
cp "/c/Users/pawar/Downloads/VEL_GST_Audit_FY2025-26_MASTER (2).xlsx" master2_snapshot_before_chain7.xlsx
echo "=== snapshot before_chain7 taken $(date +%H:%M:%S)"
bash chain5.sh || { echo "=== POST11 ABORTED (chain)"; exit 1; }
echo "=== POST11 DONE $(date +%H:%M:%S)"
