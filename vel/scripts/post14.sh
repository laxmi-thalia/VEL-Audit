#!/bin/bash
cd "/c/Users/pawar/AppData/Local/Temp/claude/c--PROJECTS-accountic/ed6b1fc9-75d0-4eb0-9100-e931702d9fa7/scratchpad"
PY="/c/PROJECTS/accountic/backend/.venv/Scripts/python.exe"
echo "=== b3_2b_date_keys START $(date +%H:%M:%S)"
"$PY" -X utf8 b3_2b_date_keys.py || { echo "=== b3_2b_date_keys FAILED"; echo "=== POST14 ABORTED"; exit 1; }
echo "=== b3_2b_date_keys END $(date +%H:%M:%S)"
cp "/c/Users/pawar/Downloads/VEL_GST_Audit_FY2025-26_MASTER (2).xlsx" master2_snapshot_before_chain8.xlsx
bash chain5.sh || { echo "=== POST14 ABORTED (chain)"; exit 1; }
echo "=== POST14 DONE $(date +%H:%M:%S)"
