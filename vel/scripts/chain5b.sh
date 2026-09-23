#!/bin/bash
cd "/c/Users/pawar/AppData/Local/Temp/claude/c--PROJECTS-accountic/ed6b1fc9-75d0-4eb0-9100-e931702d9fa7/scratchpad"
PY="/c/PROJECTS/accountic/backend/.venv/Scripts/python.exe"
for s in t6a1_2b_period fy2627_reco remarks_mirror dates_ddmmyy; do
  echo "=== $s START $(date +%H:%M:%S)"
  "$PY" -X utf8 "$s.py" || { echo "=== $s FAILED rc=$?"; echo "=== CHAIN5B ABORTED"; exit 1; }
  echo "=== $s END $(date +%H:%M:%S)"
done
echo "=== CHAIN5B DONE $(date +%H:%M:%S)"
