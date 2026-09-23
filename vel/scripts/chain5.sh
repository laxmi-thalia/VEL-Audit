#!/bin/bash
cd "/c/Users/pawar/AppData/Local/Temp/claude/c--PROJECTS-accountic/ed6b1fc9-75d0-4eb0-9100-e931702d9fa7/scratchpad"
PY="/c/PROJECTS/accountic/backend/.venv/Scripts/python.exe"
for s in cascade_fix final_countif_rule t6a1_2b_period fy2627_reco remarks_mirror; do
  echo "=== $s START $(date +%H:%M:%S)"
  "$PY" -X utf8 "$s.py" || { echo "=== $s FAILED rc=$?"; echo "=== CHAIN5 ABORTED"; exit 1; }
  echo "=== $s END $(date +%H:%M:%S)"
done
echo "=== CHAIN5 DONE $(date +%H:%M:%S)"
