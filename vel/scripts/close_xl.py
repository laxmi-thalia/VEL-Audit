import win32com.client as win32, pythoncom, time
pythoncom.CoInitialize()
try:
    xl=win32.GetActiveObject("Excel.Application")
except Exception as e:
    print("no running Excel found:", e); raise SystemExit
targets=["VEL_Sales_Register_FY2025-26_DRAFT.xlsx","VEL_Step2_Reco_GSTR1_DRAFT.xlsx"]
print("open workbooks:", [w.Name for w in xl.Workbooks])
for w in list(xl.Workbooks):
    if w.Name in targets:
        saved = bool(w.Saved)
        print(f"  {w.Name}: Saved={saved}")
        if not saved:
            w.Save(); print("    -> had unsaved changes, SAVED before closing")
        w.Close(SaveChanges=False)
        print("    -> closed")
left=[w.Name for w in xl.Workbooks]
print("remaining open:", left)
if not left:
    xl.Quit(); print("Excel had nothing else open - quit")
