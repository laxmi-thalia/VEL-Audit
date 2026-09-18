import win32com.client as win32, pythoncom
pythoncom.CoInitialize()
xl=None
for how in ("GetObject","Dispatch"):
    try:
        xl = win32.GetObject(Class="Excel.Application") if how=="GetObject" else win32.Dispatch("Excel.Application")
        n = xl.Workbooks.Count
        print(f"attached via {how}; workbooks open: {n}")
        break
    except Exception as e:
        print(f"{how} failed: {e}"); xl=None
if xl is None: raise SystemExit("could not attach")
targets={"VEL_Sales_Register_FY2025-26_DRAFT.xlsx","VEL_Step2_Reco_GSTR1_DRAFT.xlsx"}
for w in list(xl.Workbooks):
    print("  open:", w.Name, "| Saved =", bool(w.Saved))
for w in list(xl.Workbooks):
    if w.Name in targets:
        if not bool(w.Saved): w.Save(); print(f"   {w.Name}: saved unsaved changes")
        w.Close(SaveChanges=False); print(f"   {w.Name}: closed")
print("remaining:", [w.Name for w in xl.Workbooks])
