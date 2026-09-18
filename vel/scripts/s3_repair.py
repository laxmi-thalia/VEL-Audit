import win32com.client as win32, pythoncom, os
SP = os.path.dirname(os.path.abspath(__file__))
P = os.path.join(SP, "s3_test.xlsx")
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    try:
        wb = xl.Workbooks.Open(P, CorruptLoad=2)
        print("opened in REPAIR mode; sheets:", [w.Name for w in wb.Sheets])
        for w in wb.Sheets:
            try: n = w.UsedRange.SpecialCells(-4123, 16).Count
            except Exception: n = 0
            print("  ", w.Name, "used", w.UsedRange.Address, "errcells", n)
        wb.Close(False)
    except Exception as e:
        print("repair open failed too:", e)
    # bisect: drop one sheet at a time with openpyxl and try a normal open
    import openpyxl
    for drop in ["Exceptions", "Amendment Check", "Summary", "Month-on-Month", "GSTR-1 vs GSTR-3B"]:
        wb2 = openpyxl.load_workbook(P)
        if drop in wb2.sheetnames: del wb2[drop]
        t = os.path.join(SP, "s3_bisect.xlsx"); wb2.save(t)
        try:
            w = xl.Workbooks.Open(t); w.Close(False); print("without %-20s -> OPENS" % drop)
        except Exception:
            print("without %-20s -> still fails" % drop)
finally:
    xl.Quit()
