import win32com.client as win32, pythoncom, collections, re
P = r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    ws = wb.Worksheets("SR_2025-26")
    try: errs = ws.Range("A5:BA27006").SpecialCells(-4123, 16).Count
    except Exception: errs = 0
    print("formula ERROR cells:", errs)
    hdr = [ws.Cells(4, c).Value for c in range(1, 54)]
    print("headers O..T :", hdr[14:20])
    print("headers AO..AT:", hdr[40:46])
    print("headers AY..BA:", hdr[50:53])
    def col(name):
        return hdr.index(name) + 1
    rng = ws.Range("A5:BA27006").Value
    cnt = lambda ci: collections.Counter((r[ci - 1] if r[ci - 1] is not None else "") for r in rng)
    print("\nSUBTOTAL taxable (R2):", ws.Range("R2").Value, "| CGST (T2):", ws.Range("T2").Value, "| Qty (AJ2):", ws.Range("AJ2").Value)
    print("Supplier name blanks:", sum(1 for r in rng if not r[col('Supplier Legal Name') - 1]))
    pos = cnt(col("Place of Supply"))
    print("Place of Supply formats:", collections.Counter(
        "2-digit" if re.fullmatch(r"\d{2}", str(k)) else "blank" if k == "" else "other" for k in pos.elements()))
    rc = cnt(col("GST rate check")); print("GST rate check:", dict(rc))
    ar = [r[col("Actual Rate Charged (%)") - 1] for r in rng if r[col("Actual Rate Charged (%)") - 1] not in (None, "")]
    print("Actual rate sample:", ar[:6], "| distinct rounded:", sorted(collections.Counter(round(v) for v in ar).items())[:12])
    pc = cnt(col("POS check")); print("POS check:", dict(pc))
    print("GSTR-1 Type:", dict(cnt(col("GSTR-1 Type"))))
    print("Matched with GSTR-1:", dict(cnt(col("Matched with GSTR-1"))))
    mo = cnt(col("GSTR-1 Month")); print("GSTR-1 Month (top):", mo.most_common(13))
    print("G/S:", dict(cnt(col("GOOD (G) or SERVICE(S)"))))
    print("GSTR-9:", dict(cnt(col("GSTR-9"))))
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
