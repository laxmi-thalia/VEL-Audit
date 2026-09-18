"""After an openpyxl save the COM-created pivot PT_6A1 keeps its definition but loses its cache -> Excel refuses to open.
Strip the PT_6A1 parts (pivotTable, its rels, cache definition/records, workbook + sheet rels, content types) at zip level.
The pivot is re-created by cascade_fix.py. Idempotent: prints 'nothing to strip' when absent."""
import zipfile, re, shutil, sys
P = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
TMP = P + ".tmp"
zin = zipfile.ZipFile(P); names = zin.namelist()
orphan = [n for n in names if n.startswith("xl/pivotTables/pivotTable") and n.endswith(".xml") and "PT_6A1" in zin.read(n).decode()]
if not orphan: print("nothing to strip"); zin.close(); sys.exit(0)
pt = orphan[0]; ptrel = zin.read(pt.replace("pivotTables/", "pivotTables/_rels/") + ".rels").decode()
cache = "xl/" + re.search(r'Target="/?(?:xl/)?([^"]*pivotCacheDefinition\d+\.xml)"', ptrel).group(1); cid = re.search(r"(\d+)\.xml", cache).group(1)
drop = {pt, pt.replace("pivotTables/", "pivotTables/_rels/") + ".rels", cache, cache.replace("pivotCache/", "pivotCache/_rels/") + ".rels", "xl/pivotCache/pivotCacheRecords%s.xml" % cid}
wrel = zin.read("xl/_rels/workbook.xml.rels").decode()
m = re.search(r'<Relationship [^>]*?Id="(rId\d+)"[^>]*?pivotCacheDefinition%s\.xml' % cid, wrel) or re.search(r'pivotCacheDefinition%s\.xml"[^>]*?Id="(rId\d+)"' % cid, wrel); rid = m.group(1)
sheet_rel = next(n for n in names if n.startswith("xl/worksheets/_rels/") and pt.split("/")[-1] in zin.read(n).decode())
zout = zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    n = item.filename
    if n in drop: continue
    data = zin.read(n)
    if n == "xl/workbook.xml": data = re.sub(r'<pivotCache [^>]*r:id="%s"/>' % rid, "", data.decode()).encode()
    elif n == "xl/_rels/workbook.xml.rels": data = re.sub(r'<Relationship [^>]*pivotCacheDefinition%s\.xml[^>]*/>' % cid, "", data.decode()).encode()
    elif n == sheet_rel:
        d = re.sub(r'<Relationship [^>]*%s[^>]*/>' % re.escape(pt.split("/")[-1]), "", data.decode())
        if re.search(r"<Relationships[^>]*>\s*</Relationships>", d): continue
        data = d.encode()
    elif n == "[Content_Types].xml": data = re.sub(r'<Override [^>]*(pivotCacheDefinition%s|pivotCacheRecords%s|%s)[^>]*/>' % (cid, cid, re.escape(pt.split("/")[-1])), "", data.decode()).encode()
    zout.writestr(item, data)
zout.close(); zin.close(); shutil.move(TMP, P); print("stripped:", sorted(drop))
