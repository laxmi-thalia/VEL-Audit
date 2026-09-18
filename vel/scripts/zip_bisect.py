"""Zip-level sheet removal bisect: drop named sheets from the package (workbook.xml entry, rels, content type,
part, definedNames with that localSheetId; renumber later localSheetIds) and test whether Excel opens the result."""
import zipfile, re, html, sys, os, time
import win32com.client as win32, pythoncom
SRC = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
SP = os.path.dirname(os.path.abspath(__file__))
def build(drop_names, dst):
    zin = zipfile.ZipFile(SRC)
    wbx = zin.read("xl/workbook.xml").decode()
    rel = zin.read("xl/_rels/workbook.xml.rels").decode()
    rels = {}
    for m in re.finditer(r'<Relationship ([^>]*)/>', rel):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))); rels[a["Id"]] = a["Target"].lstrip("/")
    sheets = []
    for i, m in enumerate(re.finditer(r"<sheet ([^>]*)/>", wbx)):
        a = dict(re.findall(r'(\w+:?\w*)="([^"]*)"', m.group(1))); sheets.append((i, html.unescape(a["name"]), a.get("r:id") or a.get("id"), m.group(0)))
    drop_idx = {i for i, n, rid, tag in sheets if n in drop_names}
    drop_parts = set(); drop_rids = set()
    for i, n, rid, tag in sheets:
        if i in drop_idx:
            part = rels[rid]; part = part if part.startswith("xl/") else "xl/" + part
            drop_parts.add(part); drop_parts.add(part.replace("worksheets/", "worksheets/_rels/") + ".rels"); drop_rids.add(rid)
            wbx = wbx.replace(tag, "")
    # definedNames: drop those on dropped sheets, renumber the rest
    def fix_dn(m):
        lid = int(m.group(1))
        if lid in drop_idx: return ""
        new = lid - sum(1 for d in drop_idx if d < lid)
        return m.group(0).replace('localSheetId="%d"' % lid, 'localSheetId="%d"' % new)
    wbx = re.sub(r'<definedName [^>]*localSheetId="(\d+)"[^>]*>[^<]*</definedName>', fix_dn, wbx)
    # activeTab / firstSheet renumber
    def fix_view(m):
        v = int(m.group(2)); new = v - sum(1 for d in drop_idx if d < v)
        return '%s="%d"' % (m.group(1), max(new, 0))
    wbx = re.sub(r'(activeTab|firstSheet)="(\d+)"', fix_view, wbx)
    for rid in drop_rids: rel = re.sub(r'<Relationship [^>]*Id="%s"[^>]*/>' % rid, "", rel)
    ct = zin.read("[Content_Types].xml").decode()
    for part in drop_parts: ct = re.sub(r'<Override [^>]*PartName="/%s"[^>]*/>' % re.escape(part), "", ct)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        n = item.filename
        if n in drop_parts: continue
        data = zin.read(n)
        if n == "xl/workbook.xml": data = wbx.encode()
        elif n == "xl/_rels/workbook.xml.rels": data = rel.encode()
        elif n == "[Content_Types].xml": data = ct.encode()
        zout.writestr(item, data)
    zout.close(); zin.close()
def try_open(path):
    pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        try: w = xl.Workbooks.Open(path, ReadOnly=True); n = w.Worksheets.Count; w.Close(False); return "OPENS (%d sheets)" % n
        except Exception as e: return "fails"
    finally: xl.Quit()
TRIALS = [["LY 24-25 claims"], ["Table 8C vs 13-12 (FY 25-26)"], ["T12B_T12C differences"], ["Table 13 & 6A1 differences"], ["ITC Summary"],
          ["ITC Summary", "Table 13 & 6A1 differences", "T12B_T12C differences", "Table 8C vs 13-12 (FY 25-26)", "LY 24-25 claims"]]
for names in TRIALS:
    dst = os.path.join(SP, "zb.xlsx"); t0 = time.time(); build(set(names), dst)
    print("without %-90s -> %s  (%.0fs)" % (", ".join(names), try_open(dst), time.time() - t0), flush=True)
