import re
p = r"c:\PROJECTS\accountic\.claude\skills\gst-audit-vel-itc\SKILL.md"
b = open(p, "rb").read()
pat = b"Tax Comparison Reports" + bytes([0x82]) + b"5-26_<GSTIN>"
good = b"Tax Comparison Reports" + b"\\" + b"2025-26_<GSTIN>"
b2 = b.replace(pat, good)
# also handle the two-byte utf-8 form if present
pat2 = b"Tax Comparison Reports" + bytes([0xC2, 0x82]) + b"5-26_<GSTIN>"
b2 = b2.replace(pat2, good)
open(p, "wb").write(b2)
print("changed:", b2 != b)
print("now:", re.search(rb"Tax Comparison Reports.{0,6}5-26_<GSTIN>", b2).group())
