import json, os
SP=os.path.dirname(os.path.abspath(__file__))
hdrs=json.load(open(os.path.join(SP,"hdrs.json")))
# target letter, target header, kind, source (lowercased) or rule, note
M=[
("A","Month","DERIVED","document date","month name from Document Date"),
("B","My State","MAPPED","state","NEEDS state-name normalisation (Madya/Madhya etc.)"),
("C","My GSTIN","MAPPED","gstin",""),
("D","Supplier Legal Name","MAPPED","seller legal name",""),
("E","Document Date","MAPPED","document date",""),
("F","Document Number","MAPPED","document number",""),
("G","Invoice Type","DERIVED","document type","Sales / Credit Notes from doc type"),
("H","Document Type Code","MAPPED","document type","NEEDS doc-type normalisation (15 spellings)"),
("I","Supply Type Code","MAPPED","supply type","blank on advances - must be filled"),
("J","Recipient GSTIN","MAPPED","buyer gstin",""),
("K","Recipient Legal Name","MAPPED","buyer legal name",""),
("L","Place of Supply","MAPPED","buyer place of supply",""),
("M","Item Price","MAPPED","item unit price",""),
("N","Gross Amount","MAPPED","item total amount",""),
("O","GST Rate","MAPPED","item gst rate",""),
("P","GST rate check","DERIVED","-","tax / taxable lands on a valid rate"),
("Q","Taxable Value","MAPPED","item assessable amount",""),
("R","IGST Amount","MAPPED","item igst amount",""),
("S","CGST Amount","MAPPED","item cgst amount",""),
("T","SGST Amount","MAPPED","item sgst amount",""),
("U","Cess Amount","MAPPED","item cess amount",""),
("V","Total GST","FORMULA","-","=SUM(R:T); cross-check vs source 'total tax'"),
("W","Total Invoice Value","FORMULA","-","=SUM(Q:U)"),
("X","Invoice Reference No","MAPPED","invoice reference no",""),
("Y","IRN (len)","DERIVED","-","=LEN(X); must be 64 or blank"),
("Z","Status","MAPPED","status",""),
("AA","Cancel Date","MAPPED","cancel date",""),
("AB","Document Status","MAPPED","document status",""),
("AC","Ship To GSTIN","MAPPED","shipping gstin",""),
("AD","Ship To LegalName","MAPPED","shipping legalname",""),
("AE","Item Description","MAPPED","item product description",""),
("AF","GOOD (G) or SERVICE(S)","DERIVED","item hsn code","SAC 99xx = S; cross-check 'item is service'"),
("AG","HSN or SAC Code","MAPPED","item hsn code for uploading","?? or plain 'item hsn code' - ASK CA"),
("AH","Item Description for uploading","MAPPED","item product description","?? no cleaned source col - ASK CA"),
("AI","Quantity for uploading","MAPPED","item quantity",""),
("AJ","Unit of Measurement for uploading","MAPPED","item unit for uploading",""),
("AK","GL Name","NO SOURCE","-","from SAP GL - not yet received"),
("AL","Profit Centre","NO SOURCE","-","from SAP GL - not yet received"),
("AM","Cost Centre","NO SOURCE","-","from SAP GL - not yet received"),
]
print(f"{'col':4s} {'target':36s} {'kind':9s} {'resolves in':12s} source")
bad=[]
for letter,tgt,kind,src,note in M:
    if kind in ("MAPPED","DERIVED") and src not in ("-",""):
        miss=[m for m,h in hdrs.items() if src not in h]
        ok = "ALL 12" if not miss else f"MISSING {miss}"
        if miss: bad.append((letter,tgt,src,miss))
    else:
        ok = "n/a"
    print(f"{letter:4s} {tgt:36s} {kind:9s} {ok:12s} {src}")
print("\n--- unresolved ---")
print(bad if bad else "  none - every mapped source column exists in all 12 months")
