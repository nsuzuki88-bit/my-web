import os, subprocess, sys, time, uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellRangeAddress
def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
src,dst,sheet,c1,r1,c2,r2 = sys.argv[1],sys.argv[2],sys.argv[3],*map(int,sys.argv[4:8])
src,dst = os.path.abspath(src), os.path.abspath(dst)
port=2008
so=subprocess.Popen(["soffice","--headless","--norestore","--nolockcheck","--nodefault",
  "-env:UserInstallation=file:///tmp/lo_pdf3", f"--accept=socket,host=127.0.0.1,port={port};urp;"])
res=uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",uno.getComponentContext())
ctx=None
for _ in range(60):
    try:
        ctx=res.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
d=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
doc=d.loadComponentFromURL(uno.systemPathToFileUrl(src),"_blank",0,(prop("Hidden",True),))
doc.calculateAll()
sheets=doc.Sheets
idx=[sheets.getByIndex(i).Name for i in range(sheets.Count)].index(sheet)
# 対象シート以外は印刷範囲を空にして、対象シートだけ範囲を指定する
for i in range(sheets.Count):
    sheets.getByIndex(i).setPrintAreas(())
a=CellRangeAddress(); a.Sheet=idx; a.StartColumn=c1; a.StartRow=r1; a.EndColumn=c2; a.EndRow=r2
sheets.getByName(sheet).setPrintAreas((a,))
doc.storeToURL(uno.systemPathToFileUrl(dst),(prop("FilterName","calc_pdf_Export"),))
doc.close(False); d.terminate(); time.sleep(2); so.terminate()
print("pdf:",dst)
