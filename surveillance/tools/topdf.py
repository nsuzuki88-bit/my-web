import os, subprocess, sys, time, uno
from com.sun.star.beans import PropertyValue
def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
src, dst, sheet = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]), sys.argv[3]
port=2004
so=subprocess.Popen(["soffice","--headless","--norestore","--nolockcheck","--nodefault",
  "-env:UserInstallation=file:///tmp/lo_pdf", f"--accept=socket,host=127.0.0.1,port={port};urp;"])
ctx=None
res=uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",uno.getComponentContext())
for _ in range(60):
    try:
        ctx=res.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
d=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
doc=d.loadComponentFromURL(uno.systemPathToFileUrl(src),"_blank",0,(prop("Hidden",True),))
doc.calculateAll()
sheets=doc.Sheets
# 対象シートをアクティブにしてから、それ以外を非表示にして1シートだけPDF化
doc.CurrentController.setActiveSheet(sheets.getByName(sheet))
for i in range(sheets.Count):
    sh=sheets.getByIndex(i)
    if sh.Name != sheet:
        sh.IsVisible = False
doc.storeToURL(uno.systemPathToFileUrl(dst),(prop("FilterName","calc_pdf_Export"),))
doc.close(False); d.terminate(); time.sleep(2); so.terminate()
print("pdf:",dst)
