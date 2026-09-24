# -*- coding: utf-8 -*-
"""LibreOffice(UNO)でブックを開き、全再計算して保存し直す"""
import os, subprocess, sys, time, uno
from com.sun.star.beans import PropertyValue


def prop(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p


def connect(port, tries=60):
    ctx = uno.getComponentContext()
    res = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", ctx)
    url = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    for _ in range(tries):
        try:
            return res.resolve(url)
        except Exception:
            time.sleep(1)
    raise RuntimeError("could not connect to soffice")


def main(src, dst, port=2002):
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    profile = "/tmp/lo_profile_recalc"
    soffice = subprocess.Popen([
        "soffice", "--headless", "--norestore", "--nolockcheck", "--nodefault",
        f"-env:UserInstallation=file://{profile}",
        f"--accept=socket,host=127.0.0.1,port={port};urp;",
    ])
    try:
        ctx = connect(port)
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(src), "_blank", 0,
            (prop("Hidden", True), prop("UpdateDocMode", 1), prop("MacroExecutionMode", 0)))
        t = time.time()
        doc.calculateAll()
        print("calculateAll:", round(time.time() - t, 2), "s")
        doc.storeToURL(uno.systemPathToFileUrl(dst),
                       (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(False)
        print("wrote", dst)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        time.sleep(2)
        soffice.terminate()


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
