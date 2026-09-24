# -*- coding: utf-8 -*-
"""検証用：LibreOffice出力からチャート/図形を除いたコピーを作る（openpyxlで読むため）"""
import re, shutil, sys, zipfile

src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
drop = lambda n: ("/charts/" in n or "/drawings/" in n or n.endswith("drawing.xml"))
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zin.infolist():
        if drop(it.filename):
            continue
        data = zin.read(it.filename)
        if it.filename == "[Content_Types].xml":
            s = data.decode("utf-8")
            s = re.sub(r'<Override[^>]*(charts|drawings)[^>]*/>', '', s)
            data = s.encode("utf-8")
        elif it.filename.endswith(".rels"):
            s = data.decode("utf-8")
            s = re.sub(r'<Relationship[^>]*(?:charts|drawings)/[^>]*/>', '', s)
            data = s.encode("utf-8")
        elif re.match(r"xl/worksheets/sheet\d+\.xml$", it.filename):
            s = data.decode("utf-8")
            s = re.sub(r'<drawing[^>]*/>', '', s)
            data = s.encode("utf-8")
        zo.writestr(it, data)
print("stripped ->", dst)
