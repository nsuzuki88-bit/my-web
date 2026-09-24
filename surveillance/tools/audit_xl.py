# -*- coding: utf-8 -*-
"""生成したブックを Excel の目で点検する（LibreOffice では見逃す差を拾う）"""
import re, sys, zipfile, html, collections

def col2n(s):
    n = 0
    for ch in s.upper(): n = n*26 + ord(ch) - 64
    return n

z = zipfile.ZipFile(sys.argv[1])
wbx = z.read("xl/workbook.xml").decode()
names = re.findall(r'<definedName name="([^"]+)"', wbx)
sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx)
rels = {a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"',
        z.read("xl/_rels/workbook.xml.rels").decode())}
rels.update({a: b for a, b in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"',
        z.read("xl/_rels/workbook.xml.rels").decode())})

bad = 0
print("== 名前付き範囲")
for n in names:
    m = re.fullmatch(r'([A-Za-z]{1,3})(\d+)', n)
    if m and col2n(m.group(1)) <= 16384 and int(m.group(2)) >= 1:
        print("  NG セル番地と同じ名前:", n); bad += 1
user_names = {n for n in names if not n.startswith("_xlnm")}
print("  ", sorted(user_names))

EXCEL2010 = set("""ABS AND AVERAGE CHOOSE COLUMN COLUMNS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS DATE DAY
EDATE EOMONTH HYPERLINK IF IFERROR INDEX INDIRECT INT ISBLANK ISERROR ISNUMBER LEFT LEN LOOKUP MATCH MAX MIN
MOD MONTH N NOT OFFSET OR RIGHT ROUND ROUNDDOWN ROUNDUP ROW ROWS SUM SUMIF SUMIFS SUMPRODUCT TEXT TODAY
VALUE WEEKDAY YEAR MID TRIM UPPER LOWER AVERAGEIF CONCATENATE VLOOKUP HLOOKUP ISTEXT NA DATEDIF""".split())

gen = ["VAE対象一覧", "臨床所見", "VAE-01", "All"]
funcs = collections.Counter(); arrays = plain_arr = 0
tok_unknown = collections.Counter()
for nm, rid in sheets:
    if nm not in gen: continue
    x = z.read("xl/" + rels[rid].replace("/xl/", "").lstrip("/")).decode()
    cells = re.findall(r'<c r="([A-Z]+\d+)"[^>]*>\s*<f([^>]*)>(.*?)</f>', x, re.S)
    cf = re.findall(r'<formula>(.*?)</formula>', x, re.S)
    for coord, attr, f in cells + [("CF", ' t="cf"', f) for f in cf]:
        f = html.unescape(f)
        body = re.sub(r'"[^"]*"', '""', f)                 # 文字列リテラルを除く
        body = re.sub(r"'[^']+'!", "S!", body)             # 引用符つきシート名
        body = re.sub(r"[^\s(),=<>+\-*/&:;{}]+!", "S!", body)  # 引用符なしのシート名
        body = body.replace("$", "")                       # 絶対参照の $ を外す
        for fn in re.findall(r'([A-Z][A-Z0-9\.]*)\(', body):
            funcs[fn] += 1
        if any(k in f for k in ("SUMPRODUCT(", "MATCH(TRUE", "LOOKUP(2,1/")):
            if 't="array"' in attr: arrays += 1
            elif attr.strip() != 't="cf"': plain_arr += 1; print("  NG 配列数式でない:", nm, coord)
        # 関数でもセル番地でもシート参照でもない識別子
        for t in re.findall(r'(?<![A-Za-z0-9_\.!$:])([A-Za-z_぀-鿿][A-Za-z0-9_\.぀-鿿]*)(?![A-Za-z0-9_\(!])', body):
            if re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+', t) or t in ("TRUE", "FALSE", "S"): continue
            if t in user_names: continue
            tok_unknown[(nm, t)] += 1
print("== 使っている関数:", dict(funcs))
unk = [f for f in funcs if f not in EXCEL2010]
print("  Excel2010に無い／要確認の関数:", unk or "なし"); bad += len(unk)
print(f"== 配列数式 {arrays}セル / 配列数式になっていない配列計算 {plain_arr}セル"); bad += plain_arr
print("== 名前として解釈される未知の識別子:", dict(tok_unknown) or "なし")
print("\n==== 問題", bad + len(tok_unknown), "件 ====")
