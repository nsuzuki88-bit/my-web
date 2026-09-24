# -*- coding: utf-8 -*-
"""段階的なしぼり込みの検証

  All → (VAC判定) → 臨床所見に入力 → 個別判定シートが自動作成 → IVAC/PVAPへ格上げ
"""
import datetime as dt
import sys, openpyxl
from common import (LIST_A_TOP, NCLIN, NSHEET, clin_base, C0, CN)

ok = ng = 0


def chk(label, got, want):
    global ok, ng
    g = got.date() if isinstance(got, dt.datetime) else got
    w = want.date() if isinstance(want, dt.datetime) else want
    if g == w or (g in (None, "") and w in (None, "")):
        ok += 1
        print(f"  ok   {label}: {got!r}")
    else:
        ng += 1
        print(f"  NG   {label}: got {got!r} want {want!r}")


def col_of(ws_all, date):
    from build import probe          # 日付行はファイルごとに違う（2026年度版は6行目）
    date_row = probe(ws_all)["DATE_ROW"]
    for c in range(C0, CN + 1):
        v = ws_all.cell(date_row, c).value
        if isinstance(v, dt.datetime) and v.date() == date.date():
            return c
    raise ValueError(date)


def mutate(src, dst, doe):
    """VAC患者1人目（臨床所見スロット1）にIVAC相当の所見を入力する"""
    wb = openpyxl.load_workbook(src, keep_vba=src.endswith(".xlsm"))
    al, cl = wb["All"], wb["臨床所見"]
    b = clin_base(1)
    cdoe = col_of(al, doe)
    cl.cell(b + 0, cdoe, 38.6)             # 最高体温（DOE当日）
    cl.cell(b + 4, cdoe + 1, "はい")       # 新規抗菌薬開始（DOE翌日）
    cl.cell(b + 5, cdoe + 1, "MEPM 4日間")
    wb.save(dst)
    print("wrote", dst, f"（DOE {doe:%Y/%m/%d} = {cdoe}列目に入力）")


def add_pvap(src, dst, doe):
    wb = openpyxl.load_workbook(src, keep_vba=src.endswith(".xlsm"))
    al, cl = wb["All"], wb["臨床所見"]
    b = clin_base(1)
    cdoe = col_of(al, doe)
    cl.cell(b + 6, cdoe + 1, "基準1")      # PVAP基準1
    cl.cell(b + 7, cdoe + 1, "ETA 10^5 CFU/mL")
    wb.save(dst)
    print("wrote", dst)


def verify(path, block, pid, doe, want_judge):
    wb = openpyxl.load_workbook(path, data_only=True)
    lst, clin = wb["VAE対象一覧"], wb["臨床所見"]
    r = LIST_A_TOP
    print(f"\n=== 一覧セクションA スロット1 ===")
    chk("  ブロック", lst.cell(r, 2).value, block)
    chk("  患者ID", lst.cell(r, 3).value, pid)
    chk("  DOE①", lst.cell(r, 8).value, doe)
    chk("  臨床所見", lst.cell(r, 12).value, "入力あり")
    chk("  判定シート番号", lst.cell(r, 13).value, 1)
    chk("  判定①（格上げ後）", lst.cell(r, 9).value, want_judge)

    print(f"\n=== VAE-01（臨床所見に入力した1人目に割り当て）===")
    w = wb["VAE-01"]
    chk("  ブロック番号", w["C4"].value, block)
    chk("  患者ID", w["C5"].value, pid)
    chk("  DOE①", w["K8"].value, doe)
    chk("  判定①", w["K9"].value, want_judge)

    print(f"\n=== VAE-02 は未割当のまま ===")
    chk("  ブロック番号", wb["VAE-02"]["C4"].value, None)

    print(f"\n=== 臨床所見スロット2以降は空 ===")
    chk("  スロット2 ブロック", clin.cell(clin_base(2) + 1, 1).value, None)

    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "mutate":
        mutate(sys.argv[2], sys.argv[3], dt.datetime.fromisoformat(sys.argv[4]))
    elif cmd == "pvap":
        add_pvap(sys.argv[2], sys.argv[3], dt.datetime.fromisoformat(sys.argv[4]))
    else:
        sys.exit(1 if verify(sys.argv[2], int(sys.argv[3]), int(sys.argv[4]),
                             dt.datetime.fromisoformat(sys.argv[5]), sys.argv[6]) else 0)
