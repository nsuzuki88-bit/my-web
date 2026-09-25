# -*- coding: utf-8 -*-
"""再計算済みワークブックを Python独立実装（ref_vae.py）と突き合わせる。

  セクションB … 全ブロックの MV日数・DOE・VAC判定
  セクションA … VAC患者の一覧と、臨床所見・個別判定シートの割当
  臨床所見   … VAC患者だけがスロットに入っているか
"""
import datetime as dt
import sys, openpyxl
import ref_vae as R
from common import (LIST_A_TOP, LIST_A_BOT, LIST_B_TOP, LIST_B_BOT,
                    NSHEET, NCLIN, NBLK, MV_MIN, clin_base, CN)

ok = ng = 0


def chk(label, got, want):
    global ok, ng
    g = got.date() if isinstance(got, dt.datetime) else got
    w = want.date() if isinstance(want, dt.datetime) else want
    if g == w or (g in (None, "") and w in (None, "")):
        ok += 1
    else:
        ng += 1
        print(f"  NG  {label}: got {got!r} want {want!r}")


def main(calc, src):
    global ok, ng
    dates, ref = R.analyse_global(src)
    refmap = {r["k"]: r for r in ref}
    vac = R.vac_patients(ref)          # スロットの並び＝DOE①の早い順

    wb = openpyxl.load_workbook(calc, data_only=True)
    lst = wb["VAE対象一覧"]
    clin = wb["臨床所見"]

    # ---- セクションB ------------------------------------------------
    print("=== セクションB：全ブロックのスキャンとVAC判定 ===")
    want_vacno = {r["k"]: i for i, r in enumerate(vac, 1)}
    for k in range(1, NBLK + 1):
        r = LIST_B_TOP + k - 1
        w = refmap[k]
        chk(f"block{k} MV日数", lst.cell(r, 4).value or 0, w["mvtotal"])
        chk(f"block{k} 初回MV日", lst.cell(r, 5).value, w["first_mv"] if w["mvtotal"] else None)
        chk(f"block{k} 最終MV日", lst.cell(r, 6).value, w["last_mv"] if w["mvtotal"] else None)
        d = w["doe"]
        chk(f"block{k} DOE①", lst.cell(r, 7).value, d[0] if len(d) > 0 else None)
        chk(f"block{k} DOE②", lst.cell(r, 8).value, d[1] if len(d) > 1 else None)
        chk(f"block{k} 判定", lst.cell(r, 9).value, "VAC" if d else None)
        chk(f"block{k} VAC番号", lst.cell(r, 10).value, want_vacno.get(k))
    print(f"  {NBLK}ブロック照合、VAC判定 {len(vac)}名")

    # ---- セクションA ------------------------------------------------
    print("\n=== セクションA：VAC患者の一覧 ===")
    for n in range(1, NCLIN + 1):
        r = LIST_A_TOP + n - 1
        w = vac[n - 1] if n <= len(vac) else None
        if w is None:
            chk(f"slot{n} 空", lst.cell(r, 2).value, None)
            chk(f"slot{n} 臨床所見欄も空", lst.cell(r, 12).value, None)
            continue
        chk(f"slot{n} ブロック", lst.cell(r, 2).value, w["k"])
        chk(f"slot{n} 患者ID", lst.cell(r, 3).value, w["pid"])
        chk(f"slot{n} 氏名", lst.cell(r, 4).value, w["name"])
        chk(f"slot{n} MV1日目", lst.cell(r, 6).value, w["ep1_start"])
        chk(f"slot{n} MV日数", lst.cell(r, 7).value, w["mvtotal"])
        chk(f"slot{n} DOE①", lst.cell(r, 8).value, w["doe"][0])
        # 臨床所見が未入力なら基本判定のVAC、入力済みならIVAC/PVAPに格上げされうる
        judge = lst.cell(r, 9).value
        if lst.cell(r, 12).value == "入力あり":
            chk(f"slot{n} 判定①", judge in ("VAC", "IVAC", "PVAP"), True)
        else:
            chk(f"slot{n} 判定①", judge, "VAC")
        print(f"  slot{n}: block{w['k']} {w['pid']} {w['name']}  "
              f"DOE={w['doe'][0]:%m/%d}  判定={lst.cell(r,9).value}  "
              f"臨床所見={lst.cell(r,12).value}  判定シート={lst.cell(r,13).value!r}")

    # ---- 臨床所見 ----------------------------------------------------
    print("\n=== 臨床所見：VAC患者だけがスロットに入っているか ===")
    for n in range(1, NCLIN + 1):
        b = clin_base(n)
        w = vac[n - 1] if n <= len(vac) else None
        chk(f"臨床所見 slot{n} ブロック", clin.cell(b + 1, 1).value, w["k"] if w else None)
        chk(f"臨床所見 slot{n} 患者ID", clin.cell(b + 2, 1).value, w["pid"] if w else None)
        chk(f"臨床所見 slot{n} 氏名", clin.cell(b + 3, 1).value, w["name"] if w else None)
        chk(f"臨床所見 slot{n} DOE", clin.cell(b + 4, 1).value, w["doe"][0] if w else None)
    print(f"  {NCLIN}スロット照合（VAC患者{len(vac)}名のみ埋まる）")

    # ---- 個別判定シート ----------------------------------------------
    print("\n=== 個別判定シート：臨床所見の入力がある患者だけ ===")
    entered = [n for n in range(1, len(vac) + 1)
               if lst.cell(LIST_A_TOP + n - 1, 12).value == "入力あり"]
    print(f"  臨床所見に入力がある患者: {len(entered)}名")
    for i in range(1, NSHEET + 1):
        ws = wb[f"VAE-{i:02d}"]
        want = vac[entered[i - 1] - 1] if i <= len(entered) else None
        chk(f"VAE-{i:02d} ブロック", ws["C4"].value, want["k"] if want else None)
        if want:
            chk(f"VAE-{i:02d} 患者ID", ws["C5"].value, want["pid"])
            chk(f"VAE-{i:02d} DOE①", ws["K8"].value, want["doe"][0])

    # ---- VAE集計 --------------------------------------------------
    print("\n=== VAE集計：月別の分母と件数 ===")
    import ref_clabsi as RC
    _d, blks, _g = RC.blocks(src)
    vrows = R.read_blocks(src)[1]
    start = dates[0]
    months = [((start.month - 1 + i) // 12 + start.year, (start.month - 1 + i) % 12 + 1)
              for i in range(12)]
    vz = wb["VAE集計"]
    for i, (y, m) in enumerate(months):
        inm = lambda d: (d.year, d.month) == (y, m)
        pdays = sum(1 for b in blks for d, x in zip(_d, b["c"]) if x in ("C有", "C無") and inm(d))
        vdays = sum(1 for b in vrows for d, x in zip(dates, b["v"]) if x == "V有" and inm(d))
        events = sum(1 for r in vac[:NCLIN] for d in r["doe"][:2] if inm(d))
        chk(f"VAE集計 {m}月 延べ入室患者日数", vz.cell(5, i + 2).value, pdays)
        chk(f"VAE集計 {m}月 延べ人工呼吸器使用日数", vz.cell(6, i + 2).value, vdays)
        chk(f"VAE集計 {m}月 VAE合計", vz.cell(11, i + 2).value, events)
        if pdays:
            print(f"  {m:2d}月: 延べ入室患者日数 {pdays:4d}  延べ人工呼吸器使用日数 {vdays:4d}  "
                  f"使用比 {vz.cell(7, i + 2).value:.3f}  VAE {events}件")

    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1], sys.argv[2]) else 0)
