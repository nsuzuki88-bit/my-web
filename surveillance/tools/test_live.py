# -*- coding: utf-8 -*-
"""入力を書き換えたときに、しぼり込みの各段が追従するかを検証する

  All を直す → VAC判定が変わる → 臨床所見のスロットが入れ替わる
  → 個別判定シートの割当が変わる
"""
import datetime as dt
import sys, openpyxl
from common import (LIST_A_TOP, LIST_B_TOP, NCLIN, NSHEET, clin_base,
                    C0, all_v, all_f, all_p, all_id, all_name)

D = lambda m, d: dt.datetime(2025, m, d)
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


def mutate(src, dst):
    wb = openpyxl.load_workbook(src)
    al, cl = wb["All"], wb["臨床所見"]

    # ① ブロック10：FiO2の上昇を 20→5 ポイントに下げる → VACでなくなる
    for d in (3, 4):
        al.cell(all_f(10), C0 + d - 1, 45)

    # ② ブロック20に新規患者を追加 → 新しいVACとして末尾に並ぶ
    al.cell(all_id(20), 1, "T20")
    al.cell(all_name(20), 1, "検証20")
    for d, (f, p) in enumerate([(40, 5), (40, 5), (70, 5), (70, 5)]):
        al.cell(all_v(20), C0 + d, "V有")
        al.cell(all_f(20), C0 + d, f)
        al.cell(all_p(20), C0 + d, p)

    # ③ ブロック9（VAC 6人目）のPVAP基準を削除 → PVAP から IVAC へ
    cl.cell(clin_base(6) + 6, C0 + 5 - 1).value = None

    # ④ スロット7（ブロック11・DOE 4/3）にDOEから離れた日付で入力 → 警告が出るはず
    cl.cell(clin_base(7) + 0, C0 + 20 - 1, 38.2)

    wb.save(dst)
    print("wrote", dst)


def verify(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    lst, clin = wb["VAE対象一覧"], wb["臨床所見"]
    A = lambda n, c: lst.cell(LIST_A_TOP + n - 1, c).value
    Bk = lambda k, c: lst.cell(LIST_B_TOP + k - 1, c).value

    print("\n① ブロック10：FiO2の上昇を20→5ポイントに変更 → VACでなくなる")
    chk("  セクションB 判定", Bk(10, 9), None)
    chk("  セクションB VAC番号", Bk(10, 10), None)

    print("\n② ブロック20に患者を追加 → 新しいVACとして末尾に並ぶ")
    chk("  セクションB 判定", Bk(20, 9), "VAC")
    chk("  セクションB DOE①", Bk(20, 7), D(4, 3))

    print("\n VAC患者の並び（ブロック番号）")
    got = [A(n, 2) for n in range(1, NCLIN + 1) if A(n, 2) is not None]
    chk("  並び順", got, [1, 3, 4, 5, 8, 9, 11, 12, 20])

    print("\n③ ブロック9（6人目）のPVAP基準を削除 → IVACへ格下げ")
    chk("  判定①", A(6, 9), "IVAC")
    chk("  ブロック番号", A(6, 2), 9)

    print("\n 臨床所見に入力がある患者と判定シートの割当")
    chk("  5人目(ブロック8) 臨床所見", A(5, 12), "入力あり")
    chk("  5人目 判定シート番号", A(5, 13), 1)
    chk("  6人目(ブロック9) 判定シート番号", A(6, 13), 2)
    chk("  7人目(ブロック11) 判定シート番号", A(7, 13), 3)
    chk("  8人目(ブロック12) 臨床所見", A(8, 12), "未入力")
    chk("  VAE-01 のブロック", wb["VAE-01"]["C4"].value, 8)
    chk("  VAE-02 のブロック", wb["VAE-02"]["C4"].value, 9)
    chk("  VAE-03 のブロック", wb["VAE-03"]["C4"].value, 11)
    chk("  VAE-04 は未割当", wb["VAE-04"]["C4"].value, None)

    print("\n④ DOEから離れた日に入力したスロット7 → 警告が出る")
    chk("  スロット5（正常）", clin.cell(clin_base(5) + 5, 1).value, None)
    chk("  スロット7（DOE 4/3 に対し 4/20 に入力）",
        clin.cell(clin_base(7) + 5, 1).value, "⚠割当確認")

    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


if __name__ == "__main__":
    if sys.argv[1] == "mutate":
        mutate(sys.argv[2], sys.argv[3])
    else:
        sys.exit(1 if verify(sys.argv[2]) else 0)
