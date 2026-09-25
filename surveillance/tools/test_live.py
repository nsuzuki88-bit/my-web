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
    """スロットの並び（DOE①の早い順、同じ日はブロック番号順）
      変更前：1:b4 2:b5 3:b10 4:b11 5:b12 6:b3 7:b8 8:b9 9:b1
      変更後：1:b4 2:b5 3:b11 4:b12 5:b20 6:b3 7:b8 8:b9 9:b1 10:b21
    """
    wb = openpyxl.load_workbook(src)
    al, cl = wb["All"], wb["臨床所見"]

    # ① ブロック10：FiO2の上昇を 20→5 ポイントに下げる → VACでなくなる
    for d in (3, 4):
        al.cell(all_f(10), C0 + d - 1, 45)

    # ② ブロック20に新規患者を追加（DOE 4/3）→ 同じ日のVACの中でブロック番号順に入る
    al.cell(all_id(20), 1, "T20")
    al.cell(all_name(20), 1, "検証20")
    for d, (f, p) in enumerate([(40, 5), (40, 5), (70, 5), (70, 5)]):
        al.cell(all_v(20), C0 + d, "V有")
        al.cell(all_f(20), C0 + d, f)
        al.cell(all_p(20), C0 + d, p)

    # ②-2 ブロック21に5月の患者を追加（DOE 5/3）→ 末尾に付き、入力済みのスロットは動かない
    al.cell(all_id(21), 1, "T21")
    al.cell(all_name(21), 1, "検証21")
    for d, (f, p) in enumerate([(40, 5), (40, 5), (70, 5), (70, 5)]):
        al.cell(all_v(21), C0 + 30 + d, "V有")
        al.cell(all_f(21), C0 + 30 + d, f)
        al.cell(all_p(21), C0 + 30 + d, p)

    # ③ ブロック9（スロット8）のPVAP基準を削除 → PVAP から IVAC へ
    cl.cell(clin_base(8) + 6, C0 + 5 - 1).value = None

    # ④ スロット4（変更後はブロック12・DOE 4/3）にDOEから離れた日付で入力 → 警告が出るはず
    cl.cell(clin_base(4) + 0, C0 + 20 - 1, 38.2)

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

    print("\n② ブロック20・21に患者を追加")
    chk("  ブロック20 判定", Bk(20, 9), "VAC")
    chk("  ブロック20 DOE①", Bk(20, 7), D(4, 3))
    chk("  ブロック21 判定", Bk(21, 9), "VAC")
    chk("  ブロック21 DOE①", Bk(21, 7), D(5, 3))

    print("\n VAC患者の並び（ブロック番号）＝DOE①の早い順、同じ日はブロック番号順")
    got = [A(n, 2) for n in range(1, NCLIN + 1) if A(n, 2) is not None]
    chk("  並び順", got, [4, 5, 11, 12, 20, 3, 8, 9, 1, 21])
    chk("  5月のVAC（ブロック21）は末尾", A(10, 2), 21)
    chk("  臨床所見を入力済みのスロット7・8は動かない（ブロック8・9）", [A(7, 2), A(8, 2)], [8, 9])

    print("\n③ ブロック9（スロット8）のPVAP基準を削除 → IVACへ格下げ")
    chk("  判定①", A(8, 9), "IVAC")
    chk("  ブロック番号", A(8, 2), 9)
    chk("  ブロック8（スロット7）はIVACのまま", A(7, 9), "IVAC")

    print("\n 臨床所見に入力がある患者と判定シートの割当（スロット順）")
    chk("  スロット4(ブロック12) 臨床所見", A(4, 12), "入力あり")
    chk("  スロット4 判定シート番号", A(4, 13), 1)
    chk("  スロット7(ブロック8) 判定シート番号", A(7, 13), 2)
    chk("  スロット8(ブロック9) 判定シート番号", A(8, 13), 3)
    chk("  スロット5(ブロック20) 臨床所見", A(5, 12), "未入力")
    chk("  VAE-01 のブロック", wb["VAE-01"]["C4"].value, 12)
    chk("  VAE-02 のブロック", wb["VAE-02"]["C4"].value, 8)
    chk("  VAE-03 のブロック", wb["VAE-03"]["C4"].value, 9)
    chk("  VAE-04 は未割当", wb["VAE-04"]["C4"].value, None)
    chk("  VAE-01（DOEから離れた日の体温だけ）はVACのまま", A(4, 9), "VAC")

    print("\n④ DOEから離れた日に入力したスロット4 → 警告が出る")
    chk("  スロット7（正常）", clin.cell(clin_base(7) + 5, 1).value, None)
    chk("  スロット4（DOE 4/3 に対し 4/20 に入力）",
        clin.cell(clin_base(4) + 5, 1).value, "⚠割当確認")

    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


if __name__ == "__main__":
    if sys.argv[1] == "mutate":
        mutate(sys.argv[2], sys.argv[3])
    else:
        sys.exit(1 if verify(sys.argv[2]) else 0)
