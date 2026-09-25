# -*- coding: utf-8 -*-
"""空テンプレート版の検証。

check2026.main() で Python独立実装との全面照合を行ったうえで、
教科書例ごとの期待値（DOEのMV日・最終判定）を追加で確認する。
"""
import datetime as dt
import sys, openpyxl
import ref_vae as R
import check2026
from common import LIST_A_TOP, LIST_B_TOP, NBLK

D = lambda m, d: dt.datetime(2025, m, d)

# ブロック -> (MV日数, DOEの日付, 期待する最終判定)
EXPECT = {
    1:  (6,  [D(4, 5)],            "VAC"),   # PEEP経路
    2:  (6,  [],                   None),    # PEEP 0〜5は同等
    3:  (6,  [D(4, 4)],            "VAC"),   # FiO2経路
    4:  (4,  [D(4, 3)],            "VAC"),   # 最早DOEはMV3日目
    5:  (20, [D(4, 3), D(4, 17)],  "VAC"),   # イベント期間14日
    6:  (13, [],                   None),    # エピソード分割・悪化なし
    7:  (0,  [],                   None),    # NIVは対象外
    8:  (6,  [D(4, 4)],            "IVAC"),  # 体温異常＋新規抗菌薬
    9:  (6,  [D(4, 4)],            "PVAP"),  # ＋PVAP基準1
    10: (4,  [D(4, 3)],            "VAC"),   # FiO2ちょうど+20ポイント（％）
    11: (4,  [D(4, 3)],            "VAC"),   # FiO2ちょうど+0.20（小数）
    12: (4,  [D(4, 3)],            "VAC"),   # PEEPちょうど+3cmH2O
    13: (0,  [],                   None),    # V無の日はMV日にしない
}


def _vrows(src):
    """各ブロックの (日付, V有無) の並び"""
    from build import probe
    wb = openpyxl.load_workbook(src)
    al = wb["All"]
    g = probe(al)
    out = []
    for k in range(1, g["NBLK"] + 1):
        v = g["BLOCK0"] + g["BSTEP"] * (k - 1)
        out.append([(al.cell(g["DATE_ROW"], c).value, al.cell(v, c).value)
                    for c in range(g["C0"], g["CN"] + 1) if al.cell(v, c).value is not None])
    return out


def main(calc, src):
    ng = check2026.main(calc, src)
    ok_extra = ng_extra = 0

    def chk(label, got, want):
        nonlocal ok_extra, ng_extra
        g = got.date() if isinstance(got, dt.datetime) else got
        w = want.date() if isinstance(want, dt.datetime) else want
        same = g == w or (g in (None, "") and w in (None, ""))
        if not same and isinstance(g, float) and isinstance(w, (int, float)):
            same = abs(g - w) < 1e-9
        if same:
            ok_extra += 1
        else:
            ng_extra += 1
            print(f"  NG  {label}: got {got!r} want {want!r}")

    wb = openpyxl.load_workbook(calc, data_only=True)
    lst = wb["VAE対象一覧"]
    _, ref = R.analyse_global(src)
    vac = R.vac_patients(ref)
    rank = {r["k"]: i for i, r in enumerate(vac, 1)}

    print("\n=== 教科書例ごとの期待値 ===")
    for k, (mvdays, does, judge) in EXPECT.items():
        rb = LIST_B_TOP + k - 1
        chk(f"ブロック{k} MV日数", lst.cell(rb, 4).value or 0, mvdays)
        chk(f"ブロック{k} DOE①", lst.cell(rb, 7).value, does[0] if does else None)
        chk(f"ブロック{k} DOE②", lst.cell(rb, 8).value, does[1] if len(does) > 1 else None)
        n = rank.get(k)
        got_judge = lst.cell(LIST_A_TOP + n - 1, 9).value if n else None
        chk(f"ブロック{k} 最終判定", got_judge, judge)
        mark = "VAC" if does else "―"
        print(f"  ブロック{k:2d}: MV{mvdays:2d}日 "
              f"DOE={[d.strftime('%m/%d') for d in does] or '―'} "
              f"判定={got_judge or '―'}  (期待 {judge or '―'})")

    print("\n=== VAE集計（4月）===")
    # VACはブロック1,3,4,5(2件),10,11,12 の計8件（8と9はIVAC/PVAPへ格上げ）
    vz = wb["VAE集計"]
    chk("  VAC件数", vz.cell(8, 2).value, 8)
    chk("  IVAC件数", vz.cell(9, 2).value, 1)
    chk("  PVAP件数", vz.cell(10, 2).value, 1)
    chk("  VAE合計", vz.cell(11, 2).value, 10)
    # 分母：人工呼吸器使用患者数（V有）とICU入室患者数（C有＋C無）の4月の合計
    v_days = sum(1 for b in _vrows(src) for d, x in b if x == "V有" and d.month == 4)
    chk("  延べ人工呼吸器使用日数", vz.cell(6, 2).value, v_days)
    chk("  VAC発生率", vz.cell(12, 2).value, 8 / v_days * 1000 if v_days else None)
    chk("  VAE発生率 合計", vz.cell(15, 2).value, 10 / v_days * 1000 if v_days else None)
    chk("  IVAC-plus発生率", vz.cell(16, 2).value, 2 / v_days * 1000 if v_days else None)
    yz = wb["年間集計(VAE)"]
    chk("  年間集計(VAE) 4月 VAC件数", yz["B7"].value, 8)
    chk("  年間集計(VAE) 4月 IVAC件数", yz["B8"].value, 1)
    chk("  年間集計(VAE) 4月 PVAP件数", yz["B9"].value, 1)
    print(f"  延べ人工呼吸器使用日数 {v_days}日 → VAE合計10件 ＝ "
          f"{vz.cell(15, 2).value} /1,000人工呼吸器日")

    print(f"\n==== 追加確認 ok={ok_extra}  NG={ng_extra} ====")
    return ng + ng_extra


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1], sys.argv[2]) else 0)
