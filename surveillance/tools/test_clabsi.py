# -*- coding: utf-8 -*-
"""CLABSI判定シート・集計シートの検証

  python3 test_clabsi.py make  <build出力(空テンプレート)> <テスト用ブック>
  python3 test_clabsi.py check <再計算済み(strip_charts後)> <テスト用ブック>

JHAISマニュアルの中心ラインの例（抜去日・抜去翌日・1暦日の空白でリセット・同日入替え）と、
POA／転棟ルール、ICU退室翌日、RIT、LCBI-2のIWP、CSEP、2次性BSI、汚染菌、再入室（同じID）を
Allシートと CLABSI判定セクションA に入れ、
  ① 手で決めた期待値（JHAISの規則から導いたもの）
  ② ref_clabsi.py（独立実装）
の両方と照合する。
"""
import datetime as dt
import sys, openpyxl
import ref_clabsi as R
from build import probe
import common

Y = 2025
D = lambda m, d: dt.date(Y + (1 if m <= 3 else 0), m, d)

# (ブロック, 患者ID, [(開始, 終了, "C有"/"C無"), ...])
PATIENTS = [
    (1, "C01", [(D(4, 10), D(4, 11), "C無"), (D(4, 12), D(4, 25), "C有"), (D(4, 26), D(4, 30), "C無")]),
    (2, "C02", [(D(4, 20), D(5, 10), "C有")]),
    (3, "C03", [(D(5, 1), D(5, 3), "C有"), (D(5, 4), D(5, 10), "C無")]),
    (4, "C04", [(D(5, 1), D(5, 2), "C有"), (D(5, 3), D(5, 10), "C無")]),
    (5, "C05", [(D(5, 10), D(5, 14), "C有"), (D(5, 15), D(5, 15), "C無"), (D(5, 16), D(6, 10), "C有")]),
    (6, "C06", [(D(6, 1), D(6, 20), "C有")]),
    (7, "C07", [(D(6, 1), D(6, 20), "C有")]),
    (8, "C08", [(D(6, 1), D(6, 20), "C有")]),
    (9, "C09", [(D(7, 1), D(7, 20), "C有")]),
    (10, "C10", [(D(7, 1), D(7, 20), "C有")]),
    (11, "C11", [(D(7, 1), D(7, 20), "C有")]),
    (12, "C12", [(D(8, 1), D(8, 10), "C有")]),
    (13, "C13", [(D(8, 1), D(8, 10), "C有")]),
    (14, "C14", [(D(8, 1), D(8, 5), "C有")]),                                  # 1回目の入室
    (15, "C14", [(D(9, 1), D(9, 2), "C無"), (D(9, 3), D(9, 15), "C有")]),       # 再入室（同じID）
    (16, "C16", [(D(10, 1), D(10, 10), "C無")]),                                # 中心ラインなし
    (17, "C17", [(D(10, 1), D(10, 10), "C有")]),
    (18, "C18", [(D(4, 1), D(4, 5), "C有")]),                                   # 年度初日から
    (19, "C19", [(D(4, 1), D(4, 5), "C有")]),
    (20, "C20", [(D(11, 1), D(11, 20), "C有")]),
    (21, "C21", [(D(3, 25), D(3, 31), "C有")]),                                  # 年度末
]

# BSIイベント（CLABSI判定セクションAの行）と、JHAISの規則から導いた期待値
#   期待値：(基準, DOE, ICU在室日, CL留置日数, 判定)
EVENTS = [
    (dict(pid="C01", culture=D(4, 14), kind="認定病原体"),
     ("LCBI-1", D(4, 14), 5, 3, "CLABSI")),                       # CL3日目
    (dict(pid="C02", culture=D(4, 21), kind="認定病原体"),
     ("LCBI-1", D(4, 21), 2, 2, "入室2日以内（POA・転棟元に帰属）")),
    (dict(pid="C03", culture=D(5, 4), kind="認定病原体"),
     ("LCBI-1", D(5, 4), 4, 3, "CLABSI")),                        # 抜去翌日（CL3日で抜去）
    (dict(pid="C04", culture=D(5, 3), kind="認定病原体"),
     ("LCBI-1", D(5, 3), 3, 2, "BSI（CL非関連）")),              # 2日で抜去→対象のCLではない
    (dict(pid="C05", culture=D(5, 17), kind="認定病原体"),
     ("LCBI-1", D(5, 17), 8, 2, "BSI（CL非関連）")),             # 1暦日の空白で数え直し（患者D）
    (dict(pid="C05", culture=D(5, 28), kind="認定病原体"),
     ("LCBI-1", D(5, 28), 19, 13, "RIT内（前回のBSIに菌を追加）")),
    (dict(pid="C05", culture=D(6, 2), kind="認定病原体"),
     ("LCBI-1", D(6, 2), 24, 18, "CLABSI")),                      # RIT(5/17〜5/30)の後
    (dict(pid="C06", culture=D(6, 10), kind="皮膚常在菌（2セット以上）", symptom=D(6, 8)),
     ("LCBI-2", D(6, 8), 8, 8, "CLABSI")),                        # DOE＝症状の日（IWP内）
    (dict(pid="C07", culture=D(6, 10), kind="皮膚常在菌（2セット以上）", symptom=D(6, 5)),
     ("非該当（症状がIWP外）", None, None, None, "非該当（症状がIWP外）")),
    (dict(pid="C07", culture=D(6, 12), kind="皮膚常在菌（2セット以上）"),
     ("非該当（症状なし）", None, None, None, "非該当（症状なし）")),
    (dict(pid="C08", culture=D(6, 10), kind="皮膚常在菌（1セットのみ）"),
     ("非該当（汚染菌）", None, None, None, "非該当（汚染菌）")),
    (dict(pid="C09", culture=D(7, 10), kind="認定病原体", secondary="はい"),
     ("LCBI-1", D(7, 10), 10, 10, "2次性BSI（CLABSIではない）")),
    (dict(pid="C10", kind="陰性・未実施", csep="はい", symptom=D(7, 10)),
     ("CSEP", D(7, 10), 10, 10, "CLABSI")),
    (dict(pid="C11", culture=D(7, 10), kind="陰性・未実施", csep="いいえ", symptom=D(7, 10)),
     ("非該当（CSEP基準を満たさない）", None, None, None, "非該当（CSEP基準を満たさない）")),
    (dict(pid="C12", culture=D(8, 11), kind="認定病原体"),
     ("LCBI-1", D(8, 11), 11, 10, "CLABSI")),                     # ICU退室翌日→ICUに帰属
    (dict(pid="C13", culture=D(8, 13), kind="認定病原体"),
     ("LCBI-1", D(8, 13), 0, 0, "ICU外（ICUのCLABSIではない）")),
    (dict(pid="C14", culture=D(9, 8), kind="認定病原体"),
     ("LCBI-1", D(9, 8), 8, 6, "CLABSI")),                        # 再入室の2つ目のブロックを選ぶ
    (dict(pid="C16", culture=D(10, 5), kind="認定病原体"),
     ("LCBI-1", D(10, 5), 5, 0, "BSI（CL非関連）")),
    (dict(pid="ZZZ", culture=D(10, 5), kind="認定病原体"),
     ("LCBI-1", D(10, 5), None, None, "⚠この患者IDはAllにありません")),
    (dict(pid="C17"),
     ("", None, None, None, "")),                                  # IDだけ入力した途中の行
    (dict(pid="C17", symptom=D(10, 3), kind="認定病原体"),
     ("（採取日を入力）", None, None, None, "（採取日を入力）")),
    (dict(pid="C17", culture=D(10, 3)),
     ("（菌の分類を選択）", None, None, None, "（菌の分類を選択）")),
    (dict(pid="C17", culture=D(10, 3), kind="認定病原体"),
     ("LCBI-1", D(10, 3), 3, 3, "CLABSI")),
    (dict(pid="C18", culture=D(4, 3), kind="認定病原体"),
     ("LCBI-1", D(4, 3), 3, 3, "CLABSI")),                        # 年度初日から3日目
    (dict(pid="C19", culture=D(4, 2), kind="認定病原体"),
     ("LCBI-1", D(4, 2), 2, 2, "入室2日以内（POA・転棟元に帰属）")),
    (dict(pid="C20", culture=D(11, 10), kind="認定病原体"),
     ("LCBI-1", D(11, 10), 10, 10, "CLABSI")),
    (dict(pid="C20", culture=D(11, 10), kind="認定病原体"),
     ("LCBI-1", D(11, 10), 10, 10, "RIT内（前回のBSIに菌を追加）")),  # 同じ日の2本目
    (dict(pid="C21", culture=dt.date(Y + 1, 4, 2), kind="認定病原体"),
     ("LCBI-1", dt.date(Y + 1, 4, 2), None, None, "⚠DOEが年度の範囲外")),
    (dict(culture=D(10, 5), kind="認定病原体"),
     ("LCBI-1", D(10, 5), None, None, "⚠患者IDを入力")),
]
COLS = dict(pid=2, culture=3, bact_name=4, kind=5, symptom=6, secondary=7, csep=8)


def make(src, dst):
    wb = openpyxl.load_workbook(src)
    al, cs = wb["All"], wb["CLABSI判定"]
    g = probe(al)
    common.configure(**{k: g[k] for k in ("C0", "CN", "DATE_ROW", "BLOCK0", "BSTEP")})
    col = {al.cell(g["DATE_ROW"], c).value.date(): c for c in range(g["C0"], g["CN"] + 1)}
    for k, pid, spans in PATIENTS:
        al.cell(common.all_id(k), 1, pid)
        al.cell(common.all_name(k), 1, f"検証{pid}")
        for a, b, mark in spans:
            d = a
            while d <= b:
                al.cell(common.all_c(k), col[d], mark)
                al.cell(common.all_rm(k), col[d], 300 + k)
                d += dt.timedelta(days=1)
    for i, (ev, _want) in enumerate(EVENTS):
        r = common.CL_EV_TOP + i
        for key, c in COLS.items():
            if key in ev:
                cs.cell(r, c, ev[key])
    wb.save(dst)
    print("wrote", dst, f"（患者{len(PATIENTS)}ブロック・BSIイベント{len(EVENTS)}行）")


def check(calc, src):
    ok = ng = 0

    def chk(label, got, want):
        nonlocal ok, ng
        g_ = got.date() if isinstance(got, dt.datetime) else got
        w_ = want.date() if isinstance(want, dt.datetime) else want
        if isinstance(g_, float) and isinstance(w_, (int, float)) and w_ is not None:
            same = abs(g_ - w_) < 1e-9
        else:
            same = g_ == w_ or (g_ in (None, "") and w_ in (None, ""))
        if same:
            ok += 1
        else:
            ng += 1
            print(f"  NG  {label}: got {got!r} want {want!r}")

    dates, blks, g = R.blocks(src)
    wb = openpyxl.load_workbook(calc, data_only=True)
    ws = wb["CLABSI判定"]
    events = []
    for i, (ev, _w) in enumerate(EVENTS):
        events.append(ev)
    ref = R.judge_events(dates, blks, events)

    print("=== セクションA：BSIイベントの判定（手で決めた期待値／独立実装） ===")
    for i, ((ev, want), rf) in enumerate(zip(EVENTS, ref)):
        r = common.CL_EV_TOP + i
        crit, doe, icu, cl, judge = want
        got = [ws.cell(r, c).value for c in (11, 12, 13, 14, 15)]
        lab = f"行{i+1} {ev.get('pid')}"
        # ① 手で決めた期待値
        chk(f"{lab} 基準", got[0], crit)
        chk(f"{lab} DOE", got[1], doe)
        if judge.startswith(("CLABSI", "BSI", "入室", "ICU外", "RIT", "2次性")):
            chk(f"{lab} ICU在室日", got[2], icu)
            chk(f"{lab} CL留置日数", got[3], cl)
        chk(f"{lab} 判定", got[4], judge)
        # ② 独立実装
        chk(f"{lab} 判定（独立実装）", got[4], rf["judge"])
        chk(f"{lab} ブロック（独立実装）", ws.cell(r, 9).value, rf["block"])
        if rf["icu_day"] is not None and rf["crit"] in R.CRIT:
            chk(f"{lab} ICU在室日（独立実装）", got[2], rf["icu_day"])
            chk(f"{lab} CL留置日数（独立実装）", got[3], rf["cl_day"])
        print(f"  {lab:10s} 基準={got[0]!s:28s} DOE={got[1]!s:20s} ICU={got[2]!s:4s} "
              f"CL={got[3]!s:4s} → {got[4]}")

    print("\n=== セクションB：中心ライン（C有）の患者一覧 ===")
    want = R.cl_list(dates, blks)
    for n in range(1, common.NBLK + 1):
        r = common.CL_B_TOP + n - 1
        w = want[n - 1] if n <= len(want) else None
        if w is None:
            chk(f"一覧{n} 空", ws.cell(r, 4).value, None)
            continue
        chk(f"一覧{n} No", ws.cell(r, 1).value, n)
        chk(f"一覧{n} 患者ID", ws.cell(r, 2).value, w["pid"])
        chk(f"一覧{n} 氏名", ws.cell(r, 3).value, w["name"])
        chk(f"一覧{n} ブロック", ws.cell(r, 4).value, w["k"])
        chk(f"一覧{n} 部屋", ws.cell(r, 5).value, w["room"])
        chk(f"一覧{n} ICU入室日", ws.cell(r, 6).value, w["icu_first"])
        chk(f"一覧{n} CL初日", ws.cell(r, 7).value, w["cl_first"])
        chk(f"一覧{n} CL最終日", ws.cell(r, 8).value, w["cl_last"])
        chk(f"一覧{n} CL日数", ws.cell(r, 9).value, w["cl_days"])
        chk(f"一覧{n} CL3日目", ws.cell(r, 10).value, w["cl3"])
    print(f"  中心ライン患者 {len(want)}名（CL初日順）:",
          " ".join(f"{w['pid']}(b{w['k']})" for w in want[:12]), "…" if len(want) > 12 else "")

    print("\n=== CLABSI集計：月別の分母・件数・率 ===")
    m = R.monthly(dates, blks, ref)
    cz = wb["CLABSI集計"]
    R0 = 5
    for i in range(12):
        c = i + 2
        lab = f"{m['months'][i][1]}月"
        pdays, cdays, lc, cs_ = m["patient_days"][i], m["cl_days"][i], m["lcbi"][i], m["csep"][i]
        chk(f"{lab} 延べ入室患者日数", cz.cell(R0, c).value, pdays)
        chk(f"{lab} 延べ中心ライン使用日数", cz.cell(R0 + 1, c).value, cdays)
        chk(f"{lab} 中心ライン使用比", cz.cell(R0 + 2, c).value, cdays / pdays if pdays else None)
        chk(f"{lab} CLABSI(LCBI)", cz.cell(R0 + 3, c).value, lc)
        chk(f"{lab} CLABSI(CSEP)", cz.cell(R0 + 4, c).value, cs_)
        chk(f"{lab} CLABSI合計", cz.cell(R0 + 5, c).value, lc + cs_)
        chk(f"{lab} 発生率(合計)", cz.cell(R0 + 6, c).value,
            (lc + cs_) / cdays * 1000 if cdays else None)
        chk(f"{lab} 発生率(LCBI)", cz.cell(R0 + 7, c).value, lc / cdays * 1000 if cdays else None)
        print(f"  {lab:>4s}: 患者日 {pdays:4d}  CL日 {cdays:4d}  LCBI {lc}  CSEP {cs_}  "
              f"使用比 {cz.cell(R0 + 2, c).value!s:22s} 発生率 {cz.cell(R0 + 6, c).value}")
    tp, tc = sum(m["patient_days"]), sum(m["cl_days"])
    tl, ts = sum(m["lcbi"]), sum(m["csep"])
    chk("年度計 使用比（合計から）", cz.cell(R0 + 2, 14).value, tc / tp)
    chk("年度計 発生率（合計から）", cz.cell(R0 + 6, 14).value, (tl + ts) / tc * 1000)

    print("\n=== 既存の年間集計(CLABSI)・上半期 (CLABSI) へのつながり ===")
    yz = wb["年間集計(CLABSI)"]
    hz = wb["上半期 (CLABSI)"]
    for i in range(12):
        chk(f"年間集計(CLABSI) {i+4 if i < 9 else i-8}月 件数", yz.cell(7, i + 2).value,
            m["lcbi"][i] + m["csep"][i])
    for i in range(6):
        chk(f"上半期 (CLABSI) {i+4}月 件数", hz.cell(7, i + 2).value, m["lcbi"][i] + m["csep"][i])
    chk("年間集計(CLABSI) 題名", yz["A1"].value, "CLABSIサーベイランス")
    chk("年間集計(CLABSI) 年度の使用比", yz["N6"].value, tc / tp)

    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


def check_list(calc, src):
    """実データ用：中心ライン患者の一覧と、月別の分母・使用比を独立実装と照合する（BSIの入力なし）"""
    ok = ng = 0

    def chk(label, got, want):
        nonlocal ok, ng
        g_ = got.date() if isinstance(got, dt.datetime) else got
        w_ = want.date() if isinstance(want, dt.datetime) else want
        same = g_ == w_ or (g_ in (None, "") and w_ in (None, ""))
        if not same and isinstance(g_, float) and isinstance(w_, (int, float)):
            same = abs(g_ - w_) < 1e-9
        if same:
            ok += 1
        else:
            ng += 1
            print(f"  NG  {label}: got {got!r} want {want!r}")

    dates, blks_all, g = R.blocks(src)                 # 分母はAllの全ブロック
    blks = blks_all[:common.NBLK]                       # 一覧はスキャン範囲のブロック
    wb = openpyxl.load_workbook(calc, data_only=True, read_only=True)
    ws = wb["CLABSI判定"]
    want = R.cl_list(dates, blks)
    rows = list(ws.iter_rows(min_row=common.CL_B_TOP, max_row=common.CL_B_BOT,
                             max_col=11, values_only=True))
    for n, row in enumerate(rows, 1):
        w = want[n - 1] if n <= len(want) else None
        if w is None:
            chk(f"一覧{n} 空", row[3], None)
            continue
        for label, got, exp in (("No", row[0], n), ("患者ID", row[1], w["pid"]),
                                ("氏名", row[2], w["name"]), ("ブロック", row[3], w["k"]),
                                ("部屋", row[4], w["room"]), ("ICU入室日", row[5], w["icu_first"]),
                                ("CL初日", row[6], w["cl_first"]), ("CL最終日", row[7], w["cl_last"]),
                                ("CL日数", row[8], w["cl_days"]), ("CL3日目", row[9], w["cl3"]),
                                ("判定", row[10], None)):
            chk(f"一覧{n} {label}", got, exp)
    print(f"  中心ライン（C有）の患者 {len(want)}名を照合（先頭: "
          + " ".join(f"b{w['k']}" for w in want[:8]) + " …）")
    m = R.monthly(dates, blks_all, [])
    cz = wb["CLABSI集計"]
    zr = list(cz.iter_rows(min_row=5, max_row=12, min_col=2, max_col=14, values_only=True))
    for i in range(12):
        pdays, cdays = m["patient_days"][i], m["cl_days"][i]
        lab = f"{m['months'][i][1]}月"
        chk(f"{lab} 延べ入室患者日数", zr[0][i], pdays)
        chk(f"{lab} 延べ中心ライン使用日数", zr[1][i], cdays)
        chk(f"{lab} 中心ライン使用比", zr[2][i], cdays / pdays if pdays else None)
        chk(f"{lab} CLABSI件数", zr[5][i], 0)
        print(f"  {lab:>4s}: 延べ入室患者日数 {pdays:5d}  延べ中心ライン使用日数 {cdays:5d}  "
              f"使用比 {zr[2][i] if zr[2][i] is not None else '':}")
    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "make":
        make(sys.argv[2], sys.argv[3])
    elif cmd == "list":
        wb = openpyxl.load_workbook(sys.argv[3], keep_vba=sys.argv[3].endswith(".xlsm"))
        g = probe(wb["All"])
        g["NBLK"] = min(g["NBLK"], common.NBLK_MAX)
        common.configure(**g)
        sys.exit(1 if check_list(sys.argv[2], sys.argv[3]) else 0)
    else:
        wb = openpyxl.load_workbook(sys.argv[3])
        g = probe(wb["All"])
        common.configure(**g)
        sys.exit(1 if check(sys.argv[2], sys.argv[3]) else 0)
