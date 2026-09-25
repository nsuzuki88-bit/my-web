# -*- coding: utf-8 -*-
"""CLABSI判定に血液培養陽性の例を入れて動作を確かめる

  実データ（2026年度ファイル）のコピーに、架空の血液培養陽性を入れて判定を確かめる。
  ※架空の所見を実在の患者に付けたファイルになるため、テスト用のコピーは配布しない。

    python3 demo_clabsi.py real  <v3.xlsm> <テスト用コピー.xlsm>
    python3 demo_clabsi.py check <再計算済み(strip_charts後)> <テスト用コピー.xlsm>

  空テンプレートに架空の患者（検証C01〜）と例を入れた、Excelで確かめる用のデモも作る。

    python3 demo_clabsi.py demo <build.pyの出力(空テンプレート)> <デモ.xlsx>
    python3 demo_clabsi.py check-demo <再計算済み(strip_charts後)> <デモ.xlsx>
"""
import datetime as dt
import sys, openpyxl
import common
import ref_clabsi as R
import test_clabsi as TC
from build import probe

D = lambda m, d: dt.date(2026, m, d)
COLS = dict(pid=2, culture=3, bact_name=4, kind=5, symptom=6, secondary=7, csep=8, note=16)
B1, B2, B3, B4, B5 = ("認定病原体", "皮膚常在菌（2セット以上）", "皮膚常在菌（1セットのみ）",
                      "陰性・未実施", "除外病原体")

# 実データの患者（ブロック番号で指定し、患者IDはAllシートから引く）に入れる架空の例と、
# JHAISの規則から導いた期待値 (基準, DOE, ICU在室日, CL留置日数, 判定, 期待するブロック)
REAL = [
    ("CL抜去の翌日（CL 4/6〜4/12の7日間、4/13はC無）", dict(block=21, culture=D(4, 13), kind=B1, bact_name="黄色ブドウ球菌"),
     ("LCBI-1", D(4, 13), 8, 7, "CLABSI", 21)),
    ("ICU退室の翌日（CL 4/1〜4/10、4/11は空欄）", dict(block=2, culture=D(4, 11), kind=B1, bact_name="大腸菌"),
     ("LCBI-1", D(4, 11), 11, 10, "CLABSI", 2)),
    ("ICU入室2日目（4/20入室）", dict(block=52, culture=D(4, 21), kind=B1, bact_name="緑膿菌"),
     ("LCBI-1", D(4, 21), 2, 2, "入室2日以内（POA・転棟元に帰属）", 52)),
    ("同じIDが2ブロック（ブロック10と29）＋中心ラインの入れ直し（4/14〜16はC無→4/17から再挿入）",
     dict(pid_of=29, culture=D(4, 18), kind=B1, bact_name="肺炎桿菌"),
     ("LCBI-1", D(4, 18), 9, 2, "BSI（CL非関連）", 29)),
    ("中心ラインのない患者（同じIDが3ブロック）", dict(pid_of=33, culture=D(4, 14), kind=B1, bact_name="腸球菌"),
     ("LCBI-1", D(4, 14), 3, 0, "BSI（CL非関連）", 33)),
    ("長期留置のCL13日目", dict(block=179, culture=D(6, 20), kind=B1, bact_name="カンジダ"),
     ("LCBI-1", D(6, 20), None, 13, "CLABSI", 179)),
    ("同じ患者の7日後の再陽性（RIT 6/20〜7/3の中）", dict(block=179, culture=D(6, 27), kind=B1, bact_name="カンジダ"),
     ("LCBI-1", D(6, 27), None, 20, "RIT内（前回のBSIに菌を追加）", 179)),
    ("同じ患者のRIT明け（7/10）", dict(block=179, culture=D(7, 10), kind=B1, bact_name="緑膿菌"),
     ("LCBI-1", D(7, 10), None, 33, "CLABSI", 179)),
    ("LCBI-2：皮膚常在菌2セット＋採取2日前の発熱", dict(block=178, culture=D(6, 15), kind=B2, symptom=D(6, 13),
                                                          bact_name="表皮ブドウ球菌"),
     ("LCBI-2", D(6, 13), None, 7, "CLABSI", 178)),
    ("LCBI-2：発熱が採取5日前（IWPの外）", dict(block=182, culture=D(6, 15), kind=B2, symptom=D(6, 10),
                                                 bact_name="表皮ブドウ球菌"),
     ("非該当（症状がIWP外）", None, None, None, "非該当（症状がIWP外）", 182)),
    ("皮膚常在菌が1セットだけ", dict(block=248, culture=D(7, 15), kind=B3, bact_name="CoNS"),
     ("非該当（汚染菌）", None, None, None, "非該当（汚染菌）", 248)),
    ("他部位感染（例：尿路感染）の2次性BSI", dict(block=265, culture=D(7, 15), kind=B1, secondary="はい",
                                                  bact_name="大腸菌"),
     ("LCBI-1", D(7, 15), None, None, "2次性BSI（CLABSIではない）", 265)),
    ("CSEP：血培陰性＋医師が敗血症として治療", dict(block=323, culture=D(8, 10), kind=B4, csep="はい",
                                                   symptom=D(8, 10), bact_name="（陰性）"),
     ("CSEP", D(8, 10), None, 7, "CLABSI", 323)),
    ("9月のCLABSI（CL8日目）", dict(block=383, culture=D(9, 6), kind=B1, bact_name="黄色ブドウ球菌"),
     ("LCBI-1", D(9, 6), None, 8, "CLABSI", 383)),
]


def _setup(path):
    wb = openpyxl.load_workbook(path, keep_vba=path.endswith(".xlsm"))
    g = probe(wb["All"])
    g["NBLK"] = min(g["NBLK"], common.NBLK_MAX)
    common.configure(**g)
    return wb, g


def real(src, dst):
    wb, g = _setup(src)
    al, cs = wb["All"], wb["CLABSI判定"]
    for i, (desc, ev, _w) in enumerate(REAL):
        r = common.CL_EV_TOP + i
        k = ev.get("block") or ev.get("pid_of")
        pid = al.cell(common.all_id(k), 1).value
        assert pid is not None, k
        cs.cell(r, COLS["pid"], pid)
        for key in ("culture", "bact_name", "kind", "symptom", "secondary", "csep"):
            if key in ev:
                cs.cell(r, COLS[key], ev[key])
        cs.cell(r, COLS["note"], "【テスト】" + desc)
    wb.save(dst)
    print("wrote", dst, f"（架空の血液培養 {len(REAL)}行）")


def _events_from(path):
    """ブックのセクションAの入力を読む（独立実装に渡す形）"""
    wb = openpyxl.load_workbook(path, keep_vba=path.endswith(".xlsm"), read_only=True)
    cs = wb["CLABSI判定"]
    out = []
    for r in range(common.CL_EV_TOP, common.CL_EV_BOT + 1):
        ev = {}
        for key, c in COLS.items():
            v = cs.cell(r, c).value
            if v is None or key == "note":
                continue
            ev[key] = v.date() if isinstance(v, dt.datetime) else v
        out.append(ev)
    return out


def check(calc, src, expect):
    ok = ng = 0

    def chk(label, got, want):
        nonlocal ok, ng
        g_ = got.date() if isinstance(got, dt.datetime) else got
        same = g_ == want or (g_ in (None, "") and want in (None, ""))
        if not same and isinstance(g_, float) and isinstance(want, (int, float)):
            same = abs(g_ - want) < 1e-9
        if same:
            ok += 1
        else:
            ng += 1
            print(f"  NG  {label}: got {got!r} want {want!r}")

    dates, blks, g = R.blocks(src)
    events = _events_from(src)
    ref = R.judge_events(dates, blks, events)
    wb = openpyxl.load_workbook(calc, data_only=True)
    cs = wb["CLABSI判定"]
    print(f"{'行':>2} {'場面':<44} {'ブロック':>5} {'基準':<16} {'DOE':<10} {'ICU':>4} {'CL':>3}  判定")
    for i, (desc, _ev, want) in enumerate(expect):
        r = common.CL_EV_TOP + i
        crit, doe, icu, cl, judge, blk = want
        got = {c: cs.cell(r, c).value for c in (9, 11, 12, 13, 14, 15)}
        if blk is not None:
            chk(f"行{i+1} ブロック", got[9], blk)
        chk(f"行{i+1} 基準", got[11], crit)
        chk(f"行{i+1} DOE", got[12], doe)
        if icu is not None:
            chk(f"行{i+1} ICU在室日", got[13], icu)
        if cl is not None:
            chk(f"行{i+1} CL留置日数", got[14], cl)
        chk(f"行{i+1} 判定", got[15], judge)
        # 独立実装
        rf = ref[i]
        chk(f"行{i+1} 判定（独立実装）", got[15], rf["judge"])
        chk(f"行{i+1} ブロック（独立実装）", got[9], rf["block"])
        if rf["crit"] in R.CRIT and rf["icu_day"] is not None:
            chk(f"行{i+1} ICU在室日（独立実装）", got[13], rf["icu_day"])
            chk(f"行{i+1} CL留置日数（独立実装）", got[14], rf["cl_day"])
        d = got[12].strftime("%m/%d") if isinstance(got[12], dt.datetime) else "―"
        print(f"{i+1:>2} {desc[:44]:<44} {got[9]!s:>5} {got[11]!s:<16} {d:<10} "
              f"{got[13] if got[13] not in (None, '') else '―'!s:>4} "
              f"{got[14] if got[14] not in (None, '') else '―'!s:>3}  {got[15]}")

    # セクションB（患者一覧）の「血流感染の判定」
    print("\n患者一覧（セクションB）の「血流感染の判定」")
    want_k = {}
    for rf in ref:
        if rf["block"] is None:
            continue
        want_k.setdefault(rf["block"], []).append(rf["judge"])
    rows = list(cs.iter_rows(min_row=common.CL_B_TOP, max_row=common.CL_B_BOT,
                             min_col=4, max_col=11, values_only=True))
    shown = 0
    for row in rows:
        k, label = row[0], row[7]
        if k in (None, ""):
            continue
        js = want_k.get(k)
        n = sum(1 for j in (js or []) if j == "CLABSI")
        want = "" if not js else (f"CLABSI {n}件" if n else "評価済（CLABSIなし）")
        chk(f"一覧 ブロック{k}", label, want)
        if js:
            shown += 1
            print(f"  ブロック{k:>4}: {label}")
    # 中心ラインのない患者は一覧に出ない
    listed = {row[0] for row in rows if row[0] not in (None, "")}
    for k in want_k:
        if not any(x == "C有" for x in blks[k - 1]["c"]):
            chk(f"中心ラインのないブロック{k}は一覧に出ない", k in listed, False)

    # 集計
    m = R.monthly(dates, blks, ref)
    cz, yz, hz = wb["CLABSI集計"], wb["年間集計(CLABSI)"], wb["上半期 (CLABSI)"]
    print("\nCLABSI集計（月別）")
    for i in range(12):
        c = i + 2
        cdays, lc, cs_ = m["cl_days"][i], m["lcbi"][i], m["csep"][i]
        chk(f"{m['months'][i][1]}月 LCBI", cz.cell(8, c).value, lc)
        chk(f"{m['months'][i][1]}月 CSEP", cz.cell(9, c).value, cs_)
        chk(f"{m['months'][i][1]}月 合計", cz.cell(10, c).value, lc + cs_)
        chk(f"{m['months'][i][1]}月 発生率", cz.cell(11, c).value,
            (lc + cs_) / cdays * 1000 if cdays else None)
        chk(f"{m['months'][i][1]}月 年間集計(CLABSI)の件数", yz.cell(7, c).value, lc + cs_)
        if i < 6:
            chk(f"{m['months'][i][1]}月 上半期 (CLABSI)の件数", hz.cell(7, c).value, lc + cs_)
        if cdays:
            rate = cz.cell(11, c).value
            print(f"  {m['months'][i][1]:>2}月: 中心ライン日 {cdays:4d}  LCBI {lc}  CSEP {cs_}  "
                  f"→ 発生率 {rate:.1f} /1,000中心ライン日")
    tl, ts, tc = sum(m["lcbi"]), sum(m["csep"]), sum(m["cl_days"])
    chk("年度計 件数", cz.cell(10, 14).value, tl + ts)
    chk("年度計 発生率（LCBI＋CSEP）", cz.cell(11, 14).value, (tl + ts) / tc * 1000)
    chk("年度計 発生率（LCBIのみ）", cz.cell(12, 14).value, tl / tc * 1000)
    print(f"  年度計: LCBI {tl}件＋CSEP {ts}件 ／ 中心ライン日 {tc} → "
          f"{cz.cell(11, 14).value:.1f}（LCBIのみ {cz.cell(12, 14).value:.1f}） /1,000中心ライン日")
    chk("見出し行の件数", cs["A4"].value.split("CLABSI ")[1].split("件")[0], str(tl + ts))
    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


# ---- 空テンプレートに架空の患者で作るデモ（Excelで確かめる用）--------------
DEMO = [
    ("CL3日目の陽性（最も早くCLABSIになる日）", 0),
    ("CL抜去の翌日（CL3日間で抜去）", 2),
    ("CLが2日で抜去（対象のCLではない）", 3),
    ("1暦日CLなし→再挿入2日目（数え直し）", 4),
    ("同じ患者の11日後（RIT内）", 5),
    ("RIT明け（CL18日目）", 6),
    ("LCBI-2：採取2日前の発熱（DOE＝発熱日）", 7),
    ("LCBI-2：発熱が採取5日前（IWPの外）", 8),
    ("皮膚常在菌1セットのみ", 10),
    ("2次性BSI", 11),
    ("CSEP", 12),
    ("ICU退室の翌日（ICUに帰属）", 14),
    ("ICU退室の翌々日（ICU外）", 15),
    ("再入室：同じIDの2つ目のブロック", 16),
    ("ICU入室2日目", 1),
]


def demo(src, dst):
    """test_clabsi の架空の患者と例のうち、説明に向くものを備考付きで入れる"""
    TC.make(src, dst)
    wb = openpyxl.load_workbook(dst)
    cs = wb["CLABSI判定"]
    for r in range(common.CL_EV_TOP, common.CL_EV_BOT + 1):   # いったん空にして並べ直す
        for c in COLS.values():
            cs.cell(r, c).value = None
    names = {B1: ["黄色ブドウ球菌", "大腸菌", "緑膿菌", "肺炎桿菌", "カンジダ", "腸球菌"],
             B2: ["表皮ブドウ球菌"], B3: ["CoNS"], B4: ["（陰性）"], B5: ["―"]}
    used = {}
    for i, (note, j) in enumerate(DEMO):
        ev = TC.EVENTS[j][0]
        r = common.CL_EV_TOP + i
        for key, c in COLS.items():
            if key in ev:
                cs.cell(r, c, ev[key])
        kind = ev.get("kind")
        if kind in names:                            # 検出菌名（例）
            n = used.get(kind, 0)
            cs.cell(r, COLS["bact_name"], names[kind][n % len(names[kind])])
            used[kind] = n + 1
        cs.cell(r, COLS["note"], "【例】" + note)
    wb.save(dst)
    print("wrote", dst, f"（架空の患者{len(TC.PATIENTS)}ブロック・例{len(DEMO)}行）")


def demo_expect():
    return [(note, None, TC.EVENTS[j][1] + (None,)) for note, j in DEMO]


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "real":
        real(sys.argv[2], sys.argv[3])
    elif cmd == "check":
        _setup(sys.argv[3])
        sys.exit(1 if check(sys.argv[2], sys.argv[3], REAL) else 0)
    elif cmd == "demo":
        _setup(sys.argv[2])
        demo(sys.argv[2], sys.argv[3])
    elif cmd == "check-demo":
        _setup(sys.argv[3])
        sys.exit(1 if check(sys.argv[2], sys.argv[3], demo_expect()) else 0)
