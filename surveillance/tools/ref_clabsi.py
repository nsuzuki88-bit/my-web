# -*- coding: utf-8 -*-
"""CLABSI判定・集計の独立実装（表計算の数式とは別に、JHAIS Ver.2.5 の規則をそのままPythonで書いたもの）

  blocks(path)            … Allシートの各ブロックの C有無・ID・氏名・部屋
  cl_list(blocks)         … 中心ライン（C有）患者の一覧（CL初日の早い順）
  judge_events(blocks, events, dates) … BSIイベントの判定
  monthly(path, events_judged)        … 月別の分母・件数・率
"""
import datetime as dt
import openpyxl
from build import probe

CRIT = ("LCBI-1", "LCBI-2", "CSEP")
BACT = ["認定病原体", "皮膚常在菌（2セット以上）", "皮膚常在菌（1セットのみ）",
        "陰性・未実施", "除外病原体"]


def _d(v):
    return v.date() if isinstance(v, dt.datetime) else v


def blocks(path, nblk=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    ws = wb["All"]
    wb2 = openpyxl.load_workbook(path)          # probe は通常モードのシートが必要
    g = probe(wb2["All"])
    C0, CN, DR, B0, BS = g["C0"], g["CN"], g["DATE_ROW"], g["BLOCK0"], g["BSTEP"]
    rows = list(ws.iter_rows(min_row=1, max_row=g["ALL_LAST_ROW"], max_col=CN, values_only=True))
    dates = [_d(x) for x in rows[DR - 1][C0 - 1:CN]]
    n = nblk or g["NBLK"]
    out = []
    for k in range(1, n + 1):
        v = B0 + BS * (k - 1)
        crow = [x if x in ("C有", "C無") else None for x in rows[v + 2][C0 - 1:CN]]
        room = rows[v + 3][C0 - 1:CN]
        out.append(dict(k=k, pid=rows[v][0], name=rows[v + 2][0], c=crow, room=room))
    return dates, out, g


def _run(vals, p, ok):
    """p（0始まり）で終わる、ok を満たす連続日数"""
    n = 0
    while p >= 0 and ok(vals[p]):
        n += 1
        p -= 1
    return n


def block_summary(dates, b):
    c = b["c"]
    cl = [i for i, x in enumerate(c) if x == "C有"]
    icu = [i for i, x in enumerate(c) if x in ("C有", "C無")]
    s = dict(k=b["k"], pid=b["pid"], name=b["name"], cl_days=len(cl))
    s["cl_first"] = dates[cl[0]] if cl else None
    s["cl_last"] = dates[cl[-1]] if cl else None
    s["icu_first"] = dates[icu[0]] if icu else None
    s["icu_last"] = dates[icu[-1]] if icu else None
    s["room"] = b["room"][cl[0]] if cl else None
    # 3日連続の「C有」がそろった最初の日（CL3日目）
    s["cl3"] = next((dates[i] for i in range(2, len(c))
                     if c[i] == c[i - 1] == c[i - 2] == "C有"), None)
    return s


def cl_list(dates, blks):
    rows = [block_summary(dates, b) for b in blks]
    rows = [r for r in rows if r["cl_days"] > 0]
    return sorted(rows, key=lambda r: (r["cl_first"], r["k"]))


def criterion(ev):
    cdate, fdate, kind, csep = ev.get("culture"), ev.get("symptom"), ev.get("kind"), ev.get("csep")
    if cdate is None and fdate is None:
        return ""
    if kind == BACT[0]:
        return "（採取日を入力）" if cdate is None else "LCBI-1"
    if kind == BACT[1]:
        if cdate is None:
            return "（採取日を入力）"
        if fdate is None:
            return "非該当（症状なし）"
        return "LCBI-2" if abs((fdate - cdate).days) <= 3 else "非該当（症状がIWP外）"
    if kind == BACT[2]:
        return "非該当（汚染菌）"
    if kind == BACT[3]:
        if csep != "はい":
            return "非該当（CSEP基準を満たさない）"
        return "（症状の初日を入力）" if fdate is None else "CSEP"
    if kind == BACT[4]:
        return "非該当（LCBI除外病原体）"
    return "（菌の分類を選択）"


def judge_events(dates, blks, events):
    summ = {b["k"]: block_summary(dates, b) for b in blks}
    pos = {d: i for i, d in enumerate(dates)}
    out = []
    for i, ev in enumerate(events):
        r = dict(ev)
        crit = criterion(ev)
        r["crit"] = crit
        if crit == "LCBI-1":
            doe = ev["culture"]
        elif crit == "LCBI-2":
            doe = min(ev["culture"], ev["symptom"])
        elif crit == "CSEP":
            doe = ev["symptom"]
        else:
            doe = None
        r["doe"] = doe
        # 患者IDのブロック：DOEが在室期間（入室日〜退室翌日）に入るブロック。なければ最初のブロック
        pid = ev.get("pid")
        cands = [b for b in blks if pid is not None and b["pid"] is not None
                 and str(b["pid"]) == str(pid)]
        blk = None
        if doe is not None:
            for b in cands:
                s = summ[b["k"]]
                if s["icu_first"] and s["icu_first"] <= doe <= s["icu_last"] + dt.timedelta(days=1):
                    blk = b
                    break
        if blk is None and cands:
            blk = cands[0]
        r["block"] = blk["k"] if blk else None
        r["name"] = blk["name"] if blk else None
        # ICU在室日・CL留置日数
        p = pos.get(doe) if doe is not None else None
        r["p"] = p
        icu_day = cl_day = None
        if blk and p is not None:
            c = blk["c"]
            in_icu = lambda x: x in ("C有", "C無")
            if in_icu(c[p]):
                icu_day = _run(c, p, in_icu)
            elif p >= 1 and in_icu(c[p - 1]):
                icu_day = _run(c, p - 1, in_icu) + 1
            else:
                icu_day = 0
            is_cl = lambda x: x == "C有"
            if c[p] == "C有":
                cl_day = _run(c, p, is_cl)
            elif p >= 1 and c[p - 1] == "C有":
                cl_day = _run(c, p - 1, is_cl)
            else:
                cl_day = 0
        r["icu_day"], r["cl_day"] = icu_day, cl_day
        # RIT：上の行の、同じブロックの新規一次BSIで、DOEが13日以内
        rit = False
        if r["block"] and doe is not None and ev.get("secondary") != "はい":
            for prev in out:
                if (prev["block"] == r["block"] and prev["new"] and prev["doe"] is not None
                        and doe - dt.timedelta(days=13) <= prev["doe"] <= doe):
                    rit = True
                    break
        r["rit"] = rit
        r["new"] = (crit in CRIT and r["block"] is not None
                    and ev.get("secondary") != "はい" and not rit)
        # 判定
        if pid is None and ev.get("culture") is None and ev.get("symptom") is None:
            j = ""
        elif pid is None:
            j = "⚠患者IDを入力"
        elif r["block"] is None:
            j = "⚠この患者IDはAllにありません"
        elif crit not in CRIT:
            j = crit
        elif ev.get("secondary") == "はい":
            j = "2次性BSI（CLABSIではない）"
        elif rit:
            j = "RIT内（前回のBSIに菌を追加）"
        elif p is None:
            j = "⚠DOEが年度の範囲外"
        elif icu_day == 0:
            j = "ICU外（ICUのCLABSIではない）"
        elif icu_day <= 2:
            j = "入室2日以内（POA・転棟元に帰属）"
        elif cl_day < 3:
            j = "BSI（CL非関連）"
        else:
            j = "CLABSI"
        r["judge"] = j
        out.append(r)
    return out


def monthly(dates, blks, judged):
    """月（年度の最初の月から12か月）ごとの分母と件数"""
    start = dates[0]
    months = []
    for i in range(12):
        y, m = start.year + (start.month - 1 + i) // 12, (start.month - 1 + i) % 12 + 1
        months.append((y, m))
    idx = {ym: i for i, ym in enumerate(months)}
    pd_, cl = [0] * 12, [0] * 12
    for b in blks:
        for d, x in zip(dates, b["c"]):
            i = idx.get((d.year, d.month))
            if i is None:
                continue
            if x in ("C有", "C無"):
                pd_[i] += 1
            if x == "C有":
                cl[i] += 1
    lcbi, csep = [0] * 12, [0] * 12
    for r in judged:
        if r["judge"] != "CLABSI":
            continue
        i = idx.get((r["doe"].year, r["doe"].month))
        if i is None:
            continue
        if r["crit"] in ("LCBI-1", "LCBI-2"):
            lcbi[i] += 1
        elif r["crit"] == "CSEP":
            csep[i] += 1
    return dict(months=months, patient_days=pd_, cl_days=cl, lcbi=lcbi, csep=csep)
