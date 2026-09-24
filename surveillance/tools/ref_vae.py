# -*- coding: utf-8 -*-
"""NHSN/JHAIS の VAC 判定を Python で独立に実装し、ワークブックの計算結果と突き合わせる。

表計算の数式とは別実装にすることで、判定ロジックの取り違えを相互に検出する。
"""
import datetime as dt
import openpyxl

FIO2_RISE, PEEP_RISE, PEEP_FLOOR = 0.2, 3, 5
MV_MIN = 4              # 判定シートを割り当てるMV日数の下限
EVENT_PERIOD = 14       # DOEから14日間は新規VAEを判定しない
MVROWS = 60             # 判定シートが扱えるMV日数


def norm_fio2(x):
    if x is None or not isinstance(x, (int, float)):
        return None
    return x / 100 if x > 1 else x


def adj_peep(x):
    if x is None or not isinstance(x, (int, float)):
        return None
    return max(x, PEEP_FLOOR)


def is_mv(v, f, p):
    """その日がMV日か。V有無欄が優先、空欄のときだけ FiO2/PEEP で補う。"""
    if v == "V有":
        return True
    if v in (None, ""):
        return f is not None or p is not None
    return False          # 「V無」「V有（NIV）」など


def read_blocks(path, sheet="All"):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()

    def cell(r, c):                      # 1-based
        return rows[r - 1][c - 1] if r - 1 < len(rows) and c - 1 < len(rows[r - 1]) else None

    vrows = [r for r in range(1, min(len(rows), 200) + 1) if cell(r, 2) == "V有無"]
    block0, bstep = vrows[0], vrows[1] - vrows[0]
    drows = [r for r in range(1, block0) if isinstance(cell(r, 3), dt.datetime)]
    date_row = drows[-1]
    cols = [c for c in range(1, len(rows[date_row - 1]) + 1)
            if isinstance(cell(date_row, c), dt.datetime)]
    c0, cn = cols[0], cols[-1]
    nblk = (len(rows) - block0 + 1) // bstep
    dates = [cell(date_row, c) for c in range(c0, cn + 1)]

    out = []
    for k in range(1, nblk + 1):
        b = block0 + bstep * (k - 1)
        v = [cell(b, c) for c in range(c0, cn + 1)]
        f = [cell(b + 1, c) for c in range(c0, cn + 1)]
        p = [cell(b + 2, c) for c in range(c0, cn + 1)]
        out.append(dict(k=k, pid=cell(b + 1, 1), name=cell(b + 3, 1), v=v, f=f, p=p))
    return dates, out


def episodes(blk):
    """MV日の連続した区間（＝MVエピソード）を [(開始index, 長さ), ...] で返す"""
    mv = [is_mv(blk["v"][i], blk["f"][i], blk["p"][i]) for i in range(len(blk["v"]))]
    eps, i = [], 0
    while i < len(mv):
        if mv[i]:
            j = i
            while j < len(mv) and mv[j]:
                j += 1
            eps.append((i, j - i))
            i = j
        else:
            i += 1
    return eps, sum(mv)


def does_for_episode(blk, start, length):
    """1エピソード内のDOE（MV日番号、1始まり）を返す"""
    n = min(length, MVROWS)
    F = [norm_fio2(blk["f"][start + i]) for i in range(n)]
    P = [adj_peep(blk["p"][start + i]) for i in range(n)]
    doe = []
    for i in range(n):                       # i は0始まり、MV日 = i+1
        if i + 1 < 3 or i + 1 >= n:          # 最早DOEはMV3日目／2暦日持続に翌日が必要
            continue
        f2, f1, f0, fp = F[i - 2], F[i - 1], F[i], F[i + 1]
        p2, p1, p0, pp = P[i - 2], P[i - 1], P[i], P[i + 1]
        fio2_route = (None not in (f2, f1, f0, fp) and f1 <= f2
                      and f0 - f2 >= FIO2_RISE - 1e-9 and fp - f2 >= FIO2_RISE - 1e-9)
        peep_route = (None not in (p2, p1, p0, pp) and p1 <= p2
                      and p0 - p2 >= PEEP_RISE and pp - p2 >= PEEP_RISE)
        if not (fio2_route or peep_route):
            continue
        if doe and (i + 1) - doe[-1] < EVENT_PERIOD:   # イベント期間14日
            continue
        doe.append(i + 1)
    return doe


def analyse(path):
    dates, blocks = read_blocks(path)
    res = []
    for blk in blocks:
        eps, mvtotal = episodes(blk)
        if mvtotal == 0:
            continue
        ep1 = eps[0]
        doe = does_for_episode(blk, *ep1)
        res.append(dict(k=blk["k"], pid=blk["pid"], name=blk["name"],
                        mvtotal=mvtotal, neps=len(eps),
                        ep1_start=dates[ep1[0]], ep1_len=ep1[1],
                        doe=[dates[ep1[0] + d - 1] for d in doe]))
    return dates, res


if __name__ == "__main__":
    import sys
    dates, res = analyse(sys.argv[1])
    elig = [r for r in res if r["mvtotal"] >= MV_MIN]
    print(f"MV日数>=1 の患者: {len(res)}名 / MV日数>={MV_MIN} の患者: {len(elig)}名")
    print(f"DOEを持つ患者: {sum(1 for r in res if r['doe'])}名")
    print()
    for i, r in enumerate(elig, 1):
        d = ", ".join(x.strftime("%m/%d") for x in r["doe"]) or "-"
        print(f"  slot{i:2d} block{r['k']:4d} ID={r['pid']} {r['name']}  "
              f"MV{r['mvtotal']}日 ep{r['neps']}  ep1={r['ep1_start'].strftime('%m/%d')}"
              f"({r['ep1_len']}日)  DOE={d}")


def vac_days_global(blk):
    """全エピソードを通して、VAC（酸素化悪化）の条件を満たす日のindexを返す。
    一覧シートのセクションBと同じく、暦日ベースで走査する。"""
    n = len(blk["v"])
    mv = [is_mv(blk["v"][i], blk["f"][i], blk["p"][i]) for i in range(n)]
    F = [norm_fio2(x) for x in blk["f"]]
    P = [adj_peep(x) for x in blk["p"]]
    out = []
    for i in range(2, n - 1):
        if not (mv[i - 2] and mv[i - 1] and mv[i] and mv[i + 1]):
            continue           # 4日連続でMV日＝同一エピソード内かつMV3日目以降
        f2, f1, f0, fp = F[i - 2], F[i - 1], F[i], F[i + 1]
        p2, p1, p0, pp = P[i - 2], P[i - 1], P[i], P[i + 1]
        if ((None not in (f2, f1, f0, fp) and f1 <= f2
             and f0 - f2 >= FIO2_RISE - 1e-9 and fp - f2 >= FIO2_RISE - 1e-9)
                or (None not in (p2, p1, p0, pp) and p1 <= p2
                    and p0 - p2 >= PEEP_RISE - 1e-9 and pp - p2 >= PEEP_RISE - 1e-9)):
            out.append(i)
    return out


def does_global(blk, mvtotal):
    """セクションBが出すDOE（暦日ベース・イベント期間14日を適用）"""
    if mvtotal < MV_MIN:
        return []
    doe = []
    for i in vac_days_global(blk):
        if doe and i - doe[-1] < EVENT_PERIOD:
            continue
        doe.append(i)
    return doe


def analyse_global(path):
    """ブロックごとに MV日数・エピソード・暦日ベースのDOE を返す"""
    dates, blocks = read_blocks(path)
    res = []
    for blk in blocks:
        eps, mvtotal = episodes(blk)
        doe = does_global(blk, mvtotal)
        res.append(dict(k=blk["k"], pid=blk["pid"], name=blk["name"],
                        mvtotal=mvtotal, neps=len(eps),
                        ep1_start=dates[eps[0][0]] if eps else None,
                        ep1_len=eps[0][1] if eps else 0,
                        first_mv=dates[eps[0][0]] if eps else None,
                        last_mv=dates[eps[-1][0] + eps[-1][1] - 1] if eps else None,
                        doe=[dates[i] for i in doe]))
    return dates, res
