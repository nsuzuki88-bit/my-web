# -*- coding: utf-8 -*-
"""再計算済みワークブックの結果を、Python独立実装（ref_vae.py）と突き合わせる"""
import datetime as dt
import sys, openpyxl
import ref_vae as R
from common import (LIST_A_TOP, LIST_B_TOP, NSHEET, NBLK, MV_MIN)

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
    dates, ref = R.analyse(src)
    refmap = {r["k"]: r for r in ref}
    elig = [r for r in ref if r["mvtotal"] >= MV_MIN]

    wb = openpyxl.load_workbook(calc, data_only=True)
    lst = wb["VAE対象一覧"]

    # ---- セクションB：全ブロックのMV日数 ----------------------------
    print("=== セクションB：MV日数（全ブロック）===")
    mism = 0
    for k in range(1, NBLK + 1):
        r = LIST_B_TOP + k - 1
        got = lst.cell(r, 4).value or 0
        want = refmap[k]["mvtotal"] if k in refmap else 0
        if got != want:
            mism += 1
            if mism <= 5:
                print(f"  NG  block{k} MV日数: got {got} want {want}")
        else:
            ok += 1
    if mism:
        ng += mism
    print(f"  {NBLK}ブロック照合、不一致 {mism} 件")

    # ---- セクションB：割当No ----------------------------------------
    print("\n=== セクションB：判定シートの割当 ===")
    want_slot = {r["k"]: i for i, r in enumerate(elig, 1)}
    for k in range(1, NBLK + 1):
        r = LIST_B_TOP + k - 1
        chk(f"block{k} 割当No", lst.cell(r, 7).value, want_slot.get(k))
    print(f"  MV{MV_MIN}日以上: {len(elig)}名 → slot1..{len(elig)}")

    # ---- セクションA：各判定シート ----------------------------------
    print("\n=== セクションA：個別判定シート ===")
    for i in range(1, NSHEET + 1):
        r = LIST_A_TOP + i - 1
        blk = lst.cell(r, 3).value
        want = elig[i - 1] if i <= len(elig) else None
        if want is None:
            chk(f"slot{i} 未割当", blk, None)
            continue
        chk(f"slot{i} ブロック番号", blk, want["k"])
        chk(f"slot{i} 患者ID", lst.cell(r, 4).value, want["pid"])
        chk(f"slot{i} 氏名", lst.cell(r, 5).value, want["name"])
        chk(f"slot{i} MV1日目", lst.cell(r, 7).value, want["ep1_start"])
        chk(f"slot{i} MV日数", lst.cell(r, 8).value, want["mvtotal"])
        chk(f"slot{i} エピソード数", lst.cell(r, 9).value, want["neps"])
        d = want["doe"]
        chk(f"slot{i} DOE①", lst.cell(r, 11).value, d[0] if len(d) > 0 else None)
        chk(f"slot{i} 判定①", lst.cell(r, 12).value, "VAC" if len(d) > 0 else None)
        chk(f"slot{i} DOE②", lst.cell(r, 13).value, d[1] if len(d) > 1 else None)
        chk(f"slot{i} 判定②", lst.cell(r, 14).value, "VAC" if len(d) > 1 else None)

    # ---- DOEが出た患者の判定シート本体 ------------------------------
    print("\n=== DOEが出た患者の判定シート ===")
    for i, want in enumerate(elig, 1):
        if not want["doe"]:
            continue
        ws = wb[f"VAE-{i:02d}"]
        print(f"  VAE-{i:02d}: block{want['k']} ID={want['pid']} {want['name']}")
        chk(f"  VAE-{i:02d} 患者ID", ws["C5"].value, want["pid"])
        chk(f"  VAE-{i:02d} 氏名", ws["C6"].value, want["name"])
        chk(f"  VAE-{i:02d} MV1日目", ws["C10"].value, want["ep1_start"])
        chk(f"  VAE-{i:02d} 本エピソード日数", ws["K5"].value, want["ep1_len"])
        chk(f"  VAE-{i:02d} 検出VAE件数", ws["K7"].value, len(want["doe"]))
        chk(f"  VAE-{i:02d} DOE①", ws["K8"].value, want["doe"][0])
        chk(f"  VAE-{i:02d} 判定①", ws["K9"].value, "VAC")
        # 表本体のFiO2/PEEPがAllシートと一致するか
        b = [x for x in R.read_blocks(src)[1] if x["k"] == want["k"]][0]
        eps, _ = R.episodes(b)
        s, L = eps[0]
        for d in range(1, min(L, 10) + 1):
            row = 14 + d
            chk(f"  VAE-{i:02d} MV{d}日目 PEEP", ws.cell(row, 3).value, b["p"][s + d - 1])
            chk(f"  VAE-{i:02d} MV{d}日目 FiO2", ws.cell(row, 4).value, b["f"][s + d - 1])

    print(f"\n==== ok={ok}  NG={ng} ====")
    return ng


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1], sys.argv[2]) else 0)
