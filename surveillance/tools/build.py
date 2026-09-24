# -*- coding: utf-8 -*-
"""年度サーベイランスワークシートに VAE 個別判定シートの自動作成機能を組み込む

Allシートの構造（日付行・ブロック先頭行・ブロック数）はファイルごとに違うため、
実ファイルから読み取ってから各シートを生成する。
"""
import datetime as dt
import sys, time
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as gl
import common

WIDTHS = [7, 12, 10, 12, 12, 11, 11, 10, 12, 22, 11, 9, 11, 9]


def probe(ws):
    """Allシートの実構造を読み取る"""
    # ブロック先頭行＝B列が「V有無」の行
    vrows = [r for r in range(1, min(ws.max_row, 200) + 1)
             if ws.cell(r, 2).value == "V有無"]
    assert len(vrows) >= 2, "「V有無」の行が見つかりません"
    block0, bstep = vrows[0], vrows[1] - vrows[0]

    # 日付行＝ブロック先頭より上で、C列が日付になっている一番下の行
    drows = [r for r in range(1, block0)
             if isinstance(ws.cell(r, 3).value, dt.datetime)]
    assert drows, "日付の行が見つかりません"
    date_row = drows[-1]

    # 日付列の範囲
    cols = [c for c in range(1, ws.max_column + 1)
            if isinstance(ws.cell(date_row, c).value, dt.datetime)]
    c0, cn = cols[0], cols[-1]

    last = ws.max_row
    nblk = (last - block0 + 1) // bstep
    return dict(C0=c0, CN=cn, DATE_ROW=date_row, BLOCK0=block0, BSTEP=bstep,
                NBLK=nblk, ALL_LAST_ROW=last)


def count_mv_blocks(ws, g):
    """MV日数が MV_MIN 日以上のブロック数を数える（判定シートの必要枚数の目安）"""
    n = 0
    for k in range(1, g["NBLK"] + 1):
        b = g["BLOCK0"] + g["BSTEP"] * (k - 1)
        mv = 0
        for c in range(g["C0"], g["CN"] + 1):
            v = ws.cell(b, c).value
            if v == "V有":
                mv += 1
            elif v in (None, ""):
                if ws.cell(b + 1, c).value is not None or ws.cell(b + 2, c).value is not None:
                    mv += 1
        if mv >= common.MV_MIN:
            n += 1
    return n


def build_guide(wb, GUIDE):
    from common import (F_TITLE, F_SEC, fill, Font, A_L, A_LW)
    ws = wb.create_sheet("使い方")
    ws.sheet_properties.tabColor = "2E7D32"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "VAE 個別判定シート 自動作成の使い方"
    ws["A1"].font = F_TITLE
    ws["A1"].fill = fill("2E7D32")
    ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 27.75
    for i, (a, b) in enumerate(GUIDE):
        r = 3 + i
        if a and not b:
            ws.cell(r, 1, a).font = F_SEC
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        elif a:
            ws.cell(r, 2, a).font = Font(bold=True, size=9)
            ws.cell(r, 2).alignment = A_L
            ws.cell(r, 3, b).font = Font(size=9)
            ws.cell(r, 3).alignment = A_LW
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 95
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:C{2 + len(GUIDE)}"
    return ws


def guide_text():
    from common import NSHEET, MV_MIN, NBLK
    last = f"VAE-{NSHEET:02d}"
    return [
        ("", ""),
        ("1. このファイルでできること", ""),
        ("Allシート", "これまで通り FiO2・PEEP・V有無・C有無 を日付ごとに入力します。入力内容は変更していません。"),
        ("自動作成", f"FiO2／PEEP を入力すると、人工呼吸器装着患者が自動的に検出され、"
                     f"VAE-01〜{last} の個別判定シートが自動で埋まります。"),
        ("自動反映", "すべて数式でつながっているため、Allシートの値を直したその瞬間に判定シートも再計算されます。"
                     "判定にマクロは使っていないので、マクロが禁止されたPCでも動きます。"),
        ("", ""),
        ("2. 作業の流れ", ""),
        ("① Allシート", "毎日、V有無／FiO2／PEEP／C有無／部屋 を入力（従来どおり）。A列に患者ID・氏名。"),
        ("② VAE対象一覧", "人工呼吸器装着患者と、そのMV日数・DOE・判定（VAC/IVAC/PVAP）が一覧で自動表示されます。"),
        (f"③ VAE-01〜{last[-2:]}", "一覧のリンクから個別判定シートへ。PEEP・FiO2・判定が自動で入っています。"),
        ("④ 臨床所見", "VACが出た患者だけ、該当日の 体温・WBC・新規抗菌薬・PVAP基準 を入力。"
                       "→ IVAC・PVAP まで自動判定されます。"),
        ("⑤ 年間集計", "従来どおりの集計シート。VAE対象一覧のセクションC に月別の自動集計（参考値）も出ます。"),
        ("", ""),
        ("3. セルの色", ""),
        ("黄色", "入力セル（ここだけを編集します）"),
        ("薄い灰青色", "自動計算（数式。上書きしないでください）"),
        ("緑色", "判定結果"),
        ("橙色", "VAEウィンドウ期間（DOE±2日）の行"),
        ("", ""),
        ("3-2. Allシートの赤色（FiO2・PEEPの行）", ""),
        ("濃い赤（白文字）",
         "DOE（イベント発生日）。個別判定シートが報告する日で、イベント期間14日による重複排除まで適用済みです。"),
        ("薄い赤",
         "JHAISの酸素化悪化基準（VAC）に合致した日のうち、上のDOEに当たらないもの。"
         "＝直前のDOEから14日以内で新規報告の対象外か、まだ判定シートが割り当てられていない患者です。"
         "薄い赤が出たら VAE対象一覧 でその患者の判定を確認してください。"),
        ("補足",
         "元のAllシートに入っていた条件付き書式は、#REF!エラーや参照位置のずれ"
         "（入力したセルと無関係のセルが赤くなる）があったため削除し、上の2ルールに置き換えました。"),
        ("", ""),
        ("4. MV日（人工呼吸器装着日）の判定ルール", ""),
        ("MV日とみなす日", "AllシートのV有無行が「V有」の日。V有無が空欄のときは、FiO2／PEEP に入力があればMV日とみなします。"),
        ("除外する日", "「V無」の日と「V有（NIV）」の日。非侵襲的換気（NPPV・マスクCPAP等）はVAEの分子・対象に含めません。"),
        ("エピソード", "1暦日以上完全に人工呼吸を離脱したあとの再装着は別エピソードです。"
                       "判定シートのC9「エピソード番号」を 2 に変えると2回目を判定できます。"),
        ("判定シートの割当", f"MV日数が{MV_MIN}日以上のブロックに、上から順に VAE-01〜{last} が割り当てられます"
                             f"（JHAIS：VAEの対象は人工呼吸を暦日で{MV_MIN}日以上）。"),
        ("FiO2の単位", "小数（0.4）でも％（40）でも、どちらで入力しても同じ判定になります。"),
        ("", ""),
        ("5. 判定ロジック（NHSN / JHAIS デバイス関連感染サーベイランス）", ""),
        ("日最小値", "その暦日内で1時間を超えて維持された最も低い PEEP／FiO2。PEEP 0〜5cmH2O は同等（5として自動補正）。"),
        ("ベースライン期間", "日最小 FiO2 または PEEP が2暦日以上安定または低下している期間。悪化初日の直前2日。"),
        ("VAC", "ベースライン初日の値と比べて FiO2 が20ポイント以上上昇、または PEEP が3cmH2O以上上昇し、"
                "それが2暦日以上持続。最も早いDOEは MV3日目。"),
        ("IVAC", "VAC成立＋MV3日目以降かつVAEウィンドウ（DOE±2日）内に、"
                 "①体温>38℃ または <36℃、あるいは WBC≥12,000 または ≤4,000/mm³ かつ "
                 "②新規抗菌薬開始（4QAD以上継続）。"),
        ("PVAP", "IVAC成立＋ウィンドウ内に PVAP基準1〜3 のいずれか。"),
        ("イベント期間", "DOEを1日目とする14日間は新たなVAEを判定しません（次の最早DOEは15日目）。"),
        ("", ""),
        ("6. 判定シートが足りなくなったら", ""),
        ("シートを増やす", f"{last} のシートタブを右クリック →「移動またはコピー」→「コピーを作成する」。"
                           f"できたシートのC4セルの数式内の {NSHEET} を {NSHEET+1} に書き換えれば、"
                           f"{NSHEET+1}人目に自動で割り当てられます。"),
        ("特定の患者に固定", "C4に「VAE対象一覧」セクションBのブロック番号を直接入力すると、その患者に固定されます。"),
        ("スキャン範囲", f"VAE対象一覧のセクションBは Allシートの先頭{NBLK}ブロックを見ています。"),
        ("", ""),
        ("7. 出典・注意", ""),
        ("NHSN VAE", "CDC/NHSN Patient Safety Component Manual, Chapter 10: Ventilator-Associated Event (VAE). "
                     "https://www.cdc.gov/nhsn/pdfs/pscmanual/10-vae_final.pdf"),
        ("JHAIS", "日本環境感染学会 JHAIS委員会 医療器具関連感染サーベイランス部門 マニュアル Ver.2.5 "
                  "「5) 人工呼吸器関連イベント(VAE)サーベイランス」"),
        ("注意", "自動判定は判定の補助です。最終判定は原典プロトコールと診療記録に基づき感染管理担当者が行ってください。"),
    ]


def main():
    SRC = sys.argv[1] if len(sys.argv) > 1 else "ws.xlsx"
    DST = sys.argv[2] if len(sys.argv) > 2 else "out.xlsx"
    NSHEET_ARG = int(sys.argv[3]) if len(sys.argv) > 3 else None
    NBLK_CAP = int(sys.argv[4]) if len(sys.argv) > 4 else None
    t0 = time.time()
    print("load", SRC)
    is_xlsm = SRC.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(SRC, keep_vba=is_xlsm)
    assert "All" in wb.sheetnames, wb.sheetnames
    al = wb["All"]

    g = probe(al)
    if NBLK_CAP:
        g["NBLK"] = min(g["NBLK"], NBLK_CAP)
    print("Allシートの構造:", ", ".join(f"{k}={v}" for k, v in g.items()))

    common.configure(**g)
    need = count_mv_blocks(al, g)
    nsheet = NSHEET_ARG or max(40, ((need * 2 + 19) // 20) * 20)
    common.configure(NSHEET=nsheet)
    print(f"MV{common.MV_MIN}日以上の患者: 現在{need}名 → 個別判定シートを{nsheet}枚作成")

    # configure() の後に import すること（各モジュールは import 時に定数を取り込む）
    import sheet_clin, sheet_vae, sheet_list, sheet_all
    from common import (N_ALL, N_ALLD, N_CLIN, CN, ALL_LAST_ROW,
                        CLIN_LAST_ROW, NSHEET)

    for nm, ref in ((N_ALL,  f"All!$A$1:${gl(CN)}${ALL_LAST_ROW}"),
                    (N_ALLD, f"All!$C$1:${gl(CN)}${ALL_LAST_ROW}"),
                    (N_CLIN, f"臨床所見!$A$1:${gl(CN)}${CLIN_LAST_ROW}")):
        if nm in wb.defined_names:
            del wb.defined_names[nm]
        wb.defined_names.add(DefinedName(nm, attr_text=ref))

    removed = sheet_all.build(wb)
    print(f"All: 壊れていた条件付き書式 {removed} ルールを削除し、DOE強調の2ルールに置換")
    build_guide(wb, guide_text())
    sheet_clin.build(wb)
    lst = sheet_list.build(wb)
    for i, w in enumerate(WIDTHS):
        lst.column_dimensions[gl(i + 1)].width = w
    for n in range(1, NSHEET + 1):
        sheet_vae.build(wb, n)
    print("sheets built", round(time.time() - t0, 1), "s")

    order = (["使い方", "All", "VAE対象一覧", "臨床所見"]
             + [f"VAE-{n:02d}" for n in range(1, NSHEET + 1)])
    rest = [s for s in wb.sheetnames if s not in order]
    wb._sheets = [wb[s] for s in order + rest]

    wb.calculation.fullCalcOnLoad = True
    wb.active = 0

    print("save", DST)
    wb.save(DST)
    print("done", round(time.time() - t0, 1), "s")


if __name__ == "__main__":
    main()
