# -*- coding: utf-8 -*-
"""20XX年度サーベイランスワークシートに VAE 個別判定シートの自動作成機能を組み込む"""
import sys, time
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as gl
from common import *
import sheet_clin, sheet_vae, sheet_list

SRC = sys.argv[1] if len(sys.argv) > 1 else "ws.xlsx"
DST = sys.argv[2] if len(sys.argv) > 2 else "out.xlsx"

WIDTHS = [7, 12, 10, 12, 12, 11, 11, 10, 12, 22, 11, 9, 11, 9]

GUIDE = [
    ("", ""),
    ("1. このファイルでできること", ""),
    ("Allシート", "これまで通り FiO2・PEEP・V有無・C有無 を日付ごとに入力します。入力内容は変更していません。"),
    ("自動作成", "FiO2／PEEP を入力すると、人工呼吸器装着患者が自動的に検出され、"
                 "VAE-01〜VAE-40 の個別判定シートが自動で埋まります。"),
    ("自動反映", "すべて数式でつながっているため、Allシートの値を直したその瞬間に判定シートも再計算されます。"
                 "マクロは使っていないので、マクロが禁止されたPCでも動きます。"),
    ("", ""),
    ("2. 作業の流れ", ""),
    ("① Allシート", "毎日、V有無／FiO2／PEEP／C有無／部屋 を入力（従来どおり）。A列に患者ID・氏名。"),
    ("② VAE対象一覧", "人工呼吸器装着患者と、そのMV日数・DOE・判定（VAC/IVAC/PVAP）が一覧で自動表示されます。"),
    ("③ VAE-01〜40", "一覧のリンクから個別判定シートへ。PEEP・FiO2・判定が自動で入っています。"),
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
    ("4. MV日（人工呼吸器装着日）の判定ルール", ""),
    ("MV日とみなす日", "AllシートのV有無行が「V有」の日、または FiO2／PEEP のいずれかに入力がある日。"),
    ("除外する日", "「V有（NIV）」の日。非侵襲的換気（NPPV・マスクCPAP等）はVAEの分子・対象に含めません。"),
    ("エピソード", "1暦日以上完全に人工呼吸を離脱したあとの再装着は別エピソードです。"
                   "判定シートのC9「エピソード番号」を 2 に変えると2回目を判定できます。"),
    ("", ""),
    ("5. 判定ロジック（NHSN 2026 / JHAIS デバイス関連感染サーベイランス）", ""),
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
    ("シートを増やす", "VAE-40 のシートタブを右クリック →「移動またはコピー」→「コピーを作成する」。"
                       "できたシートのC4セルの数式内の 40 を 41 に書き換えれば、41人目に自動で割り当てられます。"),
    ("特定の患者に固定", "C4に「VAE対象一覧」セクションBのブロック番号を直接入力すると、その患者に固定されます。"),
    ("", ""),
    ("7. 出典・注意", ""),
    ("NHSN VAE", "CDC/NHSN Patient Safety Component Manual, Chapter 10: Ventilator-Associated Event (VAE). "
                 "https://www.cdc.gov/nhsn/pdfs/pscmanual/10-vae_final.pdf"),
    ("JHAIS", "日本環境感染学会 JHAIS委員会 医療器具関連感染サーベイランス部門 マニュアル Ver.2.5 "
              "「5) 人工呼吸器関連イベント(VAE)サーベイランス」"),
    ("注意", "自動判定は判定の補助です。最終判定は原典プロトコールと診療記録に基づき感染管理担当者が行ってください。"),
]


def build_guide(wb):
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


def main():
    t0 = time.time()
    print("load", SRC)
    wb = openpyxl.load_workbook(SRC)
    assert "All" in wb.sheetnames, wb.sheetnames

    # 名前付き範囲
    for nm, ref in ((N_ALL,  f"All!$A$1:${gl(CN)}${ALL_LAST_ROW}"),
                    (N_ALLD, f"All!$C$1:${gl(CN)}${ALL_LAST_ROW}"),
                    (N_CLIN, f"臨床所見!$A$1:${gl(CN)}$3200")):
        if nm in wb.defined_names:
            del wb.defined_names[nm]
        wb.defined_names.add(DefinedName(nm, attr_text=ref))

    build_guide(wb)
    sheet_clin.build(wb)
    lst = sheet_list.build(wb)
    for i, w in enumerate(WIDTHS):
        lst.column_dimensions[gl(i + 1)].width = w
    for n in range(1, NSHEET + 1):
        sheet_vae.build(wb, n)
    print("sheets built", round(time.time() - t0, 1), "s")

    # 並び順： 使い方 / All / VAE対象一覧 / 臨床所見 / VAE-01.. / 既存集計
    order = (["使い方", "All", "VAE対象一覧", "臨床所見"]
             + [f"VAE-{n:02d}" for n in range(1, NSHEET + 1)])
    rest = [s for s in wb.sheetnames if s not in order]
    wb._sheets = [wb[s] for s in order + rest]

    # 開いたときに全再計算させる
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0

    print("save", DST)
    wb.save(DST)
    print("done", round(time.time() - t0, 1), "s")


if __name__ == "__main__":
    main()
