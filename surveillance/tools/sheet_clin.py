# -*- coding: utf-8 -*-
"""臨床所見シート

VAE対象一覧セクションAで VAC と判定された患者だけにスロットが対応する。
スロット n ＝ VAC患者の n 人目。ここに入力した患者だけが個別判定シートを持つ。
"""
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter as gl
from common import *

META = ["スロット", "ブロック", "患者ID", "氏名", "DOE①", "警告"]


def build(wb):
    ws = wb.create_sheet("臨床所見")
    ws.sheet_properties.tabColor = "C55A11"

    ws["A1"] = "臨床所見（IVAC／PVAP 判定用）　― VACと判定された患者だけを入力します ―"
    ws["A1"].font = F_TITLE
    ws["A1"].fill = fill("C55A11")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    ws.row_dimensions[1].height = 27.75

    ws["A2"] = ("スロット番号は「VAE対象一覧」セクションAのNo.と同じです。VACと判定された患者が上から順に入り、"
                "患者ID・氏名・DOEはA列に自動表示されます。日付列はAllシートと同じ並びです。")
    ws["A3"] = ("入力するのは DOE±2日（VAEウィンドウ期間）の前後だけで足ります。"
                "入力済みなのにDOE前後が空のスロットには A列に「⚠割当確認」が出ます"
                "（患者の並びが変わった可能性があります）。"
                "ここに1つでも入力すると、その患者に個別判定シート（VAE-01〜）が自動で割り当てられ、"
                "IVAC・PVAPまで判定されます。"
                "新規抗菌薬＝NHSN適格抗菌薬を開始した日に「はい」（開始日の前2日に同薬の投与がなく、4QAD以上継続）。")
    for r in (2, 3):
        ws.cell(r, 1).font = F_NOTE
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=30)

    ws["A5"] = "患者"
    ws["B5"] = "項目"
    for a in ("A5", "B5"):
        style_cell(ws[a], NAVY, Font(bold=True, size=9, color="FFFFFF"), A_CW)
    for c in range(C0, CN + 1):
        cell = ws.cell(5, c)
        cell.value = f"=All!{gl(c)}{DATE_ROW}"
        style_cell(cell, NAVY, Font(bold=True, size=8, color="FFFFFF"), A_C, numfmt="m/d")
    ws.row_dimensions[5].height = 24

    # ---- スロット（VAC患者 n 人目）------------------------------------
    L = "'VAE対象一覧'"
    for n in range(1, NCLIN + 1):
        b = clin_base(n)
        ar = LIST_A_TOP + n - 1                  # 一覧セクションAの対応行
        meta = [
            n,
            f'=IF({L}!$B${ar}="","",{L}!$B${ar})',    # ブロック番号
            f'=IF({L}!$C${ar}="","",{L}!$C${ar})',    # 患者ID
            f'=IF({L}!$D${ar}="","",{L}!$D${ar})',    # 氏名
            f'=IF({L}!$H${ar}="","",{L}!$H${ar})',    # DOE①
        ]
        for off, v in enumerate(meta):
            c = ws.cell(b + off, 1, v)
            style_cell(c, BAND if off == 0 else AUTO,
                       Font(bold=(off == 0), size=9), A_C,
                       numfmt="yyyy/m/d" if off == 4 else "General")
        # 入力済みなのにDOE前後に値がない＝患者の割当がずれた可能性がある
        d = f"MATCH($A${b+4},All!$C${DATE_ROW}:${gl(CN)}${DATE_ROW},0)+{C0-1}"
        ws.cell(b + 5, 1, (
            f'=IF(COUNTA($C{b}:${gl(CN)}{b+7})=0,"",'
            f'IFERROR(IF(COUNTA(INDEX({N_CLIN},{b},MAX({C0},{d}-2)):'
            f'INDEX({N_CLIN},{b+7},MIN({CN},{d}+2)))>0,"",'
            f'"⚠割当確認"),"⚠DOE不明"))'))
        style_cell(ws.cell(b + 5, 1), None, Font(size=8, bold=True, color=RED), A_C)
        for off in (6, 7):
            style_cell(ws.cell(b + off, 1), None, F_BODY, A_C)
        for off, (label, _kind) in enumerate(CLIN_ROWS):
            cell = ws.cell(b + off, 2, label)
            style_cell(cell, BAND if off == 0 else None,
                       Font(size=9, bold=(off == 0)), A_L)

    # ---- 入力域の着色 -------------------------------------------------
    # VAC患者が入っているスロットだけ黄色（＝入力してよい行）にする
    ws.conditional_formatting.add(
        f"C6:{gl(CN)}{CLIN_LAST_ROW}",
        Rule(type="expression",
             formula=[f'AND($B6<>"",INDEX($A$6:$A${CLIN_LAST_ROW},'
                      f'8*INT((ROW()-6)/8)+2)<>"")'],
             dxf=DifferentialStyle(fill=cf_fill(YELLOW)), stopIfTrue=False))
    # 患者が割り当てられていないスロットは灰色にして入力対象外だと分かるようにする
    ws.conditional_formatting.add(
        f"A6:{gl(CN)}{CLIN_LAST_ROW}",
        Rule(type="expression",
             formula=[f'INDEX($A$6:$A${CLIN_LAST_ROW},8*INT((ROW()-6)/8)+2)=""'],
             dxf=DifferentialStyle(fill=cf_fill("F2F2F2")), stopIfTrue=False))

    # ---- 入力規則 ------------------------------------------------------
    dv_abx = DataValidation(type="list", formula1='"はい"', allow_blank=True)
    dv_pvap = DataValidation(type="list", formula1='"基準1,基準2,基準3"', allow_blank=True)
    ws.add_data_validation(dv_abx)
    ws.add_data_validation(dv_pvap)
    for n in range(1, NCLIN + 1):
        b = clin_base(n)
        dv_abx.add(f"C{b+4}:{gl(CN)}{b+4}")
        dv_pvap.add(f"C{b+6}:{gl(CN)}{b+6}")

    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 20
    for c in range(C0, CN + 1):
        ws.column_dimensions[gl(c)].width = 6.5
    ws.freeze_panes = "C6"
    ws.sheet_view.zoomScale = 85
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.print_title_cols = "A:B"
    ws.print_title_rows = "5:5"
    return ws
