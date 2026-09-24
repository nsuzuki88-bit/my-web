# -*- coding: utf-8 -*-
"""臨床所見シート（体温・WBC・新規抗菌薬・PVAP基準の日次入力／Allと同じ日付列）"""
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter as gl
from common import *


def build(wb):
    ws = wb.create_sheet("臨床所見")
    ws.sheet_properties.tabColor = "C55A11"

    # ---- ヘッダー ----------------------------------------------------
    ws["A1"] = "臨床所見（IVAC／PVAP 判定用）　― VAC が出た患者の該当日だけ入力すれば足ります ―"
    ws["A1"].font = F_TITLE
    ws["A1"].fill = fill("C55A11")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    ws.row_dimensions[1].height = 27.75

    ws["A2"] = ("日付列はAllシートと完全に同じ並びです。ブロック番号もAllシートと同じ番号で対応します"
                "（例：Allの1人目＝このシートの #1）。入力するとVAE-01〜のシートに即座に反映されます。")
    ws["A3"] = ("新規抗菌薬＝NHSN適格抗菌薬を開始した日に「はい」（開始日の前2日に同薬の投与がなく、4QAD以上継続した場合のみ）。"
                "PVAP基準＝該当検体の採取日の行で 基準1／基準2／基準3 を選択。")
    for r in (2, 3):
        ws.cell(r, 1).font = F_NOTE
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=30)

    ws["A5"] = "ブロック"
    ws["B5"] = "項目"
    for a in ("A5", "B5"):
        style_cell(ws[a], NAVY, Font(bold=True, size=9, color="FFFFFF"), A_CW)
    for c in range(C0, CN + 1):
        cell = ws.cell(5, c)
        cell.value = f"=All!{gl(c)}{DATE_ROW}"
        style_cell(cell, NAVY, Font(bold=True, size=8, color="FFFFFF"), A_C, numfmt="m/d")
    ws.row_dimensions[5].height = 24

    # ---- 患者ブロック ------------------------------------------------
    for k in range(1, NBLK + 1):
        b = clin_base(k)
        idr, nmr = all_id(k), all_name(k)
        ws.cell(b, 1, k)
        style_cell(ws.cell(b, 1), BAND, Font(bold=True, size=9), A_C)
        ws.cell(b + 1, 1).value = f'=IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},{idr})="","",INDEX(All!$A$1:$A${ALL_LAST_ROW},{idr}))'
        ws.cell(b + 2, 1).value = f'=IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},{nmr})="","",INDEX(All!$A$1:$A${ALL_LAST_ROW},{nmr}))'
        for off in (1, 2):
            style_cell(ws.cell(b + off, 1), AUTO, F_BODY, A_C)
        for off in (3, 4, 5, 6, 7):
            style_cell(ws.cell(b + off, 1), None, F_BODY, A_C)

        for off, (label, _kind) in enumerate(CLIN_ROWS):
            cell = ws.cell(b + off, 2, label)
            style_cell(cell, BAND if off == 0 else None, Font(size=9, bold=(off == 0)), A_L)

    # ---- 入力域の着色（条件付き書式：1ルールで軽量に）----------------
    tint_last = min(clin_base(150) + 7, CLIN_LAST_ROW)
    ws.conditional_formatting.add(
        f"C6:{gl(CN)}{tint_last}",
        Rule(type="expression", formula=['$B6<>""'],
             dxf=DifferentialStyle(fill=cf_fill(YELLOW)), stopIfTrue=False),
    )

    # ---- 入力規則（先頭150ブロック分）--------------------------------
    dv_abx = DataValidation(type="list", formula1='"はい"', allow_blank=True)
    dv_pvap = DataValidation(type="list", formula1='"基準1,基準2,基準3"', allow_blank=True)
    ws.add_data_validation(dv_abx)
    ws.add_data_validation(dv_pvap)
    for k in range(1, 151):
        b = clin_base(k)
        dv_abx.add(f"C{b+4}:{gl(CN)}{b+4}")
        dv_pvap.add(f"C{b+6}:{gl(CN)}{b+6}")

    # ---- 体裁 --------------------------------------------------------
    ws.column_dimensions["A"].width = 10
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
