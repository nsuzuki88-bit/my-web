# -*- coding: utf-8 -*-
"""VAE対象一覧：判定シート索引 ＋ 全ブロックスキャン ＋ 月別自動集計"""
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter as gl
from common import *

SEC_A = [("No", 5), ("判定シート", 11), ("ブロック", 8), ("患者ID", 12), ("氏名", 12),
         ("部屋", 7), ("MV1日目", 11), ("MV日数\n(全期間)", 9), ("エピ\nソード数", 8),
         ("対象\nエピソード", 9), ("DOE①", 11), ("判定①", 9), ("DOE②", 11), ("判定②", 9)]

SEC_B = [("ブロック", 8), ("患者ID", 12), ("氏名", 12), ("MV日数", 8), ("初回MV日", 11),
         ("最終MV日", 11), ("割当No", 8), ("年齢\n(入力)", 8),
         ("臓器提供\n同意日(入力)", 12), ("備考(入力)", 24)]


def build(wb):
    ws = wb.create_sheet("VAE対象一覧")
    ws.sheet_properties.tabColor = "C00000"
    ws.sheet_view.showGridLines = False

    ws["A1"] = "VAE 対象一覧（AllシートのFiO2/PEEP入力から自動作成）"
    ws["A1"].font = F_TITLE
    ws["A1"].fill = fill("C00000")
    ws.merge_cells("A1:N1")
    ws.row_dimensions[1].height = 27.75
    ws["A2"] = (f"AllシートにFiO2/PEEPを入力すると、人工呼吸器装着患者が自動で検出され、"
                f"MV日数が{MV_MIN}日以上のブロックに上から順に VAE-01〜VAE-{NSHEET:02d} の"
                f"個別判定シートが割り当てられます。黄色のセル（年齢・臓器提供同意日・備考）のみ手入力です。")
    ws["A3"] = ("MV日の判定：AllのV有無行が「V有」の日。「V無」「V有（NIV）」の日は除外します"
                "（非侵襲換気はVAEの対象外）。V有無が空欄の日は、FiO2／PEEPに入力があればMV日とみなします。"
                "FiO2は小数（0.4）でも％（40）でも同じ判定になります。")
    for r in (2, 3):
        ws.cell(r, 1).font = F_NOTE
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)

    # ============ セクションA：判定シート索引 =========================
    ws.cell(5, 1, f"A. 個別判定シート（VAE-01〜VAE-{NSHEET:02d}）"
                  f"　※判定シート名をクリックすると該当シートへ移動します")
    ws.cell(5, 1).font = F_SEC
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=14)

    for i, (t, w) in enumerate(SEC_A):
        style_cell(ws.cell(6, i + 1, t), NAVY, F_HDR, A_CW)
        ws.column_dimensions[gl(i + 1)].width = w
    ws.row_dimensions[6].height = 32

    for n in range(1, NSHEET + 1):
        r = LIST_A_TOP + n - 1
        s = f"'VAE-{n:02d}'"
        g = lambda ref: f'=IF({s}!$C$4="","",{s}!{ref})'
        ws.cell(r, 1, n)
        ws.cell(r, 2, f'=HYPERLINK("#\'VAE-{n:02d}\'!A1","VAE-{n:02d}")')
        for col, f_, fmt in (
            (3, g("$C$4"),  "0"), (4, g("$C$5"), "General"), (5, g("$C$6"), "General"),
            (6, g("$C$7"),  "General"), (7, g("$C$10"), "yyyy/m/d"), (8, g("$C$11"), "0"),
            (9, g("$K$6"),  "0"), (10, g("$C$9"), "0"), (11, g("$K$8"), "yyyy/m/d"),
            (12, g("$K$9"), "General"), (13, g("$K$10"), "yyyy/m/d"), (14, g("$K$11"), "General"),
        ):
            style_cell(ws.cell(r, col, f_), AUTO, F_BODY, A_C, numfmt=fmt)
        style_cell(ws.cell(r, 1), BAND, Font(bold=True, size=9), A_C)
        style_cell(ws.cell(r, 2), BAND, Font(size=9, color="0563C1", underline="single"), A_C)
        for col in (12, 14):
            style_cell(ws.cell(r, col), GREEN, F_RESULT, A_C)

    for col in (12, 14):
        rng = f"{gl(col)}{LIST_A_TOP}:{gl(col)}{LIST_A_BOT}"
        ws.conditional_formatting.add(rng, Rule(
            type="cellIs", operator="notEqual", formula=['""'],
            dxf=DifferentialStyle(fill=cf_fill(PINK)), stopIfTrue=False))

    # ============ セクションB：全ブロックスキャン ====================
    ws.cell(LIST_B_HDR, 1, f"B. Allシート 全{NBLK}ブロックのスキャン　"
                           f"※MV日数が{MV_MIN}日以上のブロックに、上から順に判定シートが割り当てられます")
    ws.cell(LIST_B_HDR, 1).font = F_SEC
    ws.merge_cells(start_row=LIST_B_HDR, start_column=1, end_row=LIST_B_HDR, end_column=14)

    for i, (t, _w) in enumerate(SEC_B):
        style_cell(ws.cell(LIST_B_HDR + 1, i + 1, t), NAVY, F_HDR, A_CW)
    ws.row_dimensions[LIST_B_HDR + 1].height = 32

    colrng = f"COLUMN(All!$C$1:${gl(CN)}$1)"
    dates = f"All!$C${DATE_ROW}:${gl(CN)}${DATE_ROW}"
    for k in range(1, NBLK + 1):
        r = LIST_B_TOP + k - 1
        vr, fr, pr = all_v(k), all_f(k), all_p(k)
        idr, nmr = all_id(k), all_name(k)
        mvb = mvb_array(vr, fr, pr)
        ws.cell(r, 1, k)
        ws.cell(r, 2, f'=IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},{idr})="","",'
                      f'INDEX(All!$A$1:$A${ALL_LAST_ROW},{idr}))')
        ws.cell(r, 3, f'=IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},{nmr})="","",'
                      f'INDEX(All!$A$1:$A${ALL_LAST_ROW},{nmr}))')
        ws.cell(r, 4, f"=SUMPRODUCT(--({mvb}))")
        ws.cell(r, 5, f'=IF($D{r}=0,"",INDEX({dates},1,MATCH(TRUE,INDEX({mvb},0),0)))')
        ws.cell(r, 6, f'=IF($D{r}=0,"",LOOKUP(2,1/{mvb},{dates}))')
        ws.cell(r, 7, f'=IF($D{r}<{MV_MIN},"",'
                      f'COUNTIF($D${LIST_B_TOP}:$D{r},">="&{MV_MIN}))')
        for col, fmt in ((1, "0"), (2, "General"), (3, "General"), (4, "0"),
                         (5, "yyyy/m/d"), (6, "yyyy/m/d"), (7, "0")):
            style_cell(ws.cell(r, col), AUTO, F_BODY, A_C, numfmt=fmt)
        style_cell(ws.cell(r, 1), BAND, Font(bold=True, size=9), A_C)
        for col, fmt in ((8, "0"), (9, "yyyy/m/d"), (10, "General")):
            style_cell(ws.cell(r, col), YELLOW, F_BODY,
                       A_C if col < 10 else A_L, numfmt=fmt)

    ws.conditional_formatting.add(
        f"A{LIST_B_TOP}:J{LIST_B_BOT}",
        Rule(type="expression", formula=[f"$D{LIST_B_TOP}>0"],
             dxf=DifferentialStyle(font=Font(bold=True)), stopIfTrue=False))

    # ============ セクションC：月別自動集計 ==========================
    ws.cell(LIST_C_HDR, 1, "C. 月別 自動集計（参考値）　※セクションAのDOE・判定から自動計算。"
                           "最終確定値は感染管理担当者の判定に基づいて決定してください")
    ws.cell(LIST_C_HDR, 1).font = F_SEC
    ws.merge_cells(start_row=LIST_C_HDR, start_column=1, end_row=LIST_C_HDR, end_column=14)

    hr = LIST_C_HDR + 1
    style_cell(ws.cell(hr, 1, "指標"), NAVY, F_HDR, A_C)
    for i in range(12):
        style_cell(ws.cell(hr, i + 2, f'=TEXT(EDATE(All!$C${DATE_ROW},{i}),"m月")'),
                   NAVY, F_HDR, A_C)
    style_cell(ws.cell(hr, 14, "計"), NAVY, F_HDR, A_C)

    KA, KB = f"$K${LIST_A_TOP}:$K${LIST_A_BOT}", f"$L${LIST_A_TOP}:$L${LIST_A_BOT}"
    MA, MB = f"$M${LIST_A_TOP}:$M${LIST_A_BOT}", f"$N${LIST_A_TOP}:$N${LIST_A_BOT}"

    def count(kind, i):
        lo = f'">="&EDATE(All!$C${DATE_ROW},{i})'
        hi = f'"<"&EDATE(All!$C${DATE_ROW},{i+1})'
        return (f'=COUNTIFS({KA},{lo},{KA},{hi},{KB},"{kind}")'
                f'+COUNTIFS({MA},{lo},{MA},{hi},{MB},"{kind}")')

    rows = [("VAC 件数", lambda i: count("VAC", i)),
            ("IVAC 件数", lambda i: count("IVAC", i)),
            ("PVAP 件数", lambda i: count("PVAP", i))]
    for j, (label, fn) in enumerate(rows):
        r = LIST_C_TOP + j
        style_cell(ws.cell(r, 1, label), None, Font(size=9), A_L)
        for i in range(12):
            style_cell(ws.cell(r, i + 2, fn(i)), AUTO, F_BODY, A_C, numfmt="0")
        style_cell(ws.cell(r, 14, f"=SUM(B{r}:M{r})"), AUTO, Font(bold=True, size=9), A_C, numfmt="0")

    r = LIST_C_TOP + 3
    style_cell(ws.cell(r, 1, "VAE 合計（VAC＋IVAC＋PVAP）"), None, Font(size=9, bold=True), A_L)
    for i in range(12):
        c = gl(i + 2)
        style_cell(ws.cell(r, i + 2, f"=SUM({c}{LIST_C_TOP}:{c}{LIST_C_TOP+2})"),
                   GREEN, F_RESULT, A_C, numfmt="0")
    style_cell(ws.cell(r, 14, f"=SUM(B{r}:M{r})"), GREEN, F_RESULT, A_C, numfmt="0")

    r = LIST_C_TOP + 4
    style_cell(ws.cell(r, 1, "人工呼吸器装着延べ患者日数"), None, Font(size=9), A_L)
    for i in range(12):
        style_cell(ws.cell(r, i + 2, f"='年間集計(VAE)'!{gl(i+2)}4"), AUTO, F_BODY, A_C, numfmt="0")
    style_cell(ws.cell(r, 14, f"=SUM(B{r}:M{r})"), AUTO, Font(bold=True, size=9), A_C, numfmt="0")

    r = LIST_C_TOP + 5
    style_cell(ws.cell(r, 1, "VAE率（/1,000人工呼吸器日）"), None, Font(size=9, bold=True), A_L)
    for i in range(13):
        c = gl(i + 2) if i < 12 else "N"
        style_cell(ws.cell(r, i + 2 if i < 12 else 14,
                           f"=IF({c}{LIST_C_TOP+4}=0,\"\",{c}{LIST_C_TOP+3}/{c}{LIST_C_TOP+4}*1000)"),
                   GREEN, F_RESULT, A_C, numfmt="0.00")

    ws.freeze_panes = "A7"
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:N{LIST_A_BOT}"
    return ws
