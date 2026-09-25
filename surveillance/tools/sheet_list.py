# -*- coding: utf-8 -*-
"""VAE対象一覧

  セクションA：VAC と判定された患者の一覧（臨床所見・個別判定シートの割当元）
  セクションB：Allシート全ブロックのスキャンと VAC 判定（Allのデータから直接判定）
  セクションC：月別 自動集計
  セクションD：個別判定シートの判定（IVAC/PVAP への格上げ結果）
"""
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter as gl
from common import *

# セクションA：VAC患者の一覧
SEC_A = [("No", 5), ("ブロック", 8), ("患者ID", 12), ("氏名", 13), ("部屋", 7),
         ("MV1日目", 11), ("MV日数", 8), ("DOE①", 11), ("判定①", 9),
         ("DOE②", 11), ("判定②", 9), ("臨床所見", 10), ("判定\nシート", 7), ("シートへ", 11)]

# セクションB：全ブロックのスキャン
SEC_B = [("ブロック", 8), ("患者ID", 12), ("氏名", 13), ("MV日数", 8), ("初回MV日", 11),
         ("最終MV日", 11), ("DOE①", 11), ("DOE②", 11), ("判定", 9), ("VAC\n番号", 7),
         ("年齢\n(入力)", 8), ("臓器提供\n同意日(入力)", 12), ("備考(入力)", 22)]

WIDTHS = [7, 12, 13, 12, 12, 11, 11, 11, 10, 11, 11, 12, 22, 11]


def build(wb):
    ws = wb.create_sheet("VAE対象一覧")
    ws.sheet_properties.tabColor = "C00000"
    ws.sheet_view.showGridLines = False

    ws["A1"] = "VAE 対象一覧（AllシートのFiO2/PEEP入力から自動判定）"
    ws["A1"].font = F_TITLE
    ws["A1"].fill = fill("C00000")
    ws.merge_cells("A1:N1")
    ws.row_dimensions[1].height = 27.75
    ws["A2"] = ("① AllシートのFiO2/PEEPから、セクションBで全ブロックのVAC判定（DOE）を自動で行います。"
                "② VACと判定された患者だけがセクションAに並び、臨床所見シートの同じ番号のスロットに入力できます。"
                "③ 臨床所見に入力した患者だけに個別判定シート（VAE-01〜）が割り当てられ、IVAC・PVAPまで判定されます。")
    ws["A3"] = ("MV日の判定：AllのV有無行が「V有」の日。「V無」「V有（NIV）」の日は除外します（非侵襲換気はVAEの対象外）。"
                "V有無が空欄の日は、FiO2／PEEPに入力があればMV日とみなします。"
                "FiO2は小数（0.4）でも％（40）でも同じ判定になります。黄色のセルのみ手入力です。")
    for r in (2, 3):
        ws.cell(r, 1).font = F_NOTE
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)

    # ============ セクションA：VAC患者の一覧 ==========================
    ws.cell(5, 1, f"A. VACと判定された患者（DOEの早い順。臨床所見シートのスロット番号と対応・最大{NCLIN}名）")
    ws.cell(5, 1).font = F_SEC
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=14)
    for i, (t, _w) in enumerate(SEC_A):
        style_cell(ws.cell(6, i + 1, t), NAVY, F_HDR, A_CW)
    ws.row_dimensions[6].height = 32

    B = lambda col: f"${col}${LIST_B_TOP}:${col}${LIST_B_BOT}"   # セクションBの列
    for n in range(1, NCLIN + 1):
        r = LIST_A_TOP + n - 1
        # VAC番号がnのブロックを引く
        m = f"MATCH({n},{B('J')},0)"
        pick = lambda col: f'=IFERROR(INDEX({B(col)},{m}),"")'
        cb = clin_base(n)
        vals = [
            (1, n, "0"),
            (2, pick("A"), "0"),                      # ブロック
            (3, pick("B"), "General"),                # 患者ID
            (4, pick("C"), "General"),                # 氏名
            (5, f'=IF($B{r}="","",IFERROR(INDEX({N_ALL},{BLOCK0}+{BSTEP}*($B{r}-1)+4,'
                f'MATCH($F{r},All!$C${DATE_ROW}:${gl(CN)}${DATE_ROW},0)+{C0-1}),""))', "General"),
            (6, pick("E"), "yyyy/m/d"),               # MV1日目
            (7, pick("D"), "0"),                      # MV日数
            (8, pick("G"), "yyyy/m/d"),               # DOE①
            (9, f'=IF($H{r}="","",IF($M{r}="","VAC",'
                f'IFERROR(INDEX($C${LIST_D_TOP}:$C${LIST_D_BOT},$M{r}),"VAC")))', "General"),
            (10, pick("H"), "yyyy/m/d"),              # DOE②
            (11, f'=IF($J{r}="","",IF($M{r}="","VAC",'
                 f'IFERROR(INDEX($E${LIST_D_TOP}:$E${LIST_D_BOT},$M{r}),"VAC")))', "General"),
            (12, f'=IF($B{r}="","",IF(COUNTA(臨床所見!$C${cb}:${gl(CN)}${cb+7})=0,"未入力","入力あり"))',
             "General"),
            (13, f'=IF($L{r}<>"入力あり","",COUNTIF($L${LIST_A_TOP}:$L{r},"入力あり"))', "0"),
            (14, f'=IF($M{r}="","",IF($M{r}>{NSHEET},"シート不足",'
                 f'HYPERLINK("#\'VAE-"&TEXT($M{r},"00")&"\'!A1","VAE-"&TEXT($M{r},"00"))))', "General"),
        ]
        for col, v, fmt in vals:
            style_cell(ws.cell(r, col, v), AUTO, F_BODY, A_C, numfmt=fmt)
        style_cell(ws.cell(r, 1), BAND, Font(bold=True, size=9), A_C)
        for col in (9, 11):
            style_cell(ws.cell(r, col), GREEN, F_RESULT, A_C)
        style_cell(ws.cell(r, 14), BAND, Font(size=9, color="0563C1", underline="single"), A_C)

    for col in (9, 11):
        ws.conditional_formatting.add(
            f"{gl(col)}{LIST_A_TOP}:{gl(col)}{LIST_A_BOT}",
            Rule(type="cellIs", operator="notEqual", formula=['""'],
                 dxf=DifferentialStyle(fill=cf_fill(PINK)), stopIfTrue=False))
    ws.conditional_formatting.add(
        f"L{LIST_A_TOP}:L{LIST_A_BOT}",
        Rule(type="cellIs", operator="equal", formula=['"未入力"'],
             dxf=DifferentialStyle(fill=cf_fill(YELLOW)), stopIfTrue=False))

    # ============ セクションB：全ブロックのスキャンとVAC判定 ==========
    ws.cell(LIST_B_HDR, 1, f"B. Allシート 先頭{NBLK}ブロックのスキャン　"
                           f"※MV日数が{MV_MIN}日以上のブロックについて、AllのFiO2/PEEPから"
                           f"DOEを直接判定します（イベント期間14日の重複排除まで適用）")
    ws.cell(LIST_B_HDR, 1).font = F_SEC
    ws.merge_cells(start_row=LIST_B_HDR, start_column=1, end_row=LIST_B_HDR, end_column=14)
    for i, (t, _w) in enumerate(SEC_B):
        style_cell(ws.cell(LIST_B_HDR + 1, i + 1, t), NAVY, F_HDR, A_CW)
    ws.row_dimensions[LIST_B_HDR + 1].height = 32

    dates = f"All!$C${DATE_ROW}:${gl(CN)}${DATE_ROW}"
    colrng = f"COLUMN({N_D0})"
    NONE_COL = CN + 1000
    for k in range(1, NBLK + 1):
        r = LIST_B_TOP + k - 1
        idr, nmr = all_id(k), all_name(k)
        mvb = mvb_array(all_v(k), all_f(k), all_p(k))
        ws.cell(r, 1, k)
        ws.cell(r, 2, f'=IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},{idr})="","",'
                      f'INDEX(All!$A$1:$A${ALL_LAST_ROW},{idr}))')
        ws.cell(r, 3, f'=IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},{nmr})="","",'
                      f'INDEX(All!$A$1:$A${ALL_LAST_ROW},{nmr}))')
        ws.cell(r, 4, f"=SUMPRODUCT(--({mvb}))")
        ws.cell(r, 5, f'=IF($D{r}=0,"",INDEX({dates},1,MATCH(TRUE,INDEX({mvb},0),0)))')
        ws.cell(r, 6, f'=IF($D{r}=0,"",LOOKUP(2,1/{mvb},{dates}))')
        # P・Q列（非表示）＝DOEの列番号。Qは1件目から14日以降で探す
        ws.cell(r, 16, doe_col_formula(k, f"$D{r}"))
        ws.cell(r, 17, doe_col_formula(k, f"$D{r}", f"$P{r}"))
        ws.cell(r, 7, f'=IF($P{r}>{CN},"",INDEX({dates},1,$P{r}-{C0-1}))')
        ws.cell(r, 8, f'=IF($Q{r}>{CN},"",INDEX({dates},1,$Q{r}-{C0-1}))')
        ws.cell(r, 9, f'=IF($G{r}="","","VAC")')
        # VAC番号＝DOE①の早い順（同じ日はブロック番号順）。Allは日付順に入力していくので、
        # 新しくVACになった患者は必ず末尾に付き、入力済みのスロットの並びが変わらない
        G, A_ = f"$G${LIST_B_TOP}:$G${LIST_B_BOT}", f"$A${LIST_B_TOP}:$A${LIST_B_BOT}"
        ws.cell(r, 10, f'=IF($I{r}="","",COUNTIFS({G},"<"&$G{r})'
                       f'+COUNTIFS({G},$G{r},{A_},"<="&$A{r}))')
        for col, fmt in ((1, "0"), (2, "General"), (3, "General"), (4, "0"),
                         (5, "yyyy/m/d"), (6, "yyyy/m/d"), (7, "yyyy/m/d"),
                         (8, "yyyy/m/d"), (9, "General"), (10, "0")):
            style_cell(ws.cell(r, col), AUTO, F_BODY, A_C, numfmt=fmt)
        style_cell(ws.cell(r, 1), BAND, Font(bold=True, size=9), A_C)
        style_cell(ws.cell(r, 9), GREEN, F_RESULT, A_C)
        for col, fmt in ((11, "0"), (12, "yyyy/m/d"), (13, "General")):
            style_cell(ws.cell(r, col), YELLOW, F_BODY,
                       A_C if col < 13 else A_L, numfmt=fmt)
    for c in ("P", "Q"):
        ws.column_dimensions[c].hidden = True

    ws.conditional_formatting.add(
        f"A{LIST_B_TOP}:M{LIST_B_BOT}",
        Rule(type="expression", formula=[f'$I{LIST_B_TOP}="VAC"'],
             dxf=DifferentialStyle(font=Font(bold=True), fill=cf_fill(PINK)),
             stopIfTrue=False))

    # スキャン範囲より後ろにデータが残っていないかの見張り
    beyond = BLOCK0 + BSTEP * NBLK
    if beyond <= ALL_LAST_ROW:
        ws.cell(LIST_B_HDR, 1).comment = None
        c = ws.cell(LIST_B_BOT + 1, 1,
                    f'=IF(COUNTA(All!$A${beyond}:$A${ALL_LAST_ROW})=0,'
                    f'"（スキャン範囲より後ろにデータはありません）",'
                    f'"⚠ ブロック{NBLK}より後ろにもデータがあります。tools/build.py の第4引数で'
                    f'スキャン範囲を広げてください")')
        c.font = Font(size=9, bold=True, color=RED)
        ws.merge_cells(start_row=LIST_B_BOT + 1, start_column=1,
                       end_row=LIST_B_BOT + 1, end_column=14)

    # ============ セクションC：月別集計は「VAE集計」シートへ ===============
    ws.cell(LIST_C_HDR, 1, "C. 月別集計（件数・使用比・発生率）は「VAE集計」シートにあります")
    ws.cell(LIST_C_HDR, 1).font = F_SEC
    ws.merge_cells(start_row=LIST_C_HDR, start_column=1, end_row=LIST_C_HDR, end_column=10)
    c = ws.cell(LIST_C_TOP, 1, '=HYPERLINK("#\'VAE集計\'!A1","▶ VAE集計シートを開く")')
    c.font = Font(size=10, color="0563C1", underline="single")

    # ============ セクションD：個別判定シートの判定 ===================
    ws.cell(LIST_D_HDR, 1, f"D. 個別判定シート（VAE-01〜VAE-{NSHEET:02d}）の判定　"
                           f"※臨床所見に入力した患者から順に割り当てられます。上のセクションAに反映されます")
    ws.cell(LIST_D_HDR, 1).font = F_SEC
    ws.merge_cells(start_row=LIST_D_HDR, start_column=1, end_row=LIST_D_HDR, end_column=14)
    for i, t in enumerate(("No", "ブロック", "判定①", "DOE①", "判定②", "DOE②", "患者ID", "氏名")):
        style_cell(ws.cell(LIST_D_HDR + 1, i + 1, t), NAVY, F_HDR, A_C)
    for n in range(1, NSHEET + 1):
        r = LIST_D_TOP + n - 1
        s = f"'VAE-{n:02d}'"
        g = lambda ref: f'=IF({s}!$C$4="","",{s}!{ref})'
        for col, v, fmt in ((1, n, "0"), (2, g("$C$4"), "0"), (3, g("$K$9"), "General"),
                            (4, g("$K$8"), "yyyy/m/d"), (5, g("$K$11"), "General"),
                            (6, g("$K$10"), "yyyy/m/d"), (7, g("$C$5"), "General"),
                            (8, g("$C$6"), "General")):
            style_cell(ws.cell(r, col, v), AUTO, F_BODY, A_C, numfmt=fmt)
        style_cell(ws.cell(r, 1), BAND, Font(bold=True, size=9), A_C)
        for col in (3, 5):
            style_cell(ws.cell(r, col), GREEN, F_RESULT, A_C)

    for i, w in enumerate(WIDTHS):
        ws.column_dimensions[gl(i + 1)].width = w
    ws.freeze_panes = "A7"
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:N{LIST_A_BOT}"
    return ws
