# -*- coding: utf-8 -*-
"""VAE-01〜VAE-40：個別判定シート（Allシート＋臨床所見シートから全自動）"""
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter as gl
from common import *

TBL = [
    ("A", "MV日",                    5.3),
    ("B", "日付",                    9.7),
    ("C", "日最小\nPEEP\n(cmH2O)",   7.1),
    ("D", "日最小\nFiO2",            7.1),
    ("E", "最高\n体温\n(℃)",        7.1),
    ("F", "最低\n体温\n(℃)",        7.1),
    ("G", "WBC\n最高\n(/μL)",       7.9),
    ("H", "WBC\n最低\n(/μL)",       7.9),
    ("I", "新規\n抗菌薬\n開始",      8.8),
    ("J", "抗菌薬名・QADメモ",       15.9),
    ("K", "PVAP\n基準",              9.7),
    ("L", "検体・菌名・結果メモ",    19.4),
    ("M", "PEEP\n補正",              5.3),
    ("N", "体温/\nWBC\n異常",        6.2),
    ("O", "FiO2\n基線\n安定",        6.2),
    ("P", "FiO2\n悪化\n2日",         6.2),
    ("Q", "PEEP\n基線\n安定",        6.2),
    ("R", "PEEP\n悪化\n2日",         6.2),
    ("S", "VAC\n基準",               6.2),
    ("T", "VAE\n新規",               6.2),
    ("U", "IVAC",                    6.2),
    ("V", "PVAP",                    6.2),
    ("W", "判定",                    7.9),
    ("X", "#",                       3.5),
]
# 臨床所見シートの行オフセット → 判定シートの列
CLIN_MAP = {"E": 0, "F": 1, "G": 2, "H": 3, "I": 4, "J": 5, "K": 6, "L": 7}

FLOW = [
    "・VAC：MV3日目以降、直前2日（ベースライン期間）が安定/改善し、ベースライン初日比で FiO2 +20ポイント以上 または PEEP +3cmH2O以上 が2暦日持続",
    "・IVAC：VAC成立＋VAEウィンドウ（DOE±2日・MV3日目以降）内に 体温/WBC異常 かつ 新規抗菌薬（4QAD以上）",
    "・PVAP：IVAC成立＋ウィンドウ内に PVAP基準1〜3 のいずれか（除外菌に注意）",
    "・DOEから14日間（イベント期間）は新規VAEを判定しない。橙色の行＝VAEウィンドウ期間。",
    "・PEEP 0〜5cmH2O は同等（5として自動補正）。FiO2は％（例：40）で入力。",
    "・MVを1暦日以上完全に離脱した後の再装着は別エピソード（C9のエピソード番号を 2 に変更）。",
    "・A〜L列は自動取得です。体温・WBC・抗菌薬・PVAP基準は「臨床所見」シートに日付で入力してください。",
    "・判定は補助です。最終判定は原典プロトコールと診療記録に基づき感染管理担当者が行ってください。",
]


def build(wb, n):
    name = f"VAE-{n:02d}"
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False

    # ================= タイトル ======================================
    ws["A1"] = f"⑤ VAE 判定ワークシート（自動作成 {n:02d}）　― 人工呼吸器装着 {n} 人目の患者に自動で割り当てられます ―"
    ws["A1"].font = F_TITLE
    ws["A1"].fill = fill(NAVY)
    ws.merge_cells("A1:X1")
    ws.row_dimensions[1].height = 27.75
    ws["A2"] = ("Allシートに FiO2／PEEP を入力すると自動で反映されます。入力が必要なのは黄色のセル（C9 エピソード番号）だけです。"
                "日最小値＝その暦日で1時間を超えて維持された最低値。FiO2は小数（0.4）でも％（40）でも同じ判定になります。")
    ws["A2"].font = F_NOTE
    ws.merge_cells("A2:X2")

    # ================= 左：患者情報 ==================================
    L = f"'VAE対象一覧'"
    hdr = [
        ("ブロック番号",
         f'=IFERROR(INDEX({L}!$A${LIST_B_TOP}:$A${LIST_B_BOT},'
         f'MATCH({n},{L}!$G${LIST_B_TOP}:$G${LIST_B_BOT},0)),"")',
         "auto", "※自動割当。特定の患者に固定したいときは番号を直接入力"),
        ("患者ID",
         f'=IF($C$4="","",IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},$C$78+1)="","",'
         f'INDEX(All!$A$1:$A${ALL_LAST_ROW},$C$78+1)))', "auto", ""),
        ("氏名",
         f'=IF($C$4="","",IF(INDEX(All!$A$1:$A${ALL_LAST_ROW},$C$78+3)="","",'
         f'INDEX(All!$A$1:$A${ALL_LAST_ROW},$C$78+3)))', "auto", ""),
        ("部屋",
         f'=IF($H$78="","",IF(INDEX({N_ALL},$F$78,$H$78)="","",INDEX({N_ALL},$F$78,$H$78)))',
         "auto", ""),
        ("年齢",
         f'=IF($C$4="","",IF(INDEX({L}!$H${LIST_B_TOP}:$H${LIST_B_BOT},$C$4)="","",'
         f'INDEX({L}!$H${LIST_B_TOP}:$H${LIST_B_BOT},$C$4)))', "auto",
         '=IF($C$8="","",IF($C$8<18,"※18歳未満：成人VAEではなくPedVAEの定義で評価","成人VAE対象"))'),
        ("エピソード番号", 1, "input",
         '=IF($K$78="","",IF($K$78<=1,"このブロックのMVエピソードは1件","このブロックには全"&$K$78&"件のMVエピソードがあります"))'),
        ("MV1日目(挿管日)", "=$J$78", "auto", ""),
        ("MV日数(全体)", "=$L$78", "auto", ""),
        ("臓器提供同意日",
         f'=IF($C$4="","",IF(INDEX({L}!$I${LIST_B_TOP}:$I${LIST_B_BOT},$C$4)="","",'
         f'INDEX({L}!$I${LIST_B_TOP}:$I${LIST_B_BOT},$C$4)))', "auto",
         "該当時のみ。同意取得日以降のDOEは報告しない（一覧シートI列で入力）"),
    ]
    for i, (label, val, kind, note) in enumerate(hdr):
        r = 4 + i
        ws.cell(r, 1, label).font = Font(bold=True, size=9)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        c = ws.cell(r, 3, val)
        style_cell(c, YELLOW if kind == "input" else AUTO, Font(bold=True, size=10), A_C)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        if note:
            ws.cell(r, 5, note).font = F_NOTE
    for a, fmt in (("C7", "General"), ("C10", "yyyy/m/d"), ("C12", "yyyy/m/d")):
        ws[a].number_format = fmt

    # ================= 中：判定サマリー ==============================
    ws["J4"] = "判定サマリー"
    ws["J4"].font = F_SEC
    summ = [
        ("MV日数（本エピソード）", "=$I$78", "0"),
        ("エピソード総数",         "=$K$78", "0"),
        ("検出VAE件数",            "=SUM($T$15:$T$74)", "0"),
        ("1件目 DOE",  '=IFERROR(INDEX($B$15:$B$74,MATCH(1,$X$15:$X$74,0)),"")', "yyyy/m/d"),
        ("1件目 判定", '=IFERROR(INDEX($W$15:$W$74,MATCH(1,$X$15:$X$74,0)),"")', "General"),
        ("2件目 DOE",  '=IFERROR(INDEX($B$15:$B$74,MATCH(2,$X$15:$X$74,0)),"")', "yyyy/m/d"),
        ("2件目 判定", '=IFERROR(INDEX($W$15:$W$74,MATCH(2,$X$15:$X$74,0)),"")', "General"),
    ]
    for i, (label, f_, fmt) in enumerate(summ):
        r = 5 + i
        ws.cell(r, 10, label).font = Font(size=9)
        style_cell(ws.cell(r, 11, f_), GREEN, F_RESULT, A_C, numfmt=fmt)

    # ================= 右：判定の流れ ================================
    ws["M4"] = "判定の流れ（NHSN/JHAIS）"
    ws["M4"].font = F_SEC
    ws.merge_cells(start_row=4, start_column=13, end_row=4, end_column=24)
    for i, t in enumerate(FLOW):
        r = 5 + i
        ws.cell(r, 13, t).font = F_NOTE
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=24)

    # ================= 表の帯 ========================================
    ws["A13"] = "自動取得（Allシート／臨床所見シート）"
    style_cell(ws["A13"], NAVY, F_HDR, A_C)
    ws.merge_cells("A13:L13")
    ws["M13"] = "自動計算（編集しないでください）"
    style_cell(ws["M13"], GRAY, F_HDR, A_C)
    ws.merge_cells("M13:X13")

    # ================= 表ヘッダー ====================================
    for i, (col, title, width) in enumerate(TBL):
        c = ws.cell(14, i + 1, title)
        style_cell(c, NAVY if i < 12 else GRAY, F_HDR, A_CW)
        ws.column_dimensions[col].width = width
    ws.row_dimensions[14].height = 48

    # ================= 本体 ==========================================
    for d in range(1, MVROWS + 1):
        r = R_TOP + d - 1
        Y = f"$Y{r}"

        ws.cell(r, 1, d)
        style_cell(ws.cell(r, 1), AUTO, F_BODY, A_C)

        # 日付
        style_cell(ws.cell(r, 2, f'=IF(OR($J$78="",{d}>$I$78),"",$J$78+{d}-1)'),
                   AUTO, F_BODY, A_C, numfmt="yyyy/m/d")

        # 日付列インデックス（Y列・非表示）
        ws.cell(r, 25, f'=IF($B{r}="","",$H$78+{d}-1)')

        # PEEP / FiO2 （Allシートから）
        for col, rowref in (("C", "$E$78"), ("D", "$D$78")):
            idx = f"INDEX({N_ALL},{rowref},{Y})"
            style_cell(ws[f"{col}{r}"], AUTO, F_BODY, A_C)
            ws[f"{col}{r}"] = f'=IF({Y}="","",IF({idx}="","",{idx}))'

        # 体温・WBC・抗菌薬・PVAP（臨床所見シートから）
        for col, off in CLIN_MAP.items():
            idx = f"INDEX({N_CLIN},$G$78+{off},{Y})"
            style_cell(ws[f"{col}{r}"], AUTO,
                       F_BODY, A_C if col not in ("J", "L") else A_L)
            ws[f"{col}{r}"] = f'=IF({Y}="","",IF({idx}="","",{idx}))'

        ws[f"C{r}"].number_format = "0.##"
        ws[f"D{r}"].number_format = "0.##"
        ws[f"E{r}"].number_format = "0.0"
        ws[f"F{r}"].number_format = "0.0"
        ws[f"G{r}"].number_format = "#,##0"
        ws[f"H{r}"].number_format = "#,##0"

        # ---- 自動計算列 --------------------------------------------
        lo, hi = max(R_TOP, r - 2), min(R_BOT, r + 2)      # VAEウィンドウ
        back = max(R_TOP, r - 13)                           # イベント期間14日

        ws[f"M{r}"] = f'=IF(C{r}="","",MAX(C{r},{PEEP_FLOOR}))'
        # Z列（非表示）＝FiO2を小数に正規化（40％入力でも0.4入力でも同じ判定になる）
        ws[f"Z{r}"] = f'=IF(D{r}="","",{fio2n(f"D{r}")})'
        ws[f"N{r}"] = (f'=IF(COUNT(E{r}:H{r})=0,"",IF(OR(AND(E{r}<>"",E{r}>38),'
                       f'AND(F{r}<>"",F{r}<36),AND(G{r}<>"",G{r}>=12000),'
                       f'AND(H{r}<>"",H{r}<=4000)),1,0))')

        if d >= 3:
            ws[f"O{r}"] = f'=IF(COUNT(Z{r-2},Z{r-1})<2,0,IF(Z{r-1}<=Z{r-2},1,0))'
            # 0.6-0.4 が 0.19999999999999996 になる二進小数の誤差を ROUND で吸収する
            ws[f"P{r}"] = (f'=IF(COUNT(Z{r-2},Z{r},Z{r+1})<3,0,'
                           f'IF(AND(ROUND(Z{r}-Z{r-2},6)>={FIO2_RISE},'
                           f'ROUND(Z{r+1}-Z{r-2},6)>={FIO2_RISE}),1,0))')
            ws[f"Q{r}"] = f'=IF(COUNT(M{r-2},M{r-1})<2,0,IF(M{r-1}<=M{r-2},1,0))'
            ws[f"R{r}"] = (f'=IF(COUNT(M{r-2},M{r},M{r+1})<3,0,'
                           f'IF(AND(ROUND(M{r}-M{r-2},6)>={PEEP_RISE},'
                           f'ROUND(M{r+1}-M{r-2},6)>={PEEP_RISE}),1,0))')
            ws[f"S{r}"] = f'=IF(OR(AND(O{r}=1,P{r}=1),AND(Q{r}=1,R{r}=1)),1,0)'
            ws[f"T{r}"] = (f'=IF(AND(S{r}=1,COUNTIF(T{back}:T{r-1},1)=0,'
                           f'OR($C$12="",B{r}<$C$12)),1,0)')
            ws[f"U{r}"] = (f'=IF(T{r}<>1,0,IF(AND(COUNTIFS(A{lo}:A{hi},">=3",N{lo}:N{hi},1)>0,'
                           f'COUNTIFS(A{lo}:A{hi},">=3",I{lo}:I{hi},"はい")>0),1,0))')
            # K列は数式のため COUNTIFS(...,"<>") では空文字も「非空白」と数えられてしまう。
            # SUMPRODUCT で <>"" を評価する。
            ws[f"V{r}"] = (f'=IF(U{r}<>1,0,IF(SUMPRODUCT((A{lo}:A{hi}>=3)*'
                           f'(K{lo}:K{hi}<>""))>0,1,0))')
        else:
            for col in "OPQRSTUV":
                ws[f"{col}{r}"] = 0

        ws[f"W{r}"] = f'=IF(V{r}=1,"PVAP",IF(U{r}=1,"IVAC",IF(T{r}=1,"VAC","")))'
        ws[f"X{r}"] = f'=IF(T{r}=1,COUNTIF(T$15:T{r},1),"")'

        for col in "MNOPQRSTUVX":
            style_cell(ws[f"{col}{r}"], AUTO, F_BODY, A_C)
        style_cell(ws[f"W{r}"], GREEN, F_RESULT, A_C)

    # ================= 内部計算 ======================================
    ws.cell(R_CALC, 1, "▼ 内部計算（編集しないでください）").font = F_NOTE
    mvb = mvb_array("$C$78", "$C$78+1", "$C$78+2")
    colrng = f"COLUMN(All!$C$1:${gl(CN)}$1)"
    # 1列左にずらした「前日がMV日か」（先頭列は常に0）
    SHIFT = f"All!$B$1:${gl(CN-1)}${ALL_LAST_ROW}"      # 1列左にずらした範囲
    prevd = (f'(((INDEX({SHIFT},$C$78,0)="V有")'
             f'+(INDEX({SHIFT},$C$78,0)="")'
             f'*((INDEX({SHIFT},$C$78+1,0)<>"")+(INDEX({SHIFT},$C$78+2,0)<>"")))>0)'
             f'*({colrng}>{C0})')
    start = f"({mvb}*(1-{prevd}))"

    # 各エピソードの開始列を N78..S78 に順に求める（最大6エピソード）。
    # AGGREGATE は LibreOffice / 旧Excel で使えないため SUMPRODUCT(MIN(IF(...))) を使う。
    ep_cols = {}
    for i in range(NEPISODE):
        col = 14 + i                       # N, O, P, Q, R, S
        if i == 0:
            f_ = f'=IF($C$78="",0,SUMPRODUCT(MIN(IF({start},{colrng}))))'
        else:
            prev_ref = f"${gl(13 + i)}$78"
            f_ = (f'=IF(OR($C$78="",{prev_ref}=0),0,'
                  f'SUMPRODUCT(MIN(IF({start}*({colrng}>{prev_ref}),{colrng}))))')
        ep_cols[col] = (f"第{i+1}エピソード開始列", f_)

    calc = {
        3:  ('V有無行',        f'=IF($C$4="","",{BLOCK0}+{BSTEP}*($C$4-1))'),
        4:  ('FiO2行',          f'=IF($C$78="","",$C$78+1)'),
        5:  ('PEEP行',          f'=IF($C$78="","",$C$78+2)'),
        6:  ('部屋行',          f'=IF($C$78="","",$C$78+4)'),
        7:  ('臨床所見基準行',  f'=IF($C$4="","",6+8*($C$4-1))'),
        8:  ('開始列',          f'=IFERROR(IF(INDEX($N$78:${gl(13+NEPISODE)}$78,1,$C$9)=0,"",'
                                f'INDEX($N$78:${gl(13+NEPISODE)}$78,1,$C$9)),"")'),
        9:  ('エピソード日数',  f'=IF($H$78="",0,MIN(SUMPRODUCT(MIN(IF((1-{mvb})*'
                                f'({colrng}>=$H$78),{colrng},{CN+1}))),{CN+1})-$H$78)'),
        10: ('開始日',          f'=IF($H$78="","",INDEX(All!$C${DATE_ROW}:${gl(CN)}${DATE_ROW},1,$H$78-{C0-1}))'),
        11: ('エピソード総数',  f'=IF($C$78="","",SUMPRODUCT(--({start}>0)))'),
        12: ('MV総日数',        f'=IF($C$78="","",SUMPRODUCT(--({mvb})))'),
        **ep_cols,
    }
    for col, (label, f_) in calc.items():
        ws.cell(R_CALC - 1, col, label).font = Font(size=7, color="999999")
        ws.cell(R_CALC, col, f_)
    ws.cell(R_CALC, 10).number_format = "yyyy/m/d"
    for r in (R_CALC - 1, R_CALC):
        ws.row_dimensions[r].hidden = True
    ws.column_dimensions["Y"].hidden = True
    ws.column_dimensions["Z"].hidden = True

    # ================= 入力規則・条件付き書式 ========================
    dv = DataValidation(type="whole", operator="between", formula1=1, formula2=NEPISODE,
                        allow_blank=True, showErrorMessage=True,
                        error=f"エピソード番号は1〜{NEPISODE}で入力してください。",
                        errorTitle="入力エラー")
    ws.add_data_validation(dv)
    dv.add("C9")

    ws.conditional_formatting.add(
        f"A{R_TOP}:L{R_BOT}",
        # DOE±2日（＝VAEウィンドウ期間）の行を橙色にする。
        # 開始・終了とも表の範囲（15〜74行）にクランプする。
        Rule(type="expression",
             formula=[f"SUM(OFFSET($T${R_TOP},MAX(0,ROW()-{R_TOP+2}),0,"
                      f"MIN({MVROWS-1},ROW()-{R_TOP-2})-MAX(0,ROW()-{R_TOP+2})+1,1))>0"],
             dxf=DifferentialStyle(fill=cf_fill(ORANGE)), stopIfTrue=False))
    ws.conditional_formatting.add(
        f"W{R_TOP}:W{R_BOT}",
        Rule(type="expression", formula=[f'$W{R_TOP}<>""'],
             dxf=DifferentialStyle(fill=cf_fill(PINK)), stopIfTrue=False))

    ws.freeze_panes = f"C{R_TOP}"
    ws.sheet_view.zoomScale = 90

    # ================= 印刷設定（記録として印刷できるように）=========
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{R_TOP-1}:{R_TOP-1}"
    ws.print_area = f"A1:X{R_BOT}"
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    return ws
