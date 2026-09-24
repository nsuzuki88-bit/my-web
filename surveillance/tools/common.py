# -*- coding: utf-8 -*-
"""共通定数・スタイル（CLABSI/VAE 自動判定ワークブック生成）

Allシートの構造（日付行・ブロック先頭行・最終行・ブロック数）は年度ファイルによって
異なるため、build.py が実ファイルから読み取って configure() で設定する。
シート生成モジュールは configure() の後に import すること。
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

# ---- Allシートの構造（既定値＝旧テンプレート。configure()で上書きされる）----
C0, CN = 3, 367            # 日付列 C(3) .. NC(367)  = 365日
NDAYS = CN - C0 + 1
DATE_ROW = 5               # 日付が入っている行
BLOCK0 = 6                 # 1ブロック目の先頭行（＝V有無行）
BSTEP = 5                  # 1ブロックの行数
NBLK = 399                 # Allシートの患者ブロック数
ALL_LAST_ROW = 2000
NSHEET = 40                # 個別判定シートの枚数 VAE-01 .. VAE-nn
MV_MIN = 4                 # 判定シートを割り当てるMV日数の下限
                           # （JHAIS：VAEの対象は人工呼吸を暦日で4日以上）


def configure(**kw):
    """Allシートの実構造を反映する"""
    g = globals()
    for k, v in kw.items():
        assert k in g, k
        g[k] = v
    g["NDAYS"] = g["CN"] - g["C0"] + 1
    g["CLIN_LAST_ROW"] = clin_base(g["NBLK"]) + 7
    g["LIST_A_BOT"] = g["LIST_A_TOP"] + g["NSHEET"] - 1
    g["LIST_B_HDR"] = g["LIST_A_BOT"] + 3
    g["LIST_B_TOP"] = g["LIST_B_HDR"] + 2
    g["LIST_B_BOT"] = g["LIST_B_TOP"] + g["NBLK"] - 1
    g["LIST_C_HDR"] = g["LIST_B_BOT"] + 3
    g["LIST_C_TOP"] = g["LIST_C_HDR"] + 2


# Allシートの各行（k＝ブロック番号）
def all_v(k):    return BLOCK0 + BSTEP * (k - 1)   # V有無 行（A列＝「ID」ラベル）
def all_f(k):    return all_v(k) + 1               # FiO2  行（A列＝患者IDの値）
def all_p(k):    return all_v(k) + 2               # PEEP  行（A列＝「氏名」ラベル）
def all_c(k):    return all_v(k) + 3               # C有無 行（A列＝氏名の値）
def all_rm(k):   return all_v(k) + 4               # 部屋  行
def all_id(k):   return all_v(k) + 1               # 患者IDが入っているA列の行
def all_name(k): return all_v(k) + 3               # 氏名が入っているA列の行

# ---- 臨床所見シートの構造（8行/ブロック）----------------------------
CLIN_ROWS = [
    ("最高体温(℃)",        "num1"),
    ("最低体温(℃)",        "num1"),
    ("WBC最高(/μL)",       "int"),
    ("WBC最低(/μL)",       "int"),
    ("新規抗菌薬開始",      "list_abx"),
    ("抗菌薬名・QADメモ",   "text"),
    ("PVAP基準",            "list_pvap"),
    ("検体・菌名・結果メモ", "text"),
]
def clin_base(k): return 6 + 8 * (k - 1)
CLIN_LAST_ROW = clin_base(NBLK) + 7

# ---- 個別判定シート -------------------------------------------------
MVROWS = 60                # MV 1日目 .. 60日目
R_TOP = 15                 # 表の先頭行
R_BOT = R_TOP + MVROWS - 1 # 74
R_CALC = 78                # 内部計算スカラー行
NEPISODE = 6               # 1患者あたり判定できるMVエピソード数

# ---- 一覧シート -----------------------------------------------------
LIST_A_TOP = 7                        # セクションA 先頭行（VAE-01..）
LIST_A_BOT = LIST_A_TOP + NSHEET - 1
LIST_B_HDR = LIST_A_BOT + 3           # セクションB 見出し
LIST_B_TOP = LIST_B_HDR + 2           # ブロック1
LIST_B_BOT = LIST_B_TOP + NBLK - 1
LIST_C_HDR = LIST_B_BOT + 3           # セクションC 見出し（月別自動集計）
LIST_C_TOP = LIST_C_HDR + 2

# ---- 名前付き範囲 ---------------------------------------------------
N_ALL  = "rALL"    # All!$A$1:$NC$<last>   … 絶対列番号でINDEXする用
N_ALLD = "rALLD"   # All!$C$1:$NC$<last>   … 日付列のみ（配列演算用）
N_CLIN = "rCLIN"   # 臨床所見!$A$1:$NC$<last>

# ---- 判定のしきい値 -------------------------------------------------
FIO2_RISE = 0.2    # FiO2 の上昇幅（小数。20ポイント＝0.2）
PEEP_RISE = 3      # PEEP の上昇幅（cmH2O）
PEEP_FLOOR = 5     # PEEP 0〜5 は 5 として扱う

# ---- 配色（元ワークシートの配色を踏襲）------------------------------
NAVY   = "1F4E79"   # 見出し（濃紺）
GRAY   = "7F7F7F"   # 自動計算の見出し
YELLOW = "FFF6C8"   # 入力セル
AUTO   = "EEF1F5"   # 自動計算セル
GREEN  = "E2F0D9"   # 判定結果
ORANGE = "F8CBAD"   # VAEウィンドウ
PINK   = "FFC7CE"   # 判定あり
RED    = "C00000"
BAND   = "DCE6F1"   # ブロック区切り

thin = Side(style="thin", color="B4C6E7")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)

def fill(hexrgb):
    return PatternFill("solid", fgColor=hexrgb)

F_TITLE  = Font(bold=True, size=14, color="FFFFFF")
F_HDR    = Font(bold=True, size=9, color="FFFFFF")
F_LBL    = Font(bold=True, size=10)
F_BODY   = Font(size=9)
F_NOTE   = Font(size=9, color="555555")
F_RESULT = Font(bold=True, size=9, color=RED)
F_SEC    = Font(bold=True, size=11, color=NAVY)

A_C = Alignment(horizontal="center", vertical="center")
A_CW = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_L = Alignment(horizontal="left", vertical="center")
A_LW = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_cell(c, f=None, font=None, align=None, border=True, numfmt=None):
    if f:
        c.fill = fill(f)
    if font:
        c.font = font
    if align:
        c.alignment = align
    if border:
        c.border = BOX
    if numfmt:
        c.number_format = numfmt
    return c


def fio2n(expr):
    """FiO2を小数に正規化する（40→0.4、0.4→0.4）。
    ％表記と小数表記のどちらで入力されていても同じ判定になるようにする。"""
    return f"IF({expr}>1,{expr}/100,{expr})"


# ---- MV日判定 -------------------------------------------------------
# V有無欄が入っていればそれが優先。空欄のときだけ FiO2/PEEP の有無で補う。
#   「V有」        → MV日
#   「V無」「V有（NIV）」→ MV日ではない（NIVはVAEの対象外）
#   空欄            → FiO2 か PEEP に入力があれば MV日
def mvb_array(vrow, frow, prow):
    """『その日がMV日か』の 1×365 ブール配列を返す式（先頭の = は付けない）"""
    v = f"INDEX({N_ALLD},{vrow},0)"
    f_ = f"INDEX({N_ALLD},{frow},0)"
    p = f"INDEX({N_ALLD},{prow},0)"
    return f'((({v}="V有")+({v}="")*(({f_}<>"")+({p}<>"")))>0)'


def mv_scalar(v, f, p):
    """『その日がMV日か』のスカラー式（条件付き書式用）"""
    return f'OR({v}="V有",AND({v}="",OR({f}<>"",{p}<>"")))'


def cf_fill(hexrgb):
    """条件付き書式（dxf）用の塗りつぶし。元ファイルと同じく bgColor のみ指定する。"""
    return PatternFill(bgColor=hexrgb)
