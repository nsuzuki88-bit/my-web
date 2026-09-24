# -*- coding: utf-8 -*-
"""共通定数・スタイル（CLABSI/VAE 自動判定ワークブック生成）"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

# ---- Allシートの構造 -------------------------------------------------
C0, CN = 3, 367            # 日付列 C(3) .. NC(367)  = 365日
NDAYS = CN - C0 + 1
NBLK = 399                 # Allシートの患者ブロック数（行6..1996、5行/ブロック）
ALL_LAST_ROW = 2000

def all_v(k):   return 6 + 5 * (k - 1)      # V有無 行（A列=ID）
def all_f(k):   return all_v(k) + 1         # FiO2 行
def all_p(k):   return all_v(k) + 2         # PEEP 行（A列=氏名）
def all_c(k):   return all_v(k) + 3         # C有無 行
def all_rm(k):  return all_v(k) + 4         # 部屋 行

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
CLIN_LAST_ROW = clin_base(NBLK) + 7          # 3197

# ---- 個別判定シート -------------------------------------------------
NSHEET = 40                # VAE-01 .. VAE-40
MVROWS = 60                # MV 1日目 .. 60日目
R_TOP = 15                 # 表の先頭行
R_BOT = R_TOP + MVROWS - 1 # 74
R_CALC = 78                # 内部計算スカラー行
NEPISODE = 6               # 1患者あたり判定できるMVエピソード数
R_MV = 80                  # 内部計算：MV日番号
R_EP = 81                  # 内部計算：エピソード番号

# ---- 一覧シート -----------------------------------------------------
LIST_A_TOP = 7                        # セクションA 先頭行（VAE-01..）
LIST_A_BOT = LIST_A_TOP + NSHEET - 1  # 46
LIST_B_HDR = LIST_A_BOT + 3           # 49  セクションB 見出し
LIST_B_TOP = LIST_B_HDR + 2           # 51  ブロック1
LIST_B_BOT = LIST_B_TOP + NBLK - 1    # 449
LIST_C_HDR = LIST_B_BOT + 3           # セクションC 見出し（月別自動集計）
LIST_C_TOP = LIST_C_HDR + 2

# ---- 名前付き範囲 ---------------------------------------------------
N_ALL  = "rALL"    # All!$A$1:$NC$2000     … 絶対列番号でINDEXする用
N_ALLD = "rALLD"   # All!$C$1:$NC$2000     … 日付列のみ（配列演算用）
N_CLIN = "rCLIN"   # 臨床所見!$A$1:$NC$3200

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


# ---- MV日判定の配列式（日付列 C..NC 全体に対するブール配列）---------
def mvb_array(vrow, frow, prow):
    """『その日がMV日か』の 1×365 ブール配列を返す式（先頭の = は付けない）"""
    v = f"INDEX({N_ALLD},{vrow},0)"
    f_ = f"INDEX({N_ALLD},{frow},0)"
    p = f"INDEX({N_ALLD},{prow},0)"
    return (f'((({v}="V有")+({v}<>"V有（NIV）")*(({f_}<>"")+({p}<>"")))>0)')


def cf_fill(hexrgb):
    """条件付き書式（dxf）用の塗りつぶし。元ファイルと同じく bgColor のみ指定する。"""
    return PatternFill(bgColor=hexrgb)
