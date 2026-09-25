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
# Allシート上部の日ごとの集計行（A列のラベルで探す。configure()で上書きされる）
ROW_NIV = 1                # 人工呼吸器（NIV）  … COUNTIF "V有（NIV）"
ROW_V = 2                  # 人工呼吸器使用患者数 … COUNTIF "V有"
ROW_C = 3                  # CV挿入患者数       … COUNTIF "C有"
ROW_ICU = 4                # ICU入室患者数      … COUNTIF "C有"＋"C無"
NSHEET = 10                # 個別判定シートの枚数 VAE-01 .. VAE-nn
NCLIN = 40                 # 臨床所見シートのスロット数（＝VAC判定された患者の上限）
NBLK_MAX = 1200            # 一覧セクションBでスキャンするブロック数の上限
MV_MIN = 4                 # VAC判定を行うMV日数の下限
                           # （JHAIS：VAEの対象は人工呼吸を暦日で4日以上）


def configure(**kw):
    """Allシートの実構造を反映する"""
    g = globals()
    for k, v in kw.items():
        assert k in g, k
        g[k] = v
    g["NDAYS"] = g["CN"] - g["C0"] + 1
    g["CLIN_LAST_ROW"] = clin_base(g["NCLIN"]) + 7
    g["LIST_A_BOT"] = g["LIST_A_TOP"] + g["NCLIN"] - 1
    g["LIST_B_HDR"] = g["LIST_A_BOT"] + 3
    g["LIST_B_TOP"] = g["LIST_B_HDR"] + 2
    g["LIST_B_BOT"] = g["LIST_B_TOP"] + g["NBLK"] - 1
    g["LIST_C_HDR"] = g["LIST_B_BOT"] + 3
    g["LIST_C_TOP"] = g["LIST_C_HDR"] + 2
    g["LIST_D_HDR"] = g["LIST_C_TOP"] + 8
    g["LIST_D_TOP"] = g["LIST_D_HDR"] + 2
    g["LIST_D_BOT"] = g["LIST_D_TOP"] + g["NSHEET"] - 1
    g["CL_B_BOT"] = g["CL_B_TOP"] + g["NBLK"] - 1


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
def clin_base(n): return 6 + 8 * (n - 1)     # n＝VAC患者の通し番号（ブロック番号ではない）
CLIN_LAST_ROW = clin_base(NCLIN) + 7

# ---- 個別判定シート -------------------------------------------------
MVROWS = 60                # MV 1日目 .. 60日目
R_TOP = 15                 # 表の先頭行
R_BOT = R_TOP + MVROWS - 1 # 74
R_CALC = 78                # 内部計算スカラー行
NEPISODE = 6               # 1患者あたり判定できるMVエピソード数

# ---- 一覧シート -----------------------------------------------------
LIST_A_TOP = 7                        # セクションA 先頭行（VAC患者の一覧）
LIST_A_BOT = LIST_A_TOP + NCLIN - 1
LIST_B_HDR = LIST_A_BOT + 3           # セクションB 見出し
LIST_B_TOP = LIST_B_HDR + 2           # ブロック1
LIST_B_BOT = LIST_B_TOP + NBLK - 1
LIST_C_HDR = LIST_B_BOT + 3           # セクションC 見出し（月別自動集計）
LIST_C_TOP = LIST_C_HDR + 2
LIST_D_HDR = LIST_C_TOP + 8           # セクションD 見出し（個別判定シートの判定）
LIST_D_TOP = LIST_D_HDR + 2
LIST_D_BOT = LIST_D_TOP + NSHEET - 1

# ---- CLABSI判定シート -------------------------------------------------
NEV = 30                              # 血流感染（BSI）イベントの入力行数
CL_EV_TOP = 7                         # セクションA（BSIイベント）先頭行
CL_EV_BOT = CL_EV_TOP + NEV - 1
CL_B_HDR = CL_EV_BOT + 3              # セクションB（中心ライン留置患者の一覧）見出し
CL_B_TOP = CL_B_HDR + 2
CL_B_BOT = CL_B_TOP + NBLK - 1
CL_SCAN_C0 = 22                       # 非表示の作業列（V列〜）：ブロックごとのスキャン
CL_EV_H0 = 36                         # 非表示の作業列（AJ列〜）：BSIイベントごとの計算

# ---- 名前付き範囲 ---------------------------------------------------
N_ALL  = "VAE_ALL"   # All!$A$1:$NC$<last>   … 絶対列番号でINDEXする用
N_ALLD = "VAE_ALLD"  # All!$C$1:$NC$<last>   … 日付列のみ（配列演算用）
N_CLIN = "VAE_CLIN"  # 臨床所見!$A$1:$NC$<last>
# VAC判定は当日・前日・前々日・翌日を同時に見るため、同じ長さの配列を4本そろえる。
# 判定対象の日は日付列の3日目以降（前々日が日付列に収まる日）なので、
# 当日の範囲を E列から始める。こうするとどの範囲もA列・B列のラベル文字
#（"FiO2" "PEEP" "氏名" 等）を含まないため、引き算で #VALUE! にならない。
N_D0 = "VAE_D0"    # 当日     E..NC
N_D1 = "VAE_D1"    # 前日     D..NB
N_D2 = "VAE_D2"    # 前々日   C..NA
N_DP = "VAE_DP"    # 翌日     F..ND
# 名前は必ず「_」を含める。以前の rD1・rD2 は Excel では RD列1行目・2行目のセル番地と
# 区別できず、名前として扱われなかった（LibreOffice では名前として扱われるため気づけない）。
for _nm in (N_ALL, N_ALLD, N_CLIN, N_D0, N_D1, N_D2, N_DP):
    assert "_" in _nm, f"名前付き範囲 {_nm} はセル番地と衝突するおそれがあります"

# ---- 判定のしきい値 -------------------------------------------------
FIO2_RISE = 0.2    # FiO2 の上昇幅（小数。20ポイント＝0.2）
PEEP_RISE = 3      # PEEP の上昇幅（cmH2O）
PEEP_FLOOR = 5     # PEEP 0〜5 は 5 として扱う
EVENT_PERIOD = 14  # DOEから14日間は新たなVAEを判定しない

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

# 臨床所見シートの色分け（項目のまとまりごと）。濃い色＝B列の項目名、淡い色＝入力欄
CLIN_GROUPS = [          # (先頭行オフセット, 行数, 名前, 項目名の色, 入力欄の色)
    (0, 2, "体温",   "F8CBAD", "FDF0E7"),
    (2, 2, "WBC",    "BDD7EE", "EEF4FB"),
    (4, 2, "抗菌薬", "C6E0B4", "F0F7EA"),
    (6, 2, "PVAP",   "D9D2E9", "F4F1F9"),
]
WINDOW = "FFE699"   # VAEウィンドウ（DOE±2日）＝ここに入力する
WINDOW_DOE = "FFC000"   # DOE当日

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


# ---- VAC判定の配列式（一覧シートがAllシートから直接判定するために使う）----
# 当日(VAE_D0)・前日(VAE_D1)・前々日(VAE_D2)・翌日(VAE_DP) の4本の配列を列ごとに突き合わせ、
# 「その日がDOE（酸素化悪化の初日）の条件を満たすか」を 0/1 の配列で返す。
def _idx(rng, row):
    return f"INDEX({rng},{row},0)"


def _mv_factor(rng, v, f, p):
    """その日がMV日か（1/0）。V有無が優先、空欄のときだけFiO2/PEEPで補う"""
    return (f'(({_idx(rng,v)}="V有")+({_idx(rng,v)}="")'
            f'*(({_idx(rng,f)}<>"")+({_idx(rng,p)}<>""))>0)')


def _fio2_norm(rng, f):
    """FiO2を小数に正規化（40→0.4、0.4→0.4）"""
    x = _idx(rng, f)
    return f"IF({x}>1,{x}/100,{x})"


def _peep_adj(rng, p):
    """PEEP 0〜5 を 5 に補正する。配列演算なので MAX は使えない
    （MAX は配列全体を1つの値に潰してしまうため、要素ごとに効く IF を使う）"""
    x = _idx(rng, p)
    return f"IF({x}>{PEEP_FLOOR},{x},{PEEP_FLOOR})"


def vac_array(k):
    """ブロックkについて、各日がVAC（酸素化悪化）の条件を満たすかの 1×365 配列式"""
    v, f, p = all_v(k), all_f(k), all_p(k)
    R0, R1, R2, RP = N_D0, N_D1, N_D2, N_DP

    # 4日連続でMV日＝同一エピソード内、かつ当日はMV3日目以降
    mv = "*".join(_mv_factor(r, v, f, p) for r in (R2, R1, R0, RP))

    # FiO2経路：ベースライン(前々日)比で +20ポイント以上が2暦日持続
    f2, f1, f0, fp = (_fio2_norm(r, f) for r in (R2, R1, R0, RP))
    have_f = "*".join(f"ISNUMBER({_idx(r,f)})" for r in (R2, R1, R0, RP))
    route_f = (f"({have_f}*({f1}<={f2})"
               f"*(ROUND({f0}-{f2},6)>={FIO2_RISE})"
               f"*(ROUND({fp}-{f2},6)>={FIO2_RISE}))")

    # PEEP経路：ベースライン比で +3cmH2O以上が2暦日持続（0〜5は5として補正）
    p2, p1, p0, pp = (_peep_adj(r, p) for r in (R2, R1, R0, RP))
    have_p = "*".join(f"ISNUMBER({_idx(r,p)})" for r in (R2, R1, R0, RP))
    route_p = (f"({have_p}*({p1}<={p2})"
               f"*(ROUND({p0}-{p2},6)>={PEEP_RISE})"
               f"*(ROUND({pp}-{p2},6)>={PEEP_RISE}))")

    # 範囲の取り方で前々日は必ず日付列に収まる。最終日の翌日（ND列）は空欄なので
    # MV日の条件で自動的に落ちるため、列の範囲チェックは不要。
    return f"({mv}*(({route_f}+{route_p})>0))"


def doe_col_formula(k, mvdays_cell, after_cell=None):
    """VAC条件を満たす最初の列。after_cell を渡すと『その列+14日以降』で探す
    （イベント期間14日：DOEから14日間は新たなVAEを判定しない）"""
    arr = vac_array(k)
    col = f"COLUMN({N_D0})"
    if after_cell:
        arr = f"{arr}*({col}>={after_cell}+{EVENT_PERIOD})"
        guard = f"{after_cell}>{CN}"
    else:
        guard = f"{mvdays_cell}<{MV_MIN}"
    return f"=IF({guard},{CN + 1000},SUMPRODUCT(MIN(IF({arr},{col},{CN + 1000}))))"


# 範囲全体をまとめて計算する数式（IF の配列評価・SUMPRODUCT・MATCH(TRUE,…)・LOOKUP(2,1/…)）。
# Excel は「Ctrl+Shift+Enter で確定した配列数式」でないと、SUMPRODUCT の中の IF を
# 範囲全体で評価せず、数式のある行・列と交差する1セルだけで計算してしまう
# （LibreOffice は通常の数式でも配列として評価するため、LibreOfficeでは差が出ない）。
# 生成したシートのこうした数式は、すべて配列数式として保存する。
ARRAY_MARKERS = ("SUMPRODUCT(", "MATCH(TRUE", "LOOKUP(2,1/")


def as_array_formulas(ws):
    """シート内の配列評価が必要な数式を、単一セルの配列数式（{=…}）に置き換える"""
    from openpyxl.worksheet.formula import ArrayFormula
    n = 0
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("=") and any(m in v for m in ARRAY_MARKERS):
                c.value = ArrayFormula(c.coordinate, v)
                n += 1
    return n
