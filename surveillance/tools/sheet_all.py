# -*- coding: utf-8 -*-
"""Allシートの条件付き書式を入れ替える。

元のAllシートには VAC 判定を意図した条件付き書式が 3,920 ルール入っていたが、
  ・半数（1,960）に #REF! が残っている
  ・FiO2 を分数として比較（>=0.2）しているが、シートには％（40・60 等）で入力する
  ・相対参照が 4 列左・7 行上にずれており、入力したセルとは無関係のセルが赤くなる
という状態だったため、すべて削除して以下の2ルールに置き換える。

  ルール1（濃い赤）：個別判定シートが報告するDOE（イベント期間14日の重複排除まで適用済み）
  ルール2（薄い赤）：JHAISの酸素化悪化基準に合致した日（ルール1に該当しないもの）
"""
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Font, PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as gl
from common import *

DOE_RED = "FF0000"      # DOE（報告対象）
CAND_RED = "FFC7CE"     # VAC基準該当（イベント期間内 or 判定シート未割当）

# Allシート上の位置を ROW() / COLUMN() から求める式
BASE = "6+5*INT((ROW()-6)/5)"          # ブロック先頭行（＝V有無行）
BLK = "INT((ROW()-6)/5)+1"             # ブロック番号
RMOD = "MOD(ROW()-6,5)"                # 0=V有無 1=FiO2 2=PEEP 3=C有無 4=部屋
IS_FP = f"OR({RMOD}=1,{RMOD}=2)"       # FiO2行 または PEEP行


def _at(row_expr, d):
    """d 日ずらした列の値"""
    off = "" if d == 0 else (f"+{d}" if d > 0 else f"{d}")
    return f"INDEX({N_ALL},{row_expr},COLUMN(){off})"


def _v(d):  return _at(BASE, d)          # V有無
def _f(d):  return _at(f"{BASE}+1", d)   # FiO2
def _p(d):  return _at(f"{BASE}+2", d)   # PEEP
def _pa(d): return f"MAX({_p(d)},5)"     # PEEP（0〜5は5として補正）


def _mv(d):
    """その日が MV日（侵襲的人工呼吸の装着日）か"""
    return (f'AND({_v(d)}<>"V有（NIV）",'
            f'OR({_v(d)}="V有",{_f(d)}<>"",{_p(d)}<>""))')


def vac_formula():
    """JHAIS/NHSN の酸素化悪化基準（VAC）に合致する日かどうか。

    ・MV日が4日連続（-2,-1,0,+1）＝同一エピソード内、かつ当日はMV3日目以降
    ・ベースライン（-2,-1日）が安定または低下
    ・ベースライン初日比で FiO2 +20ポイント以上 または PEEP +3cmH2O以上 が2暦日持続
    """
    fio2 = (f"AND(COUNT({_f(-2)},{_f(-1)},{_f(0)},{_f(1)})=4,"
            f"{_f(-1)}<={_f(-2)},"
            f"{_f(0)}-{_f(-2)}>=20,"
            f"{_f(1)}-{_f(-2)}>=20)")
    peep = (f"AND(COUNT({_p(-2)},{_p(-1)},{_p(0)},{_p(1)})=4,"
            f"{_pa(-1)}<={_pa(-2)},"
            f"{_pa(0)}-{_pa(-2)}>=3,"
            f"{_pa(1)}-{_pa(-2)}>=3)")
    body = (f"AND({_mv(-2)},{_mv(-1)},{_mv(0)},{_mv(1)},OR({fio2},{peep}))")
    guard = f"AND({IS_FP},COLUMN()>={C0+2},COLUMN()<={CN-1})"
    return f"IF(COUNT(C6)=0,FALSE,IF({guard},{body},FALSE))"


def doe_formula():
    """個別判定シートが報告したDOEと日付が一致するか（イベント期間14日を適用済み）"""
    L = "'VAE対象一覧'"
    slot = f"INDEX({L}!$G${LIST_B_TOP}:$G${LIST_B_BOT},{BLK})"
    d1 = f"INDEX({L}!$K${LIST_A_TOP}:$K${LIST_A_BOT},{slot})"
    d2 = f"INDEX({L}!$M${LIST_A_TOP}:$M${LIST_A_BOT},{slot})"
    body = f"IFERROR(OR(C$5={d1},C$5={d2}),FALSE)"
    return f"IF(COUNT(C6)=0,FALSE,IF({IS_FP},{body},FALSE))"


def build(wb):
    ws = wb["All"]

    # ---- 壊れていた条件付き書式をすべて削除 --------------------------
    removed = sum(len(r) for r in ws.conditional_formatting._cf_rules.values())
    ws.conditional_formatting._cf_rules.clear()

    rng = f"C6:{gl(CN)}{ALL_LAST_ROW}"

    # ルール1：DOE（濃い赤・白太字）。該当したら以降のルールは評価しない
    ws.conditional_formatting.add(rng, Rule(
        type="expression", formula=[doe_formula()], stopIfTrue=True,
        dxf=DifferentialStyle(fill=PatternFill(bgColor=DOE_RED),
                              font=Font(color="FFFFFF", b=True))))

    # ルール2：VAC基準該当（薄い赤・濃赤文字）
    ws.conditional_formatting.add(rng, Rule(
        type="expression", formula=[vac_formula()], stopIfTrue=False,
        dxf=DifferentialStyle(fill=PatternFill(bgColor=CAND_RED),
                              font=Font(color="9C0006"))))

    # ---- 凡例 --------------------------------------------------------
    # A1:B4 は結合済みで空きがないため、レイアウトを崩さないようセルのコメントに入れる
    ws["A5"].comment = Comment(
        "【FiO2・PEEP行の色】\n"
        "■ 濃い赤（白文字）＝ DOE（イベント発生日）\n"
        "　 個別判定シートが報告する日。イベント期間14日による重複排除まで適用済み。\n\n"
        "■ 薄い赤 ＝ JHAISの酸素化悪化基準（VAC）に合致した日\n"
        "　 ただしDOEには当たらないもの。直前のDOEから14日以内で新規報告の対象外か、\n"
        "　 まだ個別判定シートが割り当てられていない患者です。\n"
        "　 VAE対象一覧シートで該当患者の判定を確認してください。",
        "VAEサーベイランス")
    ws["A5"].comment.width = 380
    ws["A5"].comment.height = 170
    return removed
