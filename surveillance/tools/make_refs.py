# -*- coding: utf-8 -*-
"""引用文献一覧をPDFで保存する（出版社PDFはネットワーク制限により取得不可）"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, PageBreak)

pdfmetrics.registerFont(TTFont("JP", "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"))

H1 = ParagraphStyle("H1", fontName="JP", fontSize=15, leading=21,
                    textColor=colors.HexColor("#1F4E79"), spaceAfter=8)
H2 = ParagraphStyle("H2", fontName="JP", fontSize=11.5, leading=17,
                    textColor=colors.HexColor("#1F4E79"), spaceBefore=12, spaceAfter=5)
BODY = ParagraphStyle("BODY", fontName="JP", fontSize=9, leading=14.5, spaceAfter=5)
CITE = ParagraphStyle("CITE", fontName="JP", fontSize=9, leading=14,
                      spaceAfter=3, leftIndent=14, firstLineIndent=-14)
NOTE = ParagraphStyle("NOTE", fontName="JP", fontSize=8.5, leading=13,
                      textColor=colors.HexColor("#444444"), leftIndent=14, spaceAfter=9)
SMALL = ParagraphStyle("SMALL", fontName="JP", fontSize=8, leading=12,
                       textColor=colors.HexColor("#666666"))

REFS = [
    ("1",
     "Magill SS, Klompas M, Balk R, et al. <b>Developing a new, national approach to surveillance "
     "for ventilator-associated events.</b> Critical Care Medicine. 2013.（被引用 210）",
     "本ワークシートが実装した VAC → IVAC → PVAP の3階層アルゴリズムの原典。"
     "CDCのVAP監視定義ワーキンググループが、日最小FiO2・日最小PEEPの変化という"
     "客観的で入手しやすいデータ要素のみで第1層（VAC）を定義したことを示す。"
     "本ファイルがFiO2とPEEPだけでVACを自動判定できる根拠。",
     "https://consensus.app/papers/details/d17400d375b85cd2b52c7f2f6065cfa8/"),

    ("2",
     "Klein Klouwenberg PMC, van Mourik MSM, Ong DSY, et al. <b>Electronic implementation of a "
     "novel surveillance paradigm for ventilator-associated events. Feasibility and validation.</b> "
     "American Journal of Respiratory and Critical Care Medicine. 2014.（被引用 163）",
     "【最重要の注意点】電子的実装のわずかな違いで VAE の発生率と関連死亡率が大きく変動したと報告。"
     "＝自動判定ツールは「MV日をどう数えるか」「日最小値をどう取るか」の定義を明示し、"
     "施設間比較の際は実装条件を揃える必要がある。本ファイルで MV日の判定規則を"
     "『使い方』シートと一覧シートに明記し、最終判定を感染管理担当者に残しているのはこの知見による。",
     "https://consensus.app/papers/details/7f617a37d725507b9893c5146fbe5511/"),

    ("3",
     "Stevens JP, Silva G, Gillis J, et al. <b>Automated surveillance for ventilator-associated events.</b> "
     "Chest. 2014.（被引用 60）",
     "電子カルテを用いた全自動VAE判定アルゴリズムを人手の抽出者と比較し、"
     "感度 93.5%（95%CI 77.2–98.8）、特異度 100%（95%CI 98.8–100）と報告。"
     "＝VAE判定の自動化（本ファイルの方式）が信頼性をもって成立することの裏づけ。"
     "同時に、VAEはVAP旧定義とほとんど一致しなかった（κ=0.06）。",
     "https://consensus.app/papers/details/aaf56b04643e56e8818e311831bd55d2/"),

    ("4",
     "Fan Y, Gao F, Wu Y, et al. <b>Does ventilator-associated event surveillance detect "
     "ventilator-associated pneumonia in intensive care units? A systematic review and meta-analysis.</b> "
     "Critical Care. 2016.（8か国・61,489例のSR/メタ解析、被引用 75）",
     "VAC 13.8%、IVAC 6.4%、possible VAP 1.1%、probable VAP 0.9%、従来型VAP 11.9%。"
     "VAEのVAP検出に対する感度・陽性的中率はいずれも50%未満、特異度・陰性的中率は80%超。"
     "＝VAEはVAPの代替ではない。本ファイルの判定結果をVAP件数として読み替えないこと。",
     "https://consensus.app/papers/details/68ebda054dae570e89d661565919c063/"),

    ("5",
     "He Q, Wang W, Zhu S, et al. <b>The epidemiology and clinical outcomes of ventilator-associated "
     "events among 20,769 mechanically ventilated patients at intensive care units: an observational study.</b> "
     "Critical Care. 2021.（112,697人工呼吸器日、被引用 76）",
     "VAC 16.7 / 1,000人工呼吸器日、IVAC 6.4、PVAP 1.64。"
     "MV開始からVAC発生までの中央値は5日（IQR 3–8）。VAE該当例の院内死亡は非該当例の3倍超。"
     "＝自施設のVAE率を評価する際の比較値として利用できる（月次集計シートの比較値欄）。",
     "https://consensus.app/papers/details/74c3c9cf2e295e4a88bb77e4e3c628dc/"),

    ("6",
     "Klompas M. <b>Barriers to the adoption of ventilator-associated events surveillance and prevention.</b> "
     "Clinical Microbiology and Infection. 2019.（被引用 15）",
     "米国外でのVAE導入が進まない3つの障壁を整理。修正可能なVAE危険因子として"
     "深い鎮静・正の水分バランス・輸血・高い吸気圧の強制換気モードを挙げ、"
     "予防策として挿管回避、鎮静最小化、1日1回の覚醒トライアルと自発呼吸トライアルの組合せ、"
     "保守的輸液管理、保守的輸血閾値、低1回換気量、早期離床を示す。"
     "＝VAE率が高いときに検討すべき介入の根拠。",
     "https://consensus.app/papers/details/3926bc7faa1a54438ad29d635492ea1b/"),
]

PROTOCOLS = [
    ("A", "CDC / NHSN. <b>Patient Safety Component Manual, Chapter 10: Ventilator-Associated Event (VAE).</b>",
     "https://www.cdc.gov/nhsn/pdfs/pscmanual/10-vae_final.pdf"),
    ("B", "CDC / NHSN. <b>Patient Safety Component Manual, Chapter 4: Bloodstream Infection Event "
          "(CLABSI).</b>",
     "https://www.cdc.gov/nhsn/pdfs/pscmanual/4psc_clabscurrent.pdf"),
    ("C", "日本環境感染学会 JHAIS委員会 医療器具関連感染サーベイランス部門. "
          "<b>医療器具関連感染サーベイランス マニュアル Ver.2.5</b>"
          "（5) 人工呼吸器関連イベント(VAE)サーベイランス, pp.46–54）",
     "（ご提供いただいたPDFを参照。判定ロジックの日本語定義はこれに準拠）"),
]


def main(path="引用文献リスト_VAE自動判定ワークシート.pdf"):
    doc = SimpleDocTemplate(path, pagesize=A4, topMargin=18 * mm, bottomMargin=18 * mm,
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            title="引用文献リスト — VAE自動判定ワークシート",
                            author="感染管理サーベイランス")
    S = []
    S.append(Paragraph("引用文献リスト", H1))
    S.append(Paragraph("VAE 個別判定シート自動作成ワークブック（20XX年度サーベイランスワークシート）の"
                       "判定ロジックの根拠", BODY))
    S.append(Spacer(1, 4))

    S.append(Paragraph("1. 判定アルゴリズムの根拠となる査読論文", H2))
    S.append(Paragraph("いずれも Critical Care Medicine / AJRCCM / Chest / Critical Care / "
                       "Clinical Microbiology and Infection 掲載の、集中治療・感染管理領域で"
                       "インパクトファクターが高い雑誌の論文です。", SMALL))
    S.append(Spacer(1, 6))
    for num, cite, note, url in REFS:
        S.append(Paragraph(f"[{num}] {cite}", CITE))
        S.append(Paragraph(f"→ 本ワークシートとの関係：{note}", NOTE))
        S.append(Paragraph(f'<font size="7.5" color="#0563C1">{url}</font>', NOTE))

    S.append(Paragraph("2. 準拠したプロトコール（一次資料）", H2))
    for num, cite, url in PROTOCOLS:
        S.append(Paragraph(f"[{num}] {cite}", CITE))
        S.append(Paragraph(f'<font size="7.5" color="#0563C1">{url}</font>', NOTE))

    S.append(PageBreak())
    S.append(Paragraph("3. 本ワークシートの実装判断と、その根拠の対応表", H2))
    rows = [["実装した判定", "根拠"]]
    for a, b in [
        ("ベースライン期間＝悪化初日の直前2日、\n2暦日以上の安定または低下",
         "[1] Magill 2013 / JHAIS Ver.2.5 p.49"),
        ("VAC＝ベースライン初日比で\nFiO2 +20ポイント以上 または PEEP +3cmH2O以上が\n2暦日以上持続",
         "[1] Magill 2013 / JHAIS Ver.2.5 p.49"),
        ("PEEP 0〜5cmH2O は 5 として扱う\n（MAX(PEEP,5) で自動補正）",
         "JHAIS Ver.2.5 p.47「0〜5cmH2OのPEEP値は同等と見なされる」"),
        ("最早DOE＝MV3日目",
         "[1] Magill 2013 / JHAIS Ver.2.5 p.49 注"),
        ("VAEウィンドウ＝DOE±2日、ただしMV3日目以降",
         "JHAIS Ver.2.5 p.50–51"),
        ("イベント期間＝DOEから14日間、\nこの間は新規VAEも格上げも行わない",
         "JHAIS Ver.2.5 p.51"),
        ("MVエピソード＝1暦日以上の完全離脱で分割",
         "JHAIS Ver.2.5 p.51–52"),
        ("NPPV・マスクCPAP等の非侵襲換気を除外\n（「V有（NIV）」の日）",
         "JHAIS Ver.2.5 p.47「対象外患者」"),
        ("IVAC＝体温>38℃ or <36℃、または\nWBC≥12,000 or ≤4,000/mm3 ＋ 新規抗菌薬4QAD以上",
         "[1] Magill 2013 / JHAIS Ver.2.5 p.52–53"),
        ("自動判定はあくまで補助とし、\n最終判定は感染管理担当者が行う",
         "[2] Klein Klouwenberg 2014（実装差で率が変動）"),
    ]:
        rows.append([Paragraph(a, ParagraphStyle("t", fontName="JP", fontSize=8, leading=11.5)),
                     Paragraph(b, ParagraphStyle("t", fontName="JP", fontSize=8, leading=11.5))])
    t = Table(rows, colWidths=[95 * mm, 79 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "JP"),
        ("FONTSIZE", (0, 0), (-1, 0), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B4C6E7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF1F5")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    S.append(t)

    S.append(Paragraph("4. 本PDFについて", H2))
    S.append(Paragraph(
        "本セッションの実行環境はネットワーク送信先が制限されており、出版社サイト・PubMed Central・"
        "CDC のサーバーへ接続できなかったため、各論文の本文PDFそのものを保存することができませんでした。"
        "本PDFには、書誌情報・要旨の要点・本ワークシートとの対応関係、および各論文へのリンクをまとめています。"
        "本文PDFが必要な場合は、上記リンクまたは所属機関の文献取得サービスからダウンロードしてください。",
        BODY))
    doc.build(S)
    print("wrote", path)


if __name__ == "__main__":
    import sys
    main(*sys.argv[1:])
