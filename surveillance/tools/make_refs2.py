# -*- coding: utf-8 -*-
"""引用文献一覧（CLABSI判定・使用比／発生率の集計）をPDFで保存する

出版社サイト・CDC へはこの環境から接続できないため、本文PDFではなく
書誌情報・要点・ワークブックとの対応・リンクをまとめたPDFを作る。
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from make_refs import H1, H2, BODY, CITE, NOTE, SMALL   # フォント登録と書式を共有

PROTOCOLS = [
    ("A", "日本環境感染学会 JHAIS委員会 医療器具関連感染サーベイランス部門. "
          "<b>医療器具関連感染サーベイランス マニュアル Ver.2.5</b>（2025年8月19日）",
     "p.8・p.12–14「感染率と医療器具使用比の計算」「Pooled mean」：CLABSI感染率＝CLABSI発生件数÷延べ中心ライン使用日数×1000、"
     "中心ライン使用比＝延べ中心ライン使用日数÷延べ入室患者日数、集計値は合計から求める（pooled mean）。"
     "p.19–25：感染ウインドウ期間（IWP）、イベント発生日（DOE）、入院時感染（POA）、RIT（14日）、転棟ルール。"
     "p.26–33：中心ライン・対象となる中心ライン（2暦日を超えて留置）・LCBI基準1〜3・臨床的敗血症（CSEP）・"
     "中心ラインの使用とBSIイベントの判定例（患者A〜E）。報告書はLCBI＋CSEPとLCBI単独を併記（p.12–13）。",
     "（ご提供いただいたPDFを参照。CLABSI判定シート・両集計シートの定義はこれに準拠）"),
    ("B", "CDC / NHSN. <b>Patient Safety Component Manual, Chapter 4: Bloodstream Infection Event "
          "(Central Line-Associated Bloodstream Infection and Non-central Line Associated Bloodstream "
          "Infection).</b>",
     "LCBI基準、対象となる中心ライン（eligible central line）、DOE・RIT・2次性BSIの扱い。JHAISのLCBIはこれに準拠。",
     "https://www.cdc.gov/nhsn/pdfs/pscmanual/4psc_clabscurrent.pdf"),
    ("C", "CDC / NHSN. <b>Patient Safety Component Manual, Chapter 2: Identifying Healthcare-associated "
          "Infections (HAI) for NHSN Surveillance.</b>",
     "IWP・DOE・POA・RIT・Secondary BSI Attribution Period・転棟ルール（Transfer Rule）の原典。",
     "https://www.cdc.gov/nhsn/pdfs/pscmanual/2psc_identifyinghais_nhsncurrent.pdf"),
]

REFS = [
    ("1",
     "Marsteller JA, Sexton JB, Hsu YJ, et al. <b>A multicenter, phased, cluster-randomized controlled "
     "trial to reduce central line-associated bloodstream infections in intensive care units.</b> "
     "Critical Care Medicine. 2012.（35病院45 ICU のクラスターRCT、被引用 120）",
     "CLABSIを「1,000中心ライン日あたり」で測定し、四半期ごとに率を返した。介入群は4.48→1.33"
     "（調整IRR 0.19、95%CI 0.06–0.57）、19か月後も1,000中心ライン日あたり1件未満（81%減少）。"
     "＝本ブックのCLABSI集計（/1,000中心ライン日・月別）の表し方と、率を定期的に返すことの根拠。",
     "https://consensus.app/papers/details/29b8ae19327a5fa483c15f8d10e7db5d/?utm_source=claude_desktop"),

    ("2",
     "Ben-David D, et al. <b>Impact of intensified prevention measures on rates of hospital-acquired "
     "bloodstream infection in medical-surgical intensive care units, Israel, 2011 to 2019.</b> "
     "Eurosurveillance. 2023.（国内29 ICU の分割時系列解析）",
     "全国のガイドライン・サーベイランス・率のフィードバックで、CLABSIのpooled mean が 7.4→2.1／1,000中心ライン日に低下"
     "（第I期→第II期 IRR 0.63、95%CI 0.51–0.79）。"
     "＝CLABSI集計シートの注記「サーベイランス結果のフィードバックはCLABSIの減少と関連」の根拠。",
     "https://consensus.app/papers/details/89809209675a539bbd66840f0e57e7de/?utm_source=claude_desktop"),

    ("3",
     "Blot K, Bergs J, Vogelaers D, et al. <b>Prevention of central line-associated bloodstream infections "
     "through quality improvement interventions: a systematic review and meta-analysis.</b> "
     "Clinical Infectious Diseases. 2014.（41の前後比較研究のメタ解析、被引用 175）",
     "質改善介入でCLABSIが減少（OR 0.39、95%CI 0.33–0.46）。バンドル・チェックリストを用いた介入で効果がより大きい。"
     "＝CLABSI率が高い月が続くときに検討する介入の根拠。",
     "https://consensus.app/papers/details/5e7bfc1cefc15a44be0ae77b77bb43b0/?utm_source=claude_desktop"),

    ("4",
     "Pronovost P, Needham D, Berenholtz S, et al. <b>An intervention to decrease catheter-related "
     "bloodstream infections in the ICU.</b> New England Journal of Medicine. 2006;355:2725–2732. "
     "doi:10.1056/NEJMoa061115（米国ミシガン州103 ICU、Keystone ICU Project）",
     "カテーテル関連血流感染を1,000カテーテル日あたりで追跡し、中央値 2.7→0（3か月後）、平均 7.7→1.4（16〜18か月後）。"
     "＝ICUのCLABSIを「1,000中心ライン日あたり」で継続的に測ることが、改善の評価指標として定着した代表的研究。",
     "https://doi.org/10.1056/NEJMoa061115"),

    ("5",
     "Parveen R, et al. <b>Profile of central line-associated bloodstream infections in adult, paediatric, and "
     "neonatal intensive care units of hospitals participating in a health-care-associated infection "
     "surveillance network in India: a 7-year multicentric study.</b> The Lancet Global Health. 2025.",
     "約200 ICU・7年間。月ごとに患者日と中心ライン日を分母として入力し、CLABSI率（/1,000中心ライン日）と"
     "中心ライン使用比（CLUR）を年ごとに pooled で算出（全体 8.83、CLUR 0.32。成人ICU 8.68、CLUR 0.38）。"
     "＝本ブックの「年度計は月ごとの率の平均ではなく合計から計算」の実例。",
     "https://consensus.app/papers/details/efd7b55f1d6e5b958c16690c2723d6ad/?utm_source=claude_desktop"),

    ("6",
     "Hsu HE, et al. <b>Health Care-Associated Infections Among Critically Ill Children in the US, "
     "2013-2018.</b> JAMA Pediatrics. 2020.（NHSN報告176病院）",
     "器具関連感染率（/1,000器具日）と器具使用比（器具日÷患者日）を併せて経年評価。"
     "CAUTIでは率が横ばいのまま使用比の低下（年6%）が人口あたりの感染減少に寄与した。"
     "＝発生率だけでなく使用比も並べてグラフにする理由（使用を減らすこと自体が予防策）。",
     "https://consensus.app/papers/details/1ee084b811035ef09401bb9a65642130/?utm_source=claude_desktop"),

    ("7",
     "Nakahashi S, et al. <b>A reappraisal of association between ventilator-associated events and mortality "
     "among critically ill patients using marginal structural model: multicenter observational study.</b> "
     "Intensive Care Medicine. 2025.（日本の18 ICU、1,094例）",
     "CDC定義のVAEは 106件、10.0／1,000人工呼吸器日。重症度の経時変化を調整後も30日院内死亡と関連"
     "（HR 2.00、95%CI 1.23–3.26）。"
     "＝VAE集計シートに示した日本の参考値。VAEが単なる重症度の代理ではなく、測る意味のある指標であることの根拠。",
     "https://consensus.app/papers/details/cf550ab31cd45e46970c0de132323814/?utm_source=claude_desktop"),

    ("8",
     "He Q, Wang W, Zhu S, et al. <b>The epidemiology and clinical outcomes of ventilator-associated events "
     "among 20,769 mechanically ventilated patients at intensive care units: an observational study.</b> "
     "Critical Care. 2021.（112,697人工呼吸器日）",
     "VAC 16.7、IVAC 6.4、PVAP 1.64／1,000人工呼吸器日。ICUの種類で率が大きく異なる（外科系ICUのVAC 23.72）。"
     "＝VAE集計シートの参考値。比較は同じ定義・同じ種類のICUで行う必要がある。",
     "https://consensus.app/papers/details/74c3c9cf2e295e4a88bb77e4e3c628dc/?utm_source=claude_desktop"),

    ("9",
     "Klompas M, et al. <b>Multicenter Evaluation of a Novel Surveillance Paradigm for Complications of "
     "Mechanical Ventilation.</b> PLoS ONE. 2011.（3病院600例、被引用 184）",
     "VAC 21.2／1,000人工呼吸器日。VACは死亡と関連（OR 2.0）し、判定にかかる時間は1例1.8分（VAPは39分）。"
     "＝人工呼吸器の設定値だけで自動判定するVAE監視（本ブックの方式）が現実的に回ることの根拠。",
     "https://consensus.app/papers/details/b95a5574b271597bbabf25ee47843566/?utm_source=claude_desktop"),

    ("10",
     "Haley RW, Culver DH, White JW, et al. <b>The efficacy of infection surveillance and control programs "
     "in preventing nosocomial infections in US hospitals.</b> American Journal of Epidemiology. "
     "1985;121:182–205. doi:10.1093/oxfordjournals.aje.a113990（SENIC Project）",
     "サーベイランスと感染対策の体制が整った病院では院内感染が減少し、体制のない病院では増加した。"
     "＝月別の自動集計とグラフで、結果を継続的に現場へ返す仕組みを作る理由の古典的根拠。",
     "https://doi.org/10.1093/oxfordjournals.aje.a113990"),
]

MAP = [
    ("CLABSI発生率＝件数÷延べ中心ライン使用日数×1,000\n中心ライン使用比＝延べ中心ライン使用日数÷延べ入室患者日数",
     "[A] JHAIS Ver.2.5 p.8・p.12–14"),
    ("VAE発生率＝件数÷延べ人工呼吸器使用日数×1,000\n人工呼吸器使用比＝延べ人工呼吸器使用日数÷延べ入室患者日数",
     "[A] JHAIS Ver.2.5 p.8・p.12–14"),
    ("年度計は月ごとの率の平均ではなく、年間の合計から計算（pooled）\n"
     "既存の年間集計シートの AVERAGE も同じ方法に修正", "[A] JHAIS Ver.2.5 p.12–14（Pooled mean）、[5]"),
    ("CLABSI＝LCBI-1／LCBI-2／CSEP ＋ DOEか前日に\n2暦日を超えて留置された中心ライン（CL3日目以降）",
     "[A] p.26–29、[B]"),
    ("1暦日以上「C有」が途切れたら中心ライン日を数え直す\n抜去日・抜去翌日のDOEは対象", "[A] p.31–32（患者B・C・D）"),
    ("LCBI-2：症状がIWP（採取日±3日）内、DOE＝採取日と症状の早い方", "[A] p.28、[C]"),
    ("ICU入室2日以内のDOEはPOAまたは転棟元に帰属\nICU退室翌日のDOEはICUに帰属", "[A] p.20・p.25（転棟ルール）、[C]"),
    ("RIT：一次BSIのDOEから14日間は新規にしない（2次性BSIはRITを作らない）", "[A] p.23・p.29–30"),
    ("CLABSIは「LCBI＋CSEP」と「LCBIのみ」の両方を集計", "[A] p.12–13（表1-1の併記）"),
    ("使用比と発生率を並べてグラフにする", "[6]、[1]"),
    ("参考値：VAE 10.0（日本18 ICU）、VAC 16.7・IVAC 6.4・PVAP 1.64", "[7]、[8]"),
]


def main(path="引用文献リスト_CLABSI判定・使用比集計.pdf"):
    doc = SimpleDocTemplate(path, pagesize=A4, topMargin=18 * mm, bottomMargin=18 * mm,
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            title="引用文献リスト — CLABSI判定・使用比／発生率の集計",
                            author="感染管理サーベイランス")
    S = [Paragraph("引用文献リスト（追加分）", H1),
         Paragraph("CLABSI判定シート・VAE集計シート・CLABSI集計シート（使用比と発生率の月別自動集計）の"
                   "定義と、シートに示した参考値・注記の根拠", BODY),
         Spacer(1, 4),
         Paragraph("1. 準拠したプロトコール（一次資料）", H2)]
    for num, cite, note, url in PROTOCOLS:
        S.append(Paragraph(f"[{num}] {cite}", CITE))
        S.append(Paragraph(f"→ {note}", NOTE))
        S.append(Paragraph(f'<font size="7.5" color="#0563C1">{url}</font>', NOTE))

    S.append(Paragraph("2. 査読論文", H2))
    S.append(Paragraph("New England Journal of Medicine・The Lancet Global Health・JAMA Pediatrics・"
                       "Intensive Care Medicine・Critical Care Medicine・Clinical Infectious Diseases・"
                       "Critical Care など、感染管理・集中治療領域の主要誌を優先して選びました。", SMALL))
    S.append(Spacer(1, 6))
    for num, cite, note, url in REFS:
        S.append(Paragraph(f"[{num}] {cite}", CITE))
        S.append(Paragraph(f"→ 本ワークブックとの関係：{note}", NOTE))
        S.append(Paragraph(f'<font size="7.5" color="#0563C1">{url}</font>', NOTE))

    S.append(PageBreak())
    S.append(Paragraph("3. 実装した判定・計算と、その根拠の対応表", H2))
    st = ParagraphStyle("t", fontName="JP", fontSize=8, leading=11.5)
    rows = [["実装した判定・計算", "根拠"]] + [[Paragraph(a.replace("\n", "<br/>"), st), Paragraph(b, st)]
                                              for a, b in MAP]
    t = Table(rows, colWidths=[112 * mm, 62 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7030A0")),
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
        "本セッションの実行環境はネットワークの接続先が制限されており、出版社サイト（journals.plos.org、"
        "eurosurveillance.org など）・PubMed Central・CDC のサーバーへ接続できなかったため、"
        "各論文の本文PDFそのものは保存できませんでした。本PDFには書誌情報・要点・本ワークブックとの対応関係と、"
        "各論文へのリンクをまとめています。本文PDFが必要な場合は、上記リンクまたは所属機関の文献取得サービスから"
        "ダウンロードしてください。VAE判定ロジックの根拠は、別紙「引用文献リスト_VAE自動判定ワークシート.pdf」にあります。",
        BODY))
    doc.build(S)
    print("wrote", path)


if __name__ == "__main__":
    import sys
    main(*sys.argv[1:])
