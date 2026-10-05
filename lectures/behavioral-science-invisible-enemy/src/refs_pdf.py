# 引用文献リスト（PDF）を生成する
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import ParagraphStyle

OUT = sys.argv[1] if len(sys.argv) > 1 else "refs.pdf"
pdfmetrics.registerFont(TTFont("BIZ", "/root/.fonts/BIZUDGothic-Regular.ttf"))
pdfmetrics.registerFont(TTFont("BIZB", "/root/.fonts/BIZUDGothic-Bold.ttf"))
from reportlab.lib.fonts import addMapping
addMapping("BIZ", 0, 0, "BIZ"); addMapping("BIZ", 1, 0, "BIZB"); addMapping("BIZ", 0, 1, "BIZ"); addMapping("BIZ", 1, 1, "BIZB")

NAVY = colors.HexColor("#131B35"); TEAL = colors.HexColor("#0B7A70"); MUTED = colors.HexColor("#4F5B6B"); CARD = colors.HexColor("#EEF2F7")
st = lambda **k: ParagraphStyle("s", **{**dict(fontName="BIZ", fontSize=9.5, leading=15, wordWrap="CJK"), **k})
H1 = st(fontSize=16, leading=22, textColor=NAVY, fontName="BIZB")
SUB = st(fontSize=10.5, leading=16, textColor=MUTED)
BODY = st()
SMALL = st(fontSize=8.5, leading=13, textColor=MUTED)
REFT = st(fontSize=10, leading=15, fontName="BIZB", textColor=NAVY)
LINK = st(fontSize=8.5, leading=13, textColor=TEAL)

# (著者・タイトル・書誌, DOI, PMID, 研究デザイン, 掲載誌の位置づけ, 講演で用いた要点, 使用スライド, 全文アクセス)
R = [
 ("Erasmus V, Daha TJ, Brug H, et al. Systematic review of studies on compliance with hand hygiene guidelines in hospital care. Infect Control Hosp Epidemiol. 2010;31(3):283-94.",
  "10.1086/650451", "20088678", "系統的レビュー（96研究）", "感染管理の専門誌（米国医療疫学学会の公式誌）",
  "手指衛生遵守率の中央値は40%。ICU 30〜40%／その他 50〜60%、医師32%／看護師48%、患者接触前21%／接触後47%。業務量が多い場面で遵守率が低い。", "8", "購読"),
 ("Pittet D, Hugonnet S, Harbarth S, et al. Effectiveness of a hospital-wide programme to improve compliance with hand hygiene. Lancet. 2000;356(9238):1307-12.",
  "10.1016/S0140-6736(00)02814-2", "11073019", "病院全体介入の前後比較（2万回超の機会を観察）", "総合医学のトップジャーナル",
  "遵守率48%→66%、院内感染の有病率16.9%→9.9%、MRSA伝播2.16→0.93/1万患者日、アルコール手指消毒剤3.5→15.4 L/1000患者日。ベッドサイドのアルコール手指消毒の推進が遵守率向上に大きく寄与。", "9, 21, 38", "購読"),
 ("Webb TL, Sheeran P. Does changing behavioral intentions engender behavior change? A meta-analysis of the experimental evidence. Psychol Bull. 2006;132(2):249-68.",
  "10.1037/0033-2909.132.2.249", "16536643", "実験研究47件のメタ分析", "心理学のトップジャーナル",
  "意図を中〜大きく変えても（d=0.66）、行動の変化は小〜中程度（d=0.36）にとどまる。", "10", "購読"),
 ("Jenner EA, Fletcher BC, Watson P, et al. Discrepancy between self-reported and observed hand hygiene behaviour in healthcare professionals. J Hosp Infect. 2006;63(4):418-22.",
  "10.1016/j.jhin.2006.03.012", "16772101", "観察研究（71名・132時間・1,284機会）＋質問紙", "感染管理の専門誌",
  "観察された遵守率は非常に低く、実際の行動は本人の意図・自己申告と関連しなかった。態度や意図のみを標的にした介入は行動変容に失敗しやすいと考察。", "13", "購読"),
 ("Grant AM, Hofmann DA. It's not all about me: motivating hand hygiene among health care professionals by focusing on patients. Psychol Sci. 2011;22(12):1494-9.",
  "10.1177/0956797611419172", "22075239", "病院での2つのフィールド実験", "心理学の主要誌",
  "「あなたを守る」掲示では変化なし、「患者を守る」掲示では消毒剤使用量・覆面観察での手指衛生が有意に増加。医療者は自分の免疫を過信しやすい。", "14", "購読"),
 ("Srigley JA, Furness CD, Baker GR, Gardam M. Quantification of the Hawthorne effect in hand hygiene compliance monitoring using an electronic monitoring system: a retrospective cohort study. BMJ Qual Saf. 2014;23(12):974-80.",
  "10.1136/bmjqs-2014-003080", "25002555", "後ろ向きコホート（リアルタイム位置情報システムで8か月間記録）", "医療の質・安全の主要誌",
  "監査者が見える場所の手指衛生は3.75回/台/時、同時刻の見えない場所1.48回、同じ場所の前週1.07回（約3倍）。病室内では変化なし。", "15, 16", "PMCで閲覧可"),
 ("Pittet D, Simon A, Hugonnet S, et al. Hand hygiene among physicians: performance, beliefs, and perceptions. Ann Intern Med. 2004;141(1):1-8.",
  "10.7326/0003-4819-141-1-200407060-00008", "15238364", "横断研究（医師163名の観察＋質問紙）", "内科系のトップジャーナル",
  "遵守率は平均57%。遵守は「観察されている」意識、「同僚の手本である」という信念、肯定的な態度、手指消毒剤へのアクセスのしやすさと関連。高業務量等は不遵守の要因。", "16, 20", "購読"),
 ("Dai H, Milkman KL, Hofmann DA, Staats BR. The impact of time at work and time off from work on rule compliance: the case of hand hygiene in health care. J Appl Psychol. 2015;100(3):846-62.",
  "10.1037/a0038067", "25365728", "縦断観察（35病院・4,157人・約1,370万機会）", "産業・組織心理学のトップジャーナル",
  "12時間勤務の開始から終了までに遵守率が平均8.7ポイント低下。業務強度が高いほど低下が大きく、勤務間の休息が長いと次の勤務で遵守率が回復。", "17", "購読"),
 ("Pittet D, Mourouga P, Perneger TV. Compliance with handwashing in a teaching hospital. Ann Intern Med. 1999;130(2):126-30.",
  "10.7326/0003-4819-130-2-199901190-00006", "10068358", "観察研究（2,834機会）", "内科系のトップジャーナル",
  "平均遵守率48%。ケアの強度が高い（1時間41回以上の機会）と、20回以下に比べ不遵守のオッズ比は約2.1。医師・ICU・汚染リスクの高い処置でも不遵守が多い。", "18", "購読"),
 ("Lankford MG, Zembower TR, Trick WE, et al. Influence of role models and hospital design on hand hygiene of healthcare workers. Emerg Infect Dis. 2003;9(2):217-23.",
  "10.3201/eid0902.020249", "12603993", "観察研究（新旧病院の比較、721機会）", "感染症の主要誌（米国CDC発行）",
  "上級者や同僚が手指衛生をしない部屋では自分もしにくい（オッズ比0.2）。手洗いシンクを増やした新病院で遵守率は53%→23.3%に低下し、設備の増設だけでは改善しなかった。", "19, 21, 38", "オープンアクセス（CC BY 4.0）"),
 ("Monsalve MN, Pemmaraju SV, Thomas GW, et al. Do peer effects improve hand hygiene adherence among healthcare workers? Infect Control Hosp Epidemiol. 2014;35(10):1277-85.",
  "10.1086/678068", "25203182", "センサーネットワークによる観察研究（ICU・47,694機会）", "感染管理の専門誌（米国医療疫学学会の公式誌）",
  "一人のときの遵守率20.85%に対し、他の職員がいるときは27.90%。周囲の人数が多いほど上昇（逓減あり）。", "19", "PMCで閲覧可"),
 ("Gould DJ, Moralejo D, Drey N, Chudleigh JH, Taljaard M. Interventions to improve hand hygiene compliance in patient care. Cochrane Database Syst Rev. 2017;9(9):CD005186.",
  "10.1002/14651858.CD005186.pub4", "28862335", "系統的レビュー（26研究）", "エビデンス統合の国際的標準（コクラン）",
  "使用場所の近くへのアルコール手指消毒剤の配置はおそらく遵守率をわずかに改善（中等度の確実性）。フィードバック・教育・掲示や香りなどのキューは改善の可能性（低い確実性）。WHO推奨の多角的介入は低〜非常に低い確実性。", "21, 42", "PMCで閲覧可"),
 ("Michie S, van Stralen MM, West R. The behaviour change wheel: a new method for characterising and designing behaviour change interventions. Implement Sci. 2011;6:42.",
  "10.1186/1748-5908-6-42", "21513547", "既存19フレームワークの系統的検討と新枠組みの提案", "実装科学のトップジャーナル",
  "行動は能力（Capability）・機会（Opportunity）・動機（Motivation）の3条件（COM-B）で生じる。周囲に9つの介入機能と7つの政策カテゴリーを配置した「行動変容ホイール」を提案。", "23, 24", "オープンアクセス（CC BY 2.0）"),
 ("Stiefel U, Cadnum JL, Eckstein BC, et al. Contamination of hands with methicillin-resistant Staphylococcus aureus after contact with environmental surfaces and after contact with the skin of colonized patients. Infect Control Hosp Epidemiol. 2011;32(2):185-7.",
  "10.1086/657944", "21460476", "観察研究（MRSA保菌者40名）", "感染管理の専門誌（米国医療疫学学会の公式誌）",
  "患者の皮膚に触れた後40%、環境表面に触れた後45%で手がMRSAに汚染（ほぼ同等）。", "25", "購読"),
 ("Tomas ME, Kundrapu S, Thota P, et al. Contamination of health care personnel during removal of personal protective equipment. JAMA Intern Med. 2015;175(12):1904-10.",
  "10.1001/jamainternmed.2015.4535", "26457544", "点有病率調査＋準実験的介入（蛍光ローション）", "内科系のトップジャーナル",
  "435回の脱衣シミュレーションの46.0%で皮膚・衣服が汚染（手袋52.9%＞ガウン37.8%）。蛍光で即時に汚染を見せる教育で60.0%→18.9%、1・3か月後も12.0%に維持。", "26", "購読"),
 ("Carling PC, Parry MM, Rupp ME, et al. Improving cleaning of the environment surrounding patients in 36 acute care hospitals. Infect Control Hosp Epidemiol. 2008;29(11):1035-41.",
  "10.1086/591940", "18851687", "前向き準実験（前後比較、36病院）", "感染管理の専門誌（米国医療疫学学会の公式誌）",
  "蛍光マーカーで評価した退院時清掃で、清拭されていた表面は48%（20,646か所中）。教育・手順介入と清掃スタッフへの客観的フィードバック後に77%へ改善。", "27", "購読"),
 ("Hallsworth M, Chadborn T, Sallis A, et al. Provision of social norm feedback to high prescribers of antibiotics in general practice: a pragmatic national randomised controlled trial. Lancet. 2016;387(10029):1743-52.",
  "10.1016/S0140-6736(16)00215-4", "26898856", "全国規模の実践的ランダム化比較試験（1,581診療所）", "総合医学のトップジャーナル",
  "処方上位20%の診療所に主席医務官から社会規範フィードバックの手紙を送付。6か月の処方率は3.3%減少（IRR 0.967）、推計73,406件の処方減。", "28", "オープンアクセス（CC BY-NC-ND）"),
 ("Thaler RH, Sunstein CR. Nudge: Improving Decisions about Health, Wealth, and Happiness. New Haven: Yale University Press; 2008.",
  "", "", "書籍（行動経済学）", "ナッジの概念を提唱した基本文献（著者のセイラーは2017年ノーベル経済学賞）",
  "ナッジ＝選択の自由を残したまま、選択肢の示し方（選択アーキテクチャ）を工夫して、望ましい行動を予測可能な形で後押しすること。禁止や経済的インセンティブの大きな変更は含まない。", "29, 30", "書籍"),
 ("Johnson EJ, Goldstein D. Do defaults save lives? Science. 2003;302(5649):1338-9.",
  "10.1126/science.1091721", "14631022", "オンライン実験（161名）＋欧州各国の比較", "科学のトップジャーナル",
  "臓器提供の意思表示で、自分で同意を選ぶ（オプトイン）条件の同意率は約42%、最初から同意が選ばれている（オプトアウト）条件では約82%。国別でも、オプトアウトの国の実効同意率は90%超が多く、オプトインの国と重ならない。", "31", "購読"),
 ("Chapman GB, Li M, Colby H, Yoon H. Opting in vs opting out of influenza vaccination. JAMA. 2010;304(1):43-4.",
  "10.1001/jama.2010.892", "20606147", "ランダム化比較試験（大学職員478名）", "総合医学のトップジャーナル",
  "あらかじめ接種日時を予約しておく（オプトアウト）案内で接種率45%（108/239）、自分で予約する（オプトイン）案内で33%（80/239）。予約を取り消したのはオプトアウト群の8%。", "32", "購読"),
 ("Dai H, Saccardo S, Han MA, et al. Behavioural nudges increase COVID-19 vaccinations. Nature. 2021;597(7876):404-9.",
  "10.1038/s41586-021-03843-2", "34340242", "2つのランダム化比較試験（93,354名／67,092名）", "科学のトップジャーナル",
  "接種資格通知の翌日に送ったテキストのリマインダーで、予約率＋6.07ポイント、接種率＋3.57ポイント。「あなたのワクチンが用意されています」と所有感を持たせる文面で効果がより大きかった。", "32", "PMCで閲覧可"),
 ("鈴木徳洋, 赤瀬望, 清水潤三. 手指衛生遵守率向上への行動経済学的アプローチ. 日本外科感染症学会（ポスター発表）.",
  "", "", "単施設の前後比較（2018年度 vs 2024年度）", "講演者施設の学会発表",
  "視覚的ナッジ・部署間共有・個人別使用量の見える化により、病院全体の手指消毒剤使用量は12.6→26.3 L/1000患者日（約2.09倍）。全13部署で有意に増加（p<0.0001）。", "33, 34", "—"),
 ("The Behavioural Insights Team. EAST: Four simple ways to apply behavioural insights. London: The Behavioural Insights Team; 2014.",
  "", "", "実践フレームワーク（報告書）", "英国政府系の行動インサイト機関",
  "行動を促す4原則：Easy（簡単に）、Attractive（魅力的に）、Social（社会的に）、Timely（タイミングよく）。", "37", "発行元サイトで公開"),
 ("Cure L, Van Enk R. Effect of hand sanitizer location on hand hygiene compliance. Am J Infect Control. 2015;43(9):917-21.",
  "10.1016/j.ajic.2015.05.013", "26088769", "観察研究（404床の病院・12病棟、2010〜2012年）", "感染管理の専門誌（米国感染管理疫学専門家協会の公式誌）",
  "手指消毒剤ディスペンサーの使いやすさの総合点・見えやすさ・病室入口での取りやすさが、観察された遵守率の高さと有意に関連。規格の統一だけでは関連なし。", "38", "購読"),
 ("Pronovost P, Needham D, Berenholtz S, et al. An intervention to decrease catheter-related bloodstream infections in the ICU. N Engl J Med. 2006;355(26):2725-32.",
  "10.1056/NEJMoa061115", "17192537", "多施設共同コホート（103 ICU）", "総合医学のトップジャーナル",
  "カテーテル関連血流感染の中央値2.7→0/1000カテーテル日（3か月後）、平均7.7→1.4（16〜18か月後）。最大66%の持続的減少。", "41", "購読"),
 ("Dixon-Woods M, Bosk CL, Aveling EL, Goeschel CA, Pronovost PJ. Explaining Michigan: developing an ex post theory of a quality improvement program. Milbank Q. 2011;89(2):167-205.",
  "10.1111/j.1468-0009.2011.00625.x", "21676020", "事後的プログラム理論の構築（質的・理論的分析）", "医療政策・医療社会学の主要誌",
  "ミシガンの成功要因：参加への同型化圧力、水平的ネットワークによる規範的圧力、感染を社会的問題として再定義、複数介入による文化づくり、感染率データの活用、明確なルール（hard edges）。", "41", "PMCで閲覧可"),
 ("Lotfinejad N, Peters A, Tartari E, et al. Hand hygiene in health care: 20 years of ongoing advances and perspectives. Lancet Infect Dis. 2021;21(8):e209-21.",
  "10.1016/S1473-3099(21)00383-2", "34331890", "レビュー", "感染症のトップジャーナル",
  "WHO多角的戦略は、システム変更（行動変容の前提）、教育、モニタリングとパフォーマンスフィードバック、職場でのリマインダー、施設の安全風土の組み合わせを要する。", "42", "購読"),
]

def build():
    doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
                            title="引用文献リスト：行動科学は感染対策の「見えない敵」を可視化する", author="鈴木 徳洋")
    E = []
    E.append(Paragraph("引用文献リスト", H1))
    E.append(Paragraph("講演「行動科学は感染対策の『見えない敵』を可視化する」（約60分）", SUB))
    E.append(Spacer(1, 4 * mm))
    note = ("・各文献の書誌情報と数値は、PubMedの抄録・全文（PMC）で原典を確認しました。"
            "総合医学誌（Lancet, NEJM, JAMA）、内科・感染症のトップ誌（Ann Intern Med, JAMA Intern Med, Lancet Infect Dis）、"
            "コクランレビュー、心理学・実装科学の主要誌を優先して採用しています。<br/>"
            "・「講演で用いた要点」は講演者向けの日本語要約です。引用の際は必ず原典をご確認ください。<br/>"
            "・原著PDFはこの作成環境のネットワーク制限により取得できなかったため、DOI・PubMedのリンクを掲載しています。"
            "「オープンアクセス」「PMCで閲覧可」と記載した文献は、リンク先から無料で全文を閲覧できます。")
    t = Table([[Paragraph(note, SMALL)]], colWidths=[178 * mm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), CARD), ("BOX", (0, 0), (-1, -1), 0, CARD),
                           ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                           ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    E.append(t)
    E.append(Spacer(1, 5 * mm))
    for i, (cit, doi, pmid, design, tier, finding, slides, access) in enumerate(R, 1):
        links = []
        if doi:
            links.append(f'DOI: <a href="https://doi.org/{doi}" color="#0B7A70">https://doi.org/{doi}</a>')
        if pmid:
            links.append(f'PubMed: <a href="https://pubmed.ncbi.nlm.nih.gov/{pmid}/" color="#0B7A70">PMID {pmid}</a>')
        if cit.startswith("The Behavioural"):
            links.append('URL: <a href="https://www.bi.team/publications/east-four-simple-ways-to-apply-behavioural-insights/" color="#0B7A70">bi.team（EAST）</a>')
        rows = [
            [Paragraph(f"{i}.", REFT), Paragraph(cit, REFT)],
            ["", Paragraph("　".join(links), LINK)] if links else ["", Paragraph("", LINK)],
            ["", Paragraph(f"<b>研究デザイン</b>：{design}　／　<b>掲載誌</b>：{tier}　／　<b>全文</b>：{access}", SMALL)],
            ["", Paragraph(f"<b>講演で用いた要点</b>：{finding}", BODY)],
            ["", Paragraph(f"使用スライド：{slides}", SMALL)],
        ]
        tb = Table(rows, colWidths=[9 * mm, 169 * mm])
        tb.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 1),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                                ("LINEBELOW", (0, -1), (-1, -1), 0.5, colors.HexColor("#C9D2DE"))]))
        E.append(KeepTogether([tb, Spacer(1, 4 * mm)]))
    doc.build(E)
    print("wrote", OUT)

build()
