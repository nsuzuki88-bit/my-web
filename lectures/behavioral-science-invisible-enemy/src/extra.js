// 行動科学は感染対策の「見えない敵」を可視化する（約60分）
// タイトル30pt以上・本文20pt以上・BIZ UDゴシック／1スライド1メッセージ
const pptxgen = require("pptxgenjs");
const { build: buildAssets } = require("./assets");

const OUT = process.argv[2] || "extra.pptx";
const FONT = "BIZ UDゴシック";
const C = {
  NAVY: "131B35", NAVY2: "1E2A4F", GLOW: "3FE8C2", TEAL: "0B7A70",
  ORANGE: "D1480E", INK: "1B2333", MUTED: "4F5B6B", CARD: "EEF2F7",
  LINE: "C9D2DE", WHITE: "FFFFFF", MINT: "E3F7F2", PEACH: "FCEBE3",
  SLATE: "8A96A8", PALE: "C9D4E8", BLUE: "2F6FA8",
};
const W = 13.333, MX = 0.6, CW = W - 2 * MX;

// BIZ UDゴシックは等幅：全角=1em、半角=0.5em
const em = (s) => [...s].reduce((w, ch) => {
  const c = ch.codePointAt(0);
  return w + (c < 0x7f || (c >= 0xff61 && c <= 0xff9f) ? 0.5 : 1);
}, 0);
const tw = (s, pt) => (em(s) * pt) / 72;

let A; // assets
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "鈴木 徳洋";
pres.title = "行動科学は感染対策の「見えない敵」を可視化する";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

// ---------- 基本ヘルパー ----------
function T(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: FONT, color: C.INK, fontSize: 20, margin: 0,
    isTextBox: true, valign: "top",
  }, o));
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.ShapeType.roundRect, Object.assign({
    x, y, w, h, fill: { color: fill }, line: { color: o.lineColor || fill, width: o.lineWidth || 1, dashType: o.dash || "solid" },
    rectRadius: o.r ?? 0.12,
  }, {}));
}
function circle(s, x, y, d, fill, o = {}) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d,
    fill: o.noFill ? { type: "none" } : { color: fill },
    line: { color: o.lineColor || fill, width: o.lineWidth || 1 },
  });
}
function iconCircle(s, x, y, d, fill, img) {
  circle(s, x, y, d, fill);
  const p = d * 0.24;
  s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
function numCircle(s, x, y, d, fill, txt, color, pt) {
  circle(s, x, y, d, fill);
  T(s, txt, { x, y, w: d, h: d, fontSize: pt || 28, bold: true, color: color || C.WHITE, align: "center", valign: "middle" });
}
function arrow(s, x1, y1, x2, y2, color = C.SLATE, width = 3) {
  s.addShape(pres.ShapeType.line, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001,
    flipH: x2 < x1, line: { color, width, endArrowType: "triangle" },
  });
}
function chevron(s, x, y, w = 0.4, h = 0.6, color = C.SLATE) {
  s.addShape(pres.ShapeType.chevron, { x, y, w, h, fill: { color }, line: { color } });
}
function pill(s, text, color) {
  const w = tw(text, 20) + 0.5;
  box(s, MX, 0.38, w, 0.46, color, { r: 0.23 });
  T(s, text, { x: MX, y: 0.38, w, h: 0.46, fontSize: 20, bold: true, color: C.WHITE, align: "center", valign: "middle" });
}
function title(s, text) {
  let fs = 32;
  if (tw(text, 32) > CW - 0.15) fs = 30;
  if (tw(text, fs) > CW - 0.15) console.warn("!! TITLE WRAPS:", text);
  T(s, text, { x: MX, y: 0.95, w: CW, h: 0.85, fontSize: fs, bold: true, color: C.NAVY, valign: "middle" });
}
function source(s, text) {
  const t = "出典：" + text;
  if (tw(t, 20) > 10.95) console.warn("!! SOURCE TOO LONG:", t, tw(t, 20).toFixed(2));
  T(s, t, { x: MX, y: 6.8, w: 11.0, h: 0.42, fontSize: 20, color: C.MUTED, valign: "middle" });
}
function slideNo(s, dark) {
  s.slideNumber = { x: 12.03, y: 6.8, w: 0.7, h: 0.42, fontFace: FONT, fontSize: 20, color: dark ? "8FA0C0" : C.SLATE, align: "right" };
}
function lightSlide({ pillText, pillColor, titleText, src, notes }) {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };
  pill(s, pillText, pillColor || C.NAVY);
  title(s, titleText);
  if (src) source(s, src);
  slideNo(s, false);
  s.addNotes(notes);
  return s;
}
function darkSlide(notes) {
  const s = pres.addSlide();
  s.background = { color: C.NAVY };
  slideNo(s, true);
  s.addNotes(notes);
  return s;
}
// 箇条書き（段落ごとに runs 配列）
function paras(items, base = {}) {
  return items.map((it, i) => {
    const o = typeof it === "string" ? { text: it } : it;
    return { text: o.text, options: Object.assign({ breakLine: i < items.length - 1 }, base, o.options || {}) };
  });
}
// "\n" を含む項目は段落内改行として扱う（項目間だけ余白を空ける）
function expand(lines, gap) {
  const out = [];
  lines.forEach((it) => {
    const o = typeof it === "string" ? { text: it } : it;
    const segs = o.text.split("\n");
    segs.forEach((t, k) => out.push({ text: t, options: Object.assign({}, o.options || {}, { paraSpaceAfter: k < segs.length - 1 ? 0 : gap }) }));
  });
  return out;
}
function card(s, x, y, w, h, fill, lines, o = {}) {
  box(s, x, y, w, h, fill);
  T(s, paras(expand(lines, o.gap ?? 10)), {
    x: x + 0.28, y: y + 0.25, w: w - 0.56, h: h - 0.5, fontSize: o.pt || 22, valign: o.valign || "top",
  });
}
function barChart(s, o) {
  const opts = {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.dir || "col",
    chartColors: o.colors, showValue: true, dataLabelPosition: "outEnd",
    dataLabelFontSize: o.valuePt || 24, dataLabelFontBold: true, dataLabelFontFace: FONT,
    dataLabelColor: C.INK, dataLabelFormatCode: o.fmt || "General",
    catAxisLabelFontSize: 20, catAxisLabelFontFace: FONT, catAxisLabelColor: C.INK,
    catAxisLineShow: true, valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" },
    valAxisMinVal: 0, valAxisMaxVal: o.max, barGapWidthPct: o.gap ?? 55,
    showLegend: !!o.legend, legendPos: "b", legendFontSize: 20, legendFontFace: FONT, legendColor: C.INK,
    barGrouping: "clustered",
  };
  const data = o.series || [{ name: o.name || "値", labels: o.labels, values: o.values }];
  s.addChart(pres.ChartType.bar, data, opts);
}
function caption(s, text, x = MX, y = 1.95, w = 7.3) {
  T(s, text, { x, y, w, h: 0.4, fontSize: 20, color: C.MUTED });
}
function section(part, text, notes, items, lensIcon, sub) {
  const s = darkSlide(notes);
  s.addImage({ data: A.lensPlain, x: 8.55, y: 1.25, w: 4.2, h: 4.2 });
  if (lensIcon) s.addImage({ data: A.icon[lensIcon].glow, x: 9.85, y: 2.55, w: 1.6, h: 1.6 });
  const nl = text.split("\n").length;
  const y0 = items ? 1.9 : (nl > 1 ? 2.1 : 2.45);
  T(s, part, { x: 0.9, y: y0, w: 7.4, h: 0.5, fontSize: 24, bold: true, color: C.GLOW });
  T(s, paras(text.split("\n")), { x: 0.9, y: y0 + 0.6, w: 7.4, h: 0.72 * nl + 0.1, fontSize: 40, bold: true, color: C.WHITE, valign: "top" });
  if (sub) T(s, sub, { x: 0.9, y: y0 + 0.9 + 0.72 * nl, w: 7.4, h: 0.6, fontSize: 24, color: C.PALE });
  if (items) {
    items.forEach(([ic, label], i) => {
      const y = 4.35 + i * 0.72;
      iconCircle(s, 0.9, y, 0.56, C.NAVY2, A.icon[ic].glow);
      T(s, label, { x: 1.65, y, w: 6.5, h: 0.56, fontSize: 24, color: C.WHITE, valign: "middle" });
    });
  }
  // レンズ内のアイコン
  return s;
}

function build() {
  const NUD = C.BLUE, P = "ナッジって何？";
  const PART1 = C.NAVY;

  // S1 システム1とシステム2
  {
    const s = lightSlide({
      pillText: "第1部", pillColor: PART1,
      titleText: "頭の中には「速い思考」と「遅い思考」がある",
      src: "Kahneman D. Am Psychol 2003; Frederick S. J Econ Perspect 2005",
      notes:
`【目安1分15秒】
ここで、なぜ「わかっているのに、できない」のかを、頭の働きから考えてみます。
まずクイズです。バットとボールで合わせて1,100円。バットはボールより1,000円高い。ボールはいくらでしょう？……とっさに「100円」と浮かんだ方が多いと思います。でも、それだとバットは1,100円で、合計は1,200円になってしまいます。正解は50円です。
ノーベル経済学賞を受賞した心理学者カーネマンは、人の思考を2つに分けて説明しました。システム1は「速い思考」。自動的・直感的で、努力がいらず、慣れた手順や忙しいときの判断を担います。システム2は「遅い思考」。意識的・論理的ですが、エネルギーを使い、疲れると働きにくくなります。
私たちの毎日の行動の多くは、速い思考で動いています。
（出典：Kahneman D. Am Psychol 2003;58:697-720／Frederick S. J Econ Perspect 2005;19:25-42。クイズは原典の米ドルを円に置き換えています）`,
    });
    // 左：クイズ
    box(s, MX, 2.15, 5.5, 4.4, C.CARD);
    T(s, "クイズ", { x: MX + 0.3, y: 2.35, w: 4.9, h: 0.5, fontSize: 24, bold: true, color: C.ORANGE });
    T(s, paras(expand([
      "バットとボールで合計1,100円",
      "バットはボールより1,000円高い",
      { text: "ボールはいくら？", options: { bold: true, color: C.NAVY } },
    ], 4)), { x: MX + 0.3, y: 2.95, w: 4.9, h: 1.5, fontSize: 22 });
    T(s, [
      { text: "とっさに「100円」→ 速い思考", options: { breakLine: true } },
      { text: "正解は「50円」→ 遅い思考", options: { bold: true, color: C.TEAL } },
    ], { x: MX + 0.3, y: 4.85, w: 4.9, h: 1.3, fontSize: 20, paraSpaceAfter: 8, valign: "middle" });
    // 右：2つのシステム
    const x0 = 6.4, rw = W - MX - x0;
    const sys = [
      ["FaBolt", "システム1：速い思考", "自動的・直感的・努力いらず", "慣れた手順、忙しい時の判断", C.ORANGE, C.PEACH],
      ["FaHourglassHalf", "システム2：遅い思考", "意識的・論理的・疲れやすい", "研修で学ぶ、ルールを考える", C.TEAL, C.MINT],
    ];
    sys.forEach(([ic, h, d, e, col, bg], i) => {
      const y = 2.15 + i * 2.25;
      box(s, x0, y, rw, 2.1, bg);
      iconCircle(s, x0 + 0.25, y + 0.25, 0.85, col, A.icon[ic].white);
      T(s, h, { x: x0 + 1.3, y: y + 0.25, w: rw - 1.45, h: 0.85, fontSize: 24, bold: true, color: col, valign: "middle" });
      T(s, [{ text: d, options: { breakLine: true } }, { text: "例：" + e, options: { color: C.MUTED } }],
        { x: x0 + 0.3, y: y + 1.15, w: rw - 0.5, h: 0.85, fontSize: 20, valign: "middle" });
    });
  }

  // S2 感染対策と2つの思考
  {
    const s = lightSlide({
      pillText: "第1部", pillColor: PART1,
      titleText: "研修は「遅い思考」に届くが、現場は「速い思考」で動く",
      src: "Whitby M, et al. ICHE 2006; Erasmus V, et al. ICHE 2010",
      notes:
`【目安1分】
感染対策に当てはめてみます。
看護師754名を対象にした研究では、手洗いには2種類あると整理されました。汚れを感じたときや体液に触れた後に「思わず洗う」手洗いと、患者さんの健康な皮膚に触れる前のように、意識して「選んで行う」手洗いです。そして、どちらの手洗いも、子どもの頃から身についた家庭や地域での手洗いの習慣に最も強く影響されていました。
先ほどの遵守率のデータでも、患者さんに触れた後の遵守率は47%、触れる前は21%でした。「汚れた」と感じる場面では速い思考で手が動きますが、「触れる前」は遅い思考で意識しないとできないと考えると、この差が理解しやすくなります。
研修で知識を伝えることは、遅い思考に届きます。でも、忙しい現場で行動を動かしているのは速い思考です。だから、環境の側から速い思考に働きかける工夫が必要になります。それが、後半でご紹介する「ナッジ」です。
（出典：Whitby M, et al. Infect Control Hosp Epidemiol 2006;27:484-92／Erasmus V, et al. Infect Control Hosp Epidemiol 2010;31:283-94。2つの手洗いを速い思考・遅い思考に対応させる説明は講演者による解釈です）`,
    });
    caption(s, "手指衛生の遵守率（中央値）", MX, 1.95, 6.4);
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["患者に触れた後\n（思わず洗う）", "患者に触れる前\n（意識して行う）"], values: [47, 21],
      colors: [C.ORANGE, C.TEAL], fmt: '0"%"', max: 58, valuePt: 28,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "「汚れた」と感じると\n速い思考で手が動く", options: { bold: true, color: C.ORANGE } },
      { text: "「触れる前」は\n遅い思考で意識しないと\nできない", options: { bold: true, color: C.TEAL } },
      { text: "→ 研修（知る）に加えて\n速い思考に働きかける\n環境づくり（ナッジ）を", options: { bold: true, color: C.NAVY } },
    ], { gap: 14, pt: 20, valign: "middle" });
  }

  // N1 ナッジとは
  {
    const s = lightSlide({
      pillText: P, pillColor: NUD,
      titleText: "ナッジ＝ひじで「そっと押す」。禁止も強制もしない",
      src: "Thaler RH, Sunstein CR. Nudge. Yale Univ Press 2008",
      notes:
`【目安1分】
ここで、今日のキーワード「ナッジ」を確認しておきます。
ナッジは英語で「ひじで軽くつつく」という意味です。行動経済学者のセイラーとサンスティーンが広めた考え方で、禁止や強制をしたり、お金で釣ったりせずに、選ぶ自由は残したまま、望ましい行動を「選びやすく」する工夫のことです。
実は皆さんの身の回りにもあります。レジ前の足あとマークがあると、言われなくても間隔をあけて並びます。「一番人気」と書いてあると、ついそれを選びます。最初からチェックが入っている欄は、そのままにしがちです。
先ほどの「1通の手紙」も、周りと比べた自分の位置を見せるナッジでした。
第1部でお話しした「速い思考（システム1）」は、理屈より目の前の環境に反応します。ナッジは、その速い思考に働きかけて、考えなくても望ましい行動を選べるようにする方法です。
（出典：Thaler RH, Sunstein CR. Nudge: Improving Decisions about Health, Wealth, and Happiness. Yale University Press; 2008）`,
    });
    // 左：定義
    box(s, MX, 2.15, 5.6, 4.35, C.NAVY);
    iconCircle(s, MX + 0.35, 2.45, 1.0, C.NAVY2, A.icon.FaHandPointer.glow);
    T(s, "nudge（ナッジ）\n＝ひじで軽くつつく", { x: MX + 1.6, y: 2.45, w: 3.9, h: 1.0, fontSize: 22, bold: true, color: C.GLOW, valign: "middle" });
    T(s, paras(expand([
      "選ぶ自由は残したまま、",
      "望ましい行動を",
      { text: "「選びやすく」する工夫", options: { bold: true, color: C.GLOW } },
    ], 4)), { x: MX + 0.35, y: 3.6, w: 5.0, h: 1.5, fontSize: 26, color: C.WHITE, valign: "top" });
    T(s, "× 禁止・強制　× ごほうび・罰金", { x: MX + 0.35, y: 5.2, w: 5.0, h: 0.5, fontSize: 20, color: C.PALE, valign: "middle" });
    T(s, "→ 主に「速い思考」に働きかける", { x: MX + 0.35, y: 5.75, w: 5.0, h: 0.5, fontSize: 20, bold: true, color: C.GLOW, valign: "middle" });
    // 右：身近な例
    const x0 = 6.6, rw = W - MX - x0;
    T(s, "身近なナッジ", { x: x0, y: 2.15, w: rw, h: 0.5, fontSize: 24, bold: true, color: C.NAVY });
    const ex = [
      ["FaShoePrints", "レジ前の足あとマーク", "→ 自然と間隔をあけて並ぶ"],
      ["FaThumbsUp", "「一番人気」の表示", "→ みんなと同じものを選ぶ"],
      ["FaCheckSquare", "最初からチェック済みの欄", "→ そのままにしがち"],
    ];
    ex.forEach(([ic, a, b], i) => {
      const y = 2.8 + i * 1.25;
      box(s, x0, y, rw, 1.1, C.CARD);
      iconCircle(s, x0 + 0.2, y + 0.17, 0.76, C.BLUE, A.icon[ic].white);
      T(s, [{ text: a, options: { bold: true, color: C.NAVY, breakLine: true } }, { text: b }],
        { x: x0 + 1.15, y: y + 0.08, w: rw - 1.25, h: 0.94, fontSize: 20, valign: "middle" });
    });
  }


  // M 当院の実践（写真）
  {
    const s = lightSlide({
      pillText: "見える化③ 当院の実践", pillColor: C.TEAL,
      titleText: "当院の実践：3つのナッジで「思わず」消毒する場をつくる",
      src: "鈴木徳洋, 他. 日本外科感染症学会（ポスター発表）",
      notes:
`【目安1分30秒】
ここで当院の実践をご紹介します。写真の3つは、いずれも「考えなくても手が動く」ことをねらったナッジです。
一つ目は救急外来の入口です。床の足あとで手指消毒剤のスタンドへ誘導し、目立たせました。患者さんのご家族が入室時に手指消毒をする頻度が増えました。
二つ目は病室の名札の下の表示です。「①氏名を確認したら、②手指消毒を2プッシュ」と、毎回必ず行う氏名確認を合図にして、次の行動を具体的に示しています。
三つ目は病室前の宣言です。「患者さんを感染から守るため手指消毒をします」とスタッフが公言することで、言ったことを守ろうとする心理と、「患者さんのため」という目的を同時に働かせます。
このほか、救急外来で防護具の着用手順を番号で示したことや、手指消毒剤を1本使い切るごとにシールを貼って個人の使用量を見える化したことにも取り組みました。うまくいった方法やツールは、実施部署から全部署に共有しています。
（出典：鈴木徳洋, 他. 日本外科感染症学会（ポスター発表））`,
    });
    const items = [
      ["photos/p1.jpg", "入口へ誘導する", "足あとで消毒剤へ（救急外来）"],
      ["photos/p2.jpg", "確認を合図にする", "氏名確認→消毒2プッシュ"],
      ["photos/p3.jpg", "宣言して約束する", "スタッフが患者さんに宣言"],
    ];
    const cw = 3.95, g = (CW - 3 * cw) / 2, pw = 3.0, ph = 3.3;
    items.forEach(([img, h, c], i) => {
      const x = MX + i * (cw + g);
      numCircle(s, x, 1.95, 0.6, C.TEAL, String(i + 1), C.WHITE, 24);
      s.addImage({ path: img, x: x + (cw - pw) / 2 + 0.2, y: 1.95, w: pw, h: ph });
      T(s, h, { x, y: 5.3, w: cw, h: 0.42, fontSize: 22, bold: true, color: C.NAVY, align: "center", valign: "middle" });
      T(s, c, { x, y: 5.72, w: cw, h: 0.4, fontSize: 20, color: C.INK, align: "center", valign: "middle" });
    });
    box(s, MX, 6.2, CW, 0.52, C.MINT, { r: 0.08 });
    T(s, "このほか：防護具の手順の番号化／使用量シール → 全部署へ共有", { x: MX, y: 6.2, w: CW, h: 0.52, fontSize: 20, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // MI ミシガン
  {
    const s = lightSlide({
      pillText: "第4部 仕掛ける", pillColor: C.NAVY,
      titleText: "成功の鍵は、チェックリストの裏の「見えない社会の力」",
      src: "Pronovost P, et al. NEJM 2006; Dixon-Woods M, et al. Milbank Q 2011",
      notes:
`【目安1分30秒】
組織として仕掛けた有名な例が、米国ミシガン州のICUプロジェクトです。
やったことは、中心静脈カテーテルの感染を防ぐ、エビデンスのある5つの対策です。手指衛生、挿入時のマキシマル・バリアプリコーション、クロルヘキシジンによる皮膚消毒、できるだけ大腿部を避けること、不要なカテーテルを抜くこと。これを物品カートとチェックリストで支え、守られていなければ緊急時を除いて手技を止めることにしました。
その結果、103のICUで、カテーテル関連血流感染は1000カテーテル日あたり平均7.7件から1.4件へ、最大66%減りました。
5つの対策自体は、以前から知られていたものです。では、なぜここまで成功したのか。後の社会学的な分析では、ICU同士が横につながり、「みんなやっている」ことが規範になったこと、感染を「仕方ないもの」から「防げるもの」と捉え直したこと、感染率を測って共有し、互いに比べられるようにしたこと、そして手順を守らなければ止めるという、ゆずれない仕組みを持ったことが鍵だったとされています。
道具の裏にある「見えない社会の力」こそが、行動を変えるのです。
（出典：Pronovost P, et al. N Engl J Med 2006;355:2725-32／Dixon-Woods M, et al. Milbank Q 2011;89:167-205）`,
    });
    // 左：何をしたか
    const lw = 5.7;
    box(s, MX, 2.1, lw, 4.5, C.CARD);
    T(s, "やったこと：5つの基本対策", { x: MX + 0.3, y: 2.25, w: lw - 0.5, h: 0.5, fontSize: 22, bold: true, color: C.NAVY });
    T(s, paras(["① 手指衛生", "② 挿入時のマキシマル・バリア", "③ クロルヘキシジンで皮膚消毒", "④ できるだけ大腿部を避ける", "⑤ 不要なカテーテルを抜く"], { paraSpaceAfter: 2 }),
      { x: MX + 0.3, y: 2.8, w: lw - 0.5, h: 2.05, fontSize: 20 });
    T(s, "＋物品カート・チェックリスト", { x: MX + 0.3, y: 4.85, w: lw - 0.5, h: 0.4, fontSize: 20, color: C.MUTED });
    box(s, MX + 0.25, 5.35, lw - 0.5, 1.1, C.WHITE);
    T(s, [
      { text: "血流感染 最大66%減（103 ICU）", options: { bold: true, color: C.TEAL, breakLine: true } },
      { text: "平均7.7→1.4／1000カテーテル日" },
    ], { x: MX + 0.35, y: 5.4, w: lw - 0.7, h: 1.0, fontSize: 20, valign: "middle" });
    // 右：なぜ成功したか
    const x0 = MX + lw + 0.35, rw = W - MX - x0;
    T(s, "なぜ成功したか（事後の社会学的分析）", { x: x0, y: 2.1, w: rw, h: 0.5, fontSize: 22, bold: true, color: C.ORANGE });
    const why = [
      ["FaUsers", "仲間のネットワーク", "ICU同士がつながり、規範に"],
      ["FaLightbulb", "問題の捉え直し", "「仕方ない」から「防げる」へ"],
      ["FaChartLine", "データの共有", "感染率を測って返し、比べる"],
      ["FaHandPaper", "ゆずれない仕組み", "手順を守らなければ止める"],
    ];
    why.forEach(([ic, h, d], i) => {
      const y = 2.7 + i * 0.98;
      iconCircle(s, x0, y + 0.05, 0.72, C.ORANGE, A.icon[ic].white);
      T(s, [{ text: h, options: { bold: true, color: C.NAVY, breakLine: true } }, { text: d }],
        { x: x0 + 0.9, y, w: rw - 0.9, h: 0.85, fontSize: 20, valign: "middle" });
    });
  }
}

(async () => {
  A = await buildAssets();
  build();
  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
