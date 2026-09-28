// 行動科学は感染対策の「見えない敵」を可視化する（約60分）
// タイトル30pt以上・本文20pt以上・BIZ UDゴシック／1スライド1メッセージ
const pptxgen = require("pptxgenjs");
const { build: buildAssets } = require("./assets");

const OUT = process.argv[2] || "deck.pptx";
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

// ---------- スライド ----------
function build() {
  // 1 タイトル
  {
    const s = darkSlide(
`【目安1分／0:00–1:00】
皆さん、こんにちは。市立豊中病院 感染対策室、感染管理認定看護師の鈴木です。本日は「行動科学は感染対策の『見えない敵』を可視化する」というテーマで、約60分お話しします。
医師、看護師、薬剤師、検査技師、リハビリ、放射線、事務や清掃の方まで、職種を問わず明日から使える視点をお持ち帰りいただくことが目標です。専門用語はできるだけかみ砕いてご説明しますので、行動科学がはじめての方も安心してお聞きください。`);
    s.addImage({ data: A.lensTitle, x: 8.45, y: 1.3, w: 4.4, h: 4.4 });
    T(s, "感染対策講習会　2026年10月15日", { x: 0.9, y: 1.25, w: 7.4, h: 0.45, fontSize: 20, bold: true, color: C.GLOW });
    T(s, [
      { text: "行動科学は感染対策の", options: { breakLine: true } },
      { text: "「見えない敵」", options: { color: C.GLOW } },
      { text: "を可視化する" },
    ], { x: 0.9, y: 1.85, w: 7.6, h: 1.9, fontSize: 40, bold: true, color: C.WHITE, valign: "middle", lineSpacingMultiple: 1.1 });
    T(s, paras(["「わかっているのに、できない」の正体を", "多職種で読み解く"]),
      { x: 0.9, y: 3.95, w: 7.6, h: 1.0, fontSize: 24, color: C.PALE });
    T(s, paras(["市立豊中病院 感染対策室", "感染管理認定看護師　鈴木 徳洋"]),
      { x: 0.9, y: 5.35, w: 7.6, h: 0.9, fontSize: 20, color: C.WHITE });
  }

  // 2 問いかけ
  {
    const s = darkSlide(
`【目安2分／1:00–3:00】
最初に皆さんに質問です。感染対策の「敵」とは何でしょうか。隣の方と30秒ほど話してみてください。
（30秒）
ありがとうございます。細菌、ウイルス、耐性菌、といった声が多かったと思います。もちろん正解です。
ただ、今日はもう一つ、もっと手ごわい敵の話をします。それは、私たちの「行動」をゆがめてしまう、目に見えない力です。`);
    s.addImage({ data: A.lensPlain, x: 5.62, y: 0.75, w: 2.1, h: 2.1 });
    s.addImage({ data: A.icon.FaQuestion.glow, x: 6.27, y: 1.4, w: 0.8, h: 0.8 });
    T(s, paras(["感染対策の「敵」は、", "何だと思いますか？"]),
      { x: 0.9, y: 3.05, w: 11.53, h: 1.9, fontSize: 40, bold: true, color: C.WHITE, align: "center", valign: "middle" });
    T(s, "隣の方と30秒、話してみてください", { x: 0.9, y: 5.25, w: 11.53, h: 0.6, fontSize: 26, color: C.GLOW, align: "center" });
  }

  // 3 微生物は見える化してきた
  {
    const s = lightSlide({
      pillText: "はじめに", titleText: "私たちは、微生物を「見える化」する道具を持っている",
      notes:
`【目安1分／3:00–4:00】
私たちは、微生物という目に見えない敵を、道具で「見える化」してきました。顕微鏡とグラム染色で形を見て、培養と感受性検査で効く薬を見て、遺伝子検査で正体を見て、蛍光マーカーやATPで汚れを見る。
見えるからこそ、私たちは対策を立てることができます。`,
    });
    const items = [
      ["FaMicroscope", "顕微鏡", "「形」を見る"],
      ["FaVial", "培養・感受性", "「効く薬」を見る"],
      ["FaDna", "遺伝子検査", "「正体」を見る"],
      ["FaSearchPlus", "蛍光・ATP", "「汚れ」を見る"],
    ];
    const w = 2.8, g = (CW - 4 * w) / 3;
    items.forEach(([ic, h, sub], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.2, w, 3.1, C.CARD);
      iconCircle(s, x + w / 2 - 0.6, 2.5, 1.2, C.TEAL, A.icon[ic].white);
      T(s, h, { x, y: 3.9, w, h: 0.55, fontSize: 24, bold: true, color: C.NAVY, align: "center" });
      T(s, sub, { x, y: 4.5, w, h: 0.5, fontSize: 22, color: C.INK, align: "center" });
    });
    T(s, "見えるから、対策を立てられる", { x: MX, y: 5.7, w: CW, h: 0.7, fontSize: 28, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // 4 氷山
  {
    const s = lightSlide({
      pillText: "はじめに", titleText: "感染を広げる「行動の原因」は、目に見えない",
      notes:
`【目安1分30秒／4:00–5:30】
では、感染を広げてしまう「行動」の原因はどうでしょうか。
手指衛生を1回飛ばしてしまう、防護具を外すときに自分を汚染してしまう、必要のない抗菌薬を処方してしまう。これらは水面の上に見えている「行動」です。
しかし、その下には、思い込み、見られているかどうか、疲れ、周囲の空気、環境や動線といった「見えない原因」が隠れています。
今日は、この水面下の敵を、行動科学というレンズで見える化していきます。`,
    });
    s.addImage({ data: A.iceberg, x: 0.6, y: 2.0, w: 6.1, h: 4.53 });
    T(s, "行動", { x: 3.3, y: 2.95, w: 0.9, h: 0.45, fontSize: 24, bold: true, color: C.NAVY, align: "center" });
    T(s, "原因", { x: 3.15, y: 4.65, w: 1.2, h: 0.6, fontSize: 30, bold: true, color: C.ORANGE, align: "center" });
    const rx = 7.1, rw = CW - (rx - MX);
    T(s, "水面の上＝見える「行動」", { x: rx, y: 2.05, w: rw, h: 0.5, fontSize: 24, bold: true, color: C.NAVY });
    T(s, paras(["手指衛生のスキップ", "防護具を外すときの汚染", "不要な抗菌薬の処方"], { bullet: true, paraSpaceAfter: 4 }),
      { x: rx, y: 2.6, w: rw, h: 1.4, fontSize: 22 });
    T(s, "水面の下＝見えない「原因」", { x: rx, y: 4.2, w: rw, h: 0.5, fontSize: 24, bold: true, color: C.ORANGE });
    const chips = [["思い込み", "見られている効果", "疲れ"], ["周囲の空気", "環境と動線"]];
    chips.forEach((row, r) => {
      let x = rx;
      row.forEach((t) => {
        const w = tw(t, 20) + 0.4;
        box(s, x, 4.8 + r * 0.62, w, 0.5, C.PEACH, { r: 0.25 });
        T(s, t, { x, y: 4.8 + r * 0.62, w, h: 0.5, fontSize: 20, bold: true, color: C.ORANGE, align: "center", valign: "middle" });
        x += w + 0.15;
      });
    });
    T(s, "→ 今日は、この水面下を可視化します", { x: rx, y: 6.2, w: rw, h: 0.5, fontSize: 22, bold: true, color: C.TEAL });
  }

  // 5 行動科学は顕微鏡
  {
    const s = lightSlide({
      pillText: "はじめに", titleText: "行動科学は、行動を見るための「顕微鏡」",
      notes:
`【目安1分／5:30–6:30】
微生物学では、顕微鏡や培養で見えないものを見て、消毒や抗菌薬で対処します。
行動科学も同じです。理論とデータで行動の原因を見えるようにし、仕組みや環境を設計して対処します。
つまり行動科学は、行動を見るための「顕微鏡」だと考えていただくとわかりやすいと思います。`,
    });
    const rows = [
      ["微生物", C.TEAL, C.CARD, "顕微鏡・培養で見る", "消毒・抗菌薬で対処する"],
      ["行動", C.ORANGE, C.MINT, "理論とデータで原因を見る", "仕組みと環境を設計して対処"],
    ];
    rows.forEach(([lab, col, fill, a, b], i) => {
      const y = 2.35 + i * 2.05;
      numCircle(s, MX, y, 1.5, col, lab, C.WHITE, 24);
      box(s, 2.4, y + 0.1, 4.5, 1.3, fill);
      T(s, a, { x: 2.55, y: y + 0.1, w: 4.2, h: 1.3, fontSize: 24, bold: true, color: C.NAVY, align: "center", valign: "middle" });
      arrow(s, 7.0, y + 0.75, 7.6, y + 0.75, C.SLATE, 4);
      box(s, 7.75, y + 0.1, 4.98, 1.3, fill);
      T(s, b, { x: 7.9, y: y + 0.1, w: 4.68, h: 1.3, fontSize: 24, bold: true, color: C.NAVY, align: "center", valign: "middle" });
    });
    T(s, "見えないものを「見る」→ だから「対処」できる", { x: MX, y: 6.05, w: CW, h: 0.55, fontSize: 22, color: C.MUTED, align: "center" });
  }

  // 6 今日のゴール
  {
    const s = lightSlide({
      pillText: "はじめに", titleText: "今日のゴール：明日、ひとつ「仕掛ける」",
      notes:
`【目安1分／6:30–7:30】
今日のゴールは3つのステップです。まず、5つの「見えない敵」を知ること。次に、COM-B（コム・ビー）という枠組みで原因を診断できるようになること。そして最後に、明日から小さく一つ仕掛けること。
聞いて終わりではなく、明日一つだけ行動が変わることを目指しましょう。`,
    });
    const items = [["1", "知る", "5つの\n「見えない敵」を知る"], ["2", "見る", "COM-Bで\n原因を診断する"], ["3", "動く", "明日から\n小さく仕掛ける"]];
    const w = 3.6, g = (CW - 3 * w) / 2;
    items.forEach(([n, h, sub], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.25, w, 3.7, C.CARD);
      numCircle(s, x + w / 2 - 0.55, 2.55, 1.1, C.NAVY, n, C.GLOW, 36);
      T(s, h, { x, y: 3.85, w, h: 0.65, fontSize: 30, bold: true, color: C.NAVY, align: "center" });
      T(s, sub, { x: x + 0.25, y: 4.6, w: w - 0.5, h: 1.2, fontSize: 22, color: C.INK, align: "center" });
      if (i < 2) chevron(s, x + w + g / 2 - 0.18, 3.9, 0.36, 0.6);
    });
  }

  // 7 第1部
  section("第1部", "感染対策の成否は\n「行動」で決まる",
`【目安20秒／7:30–7:50】
第1部では、なぜ感染対策が「行動」の問題なのかを確認します。`, null, "FaWalking", "なぜ「行動」の問題なのか");

  // 8 遵守率40%
  {
    const s = lightSlide({
      pillText: "第1部", titleText: "手指衛生の遵守率は、世界でおよそ40%",
      src: "Erasmus V, et al. Infect Control Hosp Epidemiol 2010",
      notes:
`【目安1分30秒／7:50–9:20】
手指衛生の遵守率について、96の研究をまとめた系統的レビューがあります。遵守率の中央値は40%でした。
職種別では医師32%、看護師48%。タイミング別では、患者に触れる前が21%、触れた後が47%。ICUでは30〜40%と、他の場所の50〜60%より低い結果でした。
注目していただきたいのは「触れる前」が21%という点です。患者さんを守るためのタイミングが、最も守られていないのです。また、業務が立て込む場面ほど遵守率が低いことも報告されています。
（出典：Erasmus V, et al. Infect Control Hosp Epidemiol 2010;31:283-94）`,
    });
    circle(s, 0.95, 2.15, 3.7, C.MINT, { lineColor: C.TEAL, lineWidth: 6 });
    T(s, "40%", { x: 0.95, y: 3.05, w: 3.7, h: 1.2, fontSize: 80, bold: true, color: C.ORANGE, align: "center", valign: "middle" });
    T(s, "遵守率の中央値", { x: 0.95, y: 4.3, w: 3.7, h: 0.5, fontSize: 22, bold: true, color: C.NAVY, align: "center" });
    T(s, "96研究の系統的レビュー", { x: MX, y: 6.0, w: 4.4, h: 0.45, fontSize: 20, color: C.MUTED, align: "center" });
    const rows = [
      ["職種", ["医師", "32%"], ["看護師", "48%"]],
      ["タイミング", ["患者に触れる前", "21%", true], ["触れた後", "47%"]],
      ["場所", ["ICU", "30〜40%"], ["その他", "50〜60%"]],
    ];
    const x0 = 5.4, rw = W - MX - x0;
    rows.forEach(([lab, a, b], i) => {
      const y = 2.15 + i * 1.45;
      box(s, x0, y, rw, 1.25, C.CARD);
      T(s, lab, { x: x0 + 0.25, y, w: 1.9, h: 1.25, fontSize: 22, bold: true, color: C.NAVY, valign: "middle" });
      [a, b].forEach(([nm, v, hot], j) => {
        const gx = x0 + 2.25 + j * 2.55;
        T(s, nm, { x: gx, y: y + 0.12, w: 2.5, h: 0.42, fontSize: 20, color: C.MUTED });
        T(s, v, { x: gx, y: y + 0.52, w: 2.5, h: 0.62, fontSize: 32, bold: true, color: hot ? C.ORANGE : C.INK });
      });
    });
  }

  // 9 行動が変われば感染は減る
  {
    const s = lightSlide({
      pillText: "第1部", titleText: "行動が変われば、感染は減る",
      src: "Pittet D, et al. Lancet 2000",
      notes:
`【目安1分30秒／9:20–10:50】
では、行動が変わると何が起きるのか。ジュネーブ大学病院の有名な研究です。ベッドサイドでのアルコール手指消毒を柱に病院全体で取り組んだ結果、遵守率は48%から66%へ上がり、院内感染の有病率は16.9%から9.9%へ、MRSAの伝播は1万患者日あたり2.16件から0.93件へ減少しました。同時に、アルコール手指消毒剤の使用量は1000患者日あたり3.5Lから15.4Lに増えています。
行動が変われば、感染は減る。これが今日の出発点です。
（出典：Pittet D, et al. Lancet 2000;356:1307-12）`,
    });
    const items = [
      ["手指衛生の遵守率", "48%→66%", ""],
      ["院内感染の有病率", "16.9%→9.9%", ""],
      ["MRSAの伝播", "2.16→0.93", "（件／1万患者日）"],
    ];
    const w = 3.8, g = (CW - 3 * w) / 2;
    items.forEach(([lab, v, u], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.2, w, 3.0, C.CARD);
      T(s, lab, { x, y: 2.5, w, h: 0.5, fontSize: 24, bold: true, color: C.NAVY, align: "center" });
      T(s, v, { x, y: 3.25, w, h: 0.9, fontSize: 36, bold: true, color: C.TEAL, align: "center", valign: "middle" });
      if (u) T(s, u, { x, y: 4.3, w, h: 0.45, fontSize: 20, color: C.MUTED, align: "center" });
    });
    T(s, "ベッドサイドのアルコール手指消毒を柱に、病院全体で取り組んだ結果\n（ジュネーブ大学病院）",
      { x: MX, y: 5.45, w: CW, h: 1.0, fontSize: 22, color: C.INK, align: "center", valign: "middle" });
  }

  // 10 意図と行動のギャップ
  {
    const s = lightSlide({
      pillText: "第1部", titleText: "「やる気」を大きく変えても、行動の変化は約半分",
      src: "Webb TL, Sheeran P. Psychol Bull 2006",
      notes:
`【目安1分30秒／10:50–12:20】
では、行動を変えるには知識ややる気を高めればよいのでしょうか。
47の実験研究をまとめたメタ分析では、意図、つまり「やろう」という気持ちを中〜大きく変えても、実際の行動の変化は小〜中程度にとどまりました。効果量でいうと0.66に対して0.36、おおよそ半分です。
知識や意識を高める研修は大切です。しかし、それだけでは行動は変わりきらない。行動そのものと、それを取り巻く環境に目を向ける必要があります。
（出典：Webb TL, Sheeran P. Psychol Bull 2006;132:249-68）`,
    });
    caption(s, "効果量（d）　※47の実験研究のメタ分析");
    barChart(s, {
      x: MX, y: 2.35, w: 6.9, h: 3.6, dir: "bar",
      labels: ["行動の変化", "意図の変化"], values: [0.36, 0.66],
      colors: [C.ORANGE, C.TEAL], fmt: "0.00", max: 0.85, valuePt: 28, gap: 45,
    });
    card(s, 7.9, 2.2, W - MX - 7.9, 4.2, C.MINT, [
      { text: "知識・意図・態度に\n働きかけるだけでは、\n行動は変わりきらない", options: { bold: true, color: C.NAVY } },
      { text: "→ 行動そのものと、\nそれを取り巻く環境に\n目を向ける", options: { bold: true, color: C.TEAL } },
    ], { pt: 24, gap: 24, valign: "middle" });
  }

  // 11 人ではなく見えない力
  {
    const s = lightSlide({
      pillText: "第1部", titleText: "問題は「人」ではなく、行動を取り巻く「見えない力」",
      notes:
`【目安1分／12:20–13:20】
ここで見方を変えてみましょう。これまで私たちは、守られない理由を「意識が低い」「やる気がない」と、人の問題として捉えがちでした。
行動科学では、人の周りに働いている「見えない力」に目を向けます。思い込み、疲れ、周囲の空気、動線。これらは本人の努力だけでは変えにくいものです。
人を責める前に、仕組みを疑う。これが行動科学の基本姿勢です。`,
    });
    const cols = [
      [MX, C.PEACH, C.ORANGE, "FaTimes", "これまでの見方", ["意識が低い", "やる気がない", "何度言っても守らない"]],
      [7.33, C.MINT, C.TEAL, "FaCheck", "行動科学の見方", ["思い込みが働いている", "疲れで判断がすり減る", "周りの空気に流される", "動線が行動を邪魔する"]],
    ];
    cols.forEach(([x, fill, col, ic, h, items]) => {
      box(s, x, 2.15, 5.4, 4.35, fill);
      iconCircle(s, x + 0.3, 2.4, 0.75, col, A.icon[ic].white);
      T(s, h, { x: x + 1.25, y: 2.4, w: 3.9, h: 0.75, fontSize: 26, bold: true, color: col, valign: "middle" });
      T(s, paras(items, { bullet: true, paraSpaceAfter: 8 }), { x: x + 0.4, y: 3.45, w: 4.8, h: 2.8, fontSize: 24 });
    });
    chevron(s, 6.42, 3.95, 0.5, 0.8);
  }

  // 12 第2部：5つの敵
  {
    const s = darkSlide(
`【目安40秒／13:20–14:00】
第2部では、5つの「見えない敵」を順番に可視化していきます。①思い込み、②見られている効果、③疲れ、④周囲の空気、⑤環境と動線です。`);
    T(s, "第2部", { x: 0.9, y: 0.9, w: 6, h: 0.5, fontSize: 24, bold: true, color: C.GLOW });
    T(s, "5つの「見えない敵」", { x: 0.9, y: 1.45, w: 11.5, h: 0.9, fontSize: 40, bold: true, color: C.WHITE });
    const items = [["FaUserCheck", "思い込み"], ["FaEye", "見られている効果"], ["FaBatteryQuarter", "疲れ"], ["FaUsers", "周囲の空気"], ["FaRoute", "環境と動線"]];
    const slot = CW / 5, d = 1.7;
    items.forEach(([ic, name], i) => {
      const x = MX + i * slot + (slot - d) / 2;
      circle(s, x, 2.95, d, C.NAVY2, { lineColor: C.GLOW, lineWidth: 3 });
      s.addImage({ data: A.icon[ic].glow, x: x + d * 0.27, y: 2.95 + d * 0.27, w: d * 0.46, h: d * 0.46 });
      T(s, "敵" + "①②③④⑤"[i], { x: MX + i * slot, y: 4.85, w: slot, h: 0.5, fontSize: 22, bold: true, color: C.GLOW, align: "center" });
      T(s, name, { x: MX + i * slot + 0.1, y: 5.35, w: slot - 0.2, h: 1.0, fontSize: 24, bold: true, color: C.WHITE, align: "center" });
    });
  }

  const ENEMY = C.ORANGE;

  // 13 敵① 自己評価
  {
    const s = lightSlide({
      pillText: "見えない敵① 思い込み", pillColor: ENEMY,
      titleText: "「できている」という自己評価は、実際の行動を映さない",
      src: "Jenner EA, et al. J Hosp Infect 2006",
      notes:
`【目安1分30秒／14:00–15:30】
一つ目の敵は「思い込み」です。
71名の医療者を合計132時間観察し、1,284回の手指衛生の機会を調べた研究では、観察された遵守率は非常に低く、しかも実際の行動は、本人の「やるつもり」や「やっている」という自己申告とは関連しませんでした。MRSA保菌者のケアでさえ、リスクに応じた行動がとれていなかったと報告されています。
つまり、アンケートや自己評価では、行動の実態は見えないのです。意識調査で「できている」と答えた部署ほど、実際の観察データを返すことが大切になります。
（出典：Jenner EA, et al. J Hosp Infect 2006;63:418-22）`,
    });
    const cw = 4.9;
    [[MX, "FaCommentDots", C.NAVY, "本人の意図・自己申告", "「やっている」"], [W - MX - cw, "FaEye", C.ORANGE, "実際に観察された行動", "遵守率は非常に低い"]].forEach(([x, ic, col, h, sub]) => {
      box(s, x, 2.15, cw, 2.35, C.CARD);
      iconCircle(s, x + cw / 2 - 0.45, 2.35, 0.9, col, A.icon[ic].white);
      T(s, h, { x, y: 3.35, w: cw, h: 0.5, fontSize: 24, bold: true, color: C.NAVY, align: "center" });
      T(s, sub, { x, y: 3.85, w: cw, h: 0.45, fontSize: 22, color: C.MUTED, align: "center" });
    });
    s.addImage({ data: A.icon.FaNotEqual.orange, x: W / 2 - 0.55, y: 2.75, w: 1.1, h: 1.1 });
    card(s, MX, 4.75, CW, 1.8, C.MINT, [
      { text: "71名・1,284回の手指衛生機会を観察", options: { bold: true, color: C.NAVY } },
      { text: "→ 実際の行動は、本人の意図や自己申告と関連しなかった", options: { color: C.INK } },
      { text: "→ アンケートや意識調査だけでは、行動は見えない", options: { bold: true, color: C.TEAL } },
    ], { pt: 22, gap: 4, valign: "middle" });
  }

  // 14 敵① 過信と主語
  {
    const s = lightSlide({
      pillText: "見えない敵① 思い込み", pillColor: ENEMY,
      titleText: "「自分は大丈夫」という過信が、メッセージを無力にする",
      src: "Grant AM, Hofmann DA. Psychol Sci 2011",
      notes:
`【目安2分／15:30–17:30】
思い込みのもう一つの形が、「自分は大丈夫」という過信です。
病院で行われた2つのフィールド実験で、手指消毒剤の近くに2種類の掲示をしました。「手指衛生で、あなたが病気になるのを防げます」と、「手指衛生で、患者さんが病気になるのを防げます」。違いは主語の1語だけです。
結果、「あなた」の掲示では変化がなく、「患者さん」の掲示では、消毒剤の使用量も、覆面での観察による手指衛生も有意に増えました。医療者は日頃から病原体に触れているため「自分は感染しない」と過信しやすく、自分を守るメッセージは響きにくいのです。
皆さんの施設のポスターの主語は、誰になっていますか。
（出典：Grant AM, Hofmann DA. Psychol Sci 2011;22:1494-9）`,
    });
    const sw = 5.8;
    [[MX, "掲示A", "あなた", C.ORANGE, "→ 変化なし", C.MUTED, C.CARD], [W - MX - sw, "掲示B", "患者さん", C.TEAL, "→ 手指衛生が有意に増加", C.TEAL, C.MINT]].forEach(([x, lab, who, col, res, rcol, rfill]) => {
      box(s, x, 2.15, sw, 2.35, C.WHITE, { lineColor: C.SLATE, lineWidth: 2 });
      T(s, lab, { x: x + 0.25, y: 2.25, w: 2, h: 0.42, fontSize: 20, color: C.MUTED });
      T(s, [
        { text: "手指衛生で、", options: { breakLine: true } },
        { text: who, options: { color: col, bold: true, underline: { style: "sng" } } },
        { text: "が病気になるのを防げます" },
      ], { x: x + 0.2, y: 2.7, w: sw - 0.4, h: 1.6, fontSize: 24, align: "center", valign: "middle", color: C.INK });
      box(s, x, 4.7, sw, 0.8, rfill);
      T(s, res, { x, y: 4.7, w: sw, h: 0.8, fontSize: 26, bold: true, color: rcol, align: "center", valign: "middle" });
    });
    T(s, "違いは主語の1語だけ（病院での2つのフィールド実験）", { x: MX, y: 5.8, w: CW, h: 0.6, fontSize: 22, color: C.INK, align: "center", valign: "middle" });
  }

  // 15 敵② ホーソン効果
  {
    const s = lightSlide({
      pillText: "見えない敵② 見られている効果", pillColor: ENEMY,
      titleText: "「見られている」と、手指衛生は約3倍になる",
      src: "Srigley JA, et al. BMJ Qual Saf 2014",
      notes:
`【目安1分30秒／17:30–19:00】
二つ目の敵は「見られている効果」、いわゆるホーソン効果です。
カナダの大学病院で、電子システムを使って8か月間、全ての手指衛生を記録しました。すると、監査者が見える廊下の消毒剤では1時間あたり3.75回、同じ時間の監査者が見えない場所では1.48回、同じ場所の前の週は1.07回でした。監査者が見えると約3倍になり、しかも病室内では変化がありませんでした。
私たちが監査で測っている遵守率は、実際より高く見えている可能性があるということです。
（出典：Srigley JA, et al. BMJ Qual Saf 2014;23:974-80）`,
    });
    caption(s, "監査者の見え方別の手指衛生回数（回／台／時）", MX, 1.95, 6.4);
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["見える場所", "見えない場所", "前の週"], values: [3.75, 1.48, 1.07],
      colors: [C.ORANGE, C.SLATE, C.SLATE], fmt: "0.00", max: 4.4,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "電子システムで8か月間、\n全ての手指衛生を記録", options: { bold: true, color: C.NAVY } },
      "監査者が見える廊下でだけ急増\n病室内では変化なし",
      { text: "→ 監査の遵守率は、実際より\n高く見えている可能性", options: { bold: true, color: C.TEAL } },
    ], { gap: 16 });
  }

  // 16 敵② 物差しを組み合わせる
  {
    const s = lightSlide({
      pillText: "見えない敵② 見られている効果", pillColor: ENEMY,
      titleText: "物差しを組み合わせて、行動の「素顔」を見る",
      src: "Srigley JA, et al. 2014; Pittet D, et al. Ann Intern Med 2004",
      notes:
`【目安1分30秒／19:00–20:30】
では、どう測ればよいのでしょうか。答えは、物差しを組み合わせることです。
直接観察はタイミングや手技がわかりますが、見られている効果が出ます。手指消毒剤の使用量は、全体の傾向を客観的に追えますが、どのタイミングで使ったかはわかりません。電子モニタリングは24時間の実態に迫れますが、導入にコストがかかります。
医師の手指衛生を調べた研究でも、「観察されている」という意識が遵守と関連していました。一つの数字を信じ込まず、複数の物差しで行動の「素顔」に近づきましょう。
（出典：Srigley JA, et al. BMJ Qual Saf 2014；Pittet D, et al. Ann Intern Med 2004;141:1-8）`,
    });
    const items = [
      ["FaEye", "直接観察", "タイミング・手技", "見られている効果"],
      ["FaPumpSoap", "使用量", "全体の傾向", "タイミングは不明"],
      ["FaNetworkWired", "電子モニタリング", "24時間の実態", "導入のコスト"],
    ];
    const w = 3.8, g = (CW - 3 * w) / 2;
    items.forEach(([ic, h, good, weak], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.15, w, 3.55, C.CARD);
      iconCircle(s, x + w / 2 - 0.45, 2.35, 0.9, C.NAVY, A.icon[ic].white);
      T(s, h, { x, y: 3.35, w, h: 0.5, fontSize: 24, bold: true, color: C.NAVY, align: "center" });
      T(s, [
        { text: "わかる：", options: { bold: true, color: C.TEAL } },
        { text: good, options: { breakLine: true } },
        { text: "弱点：", options: { bold: true, color: C.ORANGE } },
        { text: weak },
      ], { x: x + 0.2, y: 4.0, w: w - 0.4, h: 1.55, fontSize: 20, paraSpaceAfter: 6 });
    });
    T(s, "1つの数字に頼らず、複数の物差しで「素顔」に近づく", { x: MX, y: 5.9, w: CW, h: 0.7, fontSize: 24, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // 17 敵③ 疲れ（勤務時間）
  {
    const s = lightSlide({
      pillText: "見えない敵③ 疲れ", pillColor: ENEMY,
      titleText: "手指衛生は、勤務時間とともに「すり減る」",
      src: "Dai H, et al. J Appl Psychol 2015",
      notes:
`【目安1分30秒／20:30–22:00】
三つ目の敵は「疲れ」です。
米国35病院、4,157人の医療者、約1,370万回の手指衛生機会を分析した研究では、12時間勤務の始めから終わりまでに、遵守率が平均8.7ポイント低下しました。業務が立て込むほど低下は大きく、逆に勤務と勤務の間の休息が長いと、次の勤務で遵守率が回復していました。
ルールを守る力は、時間とともにすり減る「資源」なのです。
（出典：Dai H, et al. J Appl Psychol 2015;100:846-62）`,
    });
    circle(s, 1.0, 2.1, 3.5, C.PEACH, { lineColor: C.ORANGE, lineWidth: 6 });
    T(s, "−8.7", { x: 1.0, y: 2.75, w: 3.5, h: 1.1, fontSize: 72, bold: true, color: C.ORANGE, align: "center", valign: "middle" });
    T(s, "ポイント", { x: 1.0, y: 4.0, w: 3.5, h: 0.5, fontSize: 24, bold: true, color: C.ORANGE, align: "center" });
    T(s, "12時間勤務の最初と最後の差", { x: MX, y: 5.8, w: 4.3, h: 0.5, fontSize: 20, color: C.MUTED, align: "center" });
    const rows = [
      ["FaHospital", C.NAVY, "35病院・4,157人を対象に\n約1,370万回の機会を分析"],
      ["FaBolt", C.ORANGE, "業務が立て込むほど、低下は大きくなる"],
      ["FaCoffee", C.TEAL, "休息が長いと、次の勤務で回復する"],
    ];
    rows.forEach(([ic, col, t], i) => {
      const y = 2.2 + i * 1.4;
      box(s, 5.3, y, W - MX - 5.3, 1.2, C.CARD);
      iconCircle(s, 5.5, y + 0.18, 0.84, col, A.icon[ic].white);
      T(s, t, { x: 6.55, y, w: W - MX - 6.75, h: 1.2, fontSize: 22, color: C.INK, valign: "middle" });
    });
  }

  // 18 敵③ 処方の時間帯
  {
    const s = lightSlide({
      pillText: "見えない敵③ 疲れ", pillColor: ENEMY,
      titleText: "医師の処方判断も、外来の後半ほどゆらぐ",
      src: "Linder JA, et al. JAMA Intern Med 2014",
      notes:
`【目安1分30秒／22:00–23:30】
疲れの影響は、医師の判断にも表れます。
急性気道感染症による約2万2千件の受診を分析した研究では、4時間の外来枠の1時間目に比べ、3時間目では抗菌薬処方のオッズが1.14倍、4時間目では1.26倍に上がっていました。判断を重ねるほど、「とりあえず出しておく」という楽な選択に流れやすくなる、いわゆる意思決定疲れと考えられています。
これは個人の資質の問題ではなく、「時間帯」という見えない要因です。
（出典：Linder JA, et al. JAMA Intern Med 2014;174:2029-31）`,
    });
    caption(s, "抗菌薬を処方するオッズ比（1時間目＝1.00）");
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["1時間目", "2時間目", "3時間目", "4時間目"], values: [1.0, 1.01, 1.14, 1.26],
      colors: [C.SLATE, C.SLATE, "E2885F", C.ORANGE], fmt: "0.00", max: 1.5,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "急性気道感染症の\n21,867受診・医師204名", options: { bold: true, color: C.NAVY } },
      "判断を重ねるほど\n「楽な選択」に流れやすい\n（意思決定疲れ）",
      { text: "→ 原因は資質ではなく\n「時間帯」", options: { bold: true, color: C.TEAL } },
    ], { gap: 16 });
  }

  // 19 敵③ 仕組みで守る
  {
    const s = lightSlide({
      pillText: "見えない敵③ 疲れ", pillColor: ENEMY,
      titleText: "忙しい時ほど守れない。だから「仕組み」で守る",
      src: "Pittet D, et al. Ann Intern Med 1999",
      notes:
`【目安1分30秒／23:30–25:00】
ジュネーブの研究でも、ケアが立て込み、1時間あたりの手指衛生の機会が41回を超えるような場面では、20回以下の場面に比べて、手指衛生をしないオッズが約2倍になっていました。
忙しさを「気合い」で乗り切ることはできません。手の届く場所に消毒剤を置いて探す手間をなくす、忙しい時間帯や勤務の後半に合図を出す、休憩を確保する、判断を助けるオーダーセットを用意する。意志ではなく、仕組みで守りましょう。
（出典：Pittet D, et al. Ann Intern Med 1999;130:126-30）`,
    });
    box(s, MX, 2.15, 4.6, 4.4, C.PEACH);
    T(s, "ケアが立て込むと", { x: MX, y: 2.4, w: 4.6, h: 0.5, fontSize: 22, bold: true, color: C.NAVY, align: "center" });
    T(s, "約2倍", { x: MX, y: 3.0, w: 4.6, h: 1.2, fontSize: 66, bold: true, color: C.ORANGE, align: "center", valign: "middle" });
    T(s, "手指衛生をしないオッズ", { x: MX, y: 4.3, w: 4.6, h: 0.5, fontSize: 22, bold: true, color: C.NAVY, align: "center" });
    T(s, "1時間に41回以上の機会がある場面（20回以下との比較）", { x: MX + 0.3, y: 4.95, w: 4.0, h: 1.2, fontSize: 20, color: C.MUTED, align: "center" });
    const x0 = 5.6, rw = W - MX - x0;
    T(s, "意志ではなく、仕組みで守る", { x: x0, y: 2.15, w: rw, h: 0.55, fontSize: 24, bold: true, color: C.TEAL });
    const rows = [
      ["FaPumpSoap", "手の届く場所に消毒剤（探す手間をなくす）"],
      ["FaBell", "繁忙時・勤務の後半に合図を出す"],
      ["FaCoffee", "休憩を確保する"],
      ["FaClipboardList", "判断を助けるオーダーセットを用意する"],
    ];
    rows.forEach(([ic, t], i) => {
      const y = 2.9 + i * 0.93;
      iconCircle(s, x0, y, 0.75, C.TEAL, A.icon[ic].white);
      T(s, t, { x: x0 + 0.95, y, w: rw - 0.95, h: 0.75, fontSize: 22, color: C.INK, valign: "middle" });
    });
  }

  // 20 敵④ 周りの人
  {
    const s = lightSlide({
      pillText: "見えない敵④ 周囲の空気", pillColor: ENEMY,
      titleText: "人は「ルール」より「周りの人」に従う",
      src: "Lankford MG, et al. Emerg Infect Dis 2003; Monsalve MN, et al. ICHE 2014",
      notes:
`【目安1分30秒／25:00–26:30】
四つ目の敵は「周囲の空気」です。
ある研究では、同じ部屋にいる上級の医療者や同僚が手指衛生をしないと、自分も手指衛生をしにくくなっていました。オッズ比は0.2、つまり手指衛生をする見込みが大きく下がるということです。
逆に、ICUでセンサーを使って約4万8千回の機会を調べた研究では、一人のときの遵守率は20.9%、周りに他の職員がいるときは27.9%に上がっていました。
私たちはルールよりも、周りの人の行動に強く影響されているのです。
（出典：Lankford MG, et al. Emerg Infect Dis 2003;9:217-23／Monsalve MN, et al. Infect Control Hosp Epidemiol 2014;35:1277-85）`,
    });
    const w = (CW - 0.33) / 2;
    [
      [MX, C.PEACH, C.ORANGE, ["上級者や同僚が", "手指衛生をしないと"], "オッズ比 0.2", "自分も手指衛生をしにくくなる"],
      [MX + w + 0.33, C.MINT, C.TEAL, ["周りに他の職員が", "いると"], "20.9%→27.9%", "遵守率が上がる\n（ICU・約4.8万回の機会）"],
    ].forEach(([x, fill, col, head, big, sub]) => {
      box(s, x, 2.15, w, 4.35, fill);
      iconCircle(s, x + 0.35, 2.4, 0.95, col, A.icon.FaUsers.white);
      T(s, paras(head), { x: x + 1.5, y: 2.35, w: w - 1.7, h: 1.05, fontSize: 24, bold: true, color: C.NAVY, valign: "middle" });
      T(s, big, { x, y: 3.65, w, h: 1.1, fontSize: 44, bold: true, color: col, align: "center", valign: "middle" });
      T(s, sub, { x: x + 0.3, y: 4.95, w: w - 0.6, h: 1.2, fontSize: 22, color: C.INK, align: "center" });
    });
  }

  // 21 敵④ リーダー
  {
    const s = lightSlide({
      pillText: "見えない敵④ 周囲の空気", pillColor: ENEMY,
      titleText: "リーダーの最初の一手が、部署の「空気」をつくる",
      src: "Pittet D, et al. Ann Intern Med 2004",
      notes:
`【目安1分30秒／26:30–28:00】
この「空気」をつくるのはリーダーです。医師163名を観察した研究では、遵守率は平均57%で、「自分は同僚の手本になっている」という意識が遵守と関連していました。観察されているという意識や、手指消毒剤へのアクセスのしやすさも関連していました。
明日からの一手はシンプルです。回診やラウンドの先頭に立つ人が、最初に、みんなに見えるように手指衛生をする。新人や研修医は、上の人の背中を見て学んでいます。
（出典：Pittet D, et al. Ann Intern Med 2004;141:1-8）`,
    });
    box(s, MX, 2.15, 6.4, 4.35, C.CARD);
    T(s, "医師163名の観察（遵守率 平均57%）", { x: MX + 0.3, y: 2.4, w: 5.9, h: 0.5, fontSize: 22, bold: true, color: C.NAVY });
    T(s, "遵守と関連していた要因", { x: MX + 0.3, y: 3.0, w: 5.9, h: 0.45, fontSize: 22, color: C.MUTED });
    T(s, paras(["「自分は同僚の手本だ」という意識", "観察されているという意識", "手指消毒剤へのアクセスのしやすさ"], { bullet: true, paraSpaceAfter: 14 }),
      { x: MX + 0.3, y: 3.6, w: 5.9, h: 2.6, fontSize: 24 });
    const x0 = 7.3, rw = W - MX - x0;
    box(s, x0, 2.15, rw, 4.35, C.NAVY);
    T(s, "明日からの一手", { x: x0 + 0.35, y: 2.4, w: rw - 0.7, h: 0.5, fontSize: 22, bold: true, color: C.GLOW });
    T(s, "回診・ラウンドの\n先頭の人が、最初に、\n見えるように手指衛生", { x: x0 + 0.35, y: 3.05, w: rw - 0.7, h: 1.9, fontSize: 26, bold: true, color: C.WHITE });
    T(s, "新人・研修医は\n「上の人」を見て学んでいる", { x: x0 + 0.35, y: 5.1, w: rw - 0.7, h: 1.1, fontSize: 22, color: C.PALE });
  }

  // 22 敵⑤ 置き場所
  {
    const s = lightSlide({
      pillText: "見えない敵⑤ 環境と動線", pillColor: ENEMY,
      titleText: "行動は、設備の「数」より「置き場所」で決まる",
      src: "Lankford MG 2003; Pittet D 2000; Gould DJ (Cochrane) 2017",
      notes:
`【目安1分30秒／28:00–29:30】
五つ目の敵は「環境と動線」です。
先ほどの研究では、手洗いシンクを増やした新しい病院に移ったところ、遵守率は53%から23%に下がっていました。設備を増やすだけでは行動は変わらなかったのです。
一方、ジュネーブの研究では、ベッドサイドにアルコール手指消毒剤を置いたことが遵守率向上に大きく寄与しました。コクランレビューでも、使用する場所の近くに置くことで遵守率が改善する可能性が、中等度の確実性で示されています。
「数」より「その場にあるか」。人は、そこにあるから手を伸ばします。
（出典：Lankford MG, et al. 2003／Pittet D, et al. 2000／Gould DJ, et al. Cochrane Database Syst Rev 2017;9:CD005186）`,
    });
    const w = (CW - 0.33) / 2;
    const xL = MX, xR = MX + w + 0.33;
    box(s, xL, 2.15, w, 3.55, C.PEACH);
    iconCircle(s, xL + 0.3, 2.35, 0.85, C.ORANGE, A.icon.FaBuilding.white);
    T(s, "手洗いシンクを増やした新病院", { x: xL + 1.35, y: 2.35, w: w - 1.55, h: 0.85, fontSize: 22, bold: true, color: C.NAVY, valign: "middle" });
    T(s, "53% → 23%", { x: xL, y: 3.35, w, h: 0.95, fontSize: 44, bold: true, color: C.ORANGE, align: "center", valign: "middle" });
    T(s, "旧病院→新病院の遵守率\nシンクを増やすだけでは改善しなかった", { x: xL + 0.3, y: 4.4, w: w - 0.6, h: 1.1, fontSize: 20, color: C.INK });
    box(s, xR, 2.15, w, 3.55, C.MINT);
    iconCircle(s, xR + 0.3, 2.35, 0.85, C.TEAL, A.icon.FaPumpSoap.white);
    T(s, "「使う場所」に置く", { x: xR + 1.35, y: 2.35, w: w - 1.55, h: 0.85, fontSize: 22, bold: true, color: C.NAVY, valign: "middle" });
    T(s, paras(expand([
      "ベッドサイドの手指消毒剤が、\n遵守率の向上に大きく寄与",
      "使用場所の近くに置くと遵守率が改善\n（中等度の確実性）",
    ], 14)), { x: xR + 0.3, y: 3.4, w: w - 0.6, h: 2.1, fontSize: 22 });
    T(s, "「そこにある」から、手が伸びる", { x: MX, y: 5.9, w: CW, h: 0.7, fontSize: 26, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // 23 第3部
  section("第3部", "見えない敵を\n「見える化」する",
`【目安20秒／29:30–29:50】
第3部では、これらの見えない敵を「見える化」する方法を、診断する・見せる・返す、の3つの視点でご紹介します。`,
    [["FaStethoscope", "診断する：COM-Bで原因を探す"], ["FaSearchPlus", "見せる：汚染を見える形に"], ["FaChartBar", "返す：数字で自分の位置を示す"]], "FaSearchPlus");

  const VIS = C.TEAL;

  // 24 COM-B
  {
    const s = lightSlide({
      pillText: "第3部 診断する", titleText: "行動は「能力」「機会」「動機」がそろって起こる",
      src: "Michie S, et al. Implement Sci 2011",
      notes:
`【目安1分40秒／29:50–31:30】
まず「診断する」ための道具が、COM-B（コム・ビー）モデルです。英国のMichieらが、既存の19の行動変容の枠組みを整理して提唱しました。
行動が起こるには、能力（Capability）、機会（Opportunity）、動機（Motivation）の3つがそろう必要があります。能力には身体的・心理的、機会には物理的・社会的、動機には内省的・自動的という、それぞれ2つの側面があります。
「なぜできないのか」を、この3つの箱に分けて考えるだけで、原因が見えやすくなります。
（出典：Michie S, et al. Implement Sci 2011;6:42）`,
    });
    const boxes = [
      [C.MINT, C.TEAL, "能力（C）", "Capability", ["身体的：技術・体力", "心理的：知識・記憶・注意"]],
      [C.CARD, C.NAVY, "機会（O）", "Opportunity", ["物理的：物・場所・時間", "社会的：周囲の規範・文化"]],
      [C.PEACH, C.ORANGE, "動機（M）", "Motivation", ["内省的：信念・目標", "自動的：習慣・感情"]],
    ];
    const w = 3.8, g = (CW - 3 * w) / 2;
    boxes.forEach(([fill, col, h, en, lines], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.05, w, 2.45, fill);
      T(s, [{ text: h, options: { fontSize: 26, bold: true, color: col, breakLine: true } }, { text: en, options: { fontSize: 20, color: C.MUTED } }],
        { x: x + 0.2, y: 2.15, w: w - 0.4, h: 0.9 });
      T(s, paras(lines, { paraSpaceAfter: 6 }), { x: x + 0.2, y: 3.15, w: w - 0.4, h: 1.25, fontSize: 20 });
      const cx = x + w / 2;
      arrow(s, cx, 4.55, W / 2 + (i - 1) * 1.6, 5.2, C.SLATE, 3);
    });
    box(s, W / 2 - 2.6, 5.25, 5.2, 1.1, C.NAVY);
    T(s, "行動（Behaviour）", { x: W / 2 - 2.6, y: 5.25, w: 5.2, h: 1.1, fontSize: 28, bold: true, color: C.WHITE, align: "center", valign: "middle" });
  }

  // 25 5つの敵 × COM-B
  {
    const s = lightSlide({
      pillText: "第3部 診断する", titleText: "5つの敵をCOM-Bに当てはめると、打ち手が見える",
      notes:
`【目安1分30秒／31:30–33:00】
今日の5つの敵をCOM-Bに当てはめてみます。
思い込みは内省的な動機、疲れは心理的な能力、周囲の空気は社会的な機会、環境と動線は物理的な機会の問題です。見られている効果は、行動そのものではなく、私たちの「測定」をゆがめる敵です。
箱が決まれば、打ち手も決まります。動機の問題に消毒剤の置き場所を変えても効きにくく、機会の問題に研修を重ねても効きにくい。原因の箱と打ち手の方向を合わせることが大切です。`,
    });
    const H = { bold: true, color: C.WHITE, fill: { color: C.NAVY } };
    const rows = [
      [{ text: "見えない敵", options: H }, { text: "COM-Bの箱", options: H }, { text: "打ち手の方向性", options: H }],
      ["① 思い込み", "動機（内省的）", "「患者さんを守る」を主語に"],
      ["② 見られている効果", "（測定をゆがめる）", "物差しを組み合わせる"],
      ["③ 疲れ", "能力（心理的）", "仕組み・休息・判断支援"],
      ["④ 周囲の空気", "機会（社会的）", "リーダーの率先・同僚比較"],
      ["⑤ 環境と動線", "機会（物理的）", "その場に置く・動線に組み込む"],
    ].map((r, i) => i === 0 ? r : r.map((t, j) => ({ text: t, options: { bold: j === 0, color: j === 0 ? C.ORANGE : C.INK, fill: { color: i % 2 ? C.WHITE : C.CARD } } })));
    s.addTable(rows, {
      x: MX, y: 2.1, w: CW, colW: [3.5, 3.3, CW - 6.8], rowH: 0.66,
      fontFace: FONT, fontSize: 22, valign: "middle", margin: [0.05, 0.15, 0.05, 0.15],
      border: { type: "solid", pt: 1, color: C.LINE },
    });
    T(s, "原因の「箱」と、打ち手の「方向」を合わせる", { x: MX, y: 6.15, w: CW, h: 0.55, fontSize: 24, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // 26 見せる：環境
  {
    const s = lightSlide({
      pillText: "見える化① 汚染を見せる", pillColor: VIS,
      titleText: "環境に触れた手は、患者に触れた手と同じくらい汚れる",
      src: "Stiefel U, et al. Infect Control Hosp Epidemiol 2011",
      notes:
`【目安1分30秒／33:00–34:30】
ここから「見せる」です。そもそも汚染は目に見えません。
MRSA保菌者40名の病室で調べた研究では、患者さんの皮膚に触れた後に手にMRSAが付着した割合は40%、ベッド柵やテーブルなど環境表面に触れた後は45%で、ほぼ同じでした。
患者さんに直接触れない職種、たとえばリハビリ、放射線、検査、清掃のスタッフも、環境に触れたら手指衛生が必要です。WHOの「手指衛生の5つのタイミング」の5番目、患者周囲の環境に触れた後ですね。
（出典：Stiefel U, et al. Infect Control Hosp Epidemiol 2011;32:185-7）`,
    });
    caption(s, "手にMRSAが付着した割合", MX, 1.95, 6.3);
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["患者の皮膚", "環境表面"], values: [40, 45],
      colors: [C.SLATE, C.ORANGE], fmt: '0"%"', max: 56, valuePt: 28,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "MRSA保菌者40名の病室で検証", options: { bold: true, color: C.NAVY } },
      "ベッド柵・テーブル・機器に\n触れるだけで、手は汚れる",
      { text: "→ 患者に触れない職種も、\n環境に触れたら手指衛生\n（リハビリ・放射線・検査・清掃）", options: { bold: true, color: C.TEAL } },
    ], { gap: 14 });
  }

  // 27 見せる：PPE
  {
    const s = lightSlide({
      pillText: "見える化① 汚染を見せる", pillColor: VIS,
      titleText: "蛍光で「その場で見せる」と、脱衣時の汚染は激減する",
      src: "Tomas ME, et al. JAMA Intern Med 2015",
      notes:
`【目安1分30秒／34:30–36:00】
見えない汚染を「その場で見せる」と、行動は変わります。
4病院で435回の防護具の脱衣シミュレーションを行ったところ、蛍光ローションによる皮膚や衣服の汚染が46%で起きていました。手袋を外すときの方が、ガウンより多く汚染していました。
そこで、蛍光ローションで汚染をその場で見せながら練習する教育を行うと、汚染は60%から18.9%に減り、1か月後、3か月後も12%に保たれていました。
自分の汚染が見えると、行動は自然と丁寧になります。
（出典：Tomas ME, et al. JAMA Intern Med 2015;175:1904-10）`,
    });
    caption(s, "皮膚・衣服の汚染が起きた割合");
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["介入前", "直後", "1か月後", "3か月後"], values: [60.0, 18.9, 12.0, 12.0],
      colors: [C.ORANGE, C.TEAL, C.TEAL, C.TEAL], fmt: '0.0"%"', max: 72,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "4病院・435回の脱衣\nシミュレーションで46%に汚染", options: { bold: true, color: C.NAVY } },
      "手袋はガウンより汚染が多い\n（52.9% vs 37.8%）",
      { text: "→ 蛍光ローションで汚染を\n見せながら練習", options: { bold: true, color: C.TEAL } },
    ], { gap: 14 });
  }

  // 28 見せる：清掃
  {
    const s = lightSlide({
      pillText: "見える化① 汚染を見せる", pillColor: VIS,
      titleText: "清掃も「見せて返す」と、拭き残しが減る",
      src: "Carling PC, et al. Infect Control Hosp Epidemiol 2008",
      notes:
`【目安1分30秒／36:00–37:30】
清掃も同じです。米国36病院で、約2万か所の高頻度接触面を蛍光マーカーで評価したところ、退院後の清掃で拭かれていた表面は48%にとどまりました。
蛍光マーカーの結果を清掃スタッフに繰り返しフィードバックし、手順を見直した結果、77%まで改善しました。
清掃スタッフが悪いのではなく、「拭き残しが見えない」ことが敵だったのです。
（出典：Carling PC, et al. Infect Control Hosp Epidemiol 2008;29:1035-41）`,
    });
    caption(s, "清拭されていた高頻度接触面の割合", MX, 1.95, 6.3);
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["介入前", "介入後"], values: [48, 77],
      colors: [C.ORANGE, C.TEAL], fmt: '0"%"', max: 92, valuePt: 28,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "米国36病院・\n約2万か所の表面を評価", options: { bold: true, color: C.NAVY } },
      "蛍光マーカーで拭き残しを\n客観的に見える化し、\n清掃スタッフにくり返し\nフィードバック",
      { text: "→ 敵は人ではなく\n「見えないこと」", options: { bold: true, color: C.TEAL } },
    ], { gap: 14 });
  }

  // 29 返す：社会規範フィードバック
  {
    const s = lightSlide({
      pillText: "見える化② 数字で返す", pillColor: VIS,
      titleText: "「あなたは上位20%」という1通の手紙が、処方を減らした",
      src: "Hallsworth M, et al. Lancet 2016",
      notes:
`【目安1分30秒／37:30–39:00】
次は「数字で返す」です。英国で1,581の診療所を対象に行われたランダム化比較試験です。
抗菌薬処方が地域の上位20%に入る診療所の医師に、イングランドの主席医務官から「あなたの診療所の処方は、地域の80%の診療所より多い」という手紙を1通送りました。
その結果、6か月間の処方は3.3%減り、推計で約7万3千件の処方が減りました。低コストで全国規模の効果です。
自分の位置が「見える」と、人は行動を修正します。
（出典：Hallsworth M, et al. Lancet 2016;387:1743-52）`,
    });
    box(s, MX, 2.2, 6.2, 4.3, C.WHITE, { lineColor: C.SLATE, lineWidth: 2 });
    iconCircle(s, MX + 0.3, 2.45, 0.8, C.TEAL, A.icon.FaEnvelopeOpenText.white);
    T(s, "主席医務官からの手紙（要旨）", { x: MX + 1.25, y: 2.45, w: 4.8, h: 0.8, fontSize: 20, color: C.MUTED, valign: "middle" });
    T(s, "「あなたの診療所の\n抗菌薬処方は、地域の80%の\n診療所より多い状況です」", { x: MX + 0.35, y: 3.45, w: 5.5, h: 1.9, fontSize: 24, bold: true, color: C.NAVY, valign: "middle" });
    T(s, "＋ 患者さん向けのリーフレット", { x: MX + 0.35, y: 5.6, w: 5.5, h: 0.5, fontSize: 20, color: C.MUTED });
    const x0 = 7.2, rw = W - MX - x0;
    T(s, "−3.3%", { x: x0, y: 2.15, w: rw, h: 1.05, fontSize: 60, bold: true, color: C.TEAL, align: "center", valign: "middle" });
    T(s, "抗菌薬の処方（6か月間）", { x: x0, y: 3.2, w: rw, h: 0.5, fontSize: 22, color: C.INK, align: "center" });
    T(s, "約7.3万件", { x: x0, y: 3.9, w: rw, h: 0.9, fontSize: 44, bold: true, color: C.TEAL, align: "center", valign: "middle" });
    T(s, "推計の処方減少", { x: x0, y: 4.8, w: rw, h: 0.5, fontSize: 22, color: C.INK, align: "center" });
    T(s, "1,581診療所のランダム化比較試験", { x: x0, y: 5.65, w: rw, h: 0.5, fontSize: 20, color: C.MUTED, align: "center" });
  }

  // 30 返す：同僚比較と理由記載
  {
    const s = lightSlide({
      pillText: "見える化② 数字で返す", pillColor: VIS,
      titleText: "同僚との比較と「理由の記載」が、不適切な処方を減らす",
      src: "Meeker D, et al. JAMA 2016",
      notes:
`【目安1分30秒／39:00–40:30】
米国の47診療所、248名の臨床医を対象にしたランダム化比較試験では、3つの行動科学的介入を比べました。
電子カルテで抗菌薬以外の治療を提案する方法は、対照群と有意な差がありませんでした。一方、抗菌薬を処方するときに理由をカルテに記載させる方法は7.0ポイント、不適切な処方が少ない上位者と比べた結果をメールで返す方法は5.2ポイント、対照群よりも不適切な処方を減らしました。
「見られる」「比べられる」ことを、うまく仕組みに組み込んだ例です。
（出典：Meeker D, et al. JAMA 2016;315:562-70）`,
    });
    const items = [
      [C.CARD, C.MUTED, "代替案の提示", "電子カルテで\n抗菌薬以外の治療を提案", "有意差なし"],
      [C.MINT, C.TEAL, "理由の記載", "処方理由を\nカルテに記載してもらう", "−7.0ポイント"],
      [C.MINT, C.TEAL, "同僚との比較", "処方率を上位者と比べて\nメールで返す", "−5.2ポイント"],
    ];
    const w = 3.8, g = (CW - 3 * w) / 2;
    items.forEach(([fill, col, h, d, v], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.15, w, 3.65, fill);
      T(s, h, { x, y: 2.35, w, h: 0.55, fontSize: 26, bold: true, color: C.NAVY, align: "center" });
      T(s, d, { x: x + 0.3, y: 3.0, w: w - 0.6, h: 1.3, fontSize: 20, color: C.INK, align: "center" });
      T(s, v, { x, y: 4.55, w, h: 0.9, fontSize: 34, bold: true, color: col, align: "center", valign: "middle" });
    });
    T(s, "47診療所・248名の臨床医のランダム化比較試験（対照群との差）", { x: MX, y: 5.95, w: CW, h: 0.6, fontSize: 22, color: C.INK, align: "center", valign: "middle" });
  }

  // 31 返す：宣言ポスター
  {
    const s = lightSlide({
      pillText: "見える化② 数字で返す", pillColor: VIS,
      titleText: "「宣言」を掲示するだけで、処方が変わる",
      src: "Meeker D, et al. JAMA Intern Med 2014",
      notes:
`【目安1分30秒／40:30–42:00】
さらにシンプルな例です。医師の写真と署名入りで「不要な抗菌薬は処方しません」という宣言ポスターを、診察室に12週間掲示しました。
対照群では不適切な処方が43.5%から52.7%に増えたのに対し、ポスター群では42.8%から33.7%に減り、対照群と比べて19.7ポイントの差がつきました。
自分の約束が患者さんに見えている。この「宣言の見える化」が、行動を支えます。
（出典：Meeker D, et al. JAMA Intern Med 2014;174:425-31）`,
    });
    caption(s, "不適切な抗菌薬処方の割合");
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.3, legend: true, valuePt: 22,
      series: [
        { name: "対照群", labels: ["介入前", "介入期間"], values: [43.5, 52.7] },
        { name: "宣言ポスター群", labels: ["介入前", "介入期間"], values: [42.8, 33.7] },
      ],
      colors: [C.SLATE, C.TEAL], fmt: '0.0"%"', max: 64, gap: 60,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "医師の写真と署名入りの\n「不要な抗菌薬は\n処方しません」ポスターを\n診察室に12週間掲示", options: { color: C.INK } },
      { text: "→ 対照群と比べて\n19.7ポイント減", options: { bold: true, color: C.TEAL } },
    ], { gap: 16 });
  }

  // 32 当院の実践
  {
    const s = lightSlide({
      pillText: "見える化③ 当院の実践", pillColor: VIS,
      titleText: "当院の実践：行動を「見える化」する3つのナッジ",
      src: "鈴木徳洋, 他. 日本外科感染症学会（ポスター発表）",
      notes:
`【目安1分30秒／42:00–43:30】
ここで当院の実践をご紹介します。3つのナッジに取り組みました。
一つ目は、病室の入口で名札を見たら、矢印で手指消毒に誘導する表示。二つ目は、救急外来で防護具の着用手順を番号で示し、最初の番号を手指消毒にしたこと。三つ目は、手指消毒剤を1本使い切るごとにシールを貼り、使用本数の多い順に並べて、個人の使用量を見える化したことです。
そして、うまくいった方法やツールは、実施部署から全部署に共有しました。
（※実際の掲示物の写真をこのスライドに追加すると、より伝わりやすくなります）`,
    });
    const items = [
      ["FaMapSigns", "入室前の誘導", "病室の名札を見たら、\n矢印で手指消毒へ"],
      ["FaListOl", "手順の番号化", "救急外来の\nPPE着用手順を番号化\n（最初は手指消毒）"],
      ["FaStickyNote", "使用量の見える化", "1本使い切るごとに\nシールを貼り、\n多い順に並べる"],
    ];
    const w = 3.8, g = (CW - 3 * w) / 2;
    items.forEach(([ic, h, d], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.15, w, 3.7, C.CARD);
      iconCircle(s, x + w / 2 - 0.5, 2.4, 1.0, C.TEAL, A.icon[ic].white);
      T(s, h, { x, y: 3.55, w, h: 0.55, fontSize: 24, bold: true, color: C.NAVY, align: "center" });
      T(s, d, { x: x + 0.3, y: 4.2, w: w - 0.6, h: 1.5, fontSize: 20, color: C.INK, align: "center" });
    });
    T(s, "うまくいった方法やツールは、実施部署から全部署へ共有", { x: MX, y: 6.0, w: CW, h: 0.6, fontSize: 22, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // 33 当院の結果
  {
    const s = lightSlide({
      pillText: "見える化③ 当院の実践", pillColor: VIS,
      titleText: "手指消毒剤の使用量は、約2.1倍に増えた",
      src: "鈴木徳洋, 他. 日本外科感染症学会（ポスター発表）",
      notes:
`【目安1分30秒／43:30–45:00】
結果です。病院全体の手指消毒剤の使用量は、2018年度の1000患者日あたり12.6Lから、2024年度は26.3Lへ、約2.1倍に増えました。全13部署で統計学的に有意な増加でした。なお、コロナ禍の影響が大きい2020〜2023年度は比較から除いています。
一方で、直接観察による遵守率には大きな伸びは見られていません。ただ、コロナ禍を経ても低下せずに維持できています。
単年度の前後比較であり、使用量イコール遵守率ではない、という限界もあります。だからこそ、先ほどお話しした「複数の物差し」で見ていくことが大切だと考えています。看護師以外の職種へのアプローチが今後の課題です。`,
    });
    caption(s, "病院全体の手指消毒剤使用量（L／1000患者日）", MX, 1.95, 6.4);
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["2018年度", "2024年度"], values: [12.6, 26.3],
      colors: [C.SLATE, C.TEAL], fmt: "0.0", max: 31, valuePt: 28,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "全13部署で有意に増加\n（p<0.0001）", options: { bold: true, color: C.NAVY } },
      "直接観察の遵守率は\n大きな伸びなし\nただしコロナ禍後も維持",
      { text: "限界：単年度の前後比較\n使用量＝遵守率ではない", options: { color: C.MUTED } },
    ], { gap: 14 });
  }

  // 34 第4部
  section("第4部", "明日から「仕掛ける」",
`【目安20秒／45:00–45:20】
第4部では、明日から「仕掛ける」ための具体的な手順をお伝えします。`, null, "FaMagic", "4ステップ・EAST・職種別の一手");

  // 35 4ステップ
  {
    const s = lightSlide({
      pillText: "第4部 仕掛ける", titleText: "見える化は4ステップで、小さく始める",
      notes:
`【目安1分20秒／45:20–46:40】
見える化は4つのステップで進めます。
①絞る：「手指衛生を頑張る」ではなく、「誰が、いつ、どこで、何をするか」を1つに絞ります。②診断する：COM-Bで、能力・機会・動機のどこに原因があるかを探します。③仕掛ける：次にご紹介するEASTの視点で、環境や伝え方を変えます。④返す：測った結果を、現場に見える形で返します。
大きな計画より、小さく始めて回すことが成功のコツです。`,
    });
    const items = [["1", "絞る", "誰が・いつ\nどこで・何を"], ["2", "診断する", "COM-Bで\n原因の箱を探す"], ["3", "仕掛ける", "EASTで\n環境と伝え方を\n変える"], ["4", "返す", "測って、\n結果を見せる"]];
    const w = 2.75, g = (CW - 4 * w) / 3;
    items.forEach(([n, h, d], i) => {
      const x = MX + i * (w + g);
      box(s, x, 2.2, w, 3.75, C.CARD);
      numCircle(s, x + w / 2 - 0.5, 2.45, 1.0, C.NAVY, n, C.GLOW, 32);
      T(s, h, { x, y: 3.6, w, h: 0.6, fontSize: 28, bold: true, color: C.NAVY, align: "center" });
      T(s, d, { x: x + 0.25, y: 4.3, w: w - 0.5, h: 1.45, fontSize: 20, color: C.INK, align: "center" });
      if (i < 3) chevron(s, x + w + g / 2 - 0.15, 3.8, 0.3, 0.5);
    });
    T(s, "大きな計画より、小さく始めて回す", { x: MX, y: 6.1, w: CW, h: 0.55, fontSize: 24, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // 36 EAST
  {
    const s = lightSlide({
      pillText: "第4部 仕掛ける", titleText: "EAST：人が思わず動く4つの条件",
      src: "The Behavioural Insights Team. EAST. 2014",
      notes:
`【目安1分30秒／46:40–48:10】
仕掛けを考えるときのチェックリストが、英国の行動インサイトチームが提唱するEASTです。
Easy、簡単に：消毒剤を動線上や手の届くところに置く、手順を番号にする。Attractive、魅力的に：「患者さんを守る」を主語にする、目を引く表示にする。Social、社会的に：部署の実施状況を共有する、リーダーが率先する。Timely、タイミングよく：入室の直前や、手袋を外した直後に合図を出す。
今日ご紹介した研究の多くは、この4つのどれかに当てはまります。
（出典：The Behavioural Insights Team. EAST: Four simple ways to apply behavioural insights. 2014）`,
    });
    const cells = [
      ["E", C.TEAL, "Easy　簡単に", "動線上・手の届く所に消毒剤\n手順を番号化"],
      ["A", C.ORANGE, "Attractive　魅力的に", "「患者さんを守る」を主語に\n目を引く表示"],
      ["S", C.NAVY, "Social　社会的に", "部署の実施状況を共有\nリーダーが率先"],
      ["T", C.BLUE, "Timely　タイミングよく", "入室の直前に合図\n手袋を外した直後に合図"],
    ];
    const w = (CW - 0.33) / 2, h = 2.1;
    cells.forEach(([L, col, head, d], i) => {
      const x = MX + (i % 2) * (w + 0.33), y = 2.1 + Math.floor(i / 2) * (h + 0.3);
      box(s, x, y, w, h, C.CARD);
      numCircle(s, x + 0.3, y + 0.5, 1.1, col, L, C.WHITE, 40);
      T(s, head, { x: x + 1.65, y: y + 0.2, w: w - 1.85, h: 0.55, fontSize: 24, bold: true, color: col });
      T(s, d, { x: x + 1.65, y: y + 0.85, w: w - 1.85, h: 1.1, fontSize: 20, color: C.INK });
    });
  }

  // 37 職種別
  {
    const s = lightSlide({
      pillText: "第4部 仕掛ける", titleText: "職種ごとに「見えない敵」と「見える化の一手」は違う",
      notes:
`【目安1分50秒／48:10–50:00】
「見えない敵」は職種によって姿を変えます。
医師は外来後半の処方のゆらぎに、同僚比較のフィードバックが有効です。看護師は繁忙時の「触れる前」のスキップに、動線上の消毒剤と合図を。薬剤師は処方の偏りを、抗菌薬の使用データとして部署に返すことができます。臨床検査技師は、耐性菌の検出状況を部署別に見える化できます。リハビリや放射線は、機器や環境に触れた後の手指衛生。清掃や看護補助は、蛍光マーカーによるフィードバック。事務・管理者は、物品の配置や使用量を定期的に共有する。
どの職種にも、見える化の担い手としての役割があります。
※表の例は、本日ご紹介した研究をもとにした応用例です。`,
    });
    const H = { bold: true, color: C.WHITE, fill: { color: C.NAVY } };
    const data = [
      ["医師", "外来後半の処方のゆらぎ", "処方データの同僚比較"],
      ["看護師", "繁忙時の「触れる前」のスキップ", "動線上の消毒剤＋合図"],
      ["薬剤師", "処方の偏りが見えない", "抗菌薬使用量を部署へ返す"],
      ["臨床検査技師", "耐性菌の広がりが見えない", "検出状況を部署別に共有"],
      ["リハビリ・放射線", "機器・環境に触れた後", "機器に触れたら手指衛生"],
      ["清掃・看護補助", "拭き残しが見えない", "蛍光マーカーで結果を返す"],
      ["事務・管理者", "配置や人員の影響が見えない", "使用量・配置を定期共有"],
    ];
    const rows = [[{ text: "職種", options: H }, { text: "見えない敵の例", options: H }, { text: "見える化の一手", options: H }]]
      .concat(data.map((r, i) => r.map((t, j) => ({ text: t, options: { bold: j === 0, color: j === 2 ? C.TEAL : C.INK, fill: { color: i % 2 ? C.CARD : C.WHITE } } }))));
    s.addTable(rows, {
      x: MX, y: 2.05, w: CW, colW: [2.9, 4.6, CW - 7.5], rowH: 0.56,
      fontFace: FONT, fontSize: 20, valign: "middle", margin: [0.04, 0.12, 0.04, 0.12],
      border: { type: "solid", pt: 1, color: C.LINE },
    });
  }

  // 38 ミシガン
  {
    const s = lightSlide({
      pillText: "第4部 仕掛ける", titleText: "成功の鍵は、チェックリストの裏の「見えない社会の力」",
      src: "Pronovost P, et al. NEJM 2006; Dixon-Woods M, et al. Milbank Q 2011",
      notes:
`【目安1分30秒／50:00–51:30】
組織として仕掛けた有名な例が、米国ミシガン州のICUプロジェクトです。103のICUで、カテーテル関連血流感染が最大66%減り、平均で1000カテーテル日あたり7.7件から1.4件に減少しました。
この成功は「チェックリストの効果」と語られがちですが、後の社会学的な分析では、ICU同士の仲間のネットワークが規範をつくったこと、感染を「防げる問題」と捉え直したこと、感染率のデータを共有して行動の力にしたこと、現場を巻き込んで文化をつくったことが鍵だったとされています。
道具の裏にある「見えない社会の力」こそが、行動を変えるのです。
（出典：Pronovost P, et al. N Engl J Med 2006;355:2725-32／Dixon-Woods M, et al. Milbank Q 2011;89:167-205）`,
    });
    box(s, MX, 2.15, 4.9, 4.35, C.NAVY);
    T(s, "ミシガン州 103のICU", { x: MX + 0.3, y: 2.4, w: 4.3, h: 0.5, fontSize: 22, bold: true, color: C.GLOW, align: "center" });
    T(s, "カテーテル関連血流感染", { x: MX + 0.3, y: 3.0, w: 4.3, h: 0.5, fontSize: 22, color: C.WHITE, align: "center" });
    T(s, "最大66%減", { x: MX + 0.2, y: 3.6, w: 4.5, h: 1.0, fontSize: 48, bold: true, color: C.GLOW, align: "center", valign: "middle" });
    T(s, "平均 7.7→1.4", { x: MX + 0.3, y: 4.75, w: 4.3, h: 0.6, fontSize: 28, bold: true, color: C.WHITE, align: "center" });
    T(s, "（件／1000カテーテル日）", { x: MX + 0.3, y: 5.4, w: 4.3, h: 0.5, fontSize: 20, color: C.PALE, align: "center" });
    const x0 = 5.85, rw = W - MX - x0;
    T(s, "なぜ成功したのか（事後の社会学的分析）", { x: x0, y: 2.15, w: rw, h: 0.55, fontSize: 22, bold: true, color: C.NAVY });
    const rows = [
      ["FaUsers", "仲間のネットワークが「規範」をつくった"],
      ["FaBullseye", "感染を「防げる問題」と捉え直した"],
      ["FaChartBar", "データを共有し、行動の力にした"],
      ["FaHandPointer", "現場を巻き込み、文化をつくった"],
    ];
    rows.forEach(([ic, t], i) => {
      const y = 2.85 + i * 0.92;
      iconCircle(s, x0, y, 0.72, C.TEAL, A.icon[ic].white);
      T(s, t, { x: x0 + 0.95, y, w: rw - 0.95, h: 0.72, fontSize: 22, color: C.INK, valign: "middle" });
    });
  }

  // 39 ナッジは万能ではない
  {
    const s = lightSlide({
      pillText: "第4部 仕掛ける", titleText: "ナッジは万能薬ではない：組み合わせて、測って、直す",
      src: "Gould DJ (Cochrane) 2017; Lotfinejad N (Lancet Infect Dis) 2021",
      notes:
`【目安1分／51:30–52:30】
最後に、科学的に誠実であるために。手指衛生の介入を検証したコクランレビューでは、消毒剤を使用場所の近くに置くことは中等度の確実性で改善が期待できる一方、フィードバック、教育、掲示や香りなどのきっかけ、多角的な介入の多くは、効果が示唆されるものの確実性は低い、とされています。
WHOの多角的戦略も、システムの変更、教育、モニタリングとフィードバック、職場でのリマインダー、安全文化を組み合わせることを求めています。
ナッジは万能薬ではありません。組み合わせて、測って、直す。この繰り返しが大切です。
（出典：Gould DJ, et al. Cochrane Database Syst Rev 2017;9:CD005186／Lotfinejad N, et al. Lancet Infect Dis 2021;21:e209-21）`,
    });
    const H = { bold: true, color: C.WHITE, fill: { color: C.NAVY } };
    const data = [
      ["使用場所の近くに配置", "中等度"],
      ["フィードバック", "低"],
      ["教育", "低"],
      ["掲示・香りなどのきっかけ", "低"],
      ["WHO推奨の多角的介入", "低〜非常に低"],
    ];
    const rows = [[{ text: "介入（手指衛生）", options: H }, { text: "確実性", options: H }]]
      .concat(data.map((r, i) => r.map((t, j) => ({ text: t, options: { bold: j === 1, color: j === 1 ? (i === 0 ? C.TEAL : C.ORANGE) : C.INK, fill: { color: i % 2 ? C.CARD : C.WHITE } } }))));
    s.addTable(rows, {
      x: MX, y: 2.1, w: 7.0, colW: [4.6, 2.4], rowH: 0.6,
      fontFace: FONT, fontSize: 22, valign: "middle", margin: [0.04, 0.15, 0.04, 0.15],
      border: { type: "solid", pt: 1, color: C.LINE },
    });
    T(s, "手指衛生の介入26研究のコクランレビュー", { x: MX, y: 5.85, w: 7.0, h: 0.45, fontSize: 20, color: C.MUTED });
    // サイクル
    const cx = 10.2, cy = 4.25, r = 1.55, d = 1.15;
    circle(s, cx - r, cy - r, 2 * r, C.WHITE, { noFill: true, lineColor: C.LINE, lineWidth: 4 });
    const nodes = [["試す", C.TEAL, 0, -1], ["測る", C.NAVY, 1, 0], ["返す", C.TEAL, 0, 1], ["直す", C.NAVY, -1, 0]];
    nodes.forEach(([t, col, dx, dy]) => numCircle(s, cx + dx * r - d / 2, cy + dy * r - d / 2, d, col, t, C.WHITE, 24));
    T(s, "小さく回す", { x: cx - 1.0, y: cy - 0.3, w: 2.0, h: 0.6, fontSize: 22, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // 40 1分ワーク
  {
    const s = lightSlide({
      pillText: "ワーク", titleText: "1分ワーク：あなたの職場の「見えない敵」は？",
      notes:
`【目安2分／52:30–54:30】
では1分間、考えてみてください。お手元の紙にメモしていただいても構いません。
①あなたの職場で変えたい行動を1つ、誰が・いつ・どこで・何をするかまで絞ってください。②その行動を邪魔している見えない敵は、5つのうちどれでしょうか。③明日できる一手は何でしょうか。見せる、返す、仕掛ける、のどれでも構いません。
（1分後）
何人かの方に伺ってみましょう。`,
    });
    const rows = [
      ["1", "変えたい行動", "誰が・いつ・どこで・何を"],
      ["2", "見えない敵", "思い込み／見られている効果\n疲れ／周囲の空気／環境と動線"],
      ["3", "明日の一手", "見せる・返す・仕掛ける"],
    ];
    rows.forEach(([n, h, sub], i) => {
      const y = 2.15 + i * 1.5;
      numCircle(s, MX, y + 0.2, 0.9, C.NAVY, n, C.GLOW, 30);
      T(s, h, { x: 1.75, y: y + 0.05, w: 4.3, h: 0.5, fontSize: 26, bold: true, color: C.NAVY });
      T(s, sub, { x: 1.75, y: y + 0.6, w: 4.4, h: 0.8, fontSize: 20, color: C.MUTED });
      box(s, 6.35, y + 0.1, W - MX - 6.35, 1.15, C.WHITE, { lineColor: C.SLATE, lineWidth: 2, dash: "dash" });
    });
  }

  // 41 まとめ
  {
    const s = lightSlide({
      pillText: "まとめ", titleText: "まとめ：見えない敵は、見える化すれば対処できる",
      notes:
`【目安1分／54:30–55:30】
まとめです。
一つ、行動の原因は目に見えません。人を責める前に、仕組みを疑いましょう。
二つ、行動科学、特にCOM-Bを使えば、見えない敵を診断できます。
三つ、見せる・返す・仕掛ける。小さく始めて、測って、直していきましょう。`,
    });
    const rows = [
      ["行動の原因は目に見えない", "人を責める前に、仕組みを疑う"],
      ["COM-Bで原因を診断する", "5つの見えない敵を探す"],
      ["見せる・返す・仕掛ける", "小さく始めて、測って、直す"],
    ];
    rows.forEach(([a, b], i) => {
      const y = 2.2 + i * 1.5;
      box(s, MX, y, CW, 1.3, C.CARD);
      numCircle(s, MX + 0.25, y + 0.2, 0.9, C.TEAL, String(i + 1), C.WHITE, 30);
      T(s, a, { x: 1.95, y, w: 5.2, h: 1.3, fontSize: 26, bold: true, color: C.NAVY, valign: "middle" });
      chevron(s, 7.15, y + 0.4, 0.32, 0.5);
      T(s, b, { x: 7.75, y, w: W - MX - 7.85, h: 1.3, fontSize: 24, bold: true, color: C.TEAL, valign: "middle" });
    });
  }

  // 42 クロージング
  {
    const s = darkSlide(
`【目安30秒／55:30–56:00】
見えない敵は、「人の中」ではなく「仕組みの中」にいます。
明日、ひとつだけ、仕掛けてみませんか。`);
    s.addImage({ data: A.lensPlain, x: 5.67, y: 0.7, w: 2.0, h: 2.0 });
    s.addImage({ data: A.icon.FaSearchPlus.glow, x: 6.27, y: 1.3, w: 0.8, h: 0.8 });
    T(s, [
      { text: "見えない敵は「人の中」ではなく", options: { breakLine: true } },
      { text: "「仕組みの中」", options: { color: C.GLOW } },
      { text: "にいる" },
    ], { x: 0.9, y: 2.95, w: 11.53, h: 1.9, fontSize: 40, bold: true, color: C.WHITE, align: "center", valign: "middle", lineSpacingMultiple: 1.1 });
    T(s, "明日、ひとつだけ仕掛けてみませんか。", { x: 0.9, y: 5.15, w: 11.53, h: 0.7, fontSize: 28, color: C.GLOW, align: "center", valign: "middle" });
  }

  // 43 質疑応答
  {
    const s = darkSlide(
`【目安4分／56:00–60:00】
ご清聴ありがとうございました。ご質問をお受けします。`);
    T(s, "ご清聴ありがとうございました", { x: 0.9, y: 2.2, w: 11.53, h: 1.0, fontSize: 40, bold: true, color: C.WHITE, align: "center", valign: "middle" });
    T(s, "質疑応答", { x: 0.9, y: 3.4, w: 11.53, h: 0.7, fontSize: 30, bold: true, color: C.GLOW, align: "center", valign: "middle" });
    T(s, paras(["市立豊中病院 感染対策室", "感染管理認定看護師　鈴木 徳洋"]),
      { x: 0.9, y: 4.9, w: 11.53, h: 0.9, fontSize: 20, color: C.PALE, align: "center" });
  }

  // 44– 引用文献
  const refs = [
    "Erasmus V, et al. Infect Control Hosp Epidemiol. 2010;31:283-94.",
    "Pittet D, et al. Lancet. 2000;356:1307-12.",
    "Webb TL, Sheeran P. Psychol Bull. 2006;132:249-68.",
    "Jenner EA, et al. J Hosp Infect. 2006;63:418-22.",
    "Grant AM, Hofmann DA. Psychol Sci. 2011;22:1494-9.",
    "Srigley JA, et al. BMJ Qual Saf. 2014;23:974-80.",
    "Pittet D, et al. Ann Intern Med. 2004;141:1-8.",
    "Dai H, et al. J Appl Psychol. 2015;100:846-62.",
    "Linder JA, et al. JAMA Intern Med. 2014;174:2029-31.",
    "Pittet D, et al. Ann Intern Med. 1999;130:126-30.",
    "Lankford MG, et al. Emerg Infect Dis. 2003;9:217-23.",
    "Monsalve MN, et al. Infect Control Hosp Epidemiol. 2014;35:1277-85.",
    "Gould DJ, et al. Cochrane Database Syst Rev. 2017;9:CD005186.",
    "Michie S, et al. Implement Sci. 2011;6:42.",
    "Stiefel U, et al. Infect Control Hosp Epidemiol. 2011;32:185-7.",
    "Tomas ME, et al. JAMA Intern Med. 2015;175:1904-10.",
    "Carling PC, et al. Infect Control Hosp Epidemiol. 2008;29:1035-41.",
    "Hallsworth M, et al. Lancet. 2016;387:1743-52.",
    "Meeker D, et al. JAMA. 2016;315:562-70.",
    "Meeker D, et al. JAMA Intern Med. 2014;174:425-31.",
    "鈴木徳洋, 他. 手指衛生遵守率向上への行動経済学的アプローチ. 日本外科感染症学会（ポスター発表）.",
    "The Behavioural Insights Team. EAST: Four simple ways to apply behavioural insights. 2014.",
    "Pronovost P, et al. N Engl J Med. 2006;355:2725-32.",
    "Dixon-Woods M, et al. Milbank Q. 2011;89:167-205.",
    "Lotfinejad N, et al. Lancet Infect Dis. 2021;21:e209-21.",
  ];
  const pages = [refs.slice(0, 9), refs.slice(9, 18), refs.slice(18)];
  let n = 1;
  pages.forEach((list, p) => {
    const s = lightSlide({
      pillText: "参考資料", titleText: `引用文献（${p + 1}/${pages.length}）`,
      notes: "引用文献の一覧です（講演中は表示のみ）。各文献のDOI・要点は別添の「引用文献リスト.pdf」にまとめています。",
    });
    T(s, list.map((r, i) => ({ text: `${n + i}. ${r}`, options: { breakLine: i < list.length - 1, paraSpaceAfter: 8 } })),
      { x: MX, y: 2.05, w: CW, h: 4.6, fontSize: 20, color: C.INK });
    n += list.length;
  });
}

(async () => {
  A = await buildAssets();
  build();
  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
