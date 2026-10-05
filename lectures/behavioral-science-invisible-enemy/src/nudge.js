// 行動科学は感染対策の「見えない敵」を可視化する（約60分）
// タイトル30pt以上・本文20pt以上・BIZ UDゴシック／1スライド1メッセージ
const pptxgen = require("pptxgenjs");
const { build: buildAssets } = require("./assets");

const OUT = process.argv[2] || "nudge.pptx";
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
    ], 4)), { x: MX + 0.35, y: 3.75, w: 5.0, h: 1.6, fontSize: 26, color: C.WHITE, valign: "top" });
    T(s, "× 禁止・強制　× ごほうび・罰金", { x: MX + 0.35, y: 5.6, w: 5.0, h: 0.6, fontSize: 20, color: C.PALE, valign: "middle" });
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

  // N2 4つの道具
  {
    const s = lightSlide({
      pillText: P, pillColor: NUD,
      titleText: "人を動かす4つの道具。ナッジは「選ぶ自由」を残す",
      src: "Thaler RH, Sunstein CR. 2008; Michie S, et al. Implement Sci 2011",
      notes:
`【目安1分】
人の行動を変える道具を、大きく4つに分けてみます。
1つ目はルール。「必ずこうする」という義務や禁止です。2つ目はお金。ごほうびや罰金です。3つ目は教育。知識を伝える研修やマニュアルです。そして4つ目がナッジ。選択肢や環境の「見せ方・置き方」を設計して、自然に選びやすくします。
例えば手指衛生なら、手順を決めるのがルール、研修をするのが教育、入口の動線上に消毒剤を置いたり、足あとで誘導したりするのがナッジです。
大切なのは、ナッジは他の道具の代わりではないということ。ルールや教育と組み合わせて、「わかっているのにできない」の最後のひと押しを担います。
（出典：Thaler RH, Sunstein CR. Nudge. 2008／Michie S, et al. Implement Sci 2011;6:42）`,
    });
    const tools = [
      ["FaClipboardList", "ルール", "義務・禁止", "手順を決める", C.SLATE],
      ["FaYenSign", "お金", "ごほうび・罰金", "表彰・手当", C.SLATE],
      ["FaBook", "教育", "知識を伝える", "研修・マニュアル", C.SLATE],
      ["FaHandPointer", "ナッジ", "選び方を設計する", "入口に消毒剤\n足あとで誘導", C.BLUE],
    ];
    const w = 2.85, g = (CW - 4 * w) / 3;
    tools.forEach(([ic, h, d, e, col], i) => {
      const x = MX + i * (w + g), hi = i === 3;
      box(s, x, 2.15, w, 3.7, hi ? C.BLUE : C.CARD);
      iconCircle(s, x + w / 2 - 0.45, 2.35, 0.9, hi ? C.WHITE : C.NAVY, hi ? A.icon[ic].navy : A.icon[ic].white);
      T(s, h, { x, y: 3.35, w, h: 0.55, fontSize: 28, bold: true, color: hi ? C.WHITE : C.NAVY, align: "center", valign: "middle" });
      T(s, d, { x: x + 0.15, y: 3.95, w: w - 0.3, h: 0.5, fontSize: 20, bold: true, color: hi ? C.WHITE : C.INK, align: "center" });
      T(s, paras(["例（手指衛生）", ...e.split("\n")]), { x: x + 0.15, y: 4.6, w: w - 0.3, h: 1.15, fontSize: 20, color: hi ? C.WHITE : C.MUTED, align: "center" });
    });
    T(s, "ナッジは代わりではなく、ルール・教育と組み合わせて使う", { x: MX, y: 6.0, w: CW, h: 0.65, fontSize: 24, bold: true, color: C.TEAL, align: "center", valign: "middle" });
  }

  // N3 デフォルトの力
  {
    const s = lightSlide({
      pillText: P, pillColor: NUD,
      titleText: "最初から「選ばれている」ものを、人はそのまま選ぶ",
      src: "Johnson EJ, Goldstein D. Science 2003",
      notes:
`【目安1分】
ナッジの中でも特に強力なのが「初期設定（デフォルト）」です。
臓器提供の意思をたずねるオンライン実験で、「提供したい人がチェックを入れる」聞き方では、同意した人は約42%でした。ところが「提供したくない人がチェックを外す」、つまり最初から同意が選ばれている聞き方では、約82%に上がりました。選ぶ自由はどちらも同じです。変えたのは「最初の設定」だけです。
国別に見ても、同意が初期設定になっている国では、実際の同意率が90%を超える国が多く見られました。
人は、最初から選ばれているものを、そのまま選びやすいのです。
（出典：Johnson EJ, Goldstein D. Science 2003;302:1338-9）`,
    });
    caption(s, "臓器提供に同意した人の割合（オンライン実験）", MX, 1.95, 6.6);
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["自分で\nチェックを入れる", "最初から\nチェック済み"], values: [42, 82],
      colors: [C.SLATE, C.BLUE], fmt: '0"%"', max: 100, valuePt: 28,
    });
    card(s, 7.1, 2.2, W - MX - 7.1, 4.3, C.CARD, [
      { text: "初期設定（デフォルト）の力", options: { bold: true, color: C.NAVY } },
      "変えたのは「最初の設定」だけ\n選ぶ自由はどちらも同じ",
      "国別でも、同意が初期設定の国は\n同意率90%超が多い",
      { text: "→ 人は、最初から選ばれて\nいるものを選びやすい", options: { bold: true, color: C.TEAL } },
    ], { gap: 14, pt: 20, valign: "middle" });
  }

  // N4 予防接種を予約済みに
  {
    const s = lightSlide({
      pillText: P, pillColor: NUD,
      titleText: "予防接種も、最初から「予約済み」にすると増える",
      src: "Chapman GB, et al. JAMA 2010; Dai H, et al. Nature 2021",
      notes:
`【目安1分】
このデフォルトは、感染対策でも使えます。
米国の大学職員478名を対象にしたランダム化比較試験で、インフルエンザワクチンの案内を2通り送りました。「接種したい人は自分で予約してください」という案内では接種率33%。一方、「あなたの接種日時はすでに予約されています。都合が悪ければ変更・取消できます」という案内では45%でした。予約を取り消したのは8%だけでした。
新型コロナワクチンでも、9万人以上を対象にした試験で、「あなたのためのワクチンが用意されています」と伝えるショートメッセージが、接種率を3.57ポイント押し上げました。
職員のワクチン接種や、物品の初期配置など、「最初の設定」を見直すことは、明日からできるナッジです。
（出典：Chapman GB, et al. JAMA 2010;304:43-4／Dai H, et al. Nature 2021;597:404-9）`,
    });
    caption(s, "インフルエンザワクチン接種率（職員478名）", MX, 1.95, 6.6);
    barChart(s, {
      x: MX, y: 2.35, w: 6.1, h: 4.25,
      labels: ["自分で予約", "最初から予約済み\n（変更・取消は自由）"], values: [33, 45],
      colors: [C.SLATE, C.BLUE], fmt: '0"%"', max: 56, valuePt: 28,
    });
    const x0 = 7.1, rw = W - MX - x0;
    box(s, x0, 2.2, rw, 1.9, C.CARD);
    iconCircle(s, x0 + 0.25, 2.45, 0.8, C.BLUE, A.icon.FaCalendarCheck.white);
    T(s, "予約を取り消した\nのは8%だけ", { x: x0 + 1.25, y: 2.4, w: rw - 1.4, h: 1.5, fontSize: 22, bold: true, color: C.NAVY, valign: "middle" });
    box(s, x0, 4.3, rw, 2.2, C.CARD);
    iconCircle(s, x0 + 0.25, 4.55, 0.8, C.BLUE, A.icon.FaSms.white);
    T(s, paras(expand([
      { text: "「あなたのワクチンが\n用意されています」", options: { bold: true, color: C.NAVY } },
      "SMSで接種率＋3.57ポイント\n（新型コロナ・約9.3万人）",
    ], 6)), { x: x0 + 1.25, y: 4.45, w: rw - 1.4, h: 1.95, fontSize: 20, valign: "middle" });
  }
}

(async () => {
  A = await buildAssets();
  build();
  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
