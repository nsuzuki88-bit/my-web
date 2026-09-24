// 引用文献リスト（PDF）を index.html のデータから生成する。
// 使い方: node hhsaf/tools/build-references-pdf.mjs
// Chromium の場所は環境変数 CHROME で変更できる。
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const appDir = join(here, '..');
const html = readFileSync(join(appDir, 'index.html'), 'utf8');
const data = html.split('/* DATA:BEGIN */')[1].split('/* DATA:END */')[0];

const env = { pts: () => 0 };
const D = new Function('env',
  'const pointsOf = (id, ans) => env.pts(id, ans);\n' + data +
  '\nreturn {COMPONENTS, QUESTIONS, LEADERSHIP, LEVELS, ACTIONS, TIERS, RUBRIC, REFS};')(env);

const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const num = Object.fromEntries(D.REFS.map((r, i) => [r.key, i + 1]));
const keysIn = s => [...String(s).matchAll(/\[\[([a-z0-9]+)\]\]/g)].map(m => m[1]);
const citeText = s => esc(String(s).replace(/\[\[([a-z0-9]+)\]\]/g, (m, k) => `[${num[k]}]`));
const qmax = q => Math.max(...q.o.map(o => o[1]));

function easeRange(a) {
  if (typeof a.ease !== 'function') return { lo: a.ease, hi: a.ease };
  env.pts = () => 0; const x = a.ease({});
  env.pts = () => 99; const y = a.ease({});
  return { lo: Math.min(x, y), hi: Math.max(x, y) };
}
const tierOf = (i, e) => (i >= 4 && e >= 4 ? 'qw' : i >= 4 ? 'major' : e >= 4 ? 'fill' : 'later');

const link = r => r.doi ? `https://doi.org/${r.doi}` : (r.url || '');
const today = new Date();
const ymd = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;

const compRows = D.COMPONENTS.map(c => {
  const qs = D.QUESTIONS.filter(q => q.c === c.no && !q.gate);
  return `<tr><td>${c.no}</td><td>${esc(c.name)}<br><span class="en">${esc(c.en)}</span></td><td>${qs.length}</td><td>${qs.reduce((a, q) => a + qmax(q), 0) > 100 ? '100（上限）' : qs.reduce((a, q) => a + qmax(q), 0)}</td><td>${esc(c.desc)}</td></tr>`;
}).join('');

const actionRows = D.ACTIONS.map(a => {
  const { lo, hi } = easeRange(a);
  const e = lo === hi ? String(lo) : `${lo}〜${hi}`;
  const p = lo === hi ? String(a.impact * lo) : `${a.impact * lo}〜${a.impact * hi}`;
  const tier = D.TIERS[tierOf(a.impact, hi)].label;
  const refs = [...new Set(keysIn(a.why))].map(k => num[k]).sort((x, y) => x - y).join(', ');
  return `<tr><td>${a.id.slice(1)}</td><td><b>${esc(a.title)}</b><br><span class="small">${citeText(a.why)}${a.easeNote ? `<br>容易さ：${esc(a.easeNote)}` : ''}</span></td>
    <td>${a.qs.filter(id => id !== 'g3.4').join('<br>')}</td><td class="c">${a.impact}</td><td class="c">${e}</td><td class="c">${p}</td><td>${tier}</td><td>${refs}</td></tr>`;
}).join('');

const refItems = D.REFS.map((r, i) => `<li><span class="rn">${i + 1}</span><div>
  <p class="cit">${esc(r.au)}. <b>${esc(r.ti)}</b>. ${esc(r.so)}.</p>
  ${link(r) ? `<p class="lnk">${esc(link(r))}</p>` : ''}
  <table class="kv"><tr><th>研究デザイン</th><td>${esc(r.de)}</td></tr><tr><th>主な知見</th><td>${esc(r.fi)}</td></tr><tr><th>アプリでの用途</th><td>${esc(r.use)}</td></tr></table>
</div></li>`).join('');

const rubric = (cap, rows) => `<table class="grid"><caption>${cap}</caption>${rows.map(r => `<tr><td class="c">${r[0]}</td><td>${esc(r[1])}</td></tr>`).join('')}</table>`;

const page = `<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>引用文献リスト HHSAF 手指衛生自己評価アプリ</title>
<style>
@page{size:A4;margin:16mm 14mm 16mm 14mm}
body{font-family:"IPAPGothic","IPAGothic","Noto Sans CJK JP","Hiragino Sans","Yu Gothic",sans-serif;font-size:9.4pt;line-height:1.6;color:#15211f}
h1{font-size:16pt;color:#07463F;margin:0 0 4pt}
h2{font-size:12pt;color:#07463F;border-bottom:1.2pt solid #0B6B60;padding-bottom:2pt;margin:16pt 0 6pt;break-after:avoid}
p{margin:0 0 4pt}
.sub{color:#4a5b57;font-size:9pt}
.note{background:#eef5f3;border-left:3pt solid #0B6B60;padding:5pt 8pt;margin:8pt 0;font-size:8.8pt}
table{border-collapse:collapse;width:100%}
.grid{margin:4pt 0 8pt}
.grid caption{text-align:left;font-weight:bold;padding-bottom:2pt}
.grid th,.grid td{border:0.6pt solid #b9c7c3;padding:3pt 5pt;vertical-align:top;text-align:left}
.grid th{background:#e4eeeb;font-weight:bold}
.c{text-align:center !important;white-space:nowrap}
.en{color:#687a75;font-size:8pt}
.small{font-size:8.2pt;color:#33433f}
ol.refs{list-style:none;padding:0;margin:0}
ol.refs li{display:grid;grid-template-columns:22pt 1fr;gap:4pt;padding:5pt 0;border-bottom:0.5pt solid #d3ddda;break-inside:avoid}
.rn{font-weight:bold;color:#0B6B60;text-align:right;font-size:10.5pt}
.cit{margin:0}
.lnk{color:#0B6B60;font-size:8.2pt;word-break:break-all;margin:1pt 0 2pt}
.kv th{width:62pt;text-align:left;font-weight:normal;color:#5b6c68;vertical-align:top;padding:1pt 4pt 1pt 0;font-size:8.4pt}
.kv td{padding:1pt 0;font-size:8.6pt}
.acts td{font-size:8.4pt}
.acts tr{break-inside:avoid}
</style></head><body>
<h1>引用文献リスト</h1>
<p class="sub">HHSAF 手指衛生自己評価アプリ（WHO Hand Hygiene Self-Assessment Framework 2010 準拠）｜作成日 ${ymd}</p>
<div class="note">このリストは、アプリの推奨事項と改善アクションの優先度（効果・実施の容易さ）の根拠とした文献の書誌情報、研究デザイン、主な知見、アプリでの用途をまとめたものです。出版社の論文PDFは作成環境のネットワーク制限により取得できなかったため、本文は各文献のDOIまたはURLから参照してください。番号はアプリ内の引用番号と一致します。</div>

<h2>1. 評価の枠組み（HHSAF 2010）</h2>
<table class="grid"><tr><th class="c">No.</th><th>構成要素</th><th class="c">設問数</th><th class="c">満点</th><th>内容</th></tr>${compRows}</table>
<table class="grid"><caption>総合レベル（計500点）</caption><tr><th>レベル</th><th class="c">得点</th><th>WHOによる解釈</th></tr>
${D.LEVELS.map(l => `<tr><td>${esc(l.name)}（${esc(l.en)}）</td><td class="c">${l.min}〜${l.max}</td><td>${esc(l.desc)}</td></tr>`).join('')}</table>
<p>設問1.7は1.1〜1.6の合計が100点未満の場合のみ、3.4a・3.4bは観察者が訓練・検証を受け「5つのタイミング」（または同等）の方法で観察している場合のみ得点します。リーダーシップ基準（20項目）は上級の施設が評価し、合計12項目以上かつ各構成要素で1項目以上「はい」の場合に「手指衛生リーダーシップ」レベルと判定します（日本語訳版の判定文は「または」と読める表記ですが、WHO英語版原典と国際調査の運用に合わせています）。構成要素ごとの得点帯（0〜25／26〜50／51〜75／76〜100点）は、総合レベルの区分を5分の1にしたアプリ独自の参考区分です。</p>

<h2>2. 改善アクションの優先度の評価方法</h2>
<p>未達の設問を改善アクションにまとめ、期待される効果（1〜5）と実施の容易さ（1〜5）の積を優先度とし、高い順に並べます。同点のときは効果の高いもの、次にHHSAF得点の伸びが大きいものを上位にします。</p>
${rubric('期待される効果', D.RUBRIC.impact)}
${rubric('実施の容易さ', D.RUBRIC.ease)}
<table class="grid"><caption>優先区分</caption>${Object.values(D.TIERS).map(t => `<tr><td>${esc(t.long)}</td><td>${esc(t.rule)}</td></tr>`).join('')}</table>

<h2>3. 引用文献</h2>
<ol class="refs">${refItems}</ol>

<h2>付録　改善アクション別の評価と根拠</h2>
<table class="grid acts"><tr><th class="c">No.</th><th>アクションと根拠の要約</th><th>設問</th><th class="c">効果</th><th class="c">容易さ</th><th class="c">優先度</th><th>区分</th><th>文献</th></tr>${actionRows}</table>
<p class="sub">容易さが「3〜4」などと幅のあるアクションは、現在の回答によって値が変わります（アプリでは回答に応じて自動で判定）。</p>
</body></html>`;

const work = mkdtempSync(join(tmpdir(), 'hhsaf-refs-'));
const src = join(work, 'references.html');
writeFileSync(src, page);
const out = join(appDir, 'references-hhsaf.pdf');
const chrome = process.env.CHROME || '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';
execFileSync(chrome, ['--headless=new', '--no-sandbox', '--disable-gpu', '--no-pdf-header-footer',
  `--print-to-pdf=${out}`, 'file://' + src], { stdio: 'ignore' });
console.log('wrote', out);
