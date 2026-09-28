// SVG → PNG の装飾素材とアイコンを生成する
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa");

const GLOW = "#3FE8C2";

async function svgToPng(svg, w, h) {
  const buf = await sharp(Buffer.from(svg)).resize(w, h).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function icon(name, color = "#FFFFFF", size = 256) {
  const Comp = fa[name];
  if (!Comp) throw new Error("icon not found: " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { size, color: color })
  );
  return svgToPng(svg, size, size);
}

// 拡大鏡（レンズ）＋光る微生物：タイトル用
function lensSvg({ microbes = true, handle = true } = {}) {
  let m = "";
  if (microbes) {
    const cocci = [
      [300, 250, 22, 0.9], [340, 262, 20, 0.75], [322, 296, 21, 0.85],
      [470, 420, 16, 0.6], [496, 438, 15, 0.5], [250, 470, 18, 0.7],
      [520, 260, 12, 0.45], [420, 520, 13, 0.55], [205, 330, 10, 0.4],
      [560, 360, 11, 0.4], [360, 600, 10, 0.35], [455, 190, 9, 0.35],
    ];
    cocci.forEach(([x, y, r, o]) => {
      m += `<circle cx="${x}" cy="${y}" r="${r}" fill="${GLOW}" fill-opacity="${o}"/>`;
    });
    const rods = [
      [410, 330, 30, 0.8], [250, 560, -25, 0.55], [540, 520, 60, 0.5],
      [300, 400, 75, 0.45], [480, 300, -40, 0.4],
    ];
    rods.forEach(([x, y, rot, o]) => {
      m += `<rect x="${x - 38}" y="${y - 13}" width="76" height="26" rx="13" fill="${GLOW}" fill-opacity="${o}" transform="rotate(${rot} ${x} ${y})"/>`;
    });
  }
  const h = handle
    ? `<line x1="610" y1="610" x2="820" y2="820" stroke="#2E4274" stroke-width="64" stroke-linecap="round"/>
       <line x1="610" y1="610" x2="820" y2="820" stroke="${GLOW}" stroke-opacity="0.9" stroke-width="18" stroke-linecap="round"/>`
    : "";
  const d = handle ? 0 : 60; // ハンドルなしは中央に寄せる
  return `<svg xmlns="http://www.w3.org/2000/svg" width="900" height="900" viewBox="0 0 900 900">
    <defs><filter id="g" x="-30%" y="-30%" width="160%" height="160%"><feGaussianBlur stdDeviation="14"/></filter></defs>
    <g transform="translate(${d} ${d})">
    ${h}
    <circle cx="390" cy="390" r="300" fill="#1E2A4F" fill-opacity="0.55"/>
    <circle cx="390" cy="390" r="300" fill="none" stroke="${GLOW}" stroke-opacity="0.6" stroke-width="30" filter="url(#g)"/>
    ${m}
    <circle cx="390" cy="390" r="300" fill="none" stroke="${GLOW}" stroke-width="14"/>
    </g>
  </svg>`;
}

// 氷山（水面の上＝行動／水面の下＝原因）
function icebergSvg() {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="1240" height="920" viewBox="0 0 1240 920">
    <defs>
      <linearGradient id="w" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0" stop-color="#CFE3F3"/><stop offset="1" stop-color="#7FA9CF"/>
      </linearGradient>
    </defs>
    <rect x="0" y="300" width="1240" height="620" rx="24" fill="url(#w)"/>
    <polygon points="360,300 900,300 1030,450 1090,640 990,830 700,900 400,850 250,660 300,450"
      fill="#F4F9FD" fill-opacity="0.92" stroke="#5E8DB8" stroke-width="5"/>
    <polygon points="470,300 540,170 610,95 665,150 720,125 800,300"
      fill="#FFFFFF" stroke="#5E8DB8" stroke-width="5"/>
    <path d="M0,300 Q60,285 120,300 T240,300 T360,300 T480,300 T600,300 T720,300 T840,300 T960,300 T1080,300 T1200,300 L1240,300"
      fill="none" stroke="#3F6E99" stroke-width="5"/>
  </svg>`;
}

async function build() {
  const out = {};
  out.lensTitle = await svgToPng(lensSvg({ microbes: true, handle: true }), 900, 900);
  out.lensPlain = await svgToPng(lensSvg({ microbes: false, handle: false }), 900, 900);
  out.iceberg = await svgToPng(icebergSvg(), 1240, 920);
  const names = [
    "FaMicroscope", "FaVial", "FaDna", "FaSearchPlus", "FaQuestion", "FaEye",
    "FaUserCheck", "FaBatteryQuarter", "FaUsers", "FaRoute", "FaHospital",
    "FaBolt", "FaCoffee", "FaPumpSoap", "FaNetworkWired", "FaMapSigns",
    "FaListOl", "FaStickyNote", "FaChartBar", "FaMagic", "FaStethoscope",
    "FaCommentDots", "FaTimes", "FaCheck", "FaHandPointer", "FaBell",
    "FaClipboardList", "FaBullseye", "FaSyncAlt", "FaEnvelopeOpenText",
    "FaUserMd", "FaUserNurse", "FaPills", "FaFlask", "FaWalking", "FaBroom",
    "FaBuilding", "FaClock", "FaNotEqual", "FaArrowRight",
  ];
  out.icon = {};
  for (const n of names) {
    out.icon[n] = {
      white: await icon(n, "#FFFFFF"),
      glow: await icon(n, GLOW),
      teal: await icon(n, "#0B7A70"),
      orange: await icon(n, "#D1480E"),
      navy: await icon(n, "#131B35"),
    };
  }
  return out;
}

module.exports = { build };
