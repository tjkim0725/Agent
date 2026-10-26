// 강의 슬라이드 생성. 실행: NODE_PATH=<node_modules> node build_pptx.js <출력 경로>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");
const K = require("./content");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Tosca MCP Training",
  headFontFace: "맑은 고딕",
  bodyFontFace: "맑은 고딕",
  colors: {
    dk1: "1E2B2F", lt1: "FFFFFF", dk2: "0F4C5C", lt2: "EAF1F2",
    accent1: "0F4C5C", accent2: "E0A03A", accent3: "5E9AA6",
    accent4: "2B7D50", accent5: "C4453C", accent6: "56666B",
    hlink: "0F4C5C", folHlink: "5E9AA6",
  },
};
const HEX = THEME.colors;
const MONO = "Consolas";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = K.meta.course;
pres.subject = K.meta.subtitle;
pres.author = "QA 교육";
pres.company = "";
const C = pres.SchemeColor;
const S = pres.shapes;

// ---------- 아이콘 ----------
const ICONS = {};
async function icon(name, hex) {
  const key = name + "_" + hex;
  if (ICONS[key]) return ICONS[key];
  const svg = RDS.renderToStaticMarkup(React.createElement(lu[name], { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  ICONS[key] = "image/png;base64," + buf.toString("base64");
  return ICONS[key];
}

// ---------- 레이아웃 ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.05, w: 11.7, h: 1.8, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 7.6, h: 0.7, fontSize: 22, color: C.background2, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "meta", type: "body", x: 0.8, y: 6.25, w: 11.7, h: 0.5, fontSize: 12, color: C.background2, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: "num", type: "body", x: 0.8, y: 1.7, w: 4, h: 1.4, fontSize: 72, bold: true, color: C.accent2, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 3.25, w: 11.5, h: 1.0, fontSize: 38, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.35, w: 11.5, h: 0.8, fontSize: 18, color: C.background2, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.38, w: 12.13, h: 0.85, fontSize: 32, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: K.meta.course, options: { x: 0.6, y: 7.02, w: 6, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 12.13, y: 7.02, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 0.55, w: 11.7, h: 0.9, fontSize: 34, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});

// ---------- 공통 그리기 함수 ----------
let objSeq = 0;
const nm = (p) => `${p}_${++objSeq}`;
const shadow = () => ({ type: "outer", color: HEX.dk1, opacity: 0.12, blur: 6, offset: 2, angle: 90 });

function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 15, valign: "top", objectName: nm("text") }, o));
}
function card(s, x, y, w, h, o = {}) {
  s.addShape(S.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: o.r ?? 0.1,
    fill: o.fill ?? { color: C.background2 },
    line: o.line ?? { type: "none" },
    shadow: o.shadow ? shadow() : undefined,
    objectName: nm(o.name || "card"),
  });
}
function badge(s, x, y, d, label, o = {}) {
  s.addText(label, {
    shape: S.OVAL, x, y, w: d, h: d, isTextBox: true, margin: 0,
    fill: { color: o.fill ?? C.accent1 }, color: o.color ?? C.background1,
    fontSize: o.fontSize ?? 14, bold: true, align: "center", valign: "middle", objectName: nm("badge"),
  });
}
function chip(s, x, y, w, h, label, o = {}) {
  s.addText(label, {
    shape: S.ROUNDED_RECTANGLE, rectRadius: 0.06, x, y, w, h, isTextBox: true, margin: [0, 4, 0, 4],
    fill: { color: o.fill ?? C.background2, transparency: o.transparency ?? 0 },
    line: o.line ?? { type: "none" },
    color: o.color ?? C.text1, fontSize: o.fontSize ?? 11, bold: o.bold ?? false,
    fontFace: o.mono ? MONO : undefined, align: o.align ?? "center", valign: "middle", objectName: nm("chip"),
  });
}
function line(s, x1, y1, x2, y2, o = {}) {
  s.addShape(S.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: o.color ?? C.accent6, width: o.width ?? 1.25, dashType: o.dash ?? "solid", endArrowType: o.arrow ? "triangle" : undefined },
    objectName: nm("line"),
  });
}
async function iconCircle(s, name, x, y, d, o = {}) {
  s.addShape(S.OVAL, { x, y, w: d, h: d, fill: { color: o.fill ?? C.accent1, transparency: o.transparency ?? 0 }, line: { type: "none" }, objectName: nm("iconbg") });
  const pad = d * 0.24;
  s.addImage({ data: await icon(name, o.iconHex ?? HEX.lt1), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: nm("icon") });
}
async function iconImg(s, name, x, y, size, hex) {
  s.addImage({ data: await icon(name, hex), x, y, w: size, h: size, objectName: nm("icon") });
}
function promptBox(s, x, y, w, h, label, text, fs = 13) {
  s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.accent2, transparency: 86 }, line: { color: C.accent2, width: 0.75 }, objectName: nm("prompt") });
  chip(s, x + 0.18, y + 0.13, 1.35, 0.3, label, { fill: C.accent2, bold: true, fontSize: 11 });
  txt(s, text, { x: x + 0.2, y: y + 0.55, w: w - 0.4, h: h - 0.68, fontSize: fs, lineSpacingMultiple: 1.05 });
}
function codeBox(s, x, y, w, h, code, fs = 12) {
  s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.text1 }, line: { type: "none" }, objectName: nm("code") });
  txt(s, code, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: h - 0.4, fontFace: MONO, fontSize: fs, color: C.background1 });
}
function content(title, notes, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  s.addNotes(notes);
  return s;
}
function labTag(s, lab, time) {
  chip(s, 10.55, 0.6, 1.0, 0.42, lab, { fill: C.accent1, color: C.background1, bold: true, fontSize: 13 });
  chip(s, 11.68, 0.6, 1.05, 0.42, time, { fill: C.background2, color: C.text2, bold: true, fontSize: 13 });
}
function checkList(s, x, y, w, items, o = {}) {
  const rowH = o.rowH ?? 0.42;
  items.forEach((t, i) => {
    s.addShape(S.ROUNDED_RECTANGLE, { x, y: y + i * rowH + 0.07, w: 0.24, h: 0.24, rectRadius: 0.04, fill: { color: C.background1 }, line: { color: C.accent1, width: 1 }, objectName: nm("checkbox") });
    txt(s, t, { x: x + 0.38, y: y + i * rowH, w: w - 0.38, h: rowH, fontSize: o.fs ?? 14, valign: "middle" });
  });
}
function sectionHead(s, x, y, w, t) {
  txt(s, t, { x, y, w, h: 0.35, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
}

// 채팅 화면 목업
function chatWindow(s, x, y, w, h, header) {
  card(s, x, y, w, h, { fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, shadow: true, name: "chatwin" });
  s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h: 0.45, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: nm("chathead") });
  s.addShape(S.RECTANGLE, { x, y: y + 0.3, w, h: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: nm("chathead2") });
  s.addShape(S.OVAL, { x: x + 0.2, y: y + 0.16, w: 0.13, h: 0.13, fill: { color: C.accent4 }, line: { type: "none" }, objectName: nm("dot") });
  txt(s, header, { x: x + 0.42, y: y + 0.07, w: w - 0.6, h: 0.32, fontSize: 11, color: C.text2, bold: true, valign: "middle" });
}
function userBubble(s, x, y, w, h, text) {
  s.addText(text, { shape: S.ROUNDED_RECTANGLE, rectRadius: 0.1, x, y, w, h, isTextBox: true, margin: [6, 10, 6, 10], fill: { color: C.accent1 }, color: C.background1, fontSize: 12, valign: "middle", objectName: nm("ububble") });
}
function aiBubble(s, x, y, w, h, toolName, text) {
  card(s, x, y, w, h, { fill: { color: C.background2 }, r: 0.1, name: "aibubble" });
  chip(s, x + 0.15, y + 0.13, Math.min(w - 0.3, 0.2 + toolName.length * 0.085), 0.28, toolName, { fill: C.accent2, transparency: 70, mono: true, fontSize: 10 });
  txt(s, text, { x: x + 0.15, y: y + 0.52, w: w - 0.3, h: h - 0.6, fontSize: 12 });
}

// 트리 목업 (Commander 화면 느낌)
async function treePanel(s, x, y, w, h, header, rows, fs = 12) {
  card(s, x, y, w, h, { fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, shadow: true, name: "tree" });
  s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h: 0.42, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: nm("treehead") });
  s.addShape(S.RECTANGLE, { x, y: y + 0.28, w, h: 0.14, fill: { color: C.background2 }, line: { type: "none" }, objectName: nm("treehead2") });
  txt(s, header, { x: x + 0.2, y: y + 0.06, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, color: C.text2, valign: "middle" });
  const rowH = fs <= 11 ? 0.34 : 0.38;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rx = x + 0.25 + r.lvl * 0.32;
    const ry = y + 0.6 + i * rowH;
    const iconName = { folder: "LuFolder", tc: "LuFileCheck", step: "LuListChecks", module: "LuLayers", note: "LuInfo" }[r.type] || "LuFolder";
    if (r.type !== "note") await iconImg(s, iconName, rx, ry + 0.06, 0.22, r.type === "tc" ? HEX.accent2 : HEX.accent1);
    const runs = [{ text: r.name, options: { bold: r.bold ?? false, color: r.type === "note" ? C.accent6 : C.text1 } }];
    if (r.detail) runs.push({ text: "   " + r.detail, options: { color: C.accent6 } });
    txt(s, runs, { x: rx + (r.type === "note" ? 0 : 0.3), y: ry, w: x + w - rx - 0.45, h: rowH, fontSize: fs, valign: "middle" });
  }
}

const SEC = { open: "오프닝", p1: "Part 1 Tosca MCP 이해", p2: "Part 2 실습", p3: "Part 3 활용과 프롬프트", p4: "Part 4 프롬프트 템플릿", end: "마무리" };

async function build() {
  // ================= 오프닝 =================
  pres.addSection({ title: SEC.open });
  {
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle: SEC.open });
    s.addText(K.meta.course, { placeholder: "title" });
    s.addText(K.meta.subtitle, { placeholder: "body" });
    s.addText(`${K.meta.baseline}   |   ${K.meta.duration}`, { placeholder: "meta" });
    // 오른쪽: AI 어시스턴트 - MCP - Commander 연결 그림
    const cy = 4.0, d = 1.25;
    const xs = [8.95, 10.45, 11.95];
    line(s, xs[0] + d / 2, cy + d / 2, xs[2] + d / 2, cy + d / 2, { color: C.accent3, width: 2 });
    const items = [["LuBot", "AI 어시스턴트"], ["LuCable", "MCP"], ["LuFolderTree", "Commander"]];
    for (let i = 0; i < 3; i++) {
      const dd = i === 1 ? 0.95 : d;
      const xx = xs[i] + (d - dd) / 2, yy = cy + (d - dd) / 2;
      s.addShape(S.OVAL, { x: xx, y: yy, w: dd, h: dd, fill: { color: i === 1 ? C.accent2 : C.background1 }, line: { type: "none" }, objectName: nm("titlecircle") });
      await iconImg(s, items[i][0], xx + dd * 0.25, yy + dd * 0.25, dd * 0.5, i === 1 ? HEX.dk1 : HEX.accent1);
      txt(s, items[i][1], { x: xs[i] - 0.3, y: cy + d + 0.12, w: d + 0.6, h: 0.35, fontSize: 12, color: C.background2, align: "center" });
    }
    s.addNotes("과정 소개. 오늘은 Tosca Commander에 들어 있는 MCP Server를 써서, AI 어시스턴트에게 말로 요청해 테스트케이스를 만드는 방법을 배운다. 기준 버전은 Commander 2026.1 LTS Patch 1 이상이고, Tricentis가 공개한 commander-mcp 스킬팩 1.2.1 내용을 기준으로 했다. Tosca를 많이 안 써 본 분도 따라올 수 있게 용어부터 짚고 간다.");
  }
  {
    const s = content("오늘 진행 순서", "전체 3시간 20분 중 절반 가까이가 실습이다. 이론은 실습에 필요한 만큼만 다룬다. Part 2 실습은 Lab 0부터 Lab 5까지 순서대로 진행하는데, 앞 Lab에서 만든 결과를 뒤 Lab에서 쓰기 때문에 건너뛰지 않는다. 휴식은 Part 2가 끝난 뒤 10분.", SEC.open);
    card(s, 0.6, 1.5, 3.6, 5.2, { fill: { color: C.background2 } });
    txt(s, "총 시간", { x: 0.95, y: 1.85, w: 3.0, h: 0.4, fontSize: 14, color: C.accent6 });
    txt(s, "3시간 20분", { x: 0.95, y: 2.25, w: 3.0, h: 0.8, fontSize: 36, bold: true, color: C.text2 });
    txt(s, "휴식 10분 포함", { x: 0.95, y: 3.05, w: 3.0, h: 0.4, fontSize: 14, color: C.accent6 });
    txt(s, "그중 실습", { x: 0.95, y: 4.15, w: 3.0, h: 0.4, fontSize: 14, color: C.accent6 });
    txt(s, "90분", { x: 0.95, y: 4.55, w: 3.0, h: 0.8, fontSize: 36, bold: true, color: C.text2 });
    txt(s, "1인 1PC, 개인 워크스페이스", { x: 0.95, y: 5.35, w: 3.0, h: 0.4, fontSize: 14, color: C.accent6 });
    const y0 = 1.55, rh = 0.72, gap = 0.13;
    K.agenda.forEach((a, i) => {
      const y = y0 + i * (rh + gap);
      const hot = a.no === "02";
      card(s, 4.6, y, 8.13, rh, { fill: hot ? { color: C.accent2, transparency: 80 } : { color: C.background1 }, line: hot ? { type: "none" } : { color: C.background2, width: 1 } });
      badge(s, 4.8, y + 0.11, 0.5, a.no, { fontSize: 13 });
      txt(s, a.title, { x: 5.55, y, w: 5.2, h: rh, fontSize: 18, bold: hot, valign: "middle" });
      if (hot) chip(s, 10.15, y + 0.2, 0.85, 0.32, "실습", { fill: C.accent1, color: C.background1, bold: true });
      txt(s, a.time, { x: 11.2, y, w: 1.35, h: rh, fontSize: 16, color: C.accent6, align: "right", valign: "middle" });
    });
  }
  {
    const s = content("오늘 쓰는 Tosca 용어 8개", "AI가 돌려주는 답에는 Tosca 용어가 그대로 나온다. 용어를 모르면 결과가 맞는지 판단할 수 없다. 8개만 기억하면 오늘 실습은 충분하다. 특히 ActionMode의 Input과 Verify 차이, ScratchBook과 ExecutionList 차이는 실습에서 계속 나오니 짚어 준다.", SEC.open);
    const w = 2.86, h = 2.45, gx = 0.23, gy = 0.22;
    K.terms.forEach((t, i) => {
      const col = i % 4, row = Math.floor(i / 4);
      const x = 0.6 + col * (w + gx), y = 1.5 + row * (h + gy);
      card(s, x, y, w, h, { fill: { color: C.background2 } });
      txt(s, t.term, { x: x + 0.22, y: y + 0.2, w: w - 0.44, h: 0.42, fontSize: 17, bold: true, color: C.text2 });
      txt(s, t.ko, { x: x + 0.22, y: y + 0.6, w: w - 0.44, h: 0.3, fontSize: 12, color: C.accent6 });
      txt(s, t.desc, { x: x + 0.22, y: y + 0.98, w: w - 0.44, h: h - 1.1, fontSize: 13 });
    });
  }
  {
    const s = content("Tosca에서 테스트가 만들어지는 흐름", "Tosca 테스트는 화면을 스캔해 모듈을 만들고, 모듈을 조합해 테스트케이스를 만들고, 실행해서 결과를 보는 순서로 만들어진다. MCP로 AI가 도와줄 수 있는 구간은 2번부터 5번까지다. 1번 화면 스캔은 XScan 화면에서 사람이 해야 한다. 그래서 실습용 모듈은 미리 준비해 두었다.", SEC.open);
    const steps = [
      ["1", "화면 스캔", "XScan으로 테스트할 화면을 읽는다"],
      ["2", "모듈", "화면 요소 정보가 모듈로 저장된다"],
      ["3", "테스트케이스", "모듈을 순서대로 놓고 값을 넣는다"],
      ["4", "실행", "ScratchBook에서 시험, ExecutionList로 정식 실행"],
      ["5", "결과 분석", "실패한 스텝과 원인을 확인한다"],
    ];
    const w = 2.12, gap = 0.38, y = 1.65, h = 2.0;
    steps.forEach((st, i) => {
      const x = 0.6 + i * (w + gap);
      card(s, x, y, w, h, { fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, shadow: true });
      badge(s, x + 0.2, y + 0.2, 0.45, st[0]);
      txt(s, st[1], { x: x + 0.2, y: y + 0.75, w: w - 0.4, h: 0.4, fontSize: 17, bold: true });
      txt(s, st[2], { x: x + 0.2, y: y + 1.15, w: w - 0.4, h: 0.8, fontSize: 13, color: C.accent6 });
      if (i < 4) line(s, x + w + 0.05, y + h / 2, x + w + gap - 0.05, y + h / 2, { arrow: true, color: C.accent3, width: 1.5 });
    });
    // 담당 구간
    chip(s, 0.6, 3.95, 2.12, 0.5, "사람이 직접", { fill: C.background2, color: C.text1, bold: true, fontSize: 13 });
    chip(s, 0.6 + w + gap, 3.95, 12.73 - (0.6 + w + gap), 0.5, "AI 어시스턴트 + MCP로 요청 가능", { fill: C.accent2, color: C.text1, bold: true, fontSize: 13, transparency: 20 });
    card(s, 0.6, 4.8, 5.9, 1.7, { fill: { color: C.background2 } });
    await iconImg(s, "LuScanLine", 0.85, 5.05, 0.45, HEX.accent1);
    txt(s, "새 화면은 사람이 스캔", { x: 1.45, y: 5.0, w: 4.8, h: 0.4, fontSize: 16, bold: true });
    txt(s, "MCP는 이미 있는 모듈만 재사용한다. 새 화면이나 바뀐 화면은 XScan에서 직접 스캔해야 한다.", { x: 1.45, y: 5.42, w: 4.85, h: 0.95, fontSize: 14 });
    card(s, 6.83, 4.8, 5.9, 1.7, { fill: { color: C.accent2, transparency: 86 } });
    await iconImg(s, "LuMessageSquare", 7.08, 5.05, 0.45, HEX.accent1);
    txt(s, "나머지는 말로 요청", { x: 7.68, y: 5.0, w: 4.8, h: 0.4, fontSize: 16, bold: true });
    txt(s, "모듈 찾기, 테스트케이스 작성, 검증 실행, 실패 원인 분석을 문장으로 요청할 수 있다.", { x: 7.68, y: 5.42, w: 4.85, h: 0.95, fontSize: 14 });
  }

  // ================= Part 1 =================
  pres.addSection({ title: SEC.p1 });
  {
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: SEC.p1 });
    s.addText("01", { placeholder: "num" });
    s.addText("Tosca MCP 이해", { placeholder: "title" });
    s.addText("MCP가 무엇인지, Commander에서 AI가 무엇을 할 수 있고 무엇은 못 하는지", { placeholder: "body" });
    s.addNotes("Part 1은 25분. MCP 자체를 깊게 파지는 않는다. 오늘 실습을 이해하는 데 필요한 정도, 그리고 AI에게 맡기면 안 되는 일이 무엇인지까지가 목표다.");
  }
  {
    const s = content("MCP는 AI와 프로그램을 잇는 표준 규격", "MCP는 Model Context Protocol의 약자다. AI 앱이 외부 프로그램의 기능을 도구처럼 불러 쓸 수 있게 해 주는 공개 규격이다. 예전에는 AI 앱과 프로그램을 잇는 연동을 조합마다 따로 만들어야 했다. MCP가 생긴 뒤로는 각자 규격 하나만 맞추면 된다. 그래서 USB-C에 자주 비유한다. 2024년 11월 Anthropic이 공개했고, 2025년 12월 Linux Foundation 산하 Agentic AI Foundation으로 넘어갔다. Tricentis도 이 규격으로 Tosca용 서버를 만들었다.", SEC.p1);
    const bullets = [
      ["AI가 외부 프로그램을 도구로 쓰게 하는 규격", "Model Context Protocol의 약자. 누구나 쓸 수 있는 공개 규격"],
      ["2024년 11월 Anthropic 공개", "2025년 12월 Linux Foundation 산하 Agentic AI Foundation으로 이관"],
      ["Claude, GitHub Copilot, Cursor 등이 지원", "AI 앱 하나에 여러 MCP 서버를 붙일 수 있다"],
      ["흔히 'AI용 USB-C'에 비유", "규격 하나만 맞추면 어느 쪽이든 연결된다"],
    ];
    bullets.forEach((b, i) => {
      const y = 1.6 + i * 1.22;
      badge(s, 0.6, y + 0.05, 0.42, String(i + 1), { fontSize: 13 });
      txt(s, b[0], { x: 1.2, y, w: 5.2, h: 0.5, fontSize: 16, bold: true, valign: "middle" });
      txt(s, b[1], { x: 1.2, y: y + 0.5, w: 5.2, h: 0.55, fontSize: 13, color: C.accent6 });
    });
    // 오른쪽 비교 그림
    const panels = [{ x: 6.75, title: "규격이 없을 때", hub: false }, { x: 9.8, title: "MCP를 쓸 때", hub: true }];
    for (const p of panels) {
      card(s, p.x, 1.55, 2.93, 4.95, { fill: { color: p.hub ? C.background2 : C.background1 }, line: p.hub ? { type: "none" } : { color: C.background2, width: 1.25 } });
      txt(s, p.title, { x: p.x + 0.2, y: 1.7, w: 2.5, h: 0.4, fontSize: 14, bold: true, color: p.hub ? C.text2 : C.accent6 });
      const lx = p.x + 0.2, rx = p.x + 2.03, bw = 0.7, bh = 0.5;
      const ys = [2.55, 3.75, 4.95];
      const left = ["AI 앱", "AI 앱", "AI 앱"], right = ["도구", "도구", "도구"];
      if (!p.hub) {
        for (const a of ys) for (const b of ys) line(s, lx + bw, a + bh / 2, rx, b + bh / 2, { color: C.accent5, width: 1 });
      } else {
        const hx = p.x + 1.06, hy = 3.6, hw = 0.8, hh = 0.8;
        for (const a of ys) { line(s, lx + bw, a + bh / 2, hx, hy + hh / 2, { color: C.accent1, width: 1.25 }); line(s, hx + hw, hy + hh / 2, rx, a + bh / 2, { color: C.accent1, width: 1.25 }); }
        chip(s, hx, hy, hw, hh, "MCP", { fill: C.accent2, bold: true, fontSize: 13 });
      }
      ys.forEach((yy, i) => {
        chip(s, lx, yy, bw, bh, left[i], { fill: C.accent1, color: C.background1, fontSize: 11 });
        chip(s, rx, yy, bw, bh, right[i], { fill: C.background1, line: { color: C.accent1, width: 1 }, fontSize: 11 });
      });
      txt(s, p.hub ? "연결 6개, 규격은 하나" : "조합마다 따로 연동 9개", { x: p.x + 0.2, y: 5.85, w: 2.55, h: 0.45, fontSize: 12, color: p.hub ? C.text2 : C.accent5, bold: true });
    }
  }
  {
    const s = content("MCP를 이루는 세 가지", "Host는 우리가 직접 쓰는 AI 앱이다. VS Code, Cursor, Claude Desktop 같은 것. Client는 Host 안에 들어 있어서 MCP 서버와 통신을 맡는다. 사용자가 따로 신경 쓸 일은 없다. Server는 기능을 도구 형태로 내놓는 쪽이고, 오늘은 Tosca MCP Server가 여기에 해당한다. 아래 다섯 단계 흐름을 실습 내내 보게 된다. 특히 2번, AI가 어떤 도구를 쓸지 스스로 고른다는 점이 중요하다. 그래서 프롬프트가 구체적이어야 한다.", SEC.p1);
    // 구조도
    card(s, 0.6, 1.55, 5.3, 2.6, { fill: { color: C.background2 } });
    txt(s, "Host  (사용자가 쓰는 AI 앱)", { x: 0.85, y: 1.7, w: 4.8, h: 0.4, fontSize: 14, bold: true, color: C.text2 });
    txt(s, "VS Code + Copilot, Cursor, Claude Desktop, Claude Code", { x: 0.85, y: 2.1, w: 4.8, h: 0.4, fontSize: 12, color: C.accent6 });
    card(s, 0.85, 2.7, 4.8, 1.2, { fill: { color: C.background1 }, line: { color: C.accent3, width: 1 } });
    txt(s, "Client", { x: 1.05, y: 2.82, w: 2, h: 0.35, fontSize: 14, bold: true });
    txt(s, "Host 안에서 서버와 통신을 맡는다. 사용자가 따로 다룰 일은 없다.", { x: 1.05, y: 3.18, w: 4.4, h: 0.65, fontSize: 12, color: C.accent6 });
    line(s, 5.95, 2.85, 7.25, 2.85, { arrow: true, color: C.accent1, width: 1.75 });
    txt(s, "도구 호출", { x: 6.0, y: 2.45, w: 1.2, h: 0.35, fontSize: 11, color: C.accent6, align: "center" });
    card(s, 7.3, 1.55, 5.43, 2.6, { fill: { color: C.accent2, transparency: 84 } });
    txt(s, "Server  (기능을 도구로 제공)", { x: 7.55, y: 1.7, w: 5, h: 0.4, fontSize: 14, bold: true, color: C.text2 });
    txt(s, "오늘은 Tosca MCP Server", { x: 7.55, y: 2.1, w: 5, h: 0.4, fontSize: 12, color: C.accent6 });
    const tools = ["get_workspace_info", "list_assets", "create_test_cases", "execute_test_suite"];
    tools.forEach((t, i) => chip(s, 7.55 + (i % 2) * 2.55, 2.7 + Math.floor(i / 2) * 0.6, 2.4, 0.45, t, { fill: C.background1, mono: true, fontSize: 11 }));
    // 호출 흐름
    sectionHead(s, 0.6, 4.45, 6, "요청 하나가 처리되는 순서");
    const flow = [["사람", "문장으로 요청"], ["AI", "쓸 도구를 고름"], ["Client", "서버에 도구 호출"], ["Server", "Commander에서 실행 후 결과 반환"], ["AI", "결과를 정리해 답변"]];
    const fw = 2.3, fg = 0.158;
    flow.forEach((f, i) => {
      const x = 0.6 + i * (fw + fg);
      card(s, x, 4.95, fw, 1.6, { fill: { color: C.background1 }, line: { color: C.background2, width: 1.25 } });
      badge(s, x + 0.18, 5.12, 0.4, String(i + 1), { fontSize: 12, fill: i === 1 ? C.accent2 : C.accent1, color: i === 1 ? C.text1 : C.background1 });
      txt(s, f[0], { x: x + 0.68, y: 5.1, w: fw - 0.8, h: 0.45, fontSize: 14, bold: true, valign: "middle" });
      txt(s, f[1], { x: x + 0.18, y: 5.65, w: fw - 0.36, h: 0.8, fontSize: 13 });
    });
  }
  {
    const s = content("Tosca MCP Server는 Commander 안에서 동작한다", "Tosca MCP Server는 별도 프로그램이 아니라 Commander에 들어 있는 애드인이다. 2026.1 LTS Patch 1부터 들어갔고, Patch 2부터는 기본으로 켜져 있다. Commander가 실행되는 동안 내 PC 안에서만 접속되는 주소로 열린다. 그래서 Commander와 워크스페이스가 열려 있어야만 쓸 수 있다. AI는 워크스페이스를 연 사람의 권한으로 움직이고, 멀티 유저라면 체크아웃 체크인 규칙도 똑같이 지켜야 한다. 한 가지 더. AI 어시스턴트가 쓰는 모델은 대부분 클라우드에서 돌기 때문에, 요청과 결과 일부가 모델 쪽으로 전달된다. 사내 AI 사용 정책을 확인하고 쓰자.", SEC.p1);
    // 클라우드 모델
    card(s, 0.9, 1.5, 2.4, 0.72, { fill: { color: C.background1 }, line: { color: C.accent6, width: 1, dashType: "dash" } });
    await iconImg(s, "LuCloud", 1.05, 1.64, 0.42, HEX.accent6);
    txt(s, "AI 모델 (클라우드)", { x: 1.55, y: 1.5, w: 1.75, h: 0.72, fontSize: 12, bold: true, color: C.accent6, valign: "middle" });
    txt(s, "요청과 결과 일부가 모델로 전달된다. 사내 AI 사용 정책을 확인한다.", { x: 3.5, y: 1.5, w: 4.3, h: 0.72, fontSize: 12, color: C.accent6, valign: "middle" });
    line(s, 2.1, 2.22, 2.1, 3.3, { dash: "dash", color: C.accent6 });
    // 내 PC 경계
    s.addShape(S.ROUNDED_RECTANGLE, { x: 0.6, y: 2.55, w: 7.25, h: 4.0, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: C.accent3, width: 1.25, dashType: "dash" }, objectName: nm("pc") });
    await iconImg(s, "LuLaptop", 0.8, 2.68, 0.3, HEX.accent3);
    txt(s, "내 PC", { x: 1.18, y: 2.66, w: 1.5, h: 0.34, fontSize: 12, bold: true, color: C.accent3, valign: "middle" });
    card(s, 0.9, 3.3, 2.4, 1.85, { fill: { color: C.background2 }, shadow: true });
    await iconImg(s, "LuBot", 1.1, 3.48, 0.45, HEX.accent1);
    txt(s, "AI 어시스턴트", { x: 1.1, y: 4.0, w: 2.1, h: 0.4, fontSize: 15, bold: true });
    txt(s, "VS Code · Cursor · Claude", { x: 1.1, y: 4.42, w: 2.1, h: 0.55, fontSize: 11, color: C.accent6 });
    card(s, 4.25, 2.95, 3.35, 3.35, { fill: { color: C.background2 }, shadow: true });
    txt(s, "Tosca Commander 2026.1", { x: 4.45, y: 3.05, w: 3.0, h: 0.4, fontSize: 13, bold: true, color: C.text2 });
    card(s, 4.45, 3.5, 2.95, 0.85, { fill: { color: C.accent2, transparency: 55 } });
    txt(s, "Tosca MCP Server Addin", { x: 4.6, y: 3.5, w: 2.7, h: 0.85, fontSize: 13, bold: true, valign: "middle" });
    line(s, 5.92, 4.38, 5.92, 4.72, { arrow: true, color: C.accent1, width: 1.5 });
    card(s, 4.45, 4.75, 2.95, 1.35, { fill: { color: C.background1 }, line: { color: C.accent3, width: 1 } });
    txt(s, "열린 워크스페이스", { x: 4.6, y: 4.82, w: 2.7, h: 0.38, fontSize: 13, bold: true });
    txt(s, "싱글 유저: 내 PC 파일\n멀티 유저: 공통 리포지토리", { x: 4.6, y: 5.22, w: 2.7, h: 0.75, fontSize: 11, color: C.accent6 });
    line(s, 3.32, 4.22, 4.22, 4.22, { arrow: true, color: C.accent1, width: 1.75 });
    line(s, 4.22, 4.42, 3.32, 4.42, { arrow: true, color: C.accent1, width: 1.75 });
    txt(s, "HTTP", { x: 3.3, y: 3.82, w: 0.95, h: 0.32, fontSize: 11, bold: true, color: C.text2, align: "center" });
    txt(s, MONOURL(), { x: 0.9, y: 5.75, w: 3.2, h: 0.5, fontSize: 11, fontFace: MONO, color: C.text2, valign: "middle" });
    // 오른쪽 사실
    const facts = [
      ["LuPackageCheck", "2026.1 LTS Patch 1부터 제공", "Patch 2부터 애드인이 기본으로 켜져 있다"],
      ["LuMonitor", "Commander가 열려 있는 동안만", "Commander나 워크스페이스를 닫으면 연결도 끊긴다"],
      ["LuUser", "내 권한 그대로", "워크스페이스를 연 사용자 권한으로 움직이고, 체크인·체크아웃 규칙도 같다"],
      ["LuCloudOff", "Tosca Cloud MCP와는 별개", "Cloud용은 원격 서버에 로그인해서 쓰는 다른 제품"],
    ];
    for (let i = 0; i < facts.length; i++) {
      const y = 1.55 + i * 1.27;
      await iconCircle(s, facts[i][0], 8.25, y, 0.55);
      txt(s, facts[i][1], { x: 9.0, y: y - 0.02, w: 3.73, h: 0.4, fontSize: 15, bold: true });
      txt(s, facts[i][2], { x: 9.0, y: y + 0.38, w: 3.73, h: 0.8, fontSize: 13, color: C.accent6 });
    }
  }
  {
    const s = content("같은 일을 할 때 무엇이 달라지나", "로그인 테스트케이스 하나를 만든다고 해 보자. 직접 하면 폴더 만들고, 모듈을 끌어다 놓고, 값을 하나씩 넣고, ActionMode를 바꾸고, 실행하고, 고치는 일을 전부 클릭으로 한다. MCP를 쓰면 사람이 하는 일은 요청하고, 계획을 확인하고, 결과를 검토하는 세 가지로 줄어든다. 단, 검토를 건너뛰면 안 된다. AI가 고른 모듈이나 값이 틀릴 수 있다.", SEC.p1);
    const L = ["폴더 만들기", "테스트케이스 만들고 이름 바꾸기", "모듈 5개를 순서대로 끌어다 놓기", "값 8개 입력", "검증 스텝의 ActionMode를 Verify로 변경", "ScratchBook에서 실행", "결과 확인 후 틀린 곳 수정"];
    const R = [["사람", "시나리오와 위치를 문장으로 요청"], ["AI", "모듈을 찾아 스텝 계획을 보여줌"], ["사람", "계획 확인 후 승인"], ["AI", "생성, 검증 실행, 결과 보고"], ["사람", "Commander에서 결과 검토"]];
    card(s, 0.6, 1.5, 5.9, 4.75, { fill: { color: C.background1 }, line: { color: C.background2, width: 1.5 } });
    txt(s, "Commander에서 직접", { x: 0.9, y: 1.68, w: 5, h: 0.45, fontSize: 18, bold: true, color: C.accent6 });
    L.forEach((t, i) => {
      const y = 2.3 + i * 0.54;
      badge(s, 0.9, y + 0.06, 0.36, String(i + 1), { fill: C.background2, color: C.text1, fontSize: 11 });
      txt(s, t, { x: 1.42, y, w: 4.9, h: 0.48, fontSize: 15, valign: "middle" });
    });
    card(s, 6.83, 1.5, 5.9, 4.75, { fill: { color: C.background2 } });
    txt(s, "MCP로 요청", { x: 7.13, y: 1.68, w: 5, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    R.forEach((t, i) => {
      const y = 2.3 + i * 0.76;
      const ai = t[0] === "AI";
      chip(s, 7.13, y + 0.1, 0.75, 0.4, t[0], { fill: ai ? C.accent2 : C.accent1, color: ai ? C.text1 : C.background1, bold: true, fontSize: 12 });
      txt(s, t[1], { x: 8.05, y, w: 4.5, h: 0.6, fontSize: 15, valign: "middle", bold: !ai });
    });
    txt(s, "반복 클릭은 AI가, 판단과 확인은 사람이 한다", { x: 0.6, y: 6.4, w: 12.13, h: 0.45, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
  }
  {
    const s = content("도구 32개, 여섯 묶음", "Tosca MCP Server가 내놓는 도구는 스킬팩 1.2.1 기준 32개다. 외울 필요는 없다. AI가 알아서 고른다. 다만 어떤 묶음이 있는지 알면 무엇을 시킬 수 있는지 감이 온다. 실습에서 주로 쓰는 것은 조회, 작성, 실행 묶음이다. 실제 목록은 Commander 패치에 따라 조금 다를 수 있으니, Lab 0에서 AI에게 도구 목록을 직접 물어본다.", SEC.p1);
    const w = 3.9, h = 2.4, gx = 0.215, gy = 0.2;
    const iconMap = { read: "LuSearch", edit: "LuPencil", author: "LuHammer", run: "LuPlay", api: "LuPlug", db: "LuDatabase" };
    for (let i = 0; i < K.toolGroups.length; i++) {
      const g = K.toolGroups[i];
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * (w + gx), y = 1.45 + row * (h + gy);
      card(s, x, y, w, h, { fill: { color: C.background2 } });
      await iconCircle(s, iconMap[g.icon] || iconMap[g.key], x + 0.22, y + 0.22, 0.6);
      txt(s, g.name, { x: x + 0.98, y: y + 0.22, w: 2.0, h: 0.6, fontSize: 19, bold: true, valign: "middle" });
      chip(s, x + w - 0.95, y + 0.34, 0.72, 0.36, `${g.tools.length}개`, { fill: C.accent2, bold: true, fontSize: 12 });
      txt(s, g.summary, { x: x + 0.22, y: y + 0.98, w: w - 0.44, h: 0.65, fontSize: 14 });
      txt(s, g.tools.slice(0, 3).map((t) => t[0]).join("\n"), { x: x + 0.22, y: y + 1.6, w: w - 0.44, h: 0.72, fontSize: 10, fontFace: MONO, color: C.accent6 });
    }
    txt(s, "commander-mcp 스킬팩 1.2.1 도구 카탈로그 기준. 전체 목록과 설명은 워크북 부록 A.", { x: 0.6, y: 6.5, w: 12.13, h: 0.35, fontSize: 11, color: C.accent6, valign: "middle" });
  }
  {
    const s = content("AI에게 맡기지 않는 일", "할 수 있는 것만큼 중요한 게 맡기면 안 되는 것이다. 화면 스캔은 AI가 못 한다. 기존 모듈만 쓴다. 비밀번호 같은 민감 정보는 사람이 넣는다. Data Integrity 연결을 AI가 만들면 Allow AI queries 옵션이 꺼진 채로 만들어지고, 켜는 건 사람이 Connection Manager에서 한다. 삭제는 되돌릴 수 없으니 목록 확인 후 따로 승인한다. 실패가 진짜 결함인지 판단하는 것도 사람 몫이다. 테스트를 고쳐서 결함을 덮으면 안 된다.", SEC.p1);
    const items = K.notForAI;
    const w = 5.95, h = 1.55, gx = 0.23, gy = 0.2;
    for (let i = 0; i < items.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.6 + col * (w + gx), y = 1.5 + row * (h + gy);
      card(s, x, y, w, h, { fill: { color: C.background2 } });
      await iconCircle(s, items[i][0], x + 0.25, y + 0.42, 0.7, { fill: C.background1, iconHex: HEX.accent5 });
      txt(s, items[i][1], { x: x + 1.18, y: y + 0.2, w: w - 1.4, h: 0.45, fontSize: 17, bold: true, valign: "middle" });
      txt(s, items[i][2], { x: x + 1.18, y: y + 0.66, w: w - 1.4, h: 0.8, fontSize: 14 });
    }
  }
  {
    const s = content("준비물과 연결 설정", "실습 전에 이 네 가지가 준비돼 있어야 한다. Commander는 2026.1 LTS Patch 1 이상이어야 하고, Patch 1이라면 MCP Server 애드인을 직접 켜야 한다. AI 클라이언트는 Tricentis가 테스트했다고 밝힌 네 가지 중 하나를 쓴다. 설정은 서버 주소 한 줄이 전부다. 오른쪽은 VS Code 예시이고, 나머지 클라이언트 설정은 워크북 2장에 있다. 스킬팩은 필수는 아니지만 Tricentis도 함께 쓰기를 권한다. 설치하면 AI가 체크아웃, 저장 같은 규칙을 알아서 지킨다.", SEC.p1);
    const items = [
      ["Tosca Commander 2026.1 LTS Patch 1 이상", "Patch 2 이상이면 MCP Server 애드인이 기본으로 켜져 있다"],
      ["실습 워크스페이스 열기", "실습 내내 열어 둔다. 닫으면 연결이 끊긴다"],
      ["AI 클라이언트 하나", "VS Code + GitHub Copilot, Cursor, Claude Code, Claude Desktop"],
      ["commander-mcp 스킬팩 (권장)", "github.com/Tricentis/mcp-skills 의 Tosca/Commander/MCP"],
    ];
    for (let i = 0; i < items.length; i++) {
      const y = 1.6 + i * 1.2;
      await iconCircle(s, "LuCheck", 0.6, y + 0.04, 0.48, { fill: C.accent4 });
      txt(s, items[i][0], { x: 1.3, y, w: 5.3, h: 0.45, fontSize: 16, bold: true, valign: "middle" });
      txt(s, items[i][1], { x: 1.3, y: y + 0.46, w: 5.3, h: 0.6, fontSize: 13, color: C.accent6 });
    }
    txt(s, "VS Code 예시  ·  .vscode/mcp.json", { x: 7.1, y: 1.5, w: 5.6, h: 0.38, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    codeBox(s, 7.1, 1.95, 5.63, 2.45, K.clients[0].code, 13);
    card(s, 7.1, 4.65, 5.63, 1.45, { fill: { color: C.accent2, transparency: 86 } });
    txt(s, [
      { text: "포트 46248", options: { bold: true, breakLine: true } },
      { text: "기본값이다. 디버그 빌드는 8080, Commander MCP 설정에서 바꿨다면 그 값을 쓴다. Cursor, Claude 설정은 워크북 2장.", options: {} },
    ], { x: 7.35, y: 4.8, w: 5.15, h: 1.2, fontSize: 13 });
  }

  // ================= Part 2 =================
  pres.addSection({ title: SEC.p2 });
  {
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: SEC.p2 });
    s.addText("02", { placeholder: "num" });
    s.addText("실습: 테스트케이스 만들기", { placeholder: "title" });
    s.addText("Lab 0 연결 확인부터 Lab 5 저장까지, 90분", { placeholder: "body" });
    s.addNotes("이제 실습. 워크북 3장을 펴게 한다. 각 Lab마다 입력할 프롬프트가 워크북에 그대로 있으니 복사해서 쓰면 된다. [내 이름] 같은 대괄호 부분만 바꾼다. 막히면 Lab 끝의 '막히면' 칸부터 보게 한다.");
  }
  {
    const s = content("실습 시나리오: Demo Web Shop 로그인", "실습 대상은 Tricentis가 공개한 연습용 쇼핑몰 Demo Web Shop이다. 시나리오는 로그인하고, 계정 이메일이 맞게 보이는지 검증하고, 로그아웃하는 것. 모듈은 미리 XScan으로 스캔해서 워크스페이스에 넣어 두었다. MCP로는 모듈을 만들 수 없기 때문이다. 실습 계정은 따로 나눠 준다. 사내 앱으로 바꿔도 흐름은 똑같다.", SEC.p2);
    card(s, 0.6, 1.5, 5.6, 5.0, { fill: { color: C.background2 } });
    await iconCircle(s, "LuShoppingCart", 0.9, 1.8, 0.7);
    txt(s, "Tricentis Demo Web Shop", { x: 1.8, y: 1.8, w: 4.2, h: 0.4, fontSize: 18, bold: true });
    txt(s, K.meta.app, { x: 1.8, y: 2.2, w: 4.2, h: 0.35, fontSize: 12, fontFace: MONO, color: C.text2 });
    txt(s, "테스트 연습용으로 공개된 쇼핑몰 사이트", { x: 0.9, y: 2.85, w: 5.0, h: 0.4, fontSize: 14, color: C.accent6 });
    sectionHead(s, 0.9, 3.45, 5, "시나리오");
    ["로그인 화면에서 실습 계정으로 로그인", "상단에 계정 이메일이 맞게 보이는지 검증", "로그아웃 후 브라우저 닫기"].forEach((t, i) => {
      badge(s, 0.9, 3.95 + i * 0.62, 0.4, String(i + 1), { fontSize: 12 });
      txt(s, t, { x: 1.45, y: 3.93 + i * 0.62, w: 4.5, h: 0.45, fontSize: 15, valign: "middle" });
    });
    txt(s, "사내 앱으로 바꿔도 진행 순서는 같다.", { x: 0.9, y: 5.85, w: 5.0, h: 0.4, fontSize: 12, color: C.accent6 });
    sectionHead(s, 6.6, 1.5, 6, "미리 준비해 둔 모듈");
    const rows = [[
      { text: "모듈", options: { bold: true, color: C.background1, fill: { color: C.accent1 } } },
      { text: "위치", options: { bold: true, color: C.background1, fill: { color: C.accent1 } } },
      { text: "주요 속성", options: { bold: true, color: C.background1, fill: { color: C.accent1 } } },
    ]];
    K.practiceModules.forEach((m, i) => rows.push([
      { text: m.module, options: { bold: true, fill: { color: i % 2 ? C.background1 : C.background2 } } },
      { text: m.where, options: { fill: { color: i % 2 ? C.background1 : C.background2 } } },
      { text: m.attrs, options: { fill: { color: i % 2 ? C.background1 : C.background2 } } },
    ]));
    s.addTable(rows, { x: 6.6, y: 1.95, w: 6.13, colW: [1.55, 1.85, 2.73], rowH: 0.55, fontSize: 13, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: C.background2 }, margin: [0.04, 0.08, 0.04, 0.08], objectName: nm("table") });
    card(s, 6.6, 5.0, 6.13, 1.5, { fill: { color: C.accent2, transparency: 86 } });
    await iconImg(s, "LuInfo", 6.85, 5.2, 0.4, HEX.accent1);
    txt(s, "Header, Login은 강사가 XScan으로 미리 스캔했다. MCP는 모듈을 새로 만들지 않으므로, 없는 화면은 사람이 먼저 스캔해야 한다.", { x: 7.4, y: 5.15, w: 5.15, h: 1.2, fontSize: 14 });
  }
  {
    const s = content("실습 진행 순서", "되돌리기 쉬운 작업부터 한다. Lab 0과 Lab 1은 조회만 하므로 아무것도 바뀌지 않는다. Lab 2와 Lab 3에서 테스트케이스를 만들고, Lab 4에서 값을 바꿔 가며 실행한다. Lab 5에서 저장한다. 시간은 대략이고, 빨리 끝난 사람은 워크북의 추가 과제를 한다.", SEC.p2);
    const nodes = [
      ["Lab 0", "연결 확인", "10분", "없음", C.background2],
      ["Lab 1", "둘러보기", "10분", "없음", C.background2],
      ["Lab 2", "수동 TC", "15분", "생성", C.accent3],
      ["Lab 3", "자동 TC", "25분", "생성", C.accent3],
      ["Lab 4", "검증 실행과 수정", "20분", "수정·실행", C.accent2],
      ["Lab 5", "저장", "10분", "저장", C.accent4],
    ];
    const cx0 = 1.35, step = 2.13, ly = 2.85;
    line(s, cx0, ly, cx0 + step * 5, ly, { color: C.accent3, width: 2.5 });
    nodes.forEach((n, i) => {
      const cx = cx0 + i * step;
      chip(s, cx - 0.5, 1.6, 1.0, 0.38, n[2], { fill: C.background2, color: C.text2, bold: true, fontSize: 12 });
      badge(s, cx - 0.42, ly - 0.42, 0.84, String(i), { fontSize: 20 });
      txt(s, n[0], { x: cx - 1.0, y: 3.5, w: 2.0, h: 0.38, fontSize: 13, color: C.accent6, align: "center" });
      txt(s, n[1], { x: cx - 1.0, y: 3.88, w: 2.0, h: 0.75, fontSize: 16, bold: true, align: "center" });
      const light = n[4] === C.background2;
      chip(s, cx - 0.85, 4.75, 1.7, 0.42, n[3], { fill: n[4], color: light || n[4] === C.accent2 ? C.text1 : C.background1, bold: true, fontSize: 12 });
    });
    txt(s, "워크스페이스에서 바뀌는 것", { x: 0.6, y: 5.3, w: 4, h: 0.35, fontSize: 11, color: C.accent6 });
    card(s, 0.6, 5.8, 12.13, 0.8, { fill: { color: C.background2 } });
    await iconImg(s, "LuUndo2", 0.85, 5.98, 0.42, HEX.accent1);
    txt(s, "조회 → 생성 → 수정 → 저장 순서. 되돌리기 쉬운 작업부터 진행한다.", { x: 1.45, y: 5.8, w: 11, h: 0.8, fontSize: 16, bold: true, valign: "middle" });
  }
  {
    const L = K.labs[0];
    const s = content("Lab 0  연결 확인", "먼저 Commander를 켜고 실습 워크스페이스를 연다. 그다음 AI 클라이언트에서 tosca-commander 서버가 켜져 있는지 본다. 첫 번째 프롬프트로 도구 목록을 받아 보고, 두 번째로 워크스페이스 정보를 받는다. 오른쪽은 화면 예시다. 응답 위에 get_workspace_info 같은 도구 이름이 보이면 MCP로 실제 Commander를 읽은 것이다. 모드가 single인지 multi인지 기록해 둔다. Lab 5에서 저장 방법이 달라진다.", SEC.p2);
    labTag(s, L.id, L.time);
    sectionHead(s, 0.6, 1.45, 6, "준비");
    L.before.forEach((t, i) => {
      badge(s, 0.6, 1.92 + i * 0.5, 0.34, String(i + 1), { fontSize: 11 });
      txt(s, t, { x: 1.08, y: 1.86 + i * 0.5, w: 5.5, h: 0.46, fontSize: 13, valign: "middle" });
    });
    promptBox(s, 0.6, 3.5, 5.95, 1.35, `프롬프트 ${L.prompts[0].label}`, L.prompts[0].text, 13);
    promptBox(s, 0.6, 5.0, 5.95, 1.5, `프롬프트 ${L.prompts[1].label}`, L.prompts[1].text, 13);
    chatWindow(s, 6.9, 1.45, 5.83, 5.05, "AI 어시스턴트   tosca-commander 연결됨");
    userBubble(s, 8.25, 2.1, 4.3, 0.95, L.prompts[1].text);
    aiBubble(s, 7.1, 3.25, 5.0, 2.65, "get_workspace_info", "현재 열린 워크스페이스입니다.\n· 이름: MCP_Training_홍길동\n· 경로: C:\\Tosca_Projects\\Tosca_Workspaces\\MCP_Training_홍길동\n· 모드: single (싱글 유저)\n체크인 없이 저장만 하면 됩니다.");
    txt(s, "화면 예시. 실제 문구와 모양은 클라이언트마다 다르다.", { x: 7.1, y: 6.03, w: 5.4, h: 0.35, fontSize: 11, color: C.accent6, valign: "middle" });
  }
  {
    const L = K.labs[1];
    const s = content("Lab 1  워크스페이스 둘러보기", "이번 Lab은 조회만 한다. 첫 번째 프롬프트로 TestCases 구조를 보고, 두 번째로 실습 모듈과 속성을 확인한다. 결과에 긴 ID가 나와도 신경 쓰지 않는다. 사람은 경로로 이야기하면 된다. Login 모듈에 Email, Password, Log in이 있는지 꼭 확인한다. Lab 3에서 이 속성들을 쓴다.", SEC.p2);
    labTag(s, L.id, L.time);
    promptBox(s, 0.6, 1.5, 6.2, 1.45, `프롬프트 ${L.prompts[0].label}`, L.prompts[0].text, 13);
    promptBox(s, 0.6, 3.1, 6.2, 1.45, `프롬프트 ${L.prompts[1].label}`, L.prompts[1].text, 13);
    txt(s, "AI가 쓰는 도구", { x: 0.6, y: 4.75, w: 3, h: 0.35, fontSize: 12, color: C.accent6, valign: "middle" });
    L.tools.forEach((t, i) => chip(s, 0.6 + i * 2.05, 5.12, 1.95, 0.38, t, { fill: C.background2, mono: true, fontSize: 10 }));
    checkList(s, 0.6, 5.65, 6.2, L.checks.slice(0, 2), { fs: 13, rowH: 0.4 });
    await treePanel(s, 7.15, 1.5, 5.58, 5.0, "결과 예시  (프롬프트 1-2)", [
      { lvl: 0, type: "folder", name: "Modules", bold: true },
      { lvl: 1, type: "folder", name: "MCP실습" },
      { lvl: 2, type: "module", name: "Header", bold: true },
      { lvl: 3, type: "step", name: "Log in", detail: "Link" },
      { lvl: 3, type: "step", name: "Log out", detail: "Link" },
      { lvl: 3, type: "step", name: "Account", detail: "Link" },
      { lvl: 2, type: "module", name: "Login", bold: true },
      { lvl: 3, type: "step", name: "Email", detail: "TextBox" },
      { lvl: 3, type: "step", name: "Password", detail: "TextBox" },
      { lvl: 3, type: "step", name: "Log in", detail: "Button" },
      { lvl: 3, type: "step", name: "Error Message", detail: "Text" },
    ], 12);
  }
  {
    const L = K.labs[2];
    const s = content("Lab 2  수동 테스트케이스 만들기", "글로 된 수동 테스트 절차를 그대로 붙여 넣어 Tosca 수동 테스트케이스로 만든다. 프롬프트 마지막 줄이 중요하다. '만들기 전에 순서를 먼저 보여줘.' AI가 계획을 보여주면 읽어 보고, 맞으면 2-2로 승인한다. 결과는 Commander에서 직접 열어 확인한다. 스텝 3개에 지시와 기대 결과가 모두 들어갔는지 본다.", SEC.p2);
    labTag(s, L.id, L.time);
    promptBox(s, 0.6, 1.5, 6.95, 4.15, `프롬프트 ${L.prompts[0].label}`, L.prompts[0].text, 13);
    promptBox(s, 0.6, 5.8, 6.95, 0.95, `프롬프트 ${L.prompts[1].label}`, L.prompts[1].text, 12);
    await treePanel(s, 7.9, 1.5, 4.83, 2.75, "Commander에서 확인", [
      { lvl: 0, type: "folder", name: "TestCases / MCP실습 / 홍길동" },
      { lvl: 1, type: "tc", name: "로그인 실패_수동", bold: true },
      { lvl: 2, type: "step", name: "1. 사이트 접속" },
      { lvl: 2, type: "step", name: "2. Log in 링크 클릭" },
      { lvl: 2, type: "step", name: "3. 틀린 비밀번호로 로그인" },
    ], 12);
    sectionHead(s, 7.9, 4.45, 4.8, "확인 포인트");
    checkList(s, 7.9, 4.85, 4.83, L.checks.slice(1), { fs: 13, rowH: 0.5 });
  }
  {
    const L = K.labs[3];
    const s = content("Lab 3  자동 테스트케이스 (1) 계획 받기", "자동화 테스트케이스는 처음부터 만들게 하지 않는다. 먼저 계획만 받는다. 프롬프트에 모듈을 찾을 범위를 정해 주고, 새로 스캔하지 말라고 적는다. 오른쪽처럼 스텝마다 모듈, 속성, 값, ActionMode가 나온 표를 받으면 하나씩 확인한다. 4번 Account 스텝이 Verify인지, 못 찾은 모듈이 없는지가 핵심이다.", SEC.p2);
    labTag(s, L.id, L.time);
    promptBox(s, 0.6, 1.5, 6.35, 5.25, `프롬프트 ${L.prompts[0].label}`, L.prompts[0].text, 12.5);
    txt(s, "AI가 돌려주는 계획 (예시)", { x: 7.25, y: 1.5, w: 5.4, h: 0.38, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
    const head = ["#", "모듈", "속성", "값", "ActionMode"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 } } }));
    const rows = [head];
    K.lab3Steps.forEach((r) => {
      const verify = r[4] === "Verify";
      rows.push(r.map((c, j) => ({ text: j === 3 ? c.replace("https://", "") : c, options: { fill: { color: verify ? "FCEFD8" : C.background1 }, bold: verify && j === 4, fontFace: j === 3 && c.startsWith("{") ? MONO : undefined } })));
    });
    s.addTable(rows, { x: 7.25, y: 1.95, w: 5.48, colW: [0.35, 1.2, 0.9, 2.05, 0.98], rowH: 0.38, fontSize: 10.5, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.5, color: C.background2 }, margin: [0.03, 0.06, 0.03, 0.06], objectName: nm("table") });
    card(s, 7.25, 5.75, 5.48, 1.0, { fill: { color: C.background2 } });
    await iconImg(s, "LuCircleCheck", 7.45, 5.97, 0.5, HEX.accent4);
    txt(s, "맞는 모듈을 못 찾은 스텝: 없음\n확인했으면 프롬프트 3-2로 승인한다", { x: 8.1, y: 5.8, w: 4.5, h: 0.9, fontSize: 13, valign: "middle" });
  }
  {
    const L = K.labs[3];
    const s = content("Lab 3  자동 테스트케이스 (2) 승인하고 확인", "계획이 맞으면 3-2로 승인한다. AI가 만든 뒤에는 스텝 구조를 다시 읽어 달라고 해서 계획과 같은지 본다. 그리고 반드시 Commander에서 직접 열어 본다. 만약 모듈을 못 찾았다는 답이 오면, 그 화면은 사람이 XScan으로 스캔한 다음 같은 프롬프트로 다시 요청하면 된다. 비밀번호는 실습이라 그대로 넣지만, 실무에서는 사내 규칙대로 관리한다.", SEC.p2);
    labTag(s, L.id, L.time);
    promptBox(s, 0.6, 1.5, 6.0, 1.3, `프롬프트 ${L.prompts[1].label}`, L.prompts[1].text, 13);
    sectionHead(s, 0.6, 3.05, 6, "확인 포인트");
    checkList(s, 0.6, 3.45, 6.0, L.checks, { fs: 14, rowH: 0.5 });
    card(s, 0.6, 5.6, 6.0, 1.15, { fill: { color: C.accent2, transparency: 86 } });
    await iconImg(s, "LuTriangleAlert", 0.82, 5.85, 0.42, HEX.accent5);
    txt(s, "모듈이 없다는 답이 오면: 사람이 XScan으로 스캔 → 같은 프롬프트로 다시 요청", { x: 1.4, y: 5.65, w: 5.05, h: 1.05, fontSize: 14, valign: "middle" });
    await treePanel(s, 6.95, 1.5, 5.78, 5.25, "Commander에서 확인", [
      { lvl: 0, type: "folder", name: "TestCases / MCP실습 / 홍길동" },
      { lvl: 1, type: "tc", name: "로그인 성공_자동", bold: true },
      { lvl: 2, type: "step", name: "OpenUrl", detail: "Url = demowebshop.tricentis.com" },
      { lvl: 2, type: "step", name: "Header", detail: "Log in = {CLICK}" },
      { lvl: 2, type: "step", name: "Login", detail: "Email, Password, Log in" },
      { lvl: 2, type: "step", name: "Header", detail: "Account = [실습 이메일]  Verify" },
      { lvl: 2, type: "step", name: "Header", detail: "Log out = {CLICK}" },
      { lvl: 2, type: "step", name: "CloseBrowser", detail: "Title = Demo Web Shop*" },
      { lvl: 0, type: "note", name: "" },
      { lvl: 0, type: "note", name: "값과 ActionMode는 스텝을 펼쳐서 확인한다." },
    ], 12);
  }
  {
    const L = K.labs[4];
    const s = content("Lab 4  ScratchBook 검증 실행과 수정", "만든 테스트케이스를 ScratchBook에서 검증 실행한다. 브라우저가 열리면 끝날 때까지 손을 떼고 기다린다. 한 번 통과하면, 일부러 Verify 값을 틀리게 바꿔서 실패를 만든다. 그리고 AI에게 원인만 설명하라고 한다. 고치라고 바로 시키지 않는 게 포인트다. 원인이 기대값 불일치, 즉 테스트 데이터 문제라는 설명을 확인한 뒤 원래 값으로 되돌리고 다시 돌린다.", SEC.p2);
    labTag(s, L.id, L.time);
    const steps = [
      [L.prompts[0], "검증 실행", "통과", C.accent4],
      [L.prompts[1], "값을 틀리게 바꿔 실행", "실패", C.accent5],
      [L.prompts[2], "원인만 설명 듣기", "기대값 불일치", C.accent2],
      [L.prompts[3], "되돌리고 다시 실행", "통과", C.accent4],
    ];
    const w = 2.83, gap = 0.27, y = 1.55, h = 3.5;
    steps.forEach((st, i) => {
      const x = 0.6 + i * (w + gap);
      card(s, x, y, w, h, { fill: { color: C.background2 } });
      chip(s, x + 0.2, y + 0.2, 0.75, 0.34, st[0].label, { fill: C.accent2, bold: true, fontSize: 12 });
      txt(s, st[1], { x: x + 0.2, y: y + 0.65, w: w - 0.4, h: 0.75, fontSize: 16, bold: true });
      txt(s, st[0].text, { x: x + 0.2, y: y + 1.4, w: w - 0.4, h: 1.4, fontSize: 13, color: C.accent6 });
      chip(s, x + 0.2, y + h - 0.65, w - 0.4, 0.42, st[2], { fill: st[3], color: st[3] === C.accent2 ? C.text1 : C.background1, bold: true, fontSize: 13 });
      if (i < 3) line(s, x + w + 0.03, y + h / 2, x + w + gap - 0.03, y + h / 2, { arrow: true, color: C.accent1, width: 1.5 });
    });
    await iconImg(s, "LuMousePointerClick", 0.6, 5.45, 0.42, HEX.accent1);
    txt(s, "실행 중에는 마우스와 키보드를 건드리지 않는다", { x: 1.15, y: 5.35, w: 5.3, h: 0.62, fontSize: 14, valign: "middle" });
    await iconImg(s, "LuCirclePause", 6.85, 5.45, 0.42, HEX.accent1);
    txt(s, "실행은 한 번에 하나. 응답이 끊겨도 다시 실행하지 말고 결과만 다시 묻는다", { x: 7.4, y: 5.35, w: 5.33, h: 0.62, fontSize: 14, valign: "middle" });
    txt(s, "AI가 쓰는 도구: execute_test_suite (is_validation_run=true), execute_test_suite_status, set_attribute", { x: 0.6, y: 6.2, w: 12.13, h: 0.35, fontSize: 11, color: C.accent6, valign: "middle" });
  }
  {
    const L = K.labs[5];
    const s = content("Lab 5  저장, 그리고 막혔을 때", "마지막으로 오늘 만든 것을 정리하고 저장한다. 싱글 유저는 저장으로 끝이다. 멀티 유저라면 저장 후 체크인까지 해야 다른 사람에게 보인다. 체크인 전에는 바뀐 객체 목록을 먼저 보는 습관을 들인다. 오른쪽은 실습에서 자주 막히는 지점이다. 전체 목록은 워크북 부록 B에 있다.", SEC.p2);
    labTag(s, L.id, L.time);
    promptBox(s, 0.6, 1.5, 5.6, 1.35, `프롬프트 ${L.prompts[0].label}`, L.prompts[0].text, 13);
    const rows = [
      [{ text: "", options: { fill: { color: C.accent1 } } }, { text: "작업 끝에 하는 일", options: { bold: true, color: C.background1, fill: { color: C.accent1 } } }],
      [{ text: "싱글 유저", options: { bold: true, fill: { color: C.background2 } } }, { text: "save_workspace", options: { fontFace: MONO, fill: { color: C.background2 } } }],
      [{ text: "멀티 유저", options: { bold: true } }, { text: "update_all → 편집 → save_workspace → check_in_all", options: { fontFace: MONO } }],
    ];
    s.addTable(rows, { x: 0.6, y: 3.1, w: 5.6, colW: [1.3, 4.3], rowH: 0.55, fontSize: 12, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: C.background2 }, margin: [0.04, 0.08, 0.04, 0.08], objectName: nm("table") });
    card(s, 0.6, 4.95, 5.6, 1.8, { fill: { color: C.background2 } });
    txt(s, "멀티 유저일 때 (5-2)", { x: 0.85, y: 5.05, w: 5.1, h: 0.38, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    txt(s, L.prompts[1].text.replace("(멀티 유저일 때) ", ""), { x: 0.85, y: 5.45, w: 5.15, h: 1.2, fontSize: 12 });
    sectionHead(s, 6.55, 1.5, 6, "막혔을 때");
    const tr = [[
      { text: "증상", options: { bold: true, color: C.background1, fill: { color: C.accent1 } } },
      { text: "이렇게 한다", options: { bold: true, color: C.background1, fill: { color: C.accent1 } } },
    ]];
    [0, 2, 3, 4, 5].forEach((idx, i) => {
      const t = K.troubleshooting[idx];
      tr.push([{ text: t[0], options: { bold: true, fill: { color: i % 2 ? C.background1 : C.background2 } } }, { text: t[2], options: { fill: { color: i % 2 ? C.background1 : C.background2 } } }]);
    });
    s.addTable(tr, { x: 6.55, y: 1.95, w: 6.18, colW: [2.1, 4.08], fontSize: 11.5, color: C.text1, valign: "middle", border: { type: "solid", pt: 0.75, color: C.background2 }, margin: [4, 6, 4, 6], objectName: nm("table") });
    card(s, 6.55, 5.35, 6.18, 1.4, { fill: { color: C.accent2, transparency: 86 } });
    await iconImg(s, "LuLifeBuoy", 6.8, 5.6, 0.42, HEX.accent1);
    txt(s, "그래도 안 되면 AI에게 받은 응답을 그대로 복사해 강사에게 보여 준다. 증상 8가지 전체는 워크북 부록 B.", { x: 7.4, y: 5.45, w: 5.15, h: 1.2, fontSize: 13, valign: "middle" });
  }

  // ================= Part 3 =================
  pres.addSection({ title: SEC.p3 });
  {
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: SEC.p3 });
    s.addText("03", { placeholder: "num" });
    s.addText("활용사례와 프롬프트 작성법", { placeholder: "title" });
    s.addText("현업에서 많이 쓰는 여섯 가지와, 결과를 바꾸는 프롬프트 쓰는 법", { placeholder: "body" });
    s.addNotes("실습에서 손으로 해 봤으니, 이제 실제 업무에 어디에 쓸지와 프롬프트를 어떻게 써야 원하는 결과가 나오는지 정리한다.");
  }
  {
    const s = content("현업에서 많이 쓰는 여섯 가지", "Commander MCP를 쓰는 팀에서 자주 나오는 쓰임새를 여섯 가지로 묶었다. 각 카드 아래에는 사람이 꼭 확인해야 할 것과, 바로 쓸 수 있는 템플릿 번호를 적어 두었다. 오늘 실습은 1번과 2번이었다. 4번 실패 원인 분석은 매일 아침 야간 실행 결과를 볼 때 특히 쓸모가 크다. 6번은 라이선스와 권한 설정이 필요하다.", SEC.p3);
    const w = 3.9, h = 2.5, gx = 0.215, gy = 0.2;
    K.useCases.forEach((u, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * (w + gx), y = 1.45 + row * (h + gy);
      card(s, x, y, w, h, { fill: { color: C.background1 }, line: { color: C.background2, width: 1.5 }, shadow: true });
      txt(s, u.no, { x: x + 0.25, y: y + 0.12, w: 0.6, h: 0.65, fontSize: 30, bold: true, color: C.accent2 });
      txt(s, u.title, { x: x + 0.85, y: y + 0.18, w: w - 1.05, h: 0.55, fontSize: 17, bold: true, valign: "middle" });
      txt(s, u.body, { x: x + 0.25, y: y + 0.82, w: w - 0.5, h: 0.95, fontSize: 13 });
      line(s, x + 0.25, y + 1.82, x + w - 0.25, y + 1.82, { color: C.background2, width: 1 });
      const cw = u.template.length > 4 ? 1.45 : 0.75;
      txt(s, [{ text: "사람이 확인  ", options: { bold: true, color: C.text2 } }, { text: u.human, options: {} }], { x: x + 0.25, y: y + 1.9, w: w - cw - 0.6, h: 0.5, fontSize: 11.5, valign: "middle" });
      chip(s, x + w - cw - 0.2, y + 1.97, cw, 0.34, u.template, { fill: C.background2, color: C.text2, bold: true, fontSize: 10 });
    });
  }
  {
    const s = content("프롬프트 문장이 도구 입력이 된다", "AI는 프롬프트를 읽고 도구에 넣을 값을 뽑아낸다. 경로를 적으면 create_test_cases의 path가 되고, 모듈 범위를 적으면 list_assets의 검색 범위가 된다. 반대로 프롬프트에 없는 정보는 AI가 추측하거나 다시 묻는다. 추측이 맞으면 다행이고, 틀리면 엉뚱한 폴더에 테스트케이스가 생긴다. 그래서 구체적으로 써야 한다.", SEC.p3);
    card(s, 0.6, 1.5, 12.13, 1.65, { fill: { color: C.accent2, transparency: 86 }, line: { color: C.accent2, width: 0.75 } });
    const B = (t) => ({ text: t, options: { bold: true, color: C.text2, underline: { style: "sng" } } });
    const N = (t) => ({ text: t, options: { color: C.text1 } });
    const M = (t) => ({ text: t, options: { bold: true, color: C.accent5 } });
    txt(s, [
      B("TestCases/MCP실습/홍길동"), M(" ①"), N(" 폴더에 "), B("'로그인 성공_자동'"), M(" ②"), N(" 테스트케이스를 만들어줘. 모듈은 "), B("Modules/MCP실습에서만"), M(" ③"), N(" 찾고, Account는 "), B("Verify"), M(" ④"), N("로 검증해줘."),
    ], { x: 0.9, y: 1.62, w: 11.6, h: 1.4, fontSize: 19, valign: "middle", lineSpacingMultiple: 1.25 });
    const calls = [
      ["③", "list_assets", "assetType = XModule\nroot = Modules/MCP실습"],
      ["AI 판단", "get_object_info", "찾은 모듈의 속성 확인\n(Header, Login, OpenUrl, CloseBrowser)"],
      ["① ②", "create_test_cases", "path = TestCases/MCP실습/홍길동\nname = 로그인 성공_자동\ntest_steps = [ ... ]"],
      ["④", "set_attribute", "Account 값의\nActionMode = Verify"],
    ];
    const w = 2.83, gap = 0.27, y = 3.65, h = 2.2;
    calls.forEach((c, i) => {
      const x = 0.6 + i * (w + gap);
      card(s, x, y, w, h, { fill: { color: C.background2 } });
      chip(s, x + 0.2, y + 0.18, c[0] === "AI 판단" ? 0.95 : 0.75, 0.34, c[0], { fill: C.background1, color: c[0] === "AI 판단" ? C.accent6 : C.accent5, bold: true, fontSize: 12 });
      txt(s, c[1], { x: x + 0.2, y: y + 0.6, w: w - 0.4, h: 0.4, fontSize: 14, bold: true, fontFace: MONO, color: C.text2 });
      txt(s, c[2], { x: x + 0.2, y: y + 1.05, w: w - 0.4, h: 1.05, fontSize: 11.5, color: C.text1 });
      if (i < 3) line(s, x + w + 0.03, y + h / 2, x + w + gap - 0.03, y + h / 2, { arrow: true, color: C.accent1, width: 1.5 });
    });
    txt(s, "프롬프트에 없는 정보는 AI가 추측하거나 다시 묻는다. 추측이 틀리면 엉뚱한 곳에 만들어진다.", { x: 0.6, y: 6.1, w: 12.13, h: 0.5, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
  }
  {
    const s = content("프롬프트는 네 가지 요소로 쓴다", "Tricentis는 MCP 프롬프트를 쓸 때 구체적으로, 맥락을 주고, 단계를 나누라고 권한다. 이걸 실무에서 바로 쓰기 좋게 대상, 작업, 조건, 확인 네 가지로 정리했다. 대상은 어디서, 작업은 무엇을, 조건은 어떻게, 확인은 결과를 어떻게 받을지다. 아래 예시처럼 네 가지가 다 들어가면 AI가 추측할 여지가 거의 없다.", SEC.p3);
    const w = 2.86, gap = 0.23, y = 1.5, h = 3.25;
    K.promptElements.forEach((e, i) => {
      const x = 0.6 + i * (w + gap);
      card(s, x, y, w, h, { fill: { color: C.background2 } });
      txt(s, e.key, { x: x + 0.25, y: y + 0.2, w: 1.5, h: 0.6, fontSize: 26, bold: true, color: C.text2, valign: "middle" });
      txt(s, e.q, { x: x + 1.5, y: y + 0.25, w: w - 1.7, h: 0.5, fontSize: 13, color: C.accent6, align: "right", valign: "middle" });
      txt(s, e.desc, { x: x + 0.25, y: y + 0.95, w: w - 0.5, h: 1.25, fontSize: 13.5 });
      txt(s, e.ex, { x: x + 0.25, y: y + 2.25, w: w - 0.5, h: 0.85, fontSize: 12, color: C.text2, italic: true });
    });
    card(s, 0.6, 5.0, 12.13, 1.55, { fill: { color: C.accent2, transparency: 86 }, line: { color: C.accent2, width: 0.75 } });
    const T = (t) => ({ text: `[${t}] `, options: { bold: true, color: C.accent5, fontSize: 12 } });
    const N = (t) => ({ text: t + " ", options: {} });
    txt(s, [
      T("대상"), N("TestCases/MCP실습/홍길동 폴더에"), T("작업"), N("'로그인 성공_자동' 테스트케이스를 만들어줘."),
      T("조건"), N("모듈은 Modules/MCP실습에서만 찾고 새로 스캔하지 마."), T("확인"), N("만들기 전에 스텝 계획을 표로 보여줘."),
    ], { x: 0.9, y: 5.1, w: 11.6, h: 1.35, fontSize: 15, valign: "middle", lineSpacingMultiple: 1.2 });
  }
  {
    const s = content("같은 요청, 다른 결과", "왼쪽은 실제로 자주 보는 요청이다. 짧고 편하지만 AI가 추측해야 할 게 많다. 오른쪽은 같은 의도를 네 요소로 고쳐 쓴 것이다. 두 번째와 세 번째 예시는 특히 위험한 경우다. 분석과 수정을 한 번에 시키면 결함을 테스트 수정으로 덮을 수 있고, 정리해 달라는 말은 삭제로 이어질 수 있다.", SEC.p3);
    const y0 = 1.5, rh = 1.62, gap = 0.17;
    for (let i = 0; i < K.promptPairs.length; i++) {
      const p = K.promptPairs[i];
      const y = y0 + i * (rh + gap);
      card(s, 0.6, y, 3.35, rh, { fill: { color: C.background2 } });
      await iconImg(s, "LuCircleX", 0.8, y + 0.2, 0.34, HEX.accent5);
      txt(s, "아쉬운 요청", { x: 1.22, y: y + 0.17, w: 2.5, h: 0.4, fontSize: 12, color: C.accent6, valign: "middle" });
      txt(s, `"${p.bad}"`, { x: 0.8, y: y + 0.65, w: 3.0, h: 0.85, fontSize: 15, bold: true });
      line(s, 4.0, y + rh / 2, 4.35, y + rh / 2, { arrow: true, color: C.accent1, width: 1.5 });
      card(s, 4.4, y, 5.7, rh, { fill: { color: C.accent2, transparency: 86 }, line: { color: C.accent2, width: 0.75 } });
      await iconImg(s, "LuCircleCheck", 4.6, y + 0.2, 0.34, HEX.accent4);
      txt(s, "고친 요청", { x: 5.02, y: y + 0.17, w: 2.5, h: 0.4, fontSize: 12, color: C.accent6, valign: "middle" });
      txt(s, p.good, { x: 4.6, y: y + 0.6, w: 5.3, h: 0.95, fontSize: 12 });
      txt(s, p.why, { x: 10.35, y, w: 2.38, h: rh, fontSize: 12, color: C.accent6, valign: "middle" });
    }
  }
  {
    const s = content("결과를 바꾸는 습관 일곱 가지", "일곱 가지 중 하나만 고르라면 세 번째, 바꾸는 일은 계획부터다. 계획을 보면 AI가 무엇을 오해했는지 실행 전에 알 수 있다. 나머지는 계획을 더 정확하게 만들기 위한 습관이다. 공통 머리말 템플릿 T00에 이 습관들을 규칙으로 넣어 두었다. 대화 시작할 때 한 번 붙여 넣으면 된다.", SEC.p3);
    const colW = 5.95;
    K.tips.forEach((t, i) => {
      const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4;
      const x = 0.6 + col * (colW + 0.23), y = 1.5 + row * 1.2;
      const hot = i === 2;
      badge(s, x, y + 0.08, 0.55, String(i + 1), { fontSize: 16, fill: hot ? C.accent2 : C.accent1, color: hot ? C.text1 : C.background1 });
      txt(s, t[0], { x: x + 0.75, y: y, w: colW - 0.8, h: 0.45, fontSize: 17, bold: true, valign: "middle" });
      txt(s, t[1], { x: x + 0.75, y: y + 0.47, w: colW - 0.8, h: 0.6, fontSize: 13.5, color: C.accent6 });
    });
    card(s, 6.78, 5.15, 5.95, 1.4, { fill: { color: C.accent2, transparency: 84 } });
    txt(s, [
      { text: "T00 공통 머리말", options: { bold: true, color: C.text2, breakLine: true } },
      { text: "위 습관을 규칙으로 정리한 템플릿. 대화를 시작할 때 한 번 붙여 넣는다.", options: {} },
    ], { x: 7.03, y: 5.3, w: 5.5, h: 1.15, fontSize: 13.5 });
  }

  // ================= Part 4 =================
  pres.addSection({ title: SEC.p4 });
  {
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: SEC.p4 });
    s.addText("04", { placeholder: "num" });
    s.addText("프롬프트 템플릿", { placeholder: "title" });
    s.addText("현재 제공되는 18종과 쓰는 방법", { placeholder: "body" });
    s.addNotes("마지막 파트. 지금 바로 쓸 수 있는 템플릿을 소개한다. 전체 문구는 워크북 5장에 있으니 슬라이드에서는 구성과 쓰는 법만 본다.");
  }
  {
    const s = content("현재 제공되는 템플릿 18종", "Tricentis가 공개한 commander-mcp 스킬팩 1.2.1에는 AI가 따라 하는 작업 순서, 즉 워크플로 템플릿 18종이 들어 있다. 이건 AI용이다. 우리는 이 18종에 하나씩 대응하는 사람용 프롬프트 템플릿을 만들었다. 번호와 순서가 공식 워크플로와 같아서, 스킬팩을 설치한 상태라면 AI가 해당 워크플로를 그대로 따른다. 여섯 묶음 중 실행·분석이 7종으로 가장 많다.", SEC.p4);
    txt(s, "commander-mcp 스킬팩(1.2.1)의 워크플로 18종에 1:1로 맞춘 사람용 프롬프트 템플릿. 전체 문구는 워크북 5장.", { x: 0.6, y: 1.42, w: 12.13, h: 0.45, fontSize: 14, color: C.accent6, valign: "middle" });
    chip(s, 0.6, 2.0, 1.75, 0.5, "T00", { fill: C.accent2, bold: true, fontSize: 13 });
    txt(s, "공통 머리말: 대화를 시작할 때 한 번 붙여 넣는 규칙 6줄", { x: 2.5, y: 2.0, w: 10, h: 0.5, fontSize: 14, bold: true, valign: "middle" });
    // 세 열에 묶음 카드를 쌓는다. 실행·분석(7종)은 한 열을 혼자 쓴다.
    const cols = [["조회", "편집", "관리"], ["작성", "확장"], ["실행·분석"]];
    const colW = 3.9, gapX = 0.215, headH = 0.5, lineH = 0.34, gapY = 0.18, y0 = 2.72;
    cols.forEach((groups, ci) => {
      const x = 0.6 + ci * (colW + gapX);
      let y = y0;
      groups.forEach((g) => {
        const list = K.templates.filter((t) => t.group === g);
        const h = headH + list.length * lineH + 0.15;
        card(s, x, y, colW, h, { fill: { color: C.background2 } });
        txt(s, g, { x: x + 0.22, y: y + 0.08, w: 2.5, h: 0.38, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
        chip(s, x + colW - 0.72, y + 0.12, 0.5, 0.3, String(list.length), { fill: C.accent1, color: C.background1, bold: true, fontSize: 11 });
        list.forEach((t, i) => {
          txt(s, [{ text: t.id + "   ", options: { bold: true, color: C.text2 } }, { text: t.title, options: {} }], { x: x + 0.22, y: y + headH + i * lineH, w: colW - 0.4, h: lineH, fontSize: 13, valign: "middle" });
        });
        y += h + gapY;
      });
    });
  }
  {
    const t = K.templates.find((x) => x.id === "T06");
    const s = content("템플릿 하나 뜯어보기: T06 자동 테스트케이스 작성", "템플릿마다 같은 구성이다. 언제 쓰는지, 프롬프트 본문, 작성 예시, AI가 주로 쓰는 도구, 확인할 점. 본문의 대괄호 부분만 바꿔 쓰면 된다. T06은 Lab 3에서 쓴 프롬프트를 일반화한 것이다. 대응하는 공식 워크플로는 author-automated-test-case다. 확인할 점은 꼭 읽게 한다. 템플릿을 써도 사람이 봐야 할 지점은 남는다.", SEC.p4);
    card(s, 0.6, 1.5, 7.1, 5.25, { fill: { color: C.accent2, transparency: 86 }, line: { color: C.accent2, width: 0.75 } });
    chip(s, 0.8, 1.65, 1.35, 0.32, "프롬프트 본문", { fill: C.accent2, bold: true });
    // 줄 단위로 나눈 뒤 줄마다 빈칸 [ ] 부분만 강조한다. 런 안에 줄바꿈을 넣으면 줄이 엉뚱하게 갈라진다.
    const runs = [];
    const lines = t.prompt.split("\n");
    lines.forEach((ln, li) => {
      const parts = ln.split(/(\[[^\]]+\])/).filter((p) => p !== "");
      parts.forEach((part, pi) => {
        const o = part.startsWith("[") ? { bold: true, color: C.accent5 } : {};
        if (pi === parts.length - 1 && li < lines.length - 1) o.breakLine = true;
        runs.push({ text: part, options: o });
      });
    });
    txt(s, runs, { x: 0.85, y: 2.15, w: 6.6, h: 4.5, fontSize: 14.5, lineSpacingMultiple: 1.12 });
    const info = [
      ["언제 쓰나", t.when],
      ["공식 워크플로", `${t.workflow}.md`],
      ["AI가 주로 쓰는 도구", t.tools],
      ["확인할 점", t.caution],
    ];
    let y = 1.5;
    const hs = [0.95, 0.85, 1.55, 1.6];
    info.forEach((it, i) => {
      card(s, 7.95, y, 4.78, hs[i], { fill: { color: C.background2 } });
      txt(s, it[0], { x: 8.15, y: y + 0.1, w: 4.4, h: 0.32, fontSize: 12, bold: true, color: C.text2, valign: "middle" });
      txt(s, it[1], { x: 8.15, y: y + 0.43, w: 4.4, h: hs[i] - 0.5, fontSize: i === 2 ? 10.5 : 12, fontFace: i === 1 || i === 2 ? MONO : undefined });
      y += hs[i] + 0.1;
    });
  }
  {
    const s = content("템플릿 쓰는 순서", "실제로 쓸 때는 이 순서를 따른다. 대화를 시작하면 T00 머리말을 먼저 붙여 넣는다. 하려는 일에 맞는 템플릿을 고르고, 대괄호를 채운다. AI가 계획을 보여주면 확인하고 승인한다. 결과를 검토하고 저장한다. 그리고 써 보니 잘 된 프롬프트는 팀 템플릿에 추가해서 같이 쓴다. 템플릿은 계속 고쳐 가는 것이다.", SEC.p4);
    const steps = [
      ["T00 붙여 넣기", "대화 시작할 때 한 번"],
      ["템플릿 고르기", "워크북 5장 목록에서 하려는 일과 맞는 것"],
      ["[ ] 채우기", "경로와 이름은 Commander에 보이는 그대로"],
      ["계획 확인·승인", "모듈, 값, 대상이 맞는지 보고 승인"],
      ["결과 검토·저장", "Commander에서 직접 열어 보고 저장"],
      ["팀 템플릿에 추가", "잘 된 프롬프트는 공유해서 같이 고친다"],
    ];
    const w = 1.88, gap = 0.17, y = 1.75, h = 2.75;
    steps.forEach((st, i) => {
      const x = 0.6 + i * (w + gap);
      card(s, x, y, w, h, { fill: i === 3 ? { color: C.accent2, transparency: 80 } : { color: C.background2 } });
      badge(s, x + 0.2, y + 0.22, 0.55, String(i + 1), { fontSize: 16, fill: i === 3 ? C.accent2 : C.accent1, color: i === 3 ? C.text1 : C.background1 });
      txt(s, st[0], { x: x + 0.2, y: y + 0.95, w: w - 0.35, h: 0.75, fontSize: 15, bold: true });
      txt(s, st[1], { x: x + 0.2, y: y + 1.7, w: w - 0.35, h: 0.95, fontSize: 12, color: C.accent6 });
    });
    card(s, 0.6, 4.85, 12.13, 1.65, { fill: { color: C.background1 }, line: { color: C.background2, width: 1.5 } });
    await iconCircle(s, "LuPuzzle", 0.9, 5.2, 0.8);
    txt(s, [
      { text: "스킬팩과 템플릿의 역할 분담", options: { bold: true, color: C.text2, breakLine: true } },
      { text: "스킬팩을 설치하면 AI가 공식 워크플로 순서를 스스로 따른다. 템플릿은 그 워크플로에 필요한 정보(경로, 범위, 값, 확인 방식)를 빠짐없이 넘겨주는 역할이다.", options: {} },
    ], { x: 1.95, y: 5.0, w: 10.55, h: 1.35, fontSize: 14, valign: "middle" });
  }

  // ================= 마무리 =================
  pres.addSection({ title: SEC.end });
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: SEC.end });
    s.addText("오늘 정리", { placeholder: "title" });
    const take = [
      ["MCP는 AI가 Commander를 다루게 해 주는 연결 규격", "Commander 2026.1 Patch 1부터 애드인으로 들어 있다"],
      ["AI는 기존 모듈로 만들고, 사람은 계획과 결과를 본다", "스캔, 비밀번호, 삭제, 결함 판단은 사람 몫"],
      ["템플릿으로 시작해서 우리 업무에 맞게 고친다", "T00 머리말 + 공식 워크플로 18종 대응 템플릿"],
    ];
    take.forEach((t, i) => {
      const y = 1.75 + i * 1.18;
      badge(s, 0.8, y + 0.05, 0.6, String(i + 1), { fill: C.accent2, color: C.text1, fontSize: 18 });
      txt(s, t[0], { x: 1.65, y, w: 6.4, h: 0.48, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
      txt(s, t[1], { x: 1.65, y: y + 0.5, w: 6.4, h: 0.45, fontSize: 13, color: C.background2 });
    });
    s.addShape(S.ROUNDED_RECTANGLE, { x: 8.55, y: 1.65, w: 4.18, h: 3.55, rectRadius: 0.1, fill: { color: C.background1, transparency: 90 }, line: { type: "none" }, objectName: nm("next") });
    txt(s, "다음 단계", { x: 8.85, y: 1.85, w: 3.6, h: 0.45, fontSize: 16, bold: true, color: C.accent2 });
    txt(s, [
      { text: "사내 테스트 대상 앱의 모듈을 XScan으로 준비한다", options: { bullet: true, breakLine: true } },
      { text: "반복 업무 하나를 골라 템플릿을 적용해 본다", options: { bullet: true, breakLine: true } },
      { text: "잘 된 프롬프트를 팀 템플릿에 추가한다", options: { bullet: true } },
    ], { x: 8.85, y: 2.4, w: 3.65, h: 2.6, fontSize: 14, color: C.background1, paraSpaceAfter: 10 });
    txt(s, "참고 자료", { x: 0.8, y: 5.55, w: 4, h: 0.35, fontSize: 12, bold: true, color: C.accent2 });
    txt(s, K.references.map((r) => `${r[0]}  ${r[1]}`).join("\n"), { x: 0.8, y: 5.9, w: 11.9, h: 1.3, fontSize: 10.5, color: C.background2 });
    s.addNotes("오늘 내용을 세 줄로 정리한다. 다음 단계로 사내 앱에 바로 적용해 보길 권한다. 모듈이 준비되지 않은 앱은 MCP로 할 수 있는 게 거의 없으니, 모듈 준비가 첫 번째다. 질문 받는다.");
  }

  await pres.writeFile({ fileName: OUT });
  await writeThemeColors(OUT);
  console.log("written", OUT);
}

// pptxgenjs는 테마 색을 Office 기본값으로 쓴다. 저장 후 theme1.xml의 색 구성표를 우리 색으로 바꾼다.
async function writeThemeColors(file) {
  const fs = require("fs");
  const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const part = "ppt/theme/theme1.xml";
  const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const scheme = `<a:clrScheme name="${THEME.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${HEX[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
  const xml = (await zip.file(part).async("string"))
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${THEME.name}"`);
  if (!xml.includes(scheme)) throw new Error("theme1.xml에 색 구성표가 없음");
  zip.file(part, xml);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

function MONOURL() { return K.meta.mcpUrl; }

build().catch((e) => { console.error(e); process.exit(1); });
