// 실습 워크북 생성. 실행: NODE_PATH=<node_modules> node build_docx.js <출력 경로>
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, HeadingLevel, LevelFormat, Header, Footer, PageNumber,
  VerticalAlign, PageBreak, TableLayoutType, HeightRule,
} = require("docx");
const K = require("./content");

const OUT = process.argv[2] || "workbook.docx";

const FONT = "맑은 고딕";
const MONO = { ascii: "Consolas", hAnsi: "Consolas", cs: "Consolas", eastAsia: "맑은 고딕" };
const COL = {
  petrol: "0F4C5C", text: "1E2B2F", muted: "56666B", tint: "EAF1F2", line: "C9D7DA",
  amberFill: "FCF1DC", amber: "E0A03A", amberText: "8A5A0B", red: "C4453C", codeFill: "F2F5F6", white: "FFFFFF",
};
const W = 9638; // A4 본문 폭 (좌우 여백 20mm)

// ---------- 기본 블록 ----------
const run = (text, o = {}) => new TextRun({ text, font: o.mono ? MONO : FONT, ...o, mono: undefined });
function P(content, o = {}) {
  const children = typeof content === "string" ? [run(content, o.run || {})] : content;
  return new Paragraph({ children, spacing: { after: o.after ?? 120, before: o.before ?? 0, line: o.line ?? 300 }, alignment: o.align, keepNext: o.keepNext, indent: o.indent, style: o.style });
}
const H1 = (t, brk = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: brk, spacing: brk ? undefined : { before: 560, after: 280 }, children: [run(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [run(t)] });
const muted = (t, o = {}) => P([run(t, { color: COL.muted, size: 18 })], o);
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

let listInstance = 0;
function bullets(items, o = {}) {
  return items.map((it) => new Paragraph({
    numbering: { reference: "bullet", level: 0 },
    spacing: { after: 60, line: 290 },
    children: typeof it === "string" ? [run(it)] : it,
  }));
}
function numbered(items) {
  const inst = ++listInstance;
  return items.map((it) => new Paragraph({
    numbering: { reference: "num", level: 0, instance: inst },
    spacing: { after: 80, line: 290 },
    children: typeof it === "string" ? [run(it)] : it,
  }));
}
function checks(items) {
  return items.map((t) => new Paragraph({
    spacing: { after: 70, line: 290 }, indent: { left: 360, hanging: 360 },
    children: [run("☐", { color: COL.petrol }), run("  " + t)],
  }));
}

const border = (c = COL.line, sz = 4) => ({ style: BorderStyle.SINGLE, size: sz, color: c });
const borders = (c, sz) => ({ top: border(c, sz), bottom: border(c, sz), left: border(c, sz), right: border(c, sz) });
const noBorders = () => ({ top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } });

function cell(content, width, o = {}) {
  const paras = Array.isArray(content) && content[0] instanceof Paragraph ? content
    : (Array.isArray(content) ? content : [content]).map((c) => (c instanceof Paragraph ? c : new Paragraph({
      children: typeof c === "string" ? [run(c, { bold: o.bold, color: o.color, size: o.size ?? 18, mono: o.mono })] : c,
      spacing: { after: 40, line: 270 }, alignment: o.align, keepNext: o.keep,
    })));
  return new TableCell({
    children: paras,
    width: { size: width, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    margins: { top: o.padV ?? 70, bottom: o.padV ?? 70, left: 110, right: 110 },
    verticalAlign: o.vAlign ?? VerticalAlign.CENTER,
    borders: o.borders ?? borders(),
    rowSpan: o.rowSpan,
  });
}
// 표: headers + rows (문자열 2차원 배열). widths 합은 W.
function table(headers, rows, widths, o = {}) {
  const trs = [];
  if (headers) trs.push(new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { fill: COL.petrol, color: COL.white, bold: true })) }));
  rows.forEach((r, ri) => trs.push(new TableRow({
    cantSplit: true,
    height: o.rowHeight ? { value: o.rowHeight, rule: HeightRule.ATLEAST } : undefined,
    children: r.map((c, i) => cell(c, widths[i], {
      fill: o.zebra === false ? undefined : (ri % 2 === 1 ? COL.tint : undefined),
      bold: o.boldFirst && i === 0, mono: o.monoCols && o.monoCols.includes(i), vAlign: o.vAlign, keep: o.keep,
    })),
  })));
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows: trs, layout: TableLayoutType.FIXED });
}
// 프롬프트 상자: 라벨 + 본문(줄마다 문단)
function promptBox(label, text, o = {}) {
  const lines = text.split("\n");
  const paras = [new Paragraph({ children: [run(label, { bold: true, color: COL.amberText, size: 17 })], spacing: { after: 60 }, keepNext: o.keep })];
  lines.forEach((ln) => {
    const parts = ln.split(/(\[[^\]]+\])/).filter((p) => p !== "");
    paras.push(new Paragraph({
      spacing: { after: 30, line: 290 }, keepNext: o.keep,
      children: parts.length ? parts.map((p) => run(p, p.startsWith("[") && o.highlight !== false ? { bold: true, color: COL.red } : {})) : [run("")],
    }));
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: lines.length < 14, children: [new TableCell({
      children: paras, width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: o.fill ?? COL.amberFill },
      margins: { top: 120, bottom: 120, left: 200, right: 200 }, borders: borders(o.border ?? COL.amber, 6),
    })] })],
  });
}
function codeBlock(code) {
  const paras = code.split("\n").map((ln) => new Paragraph({ spacing: { after: 0, line: 260 }, children: [run(ln, { mono: true, size: 17 })] }));
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      children: paras, width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: COL.codeFill },
      margins: { top: 120, bottom: 120, left: 200, right: 200 }, borders: borders(COL.line, 4),
    })] })],
  });
}
// 기록란: 항목 | 빈칸
function recordTable(fields) {
  const w = [3000, W - 3000];
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: w, layout: TableLayoutType.FIXED,
    rows: fields.map((f) => new TableRow({ cantSplit: true, height: { value: 460, rule: HeightRule.ATLEAST }, children: [
      cell(f, w[0], { fill: COL.tint, bold: true }), cell("", w[1]),
    ] })),
  });
}
// 메모용 빈 상자
function blankBox(height = 1100, hint) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: true, height: { value: height, rule: HeightRule.ATLEAST }, children: [
      new TableCell({ children: [new Paragraph({ children: hint ? [run(hint, { color: "9AA7AA", size: 16 })] : [] })], width: { size: W, type: WidthType.DXA }, borders: borders(COL.line, 4), verticalAlign: VerticalAlign.TOP, margins: { top: 80, bottom: 80, left: 160, right: 160 } }),
    ] })],
  });
}
const gap = (after = 120, keep = false) => new Paragraph({ children: [], spacing: { after }, keepNext: keep });
const label = (t) => P([run(t, { bold: true, color: COL.petrol, size: 19 })], { after: 80, before: 160, keepNext: true });
const mono = (t) => run(t, { mono: true, size: 17, color: COL.petrol });

// ---------- 본문 ----------
const body = [];
const add = (...xs) => xs.flat().forEach((x) => body.push(x));

// 표지
add(
  gap(3200),
  P([run("실습 워크북", { bold: true, color: COL.amberText, size: 24 })], { after: 200 }),
  P([run(K.meta.course, { bold: true, color: COL.petrol, size: 56 })], { after: 160 }),
  P([run(K.meta.subtitle, { color: COL.text, size: 30 })], { after: 600 }),
  P([run(K.meta.baseline, { color: COL.muted, size: 19 })], { after: 60 }),
  P([run(`교육 시간 ${K.meta.duration}`, { color: COL.muted, size: 19 })], { after: 3000 }),
  label("수강자 정보"),
  recordTable(["이름", "소속", "교육일", "워크스페이스 이름", "사용한 AI 클라이언트"]),
  pageBreak(),
);

// 사용법
add(
  P([run("이 워크북 사용법", { bold: true, color: COL.petrol, size: 32 })], { after: 200 }),
  P("강의 슬라이드와 함께 쓰는 실습용 워크북입니다. 실습 중에 입력할 프롬프트, 확인할 점, 결과를 적는 칸이 Lab마다 들어 있습니다. 교육이 끝난 뒤에는 5장 프롬프트 템플릿을 업무에 그대로 가져다 쓰면 됩니다."),
  table(["표시", "뜻"], [
    ["프롬프트 상자", "AI 어시스턴트에 그대로 입력할 문장입니다. 빨간색 [ ] 부분만 자기 상황에 맞게 바꿉니다."],
    ["☐ 확인 포인트", "결과를 보고 맞으면 체크합니다. 하나라도 안 맞으면 다음 단계로 넘어가지 않습니다."],
    ["기록란", "실행 결과를 직접 적습니다. 막혔을 때 강사에게 보여 주면 원인을 빨리 찾을 수 있습니다."],
    ["막히면", "실습에서 자주 생기는 문제와 해결 방법입니다. 전체 목록은 부록 B에 있습니다."],
  ], [2400, W - 2400], { boldFirst: true }),
  gap(200),
  label("워크북 구성"),
  table(["장", "내용", "시간", "슬라이드"], [
    ["1장", "개념 정리: MCP, Tosca MCP Server, Tosca 용어, AI에게 맡기지 않는 일", "35분", "1~12"],
    ["2장", "실습 준비: 사전 점검, 클라이언트별 연결 설정, 스킬팩 설치", "실습 전", "12"],
    ["3장", "실습 Lab 0~5: 연결 확인부터 저장까지", "90분", "13~22"],
    ["4장", "활용사례와 프롬프트 작성법, 고쳐 쓰기 연습", "30분", "23~28"],
    ["5장", "프롬프트 템플릿 T00~T18", "25분", "29~32"],
    ["부록", "A 도구 32종 · B 문제 해결 · C 예시 답안 · D 참고 자료", "", ""],
  ], [1000, 6238, 1200, 1200], { boldFirst: true }),
  gap(200),
  label("시작 전에 알아 둘 것"),
  bullets([
    "실습에 쓰는 워크스페이스는 1인 1개 싱글 유저 워크스페이스입니다. 다른 사람 작업에 영향을 주지 않습니다.",
    "AI 어시스턴트의 답은 매번 문구가 조금씩 다릅니다. 워크북의 화면 예시와 똑같지 않아도 확인 포인트만 맞으면 됩니다.",
    "AI가 쓰는 모델은 대부분 클라우드에서 동작합니다. 실습 밖에서 쓸 때는 사내 AI 사용 정책을 먼저 확인하세요.",
  ]),
);

// 1장
add(
  H1("1장  개념 정리"),
  H2("1.1  MCP란"),
  P("MCP(Model Context Protocol)는 AI 앱이 외부 프로그램의 기능을 '도구(Tool)'로 불러 쓸 수 있게 해 주는 공개 규격입니다. 2024년 11월 Anthropic이 공개했고, 2025년 12월 Linux Foundation 산하 Agentic AI Foundation으로 넘어갔습니다."),
  P("예전에는 AI 앱과 프로그램을 연결하려면 조합마다 연동을 따로 만들어야 했습니다. MCP는 규격이 하나이므로 AI 앱과 프로그램이 각자 MCP만 지원하면 서로 연결됩니다. 그래서 'AI용 USB-C'에 자주 비유합니다."),
  table(["구성 요소", "하는 일", "오늘 실습에서"], [
    ["Host", "사용자가 직접 쓰는 AI 앱", "VS Code + GitHub Copilot, Cursor, Claude Desktop, Claude Code"],
    ["Client", "Host 안에서 MCP 서버와 통신", "자동으로 동작하므로 따로 다룰 일 없음"],
    ["Server", "프로그램의 기능을 도구로 제공", "Tosca MCP Server (Commander 애드인)"],
  ], [1600, 3400, W - 5000], { boldFirst: true }),
  gap(),
  label("요청 하나가 처리되는 순서"),
  numbered([
    "사람이 AI 어시스턴트에 문장으로 요청합니다.",
    "AI가 요청을 읽고 어떤 도구를 어떤 순서로 쓸지 고릅니다.",
    "Client가 MCP 서버에 도구 호출을 보냅니다.",
    "서버(Tosca MCP Server)가 Commander에서 실제로 실행하고 결과를 돌려줍니다.",
    "AI가 결과를 정리해서 답합니다. 여러 도구가 필요하면 2~4를 반복합니다.",
  ]),
  P([run("2번이 핵심입니다. ", { bold: true }), run("AI는 프롬프트에 적힌 정보로 도구 입력값을 정합니다. 프롬프트에 없는 정보는 추측하거나 다시 묻습니다. 그래서 경로와 범위를 구체적으로 적어야 합니다.")]),

  H2("1.2  Tosca MCP Server"),
  P("Tosca MCP Server는 별도 프로그램이 아니라 Tosca Commander에 들어 있는 애드인(Tosca MCP Server Addin)입니다."),
  table(["항목", "내용"], [
    ["제공 버전", "Tosca Commander 2026.1 LTS Patch 1부터. Patch 2부터는 애드인이 기본으로 켜져 있음"],
    ["접속 주소", "http://127.0.0.1:46248/mcp (내 PC 안에서만 접속, 디버그 빌드는 8080)"],
    ["동작 조건", "Commander가 실행 중이고 워크스페이스가 열려 있어야 함. 닫으면 연결이 끊김"],
    ["권한", "워크스페이스를 연 사용자의 권한으로 동작. 체크인·체크아웃 규칙도 똑같이 적용"],
    ["테스트된 클라이언트", "Claude Code, Claude Desktop, Cursor, Visual Studio Code (Tricentis 문서 기준)"],
    ["Tosca Cloud MCP와 관계", "별개 제품. Cloud용은 원격 서버에 로그인해서 쓰며 설정 방법이 다름"],
  ], [2600, W - 2600], { boldFirst: true }),
  gap(),
  P("지금 연결된 서버가 어떤 도구를 제공하는지는 AI에게 직접 물어보면 됩니다. Tricentis 문서도 'Give me a list of all available tools'처럼 도구 목록을 요청하라고 안내합니다. 스킬팩 1.2.1 기준 도구는 32개이며, 전체 목록은 부록 A에 있습니다."),

  H2("1.3  오늘 쓰는 Tosca 용어"),
  table(["용어", "읽는 법", "뜻"], K.terms.map((t) => [t.term, t.ko, t.desc]), [2100, 1700, W - 3800], { boldFirst: true }),

  H2("1.4  AI에게 맡기지 않는 일"),
  P("할 수 있는 일만큼 중요한 것이 맡기면 안 되는 일입니다. 아래 여섯 가지는 사람이 직접 하거나 반드시 사람이 확인합니다."),
  table(["항목", "설명"], K.notForAI.map((x) => [x[1], x[2]]), [2600, W - 2600], { boldFirst: true }),

  H2("1.5  확인 문제"),
  muted("답은 부록 C에 있습니다."),
  P([run("Q1. ", { bold: true }), run("Tosca MCP Server를 쓰는 동안 Commander에서 계속 열려 있어야 하는 것은 무엇인가요?")], { after: 60 }),
  blankBox(600),
  gap(80),
  P([run("Q2. ", { bold: true }), run("필요한 화면의 모듈이 없을 때 AI 대신 사람이 써야 하는 Tosca 기능은 무엇인가요?")], { after: 60 }),
  blankBox(600),
  gap(80),
  P([run("Q3. ", { bold: true }), run("ScratchBook 실행과 ExecutionList 실행의 차이를 한 줄로 적어 보세요.")], { after: 60 }),
  blankBox(800),
  gap(80),
  P([run("Q4. ", { bold: true }), run("멀티 유저 워크스페이스에서 바꾼 내용을 다른 사람도 보게 하려면 저장 다음에 무엇을 해야 하나요?")], { after: 60 }),
  blankBox(600),
);

// 2장
add(
  H1("2장  실습 준비"),
  H2("2.1  사전 점검"),
  P("실습 시작 전에 아래 항목을 하나씩 확인합니다. 체크되지 않은 항목이 있으면 강사에게 알려 주세요."),
  checks([
    "Tosca Commander가 2026.1 LTS Patch 1 이상이다",
    "Patch 1이라면 Tosca MCP Server Addin을 켰다 (Patch 2 이상은 기본으로 켜져 있음)",
    "강사가 배포한 실습 워크스페이스를 열었다",
    "사내에서 허용된 AI 클라이언트를 설치하고 로그인했다",
    "브라우저에서 https://demowebshop.tricentis.com 이 열린다",
    "강사에게 실습 계정을 받았다",
    "(권장) commander-mcp 스킬팩을 설치했다 (2.3 참고)",
  ]),

  H2("2.2  클라이언트별 연결 설정"),
  P("어느 클라이언트든 Commander MCP Server 주소 한 줄을 등록하면 됩니다. 서버 이름은 스킬팩 안내와 같게 tosca-commander로 맞춥니다."),
);
K.clients.forEach((c) => {
  add(
    H3(c.name),
    P([run("설정 위치  ", { bold: true, color: COL.petrol }), run(c.where)], { after: 80, keepNext: true }),
    codeBlock(c.code),
    P([run("참고  ", { bold: true, color: COL.petrol }), run(c.note)], { before: 80, after: 160 }),
  );
});
add(
  P([run("포트 안내  ", { bold: true }), run("46248은 기본값입니다. 디버그 빌드는 8080을 쓰고, Commander MCP 설정에서 포트를 바꿨다면 그 값을 넣습니다.")]),

  H2("2.3  스킬팩 설치 (권장)"),
  P("Tricentis가 공개한 commander-mcp 스킬팩은 AI 어시스턴트에게 Commander MCP를 다루는 규칙과 작업 순서를 알려 줍니다. 체크아웃과 저장 순서, 대화상자를 여는 메뉴 작업 피하기, 워크플로 템플릿 18종 같은 내용이 들어 있습니다. MCP 연결과는 별개이며, Tricentis는 둘을 함께 쓰기를 권합니다."),
  label("Git을 쓸 수 있을 때"),
  codeBlock("git clone https://github.com/Tricentis/mcp-skills.git\ncd mcp-skills/Tosca/Commander/MCP\n.\\Install-CommanderMcpPack.ps1 -Ide Claude"),
  P([run("-Ide 값은 "), mono("Cursor"), run(", "), mono("Claude"), run(", "), mono("VSCode"), run(", "), mono("Windsurf"), run(" 중 하나입니다. VS Code는 "), mono("-Scope Project"), run("를 함께 지정해야 합니다.")], { before: 80 }),
  label("스크립트 실행이 막혀 있을 때 (폴더 복사)"),
  table(["클라이언트", "복사할 것", "붙여 넣을 위치"], [
    ["Claude Code", "claude/skills/commander-mcp/", "%USERPROFILE%\\.claude\\skills\\commander-mcp\\"],
    ["Cursor", "cursor/skills/commander-mcp/\ncursor/rules/commander-mcp.mdc", "%USERPROFILE%\\.cursor\\skills\\commander-mcp\\\n%USERPROFILE%\\.cursor\\rules\\"],
    ["VS Code + Copilot", "vscode/instructions/commander-mcp.instructions.md", "프로젝트 폴더의 .github/instructions/"],
  ].map((r) => r.map((c) => c.split("\n"))), [1900, 3700, W - 5600], { boldFirst: true }),
  gap(),
  P([run("설치 확인  ", { bold: true }), run("클라이언트에서 /commander-mcp 스킬이 보이고, Lab 0의 워크스페이스 정보 요청이 성공하면 됩니다.")]),
  muted("스킬팩은 Apache 2.0 오픈소스로 제공되며 Tricentis 제품 지원 계약 범위 밖입니다. 문제는 GitHub Issues로 제보합니다."),

  H2("2.4  실습 환경"),
  P([run("대상 앱  ", { bold: true }), run("Tricentis Demo Web Shop  "), mono(K.meta.app)]),
  P("테스트 연습용으로 공개된 쇼핑몰 사이트입니다. 실습에 쓰는 모듈은 강사가 미리 XScan으로 스캔해 워크스페이스에 넣어 두었습니다."),
  table(["모듈", "위치", "주요 속성", "설명"], K.practiceModules.map((m) => [m.module, m.where, m.attrs, m.note]), [1500, 2000, 2900, W - 6400], { boldFirst: true }),
  muted("표준 모듈(OpenUrl, CloseBrowser)의 폴더 위치와 속성 이름은 Tosca 버전에 따라 조금 다를 수 있습니다. 강사가 실습 전에 확인해 둡니다.", { before: 60 }),
  gap(),
  label("실습 계정 기록"),
  recordTable(["실습 이메일", "비밀번호 보관 방법 (적지 말 것)"]),
);

// 3장
add(
  H1("3장  실습"),
  P("Lab 0부터 순서대로 진행합니다. 앞 Lab의 결과를 뒤 Lab에서 쓰므로 건너뛰지 않습니다. 모든 프롬프트에서 [내 이름], [실습 이메일], [실습 비밀번호]는 자기 값으로 바꿉니다."),
  table(["Lab", "내용", "시간", "워크스페이스 변화"], [
    ["Lab 0", "연결 확인", "10분", "없음 (조회)"],
    ["Lab 1", "워크스페이스 둘러보기", "10분", "없음 (조회)"],
    ["Lab 2", "수동 테스트케이스 만들기", "15분", "테스트케이스 생성"],
    ["Lab 3", "자동 테스트케이스 만들기", "25분", "테스트케이스 생성"],
    ["Lab 4", "ScratchBook 검증 실행과 수정", "20분", "값 수정, 실행"],
    ["Lab 5", "저장과 정리", "10분", "저장"],
  ], [1200, 4238, 1400, 2800], { boldFirst: true }),
);
K.labs.forEach((L, idx) => {
  add(
    idx === 0 ? gap(0) : pageBreak(),
    H2(`${L.id}  ${L.title}  (${L.time})`),
    P([run("목표  ", { bold: true, color: COL.petrol }), run(L.goal)]),
    label("준비"),
    L.before.length > 1 ? numbered(L.before) : P(L.before[0]),
  );
  L.prompts.forEach((p) => add(gap(60), promptBox(`프롬프트 ${p.label}`, p.text)));
  add(
    gap(60),
    P([run("AI가 주로 쓰는 도구  ", { bold: true, color: COL.petrol }), ...L.tools.flatMap((t, i) => [mono(t), run(i < L.tools.length - 1 ? ", " : "")])], { before: 100 }),
  );
  if (L.id === "Lab 0") {
    add(
      label("화면 예시 (프롬프트 0-2)"),
      promptBox("AI 응답 예시", "[get_workspace_info 실행]\n현재 열린 워크스페이스입니다.\n· 이름: MCP_Training_홍길동\n· 경로: C:\\Tosca_Projects\\Tosca_Workspaces\\MCP_Training_홍길동\n· 모드: single (싱글 유저)\n체크인 없이 저장만 하면 됩니다.", { fill: COL.tint, border: COL.line, highlight: false }),
      muted("실제 문구와 모양은 클라이언트마다 다릅니다. 응답에 도구 이름이 보이면 MCP로 Commander를 실제로 읽은 것입니다.", { before: 60 }),
    );
  }
  if (L.id === "Lab 3") {
    add(
      label("기대 결과 (계획 표)"),
      table(["#", "모듈", "속성", "값", "ActionMode"], K.lab3Steps, [500, 1700, 1500, 4038, 1900], { monoCols: [3] }),
      muted("맞는 모듈을 못 찾았다는 스텝이 있으면, 그 화면은 사람이 XScan으로 스캔한 뒤 같은 프롬프트로 다시 요청합니다.", { before: 60 }),
    );
  }
  if (L.id === "Lab 4") {
    add(
      label("왜 일부러 실패를 만드나요"),
      P("실무에서 AI에게 가장 많이 맡기는 일 중 하나가 실패 원인 분석입니다. 원인을 미리 알고 있는 실패를 만들어 보면, AI의 설명이 맞는지 직접 확인할 수 있습니다. 4-3에서 '아직 고치지는 마'라고 적는 이유도 같습니다. 분석과 수정을 나눠야 실제 결함을 테스트 수정으로 덮는 일을 막을 수 있습니다."),
    );
  }
  add(
    label("확인 포인트"),
    checks(L.checks),
    label("결과 기록"),
    recordTable(L.record),
    label("막히면"),
    table(["증상", "이렇게 해 보세요"], L.stuck, [2800, W - 2800], { boldFirst: true, zebra: false }),
  );
  if (L.id === "Lab 4") {
    add(label("메모 (실행하면서 눈에 띈 것)"), blankBox(2200, "예: 브라우저가 열리는 데 걸린 시간, AI가 결과를 보고하기까지 몇 번 상태를 확인했는지"));
  }
  if (L.id === "Lab 3") {
    add(
      label("추가 과제 (빨리 끝났다면)"),
      P("5장의 T07 템플릿으로 '로그인 성공_자동'을 설명해 달라고 요청해 보세요. AI의 설명이 내가 만든 테스트케이스와 같은지 비교합니다."),
    );
  }
});

// 4장
add(
  H1("4장  활용사례와 프롬프트 작성"),
  H2("4.1  현업에서 많이 쓰는 여섯 가지"),
  table(["", "활용사례", "내용", "사람이 확인할 것", "템플릿"], K.useCases.map((u) => [u.no, u.title, u.body, u.human, u.template]), [450, 1900, 3988, 2200, 1100], { vAlign: VerticalAlign.CENTER }),

  H2("4.2  프롬프트 문장이 도구 입력이 되는 과정"),
  P("AI는 프롬프트를 읽고 도구에 넣을 값을 뽑아냅니다. Lab 3의 프롬프트를 예로 들면 다음과 같습니다."),
  table(["프롬프트 속 표현", "AI가 쓰는 도구", "들어가는 입력"], [
    ["Modules/MCP실습에서만 찾아", "list_assets", "assetType = XModule, root = Modules/MCP실습"],
    ["(AI가 스스로 판단)", "get_object_info", "찾은 모듈의 속성 확인"],
    ["TestCases/MCP실습/홍길동 폴더에 '로그인 성공_자동'", "create_test_cases", "path, name, test_steps"],
    ["Account가 같은지 Verify", "set_attribute", "ActionMode = Verify"],
  ], [3400, 2200, W - 5600], { monoCols: [1] }),
  gap(),
  P("프롬프트에 없는 정보는 AI가 추측하거나 다시 묻습니다. 추측이 틀리면 엉뚱한 폴더에 테스트케이스가 생기거나 다른 모듈이 쓰입니다."),

  H2("4.3  네 가지 요소로 쓰기"),
  P("Tricentis는 MCP 프롬프트를 쓸 때 구체적으로 쓰고, 프로젝트·폴더·테스트케이스 이름 같은 맥락을 주고, 요청을 단계로 나누라고 권합니다. 이 워크북에서는 이를 대상, 작업, 조건, 확인 네 가지로 정리했습니다."),
  table(["요소", "질문", "적을 내용", "예"], K.promptElements.map((e) => [e.key, e.q, e.desc, e.ex]), [900, 1300, 4038, 3400], { boldFirst: true }),
  gap(),
  promptBox("네 요소를 모두 넣은 예", "[대상] TestCases/MCP실습/홍길동 폴더에 [작업] '로그인 성공_자동' 테스트케이스를 만들어줘.\n[조건] 모듈은 Modules/MCP실습에서만 찾고 새로 스캔하지 마. [확인] 만들기 전에 스텝 계획을 표로 보여줘."),

  H2("4.4  결과를 바꾸는 습관 일곱 가지"),
  numbered(K.tips.map((t) => [run(t[0] + "  ", { bold: true }), run(t[1])])),
  P([run("하나만 고른다면 3번입니다. ", { bold: true }), run("계획을 먼저 보면 AI가 무엇을 오해했는지 실행 전에 알 수 있습니다. 이 습관들은 5장 T00 공통 머리말에 규칙으로 들어 있습니다.")]),

  H2("4.5  같은 요청, 다른 결과"),
  table(["아쉬운 요청", "고친 요청", "이유"], K.promptPairs.map((p) => [p.bad, p.good, p.why]), [2200, 4638, 2800], { vAlign: VerticalAlign.TOP }),

  H2("4.6  고쳐 쓰기 연습"),
  P("아래 요청을 네 가지 요소가 들어가게 고쳐 써 보세요. 예시 답안은 부록 C에 있습니다.", { keepNext: true }),
);
const exercises = [
  ["연습 1", "회원가입 테스트케이스 만들어줘", "어느 폴더에? 어떤 모듈로? 바로 만들까, 계획부터 볼까?"],
  ["연습 2", "실행 결과 좀 봐줘", "어느 ExecutionList? 무엇을 알고 싶은가? 고쳐도 되는가?"],
  ["연습 3", "테스트케이스 이름들 다 바꿔줘", "어느 범위? 어떤 규칙으로? 바꾸기 전에 확인할 것은?"],
];
exercises.forEach((e) => add(
  P([run(e[0] + "  ", { bold: true, color: COL.petrol }), run(`"${e[1]}"`, { bold: true })], { before: 120, after: 40, keepNext: true }),
  muted("생각할 점: " + e[2], { after: 60, keepNext: true }),
  blankBox(1300),
));

// 5장
add(
  H1("5장  프롬프트 템플릿"),
  P("Tricentis가 공개한 commander-mcp 스킬팩 1.2.1에는 AI가 따라 하는 워크플로 템플릿 18종이 들어 있습니다. 이 장의 T01~T18은 그 18종에 하나씩 대응하는 사람용 프롬프트 템플릿입니다. 번호 순서도 공식 워크플로 목록(reference/workflows-index.md)과 같습니다."),
  P("스킬팩을 설치하면 AI가 공식 워크플로 순서를 스스로 따릅니다. 템플릿은 그 워크플로에 필요한 정보(경로, 범위, 값, 확인 방식)를 빠짐없이 넘겨주는 역할을 합니다."),
  H2("5.1  템플릿 쓰는 순서"),
  numbered([
    "대화를 시작할 때 T00 공통 머리말을 한 번 붙여 넣습니다.",
    "아래 목록에서 하려는 일과 맞는 템플릿을 고릅니다.",
    "빨간색 [ ] 부분을 채웁니다. 경로와 이름은 Commander에 보이는 그대로 적습니다.",
    "AI가 보여 주는 계획에서 대상, 모듈, 값이 맞는지 보고 승인합니다.",
    "Commander에서 결과를 직접 열어 확인하고 저장합니다.",
    "잘 된 프롬프트는 팀 템플릿에 추가해 같이 고쳐 갑니다.",
  ]),
  H2("5.2  템플릿 목록"),
  table(["번호", "템플릿", "묶음", "공식 워크플로 파일"], [["T00", K.preamble.title, "공통", "(해당 없음)"], ...K.templates.map((t) => [t.id, t.title, t.group, t.workflow + ".md"])], [900, 3600, 1300, W - 5800], { boldFirst: true, monoCols: [3] }),
  H2("T00  공통 머리말"),
  P([run("언제 쓰나  ", { bold: true, color: COL.petrol }), run(K.preamble.when)]),
  promptBox("템플릿", K.preamble.prompt),
);
K.templates.forEach((t) => {
  const exampleIsRef = !t.example.includes("\n") && t.example.length < 40;
  add(
    gap(160),
    H2(`${t.id}  ${t.title}`),
    table(null, [
      ["언제 쓰나", t.when],
      ["공식 워크플로", t.workflow + ".md"],
      ["AI가 주로 쓰는 도구", t.tools],
      ["확인할 점", t.caution],
    ], [2100, W - 2100], { boldFirst: true, zebra: false, keep: true }),
    gap(80, true),
    promptBox("템플릿", t.prompt, { keep: true }),
    gap(80, true),
    exampleIsRef
      ? P([run("작성 예시  ", { bold: true, color: COL.petrol }), run(t.example)])
      : promptBox("작성 예시", t.example, { fill: COL.tint, border: COL.line, highlight: false }),
  );
});

// 부록
add(H1("부록 A  도구 32종"), muted("commander-mcp 스킬팩 1.2.1의 도구 카탈로그(reference/tools-catalog.md) 기준입니다. 실제 목록은 Commander 패치에 따라 다를 수 있으니 AI에게 직접 물어 확인하세요.", { after: 120 }));
{
  const w = [1500, 3700, W - 5200];
  const rows = [new TableRow({ tableHeader: true, children: ["묶음", "도구", "설명"].map((h, i) => cell(h, w[i], { fill: COL.petrol, color: COL.white, bold: true })) })];
  K.toolGroups.forEach((g) => {
    g.tools.forEach((tl, i) => {
      const cells = [];
      if (i === 0) cells.push(cell(`${g.name}\n${g.tools.length}개`.split("\n"), w[0], { bold: true, fill: COL.tint, rowSpan: g.tools.length }));
      cells.push(cell(tl[0], w[1], { mono: true, size: 16 }));
      cells.push(cell(tl[1], w[2]));
      rows.push(new TableRow({ cantSplit: true, children: cells }));
    });
  });
  add(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: w, rows, layout: TableLayoutType.FIXED }));
}
add(
  H1("부록 B  문제 해결", false),
  table(["증상", "흔한 원인", "해결"], K.troubleshooting, [2600, 3000, W - 5600], { boldFirst: true }),
  gap(),
  P("그래도 해결되지 않으면 AI에게 받은 응답을 그대로 복사해 강사에게 보여 주세요. 어떤 도구가 어떤 오류를 냈는지가 원인을 찾는 가장 빠른 단서입니다."),

  H1("부록 C  예시 답안", false),
  H2("C.1  확인 문제"),
  table(["문제", "답"], [
    ["Q1", "워크스페이스. MCP는 열린 Commander 안에서 동작하므로 Commander와 워크스페이스가 계속 열려 있어야 합니다."],
    ["Q2", "XScan. AI는 기존 모듈만 재사용하므로 새 화면은 사람이 XScan으로 스캔합니다."],
    ["Q3", "ScratchBook은 시험 실행이라 결과가 정식 기록으로 남지 않고, ExecutionList 실행은 실행 결과(ExecutionLog)가 남습니다."],
    ["Q4", "체크인(check_in_all). 저장(save_workspace)만으로는 다른 사용자에게 반영되지 않습니다."],
  ], [1000, W - 1000], { boldFirst: true }),
  H2("C.2  고쳐 쓰기 연습"),
  muted("정답은 하나가 아닙니다. 대상, 작업, 조건, 확인이 다 들어 있으면 됩니다.", { after: 100 }),
  P([run("연습 1", { bold: true, color: COL.petrol })], { after: 60, keepNext: true }),
  promptBox("예시", "TestCases/MCP실습/[내 이름] 폴더에 '회원가입_정상_자동' 자동화 테스트케이스를 만들려고 해. 아직 만들지 말고 계획만 보여줘.\n모듈은 Modules/MCP실습에서만 찾고 새로 스캔하지 마. 시나리오는 [단계]이고, 가입 완료 문구를 Verify 해줘.\n맞는 모듈이 없는 스텝은 따로 표시해줘."),
  muted("실습 워크스페이스에는 회원가입 화면 모듈이 없으므로, AI가 '맞는 모듈 없음'을 보고하는 것이 정상입니다.", { before: 60, after: 160 }),
  P([run("연습 2", { bold: true, color: COL.petrol })], { after: 60, keepNext: true }),
  promptBox("예시", "Execution/[ExecutionList 이름]의 마지막 실행 결과를 보여줘. 통과/실패/미실행 개수를 먼저 알려 주고,\n실패한 테스트케이스는 실패 스텝과 메시지를 표로 정리해줘. 수정은 하지 마."),
  gap(160),
  P([run("연습 3", { bold: true, color: COL.petrol })], { after: 60, keepNext: true }),
  promptBox("예시", "TestCases/MCP실습/[내 이름] 폴더의 테스트케이스 이름을 '[화면]_[시나리오]_[수동/자동]' 규칙으로 바꾸려고 해.\n바꿀 이름을 현재 이름과 나란히 표로 먼저 보여줘. 내가 확인하면 그때 바꿔줘. 이름은 속성 변경(set_attribute)으로 바꿔."),

  H1("부록 D  참고 자료", false),
  table(["자료", "주소"], K.references, [4200, W - 4200], { monoCols: [1] }),
  gap(),
  muted("이 워크북의 도구 목록과 워크플로 이름은 2026년 10월 기준 commander-mcp 스킬팩 1.2.1을 따릅니다. 스킬팩이 갱신되면 5장 템플릿 목록과 부록 A를 함께 고칩니다."),
);

// ---------- 문서 ----------
const doc = new Document({
  creator: "QA 교육",
  lastModifiedBy: "QA 교육",
  title: `${K.meta.course} 실습 워크북`,
  description: K.meta.subtitle,
  styles: {
    default: { document: { run: { font: FONT, size: 20, color: COL.text } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 36, bold: true, color: COL.petrol },
        paragraph: { spacing: { before: 0, after: 280 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 27, bold: true, color: COL.petrol },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: COL.text },
        paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullet", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 320 } } } }] },
    ],
  },
  sections: [{
    properties: {
      titlePage: true,
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134, header: 567, footer: 567 } },
    },
    headers: {
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run(`${K.meta.course}  ·  실습 워크북`, { size: 16, color: COL.muted })] })] }),
      first: new Header({ children: [new Paragraph({ children: [] })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 17, color: COL.muted })] })] }),
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("written", OUT); });
