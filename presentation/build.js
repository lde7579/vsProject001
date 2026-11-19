// 발표자료 생성: presentation/data.json → out/발표_적정재고일수.pptx
// 숫자는 모두 data.json(= extract.py가 결과 xlsx에서 읽은 값)에서 가져온다. 손으로 적은 숫자는 없다.
// 실행: node presentation/build.js [--data presentation/data.json] [--out out/발표_적정재고일수.pptx]
'use strict';
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

// ---------------------------------------------------------------- 인자·데이터
const ROOT = path.resolve(__dirname, '..');
function arg(name, def) {
  const i = process.argv.indexOf(name);
  return i > 0 ? path.resolve(process.argv[i + 1]) : def;
}
const DATA_PATH = arg('--data', path.join(__dirname, 'data.json'));
const OUT_PATH = arg('--out', path.join(ROOT, 'out', '발표_적정재고일수.pptx'));
const D = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));

// ---------------------------------------------------------------- 양식 상수
const FONT = '맑은 고딕';
const TASK = 'MP 시스템 상품 안전재고 산출 Logic 변경';
const W = 13.333, MX = 0.6, CW = W - 2 * MX; // LAYOUT_WIDE 13.333 x 7.5 in
const C = {
  navy: '1B2E44', navy2: '27405C', teal: '0B7F80', tealL: 'DCEEEE', slate: '52657A', gray: '5F6F80',
  light: 'F0F3F6', line: 'D3DBE3', text: '1F2D3D', amber: 'C8891A', amberL: 'F6ECDA', red: 'C0392B',
  redL: 'F9E6E3', white: 'FFFFFF', sky: 'A9B8C8', mute: '9AA8B6',
};
// 기준별 색(모든 차트에서 같게)
const BASIS_COLOR = { '기존수기': '9AA8B6', '가중평균': C.teal, '실적기준': C.slate, 'FCST기준': 'C8A25E' };
const BASIS_TEXT = { '기존수기': C.slate, '가중평균': C.teal, '실적기준': C.slate, 'FCST기준': 'A27A2E' };
const BASES = ['기존수기', '가중평균', '실적기준', 'FCST기준'];
const formulaBases = ['가중평균', '실적기준', 'FCST기준'];

// ---------------------------------------------------------------- 숫자·주차 표기
const isNum = (v) => typeof v === 'number' && !Number.isNaN(v);
const pct = (v, d = 1) => (isNum(v) ? (v * 100).toFixed(d) + '%' : '–');
// 차이는 화면에 보이는 반올림 값끼리 뺀다(예: 69.0% − 60.6% = 8.4%p)
const pp = (a, b) => ((Math.round(a * 1000) - Math.round(b * 1000)) / 10).toFixed(1) + '%p';
const f1 = (v) => (isNum(v) ? v.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : '–');
const int = (v) => (isNum(v) ? Math.round(v).toLocaleString('en-US') : '–');
const qty = (v) => (isNum(v) && Math.abs(v) < 100 ? f1(v) : int(v));
const wfmt = (v) => (isNum(v) ? v.toFixed(1) : '–');
const PENDING = '산출 예정';
// 'WK35~WK43' 같은 주차 범위를 'W35~43'으로 통일한다
function wr(s) {
  const m = String(s || '').match(/W[K]?(\d+)\D+W?[K]?(\d+)/);
  return m ? `W${m[1]}~${m[2]}` : String(s || '');
}

// ---------------------------------------------------------------- 파생 값
const M = D.metrics['공통'];
const MALL = D.metrics['전체'];
const SRC = D.sources.result;
const WS = D.weight_search; // 없으면 null
const VAL = wr(D.criteria['검증 주차']); // 사후 검증 구간 W35~43
const HIST = wr(D.criteria['실적 주차']); // 실적 W01~31
const SEL = WS ? wr(WS.criteria['선택 구간']) : 'W35~38';
const HOLD = WS ? wr(WS.criteria['검증 구간']) : 'W39~43';
const groupsOrdered = Object.keys(D.metrics)
  .filter((g) => !['전체', '공통', '미분류'].includes(g))
  .sort((a, b) => (D.metrics[b]['가중평균'].n_demand - D.metrics[a]['가중평균'].n_demand) || a.localeCompare(b));
const bigGroups = groupsOrdered.filter((g) => D.metrics[g]['가중평균'].n_demand >= 5);
const worstGroup = bigGroups.reduce((a, b) => (D.metrics[a]['가중평균'].wape >= D.metrics[b]['가중평균'].wape ? a : b));
const bestGroup = bigGroups.reduce((a, b) => (D.metrics[a]['가중평균'].wape <= D.metrics[b]['가중평균'].wape ? a : b));
const bestWape = formulaBases.reduce((a, b) => (M[a].wape <= M[b].wape ? a : b));
const coverGap = M['기존수기'].cover - M['가중평균'].cover; // 양수면 기존수기가 높음
const SS = D.safety;
const oldLevels = Object.keys(D.days.old_safety_levels);
function wsRow(period, scope, weights, basis) {
  if (!WS) return null;
  return WS.compare.find((r) => r.period === period && r.scope === scope && r.weights === weights && r.basis === basis) || null;
}
const wsCur = wsRow('검증', '전체', '현재', '가중평균');
const wsNew = wsRow('검증', '전체', '탐색', '가중평균');
const wsReady = !!(wsCur && wsNew && isNum(wsCur.wape) && isNum(wsNew.wape));
const contrib = wsReady && WS.contributions ? WS.contributions.groups : null;
// 오차 감소 몫이 큰 순서의 그룹 두 개(예: WPU 71% · SEN 16%)
const topContrib = contrib ? Object.entries(contrib).filter(([, v]) => isNum(v.share)).sort((a, b) => b[1].share - a[1].share).slice(0, 2) : [];
const topContribText = topContrib.map(([g, v]) => `${g} ${pct(v.share, 0)}`).join(' · ');
const METHOD = { '그룹최적': '그룹', '공통최적': '전체 합산', '현재유지': '현재 유지' };
// AI 가중치 탐색: 후보 수·그룹 수·탐색 공간. 그룹 오차는 그 그룹의 w에만 달려 있어서
// 그룹 × 후보 계산으로 전체 조합(후보 수의 그룹 수 제곱)에서 고른 것과 같은 답을 얻는다.
const CANDS = WS ? [...new Set(WS.candidates.map((c) => c.w))].sort((a, b) => a - b) : [];
const NG = WS ? WS.choices.length : 0;
const SPACE = WS ? Math.pow(CANDS.length, NG) : null;
const man = (v) => (isNum(v) ? Math.round(v / 1e4).toLocaleString('en-US') + '만' : '–');
const TESTS = isNum(D.tests) ? D.tests : null;
const SMALL = WS ? WS.choices.filter((c) => c.method === '공통최적').map((c) => c.group) : [];
const MINSKU = (() => {
  const m = WS && String(WS.criteria['작은 그룹 규칙'] || '').match(/(\d+)\s*개?\s*미만/);
  return m ? Number(m[1]) : 5;
})();
const WRANGE = D.weights.length
  ? `${wfmt(Math.min(...D.weights.map((w) => w.wh)))}~${wfmt(Math.max(...D.weights.map((w) => w.wh)))}` : '–';

// ---------------------------------------------------------------- 프레젠테이션
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = TASK;
pres.subject = 'AI 탐색으로 찾은 가중평균 수요로 안전재고일수 산출·검증';
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
let page = 0;

function T(s, text, o) {
  s.addText(text, Object.assign({ fontFace: FONT, color: C.text, fontSize: 12, margin: 0, valign: 'top' }, o));
}
function R(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } }, o));
}
function RR(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: Math.min(0.08, h / 2), fill: { color: fill }, line: { color: fill, width: 0 } }, o));
}
function box(s, text, x, y, w, h, o = {}) {
  const shape = o.shape || (o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE);
  const opt = Object.assign({
    x, y, w, h, fontFace: FONT, fontSize: 12, color: C.text, align: 'center', valign: 'middle', margin: 4,
    fill: { color: o.fill || C.light }, line: { color: o.lineColor || o.fill || C.light, width: o.lineWidth || 0.75 },
  }, o, { shape });
  if (o.round) opt.rectRadius = typeof o.round === 'number' ? o.round : Math.min(0.08, h / 2);
  if (o.dash) opt.line = { color: o.lineColor || C.amber, width: 1.25, dashType: 'dash' };
  delete opt.round; delete opt.lineColor; delete opt.lineWidth; delete opt.dash;
  s.addText(text, opt);
}
function arrow(s, x, y, dx, dy, color = C.slate, o = {}) {
  s.addShape(pres.shapes.LINE, {
    x: dx >= 0 ? x : x + dx, y: dy >= 0 ? y : y + dy, w: Math.abs(dx), h: Math.abs(dy),
    flipH: dx < 0, flipV: dy < 0,
    line: Object.assign({ color, width: 1.5, endArrowType: 'triangle' }, o),
  });
}
function line(s, x, y, dx, dy, color = C.line, o = {}) {
  s.addShape(pres.shapes.LINE, {
    x: dx >= 0 ? x : x + dx, y: dy >= 0 ? y : y + dy, w: Math.abs(dx), h: Math.abs(dy), flipH: dx < 0, flipV: dy < 0,
    line: Object.assign({ color, width: 1 }, o),
  });
}
function num(s, n, x, y, d, fill = C.navy, color = C.white, size) {
  box(s, String(n), x, y, d, d, { shape: pres.shapes.OVAL, fill, color, bold: true, fontSize: size || Math.round(d * 30), margin: 0 });
}
// 줄 목록(각 줄을 문단으로)
const lines = (arr, o = {}) => arr.map((t, i) => ({ text: t, options: Object.assign({ breakLine: i < arr.length - 1 }, o) }));

// 도형으로 그린 아이콘(원 안). 외부 그림 파일 없이 사내 PC에서도 깨지지 않는다.
function icon(s, kind, x, y, d, bg = C.navy, fg = C.white) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg, width: 0 } });
  const g = d * 0.5, gx = x + d * 0.25, gy = y + d * 0.25;
  const fill = { color: fg }, ln = { color: fg, width: 0 };
  const sh = (type, o) => s.addShape(type, Object.assign({ fill, line: ln }, o));
  const hole = (o) => s.addShape(pres.shapes.RECTANGLE, Object.assign({ fill: { color: bg }, line: { color: bg, width: 0 } }, o));
  const person = (px, py, sc) => {
    sh(pres.shapes.OVAL, { x: px + g * 0.32 * sc, y: py, w: g * 0.36 * sc, h: g * 0.36 * sc });
    sh(pres.shapes.ROUNDED_RECTANGLE, { x: px + g * 0.14 * sc, y: py + g * 0.44 * sc, w: g * 0.72 * sc, h: g * 0.5 * sc, rectRadius: g * 0.22 * sc });
  };
  switch (kind) {
    case 'user': person(gx, gy + g * 0.02, 1); break;
    case 'users': person(gx - g * 0.28, gy + g * 0.12, 0.85); person(gx + g * 0.42, gy + g * 0.12, 0.85); break;
    case 'chart': [0.45, 0.75, 1.0].forEach((hh, i) => sh(pres.shapes.RECTANGLE, { x: gx + g * (0.05 + i * 0.34), y: gy + g * (1 - hh), w: g * 0.24, h: g * hh })); break;
    case 'table':
      for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) sh(pres.shapes.RECTANGLE, { x: gx + g * (0.02 + c * 0.34), y: gy + g * (0.05 + r * 0.32), w: g * 0.28, h: g * 0.24 });
      break;
    case 'check':
      s.addShape(pres.shapes.LINE, { x: gx + g * 0.08, y: gy + g * 0.5, w: g * 0.3, h: g * 0.3, line: { color: fg, width: d * 7 } });
      s.addShape(pres.shapes.LINE, { x: gx + g * 0.38, y: gy + g * 0.12, w: g * 0.55, h: g * 0.68, flipV: true, line: { color: fg, width: d * 7 } });
      break;
    case 'alert':
      sh(pres.shapes.ISOSCELES_TRIANGLE, { x: gx - g * 0.04, y: gy + g * 0.02, w: g * 1.08, h: g * 0.92 });
      T(s, '!', { x: gx, y: gy + g * 0.22, w: g, h: g * 0.7, fontSize: Math.round(d * 22), bold: true, color: bg, align: 'center', valign: 'middle' });
      break;
    case 'gear':
      sh(pres.shapes.GEAR_6, { x: gx, y: gy, w: g, h: g });
      s.addShape(pres.shapes.OVAL, { x: gx + g * 0.33, y: gy + g * 0.33, w: g * 0.34, h: g * 0.34, fill: { color: bg }, line: { color: bg, width: 0 } });
      break;
    case 'db': sh(pres.shapes.CAN, { x: gx + g * 0.1, y: gy, w: g * 0.8, h: g }); break;
    case 'doc':
      // FOLDED_CORNER는 pptxgenjs가 잘못된 이름(folderCorner)으로 써서 PowerPoint가 열지 못한다 → 사각형으로 그림
      sh(pres.shapes.RECTANGLE, { x: gx + g * 0.12, y: gy, w: g * 0.76, h: g });
      [0.3, 0.5, 0.7].forEach((yy) => hole({ x: gx + g * 0.25, y: gy + g * yy, w: g * 0.5, h: g * 0.07 }));
      break;
    case 'search':
      sh(pres.shapes.DONUT, { x: gx, y: gy, w: g * 0.72, h: g * 0.72 });
      s.addShape(pres.shapes.LINE, { x: gx + g * 0.62, y: gy + g * 0.62, w: g * 0.33, h: g * 0.33, line: { color: fg, width: d * 8 } });
      break;
    case 'flow':
      sh(pres.shapes.RECTANGLE, { x: gx, y: gy + g * 0.05, w: g * 0.36, h: g * 0.3 });
      sh(pres.shapes.RECTANGLE, { x: gx + g * 0.64, y: gy + g * 0.05, w: g * 0.36, h: g * 0.3 });
      sh(pres.shapes.RECTANGLE, { x: gx + g * 0.32, y: gy + g * 0.65, w: g * 0.36, h: g * 0.3 });
      s.addShape(pres.shapes.LINE, { x: gx + g * 0.36, y: gy + g * 0.2, w: g * 0.28, h: 0, line: { color: fg, width: d * 3 } });
      s.addShape(pres.shapes.LINE, { x: gx + g * 0.5, y: gy + g * 0.2, w: 0, h: g * 0.45, line: { color: fg, width: d * 3 } });
      break;
    case 'calendar':
      sh(pres.shapes.RECTANGLE, { x: gx, y: gy + g * 0.08, w: g, h: g * 0.86 });
      hole({ x: gx + g * 0.1, y: gy + g * 0.36, w: g * 0.8, h: g * 0.48 });
      for (let i = 0; i < 3; i++) sh(pres.shapes.RECTANGLE, { x: gx + g * (0.18 + i * 0.25), y: gy + g * 0.46, w: g * 0.14, h: g * 0.12 });
      break;
    case 'lock':
      sh(pres.shapes.RECTANGLE, { x: gx + g * 0.1, y: gy + g * 0.45, w: g * 0.8, h: g * 0.55 });
      s.addShape(pres.shapes.BLOCK_ARC, { x: gx + g * 0.22, y: gy, w: g * 0.56, h: g * 0.9, fill, line: ln, angleRange: [180, 0], arcThicknessRatio: 0.25 });
      break;
    case 'target':
      sh(pres.shapes.DONUT, { x: gx, y: gy, w: g, h: g });
      sh(pres.shapes.OVAL, { x: gx + g * 0.32, y: gy + g * 0.32, w: g * 0.36, h: g * 0.36 });
      break;
    default: break;
  }
}

function tag(s, text, kind = 'light') {
  const st = {
    dark: { fill: C.navy, color: C.white }, teal: { fill: C.teal, color: C.white }, light: { fill: 'E6EBF0', color: C.navy },
    amber: { fill: C.amber, color: C.white }, outline: { fill: C.navy, color: C.white, lineColor: C.sky },
  }[kind];
  box(s, text, 10.78, 0.5, 1.95, 0.44, { round: 0.22, fill: st.fill, color: st.color, lineColor: st.lineColor || st.fill, bold: true, fontSize: 12 });
}
function footer(s, source, dark = false) {
  const col = dark ? C.sky : C.gray;
  T(s, TASK, { x: MX, y: 7.05, w: 4.1, h: 0.26, fontSize: 9, color: col, valign: 'middle' });
  T(s, '근거: ' + source, { x: 4.7, y: 7.05, w: 7.45, h: 0.26, fontSize: 9, color: col, align: 'right', valign: 'middle' });
  T(s, String(page), { x: 12.25, y: 7.05, w: 0.48, h: 0.26, fontSize: 10, color: col, align: 'right', valign: 'middle', bold: true });
}
function slide({ title, sub, tagText, tagKind, source }) {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.white };
  T(s, title, { x: MX, y: 0.4, w: 10.0, h: 0.62, fontSize: 24, bold: true, color: C.navy, valign: 'middle' });
  if (sub) T(s, sub, { x: MX, y: 1.04, w: 10.0, h: 0.36, fontSize: 13.5, color: C.slate, valign: 'middle' });
  tag(s, tagText, tagKind);
  footer(s, source);
  return s;
}
function label(s, text, x, y, w, color = C.navy, size = 14) {
  T(s, text, { x, y, w, h: 0.32, fontSize: size, bold: true, color, valign: 'middle' });
}
function bars(s, { x, y, w, h, title, items, max, fmt = pct, better }) {
  T(s, title, { x, y, w, h: 0.3, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
  if (better) T(s, better, { x, y: y + 0.3, w, h: 0.26, fontSize: 10.5, color: C.gray, valign: 'middle' });
  const top = y + 0.85, base = y + h - 0.62, ph = base - top;
  const slot = w / items.length, bw = Math.min(0.72, slot * 0.58);
  line(s, x, base, w, 0, C.mute, { width: 1 });
  items.forEach((it, i) => {
    const bx = x + i * slot + (slot - bw) / 2;
    if (isNum(it.value)) {
      const bh = Math.max(0.02, (it.value / max) * ph);
      R(s, bx, base - bh, bw, bh, it.color);
      T(s, fmt(it.value), { x: x + i * slot, y: base - bh - 0.32, w: slot, h: 0.28, fontSize: 13, bold: true, color: it.strong ? it.color : C.text, align: 'center', valign: 'bottom' });
    } else {
      box(s, it.empty || '없음', bx - 0.05, base - 0.75, bw + 0.1, 0.72, { fill: C.white, lineColor: C.line, color: C.gray, fontSize: 10.5 });
    }
    T(s, it.label, { x: x + i * slot, y: base + 0.06, w: slot, h: 0.5, fontSize: 11, color: C.text, align: 'center', valign: 'top' });
  });
}
function table(s, rows, o) {
  const head = rows[0].map((c) => (typeof c === 'string'
    ? { text: c, options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: o.headSize || 12 } } : c));
  s.addTable([head, ...rows.slice(1)], Object.assign({
    fontFace: FONT, fontSize: 11, color: C.text, valign: 'middle',
    border: { type: 'solid', pt: 0.75, color: C.line }, margin: [0.04, 0.1, 0.04, 0.1],
  }, o));
}
const cellOpt = (text, opt = {}) => ({ text, options: opt });
// 검증 시트의 '결품'은 실제 결품이 아니라 '목표재고량 < 실제 리드타임 수요'(대리 지표)라서 화면에는 '미방어'로 쓴다
const cov = (v) => (v === '결품' ? '미방어' : v);

// 간지: 세 질문(왜 / 무엇을 / 좋아졌나)과 앞으로를 나눈다
const SECTIONS = [
  ['왜 만들었나', '사용자 · 현재 방식의 한계 · 바뀌는 것 · 계획 대비'],
  ['어떤 기술로 만들었나', 'AI 기술 구성 · 파이프라인 · 산식 · 사례 · 재고등급 · 품질 장치'],
  ['최적 가중치를 어떻게 찾았나', '탐색 설계 · 후보 전체 채점 · 탐색 결과'],
  ['실제 출고로 확인했나', '실험 설계 · 결과 1~3 · 해석과 한계'],
  ['앞으로 무엇을 하나', 'MLOps 운영안 · 다음 단계 · 달성도 · 정리'],
];
function divider(i) {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.navy };
  T(s, String(i + 1).padStart(2, '0'), { x: 0.9, y: 2.0, w: 2.2, h: 1.2, fontSize: 66, bold: true, color: C.teal, valign: 'middle' });
  T(s, SECTIONS[i][0], { x: 0.9, y: 3.2, w: 11, h: 0.9, fontSize: 36, bold: true, color: C.white, valign: 'middle' });
  T(s, SECTIONS[i][1], { x: 0.9, y: 4.15, w: 11, h: 0.45, fontSize: 16, color: C.sky, valign: 'middle' });
  // 진행 표시
  SECTIONS.forEach((sec, k) => {
    const x = 0.9 + k * 2.34;
    R(s, x, 5.55, 2.2, 0.07, k === i ? C.teal : '3A5068');
    T(s, `${k + 1}. ${sec[0]}`, { x, y: 5.7, w: 2.2, h: 0.6, fontSize: 11, color: k === i ? C.white : C.sky, bold: k === i, valign: 'top' });
  });
  tag(s, `${i + 1} / ${SECTIONS.length}`, 'outline');
  footer(s, '발표자료 설계 §3 흐름', true);
}

// ======================================================================= 표지
(function cover() {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.navy };
  T(s, '결과 보고 · 구매·SCM·MP 시스템 담당', { x: 0.9, y: 1.05, w: 9, h: 0.4, fontSize: 15, color: C.sky });
  T(s, lines(['MP 시스템 상품 안전재고', '산출 Logic 변경']),
    { x: 0.9, y: 1.65, w: 11.5, h: 1.9, fontSize: 42, bold: true, color: C.white, valign: 'middle', lineSpacingMultiple: 1.05 });
  T(s, 'AI 탐색으로 찾은 가중평균 수요로 SKU별 안전재고일수를 산출·검증', { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, color: 'D5DEE7' });
  R(s, 0.9, 4.75, 2.6, 0.06, C.teal);
  T(s, D.run_date, { x: 0.9, y: 5.0, w: 6, h: 0.4, fontSize: 15, bold: true, color: C.white });
  T(s, `가중치 ${man(SPACE)} 조합 가운데 실제 출고에 가장 가까운 값을 찾고, 고를 때 쓰지 않은 실제 출고로 확인했습니다.`,
    { x: 0.9, y: 5.45, w: 11.5, h: 0.36, fontSize: 12.5, color: C.sky });
  tag(s, '결과 보고', 'outline');
  footer(s, 'presentation/data.json (out/ 결과 파일에서 추출)', true);
})();

// ======================================================================= 한 장 요약
(function summary() {
  const s = slide({
    title: `가중치 ${man(SPACE)} 조합 가운데 최적 가중평균을 찾았습니다`,
    sub: '누구를 위해, 무엇을, 어떻게 만들었고 지금 어디까지 왔는가',
    tagText: '요약', tagKind: 'dark', source: `${WS ? WS.source + ' › 선택가중치·검증비교 · ' : ''}${SRC} › 산출결과`,
  });
  const cards = [
    ['users', '누구를 위해', ['구매 담당과 SCM(MP 시뮬레이션) 담당', '매주 SKU별 안전재고 일수를 정해 MP에 올리는 사람']],
    ['table', '무엇을', [`고정값(${oldLevels.join('·')}일)인 MP 일별 안전재고를`, 'SKU별 최적 가중평균 수요로 산출한 값으로']],
    ['search', '어떻게 (AI 탐색)', [`그룹별 가중치 후보 ${CANDS.length}개를 모두 채점해 최적값 선택`, '고를 때 쓰지 않은 구간의 실제 출고로 확인']],
  ];
  const cw = (CW - 0.4) / 3;
  cards.forEach(([ic, head, body], i) => {
    const x = MX + i * (cw + 0.2), y = 1.7;
    RR(s, x, y, cw, 2.05, C.light);
    icon(s, ic, x + 0.28, y + 0.26, 0.66, i === 2 ? C.teal : C.navy);
    T(s, head, { x: x + 1.1, y: y + 0.34, w: cw - 1.3, h: 0.5, fontSize: 17, bold: true, color: C.navy, valign: 'middle' });
    T(s, lines(body), { x: x + 0.3, y: y + 1.06, w: cw - 0.5, h: 0.92, fontSize: 12.5, color: C.text, paraSpaceAfter: 4 });
  });
  label(s, `지금까지 · 탐색 검증 구간 ${HOLD} · 안전재고일수는 기존값이 있는 ${SS.n} SKU`, MX, 3.92, 10);
  const kw = (CW - 0.6) / 4;
  const kpis = [
    { v: man(SPACE), l: '탐색한 가중치 조합', n: `품목그룹 ${NG}개 × 후보 ${CANDS.length}개 (${CANDS.length}의 ${NG}제곱)`, c: C.teal },
    { v: wsReady ? pct(wsNew.wape) : PENDING, l: 'AI 탐색 후 수요 오차율', n: wsReady ? `현재 가중치 ${pct(wsCur.wape)}보다 ${pp(wsCur.wape, wsNew.wape)} 낮음` : '', c: C.teal },
    { v: `${f1(SS.new_median)}일`, l: '산식 안전재고일수 중앙값', n: `기존 고정값 중앙값 ${f1(SS.old_median)}일 · SKU마다 다른 값`, c: C.teal },
    { v: wsReady ? pct(wsNew.cover) : PENDING, l: '결품 방어율 (탐색 후)', n: wsReady ? `탐색 전 ${pct(wsCur.cover)} · 서비스 수준(Z) 조정이 다음 과제` : '', c: C.red, warn: true },
  ];
  kpis.forEach((k, i) => {
    const x = MX + i * (kw + 0.2), y = 4.3;
    RR(s, x, y, kw, 2.08, k.warn ? C.redL : C.white, { line: { color: k.warn ? 'E8B4AD' : C.line, width: 1 } });
    if (k.warn) box(s, '주의', x + kw - 0.85, y + 0.16, 0.65, 0.3, { round: 0.15, fill: C.red, color: C.white, fontSize: 10.5, bold: true });
    T(s, k.v, { x: x + 0.25, y: y + 0.2, w: kw - 0.5, h: 0.8, fontSize: 32, bold: true, color: k.c, valign: 'middle' });
    T(s, k.l, { x: x + 0.25, y: y + 1.02, w: kw - 0.5, h: 0.38, fontSize: 13.5, bold: true, color: C.navy, valign: 'middle' });
    T(s, k.n, { x: x + 0.25, y: y + 1.42, w: kw - 0.5, h: 0.58, fontSize: 11.5, color: C.slate });
  });
  T(s, lines([
    wsReady ? `오차율 개선은 ${topContribText}에 몰려 있고, 결품 방어율은 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}입니다 · 오차율(WAPE) = Σ|예측−실제| ÷ Σ실제` : '오차율(WAPE) = Σ|예측−실제| ÷ Σ실제',
    '결품 방어율(대리 지표): 목표재고량이 실제 리드타임 수요를 덮은 SKU 비율, 실제 결품 기록이 아님 (용어는 부록 1)',
  ]), { x: MX, y: 6.44, w: CW, h: 0.52, fontSize: 10.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 간지 1
divider(0);

// ======================================================================= 사용자
(function users() {
  const s = slide({
    title: '두 담당자 모두 가중치를 얼마로 둘지 근거가 필요합니다',
    sub: '결과를 쓰는 사람과 그들이 매주 겪는 불편 (요청 자료와 설계 문서에서 정리한 업무 가정)',
    tagText: '업무 가정', tagKind: 'light', source: '설계 문서 2026-11-19 §1~2 · 기준 파일 시트 구성',
  });
  const people = [
    { ic: 'user', name: '구매 담당', role: '구매마스터·발주 점검', dark: true, rows: [
      ['맡은 일', '구매마스터 관리: 가중치, 재고등급, 리드타임, 일별 안전재고 값'],
      ['반복 업무', '품목그룹별 실적·FCST 가중치와 재고등급을 정해 발주 점검'],
      ['불편', '가중치를 바꾸면 결과가 어떻게 달라지는지 미리 확인할 방법이 없음'],
    ] },
    { ic: 'gear', name: 'SCM 담당', role: 'MP 시뮬레이션 구동', dark: false, rows: [
      ['맡은 일', '매주 MP 시뮬레이션 구동: SCM FCST, 기말재고 예상'],
      ['반복 업무', '43주차 구동 결과를 44주차 발주 점검에 사용'],
      ['불편', `일별 안전재고가 고정값(${oldLevels.join('·')}일)이라 수요 변화가 반영되지 않음`],
    ] },
  ];
  const cw = (CW - 0.3) / 2;
  people.forEach((p, i) => {
    const x = MX + i * (cw + 0.3), y = 1.7, h = 3.55;
    RR(s, x, y, cw, h, p.dark ? C.navy : C.light);
    icon(s, p.ic, x + 0.35, y + 0.3, 0.85, C.teal);
    T(s, p.name, { x: x + 1.45, y: y + 0.3, w: cw - 1.7, h: 0.5, fontSize: 21, bold: true, color: p.dark ? C.white : C.navy, valign: 'middle' });
    T(s, p.role, { x: x + 1.45, y: y + 0.78, w: cw - 1.7, h: 0.36, fontSize: 13, color: p.dark ? C.sky : C.slate, valign: 'middle' });
    p.rows.forEach(([k, v], r) => {
      const ry = y + 1.45 + r * 0.68;
      T(s, k, { x: x + 0.38, y: ry, w: 1.2, h: 0.6, fontSize: 13, bold: true, color: p.dark ? '8FC9C9' : C.teal, valign: 'middle' });
      T(s, v, { x: x + 1.62, y: ry, w: cw - 1.9, h: 0.6, fontSize: 12.5, color: p.dark ? C.white : C.text, valign: 'middle' });
    });
  });
  const y2 = 5.5;
  RR(s, MX, y2, cw, 1.2, C.tealL);
  icon(s, 'db', MX + 0.3, y2 + 0.27, 0.66, C.teal);
  T(s, 'MP 시스템 담당', { x: MX + 1.15, y: y2 + 0.18, w: cw - 1.4, h: 0.38, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  T(s, '산출한 안전재고일수를 일별 안전재고 값으로 MP에 반영', { x: MX + 1.15, y: y2 + 0.58, w: cw - 1.4, h: 0.45, fontSize: 12.5, color: C.text, valign: 'middle' });
  const x3 = MX + cw + 0.3;
  RR(s, x3, y2, cw, 1.2, C.redL);
  icon(s, 'alert', x3 + 0.3, y2 + 0.27, 0.66, C.red);
  T(s, '이번 범위가 아닌 것', { x: x3 + 1.15, y: y2 + 0.18, w: cw - 1.4, h: 0.38, fontSize: 14, bold: true, color: C.red, valign: 'middle' });
  T(s, '결품 확률·악성 재고 위험·발주 우선순위, MP DB 직접 연동, 화면 개발', { x: x3 + 1.15, y: y2 + 0.58, w: cw - 1.4, h: 0.45, fontSize: 12.5, color: C.text, valign: 'middle' });
})();

// ======================================================================= 현재 방식의 한계
(function asis() {
  const s = slide({
    title: '지금 방식은 수요가 바뀌어도 안전재고가 그대로입니다',
    sub: '현재 방식의 한계 세 가지와, 이번 과제가 답하려는 질문',
    tagText: '현재 방식', tagKind: 'light',
    source: `${SRC} › 산출결과(기존안전재고일수·실적가중치)${WS ? ' · ' + WS.source + ' › 후보성능' : ''}`,
  });
  const cards = [
    { ic: 'calendar', h: '고정 안전재고일수', big: `${oldLevels.length}종`, cap: `${oldLevels.join('·')}일 (${D.days.old_safety_n} SKU)`,
      body: '수요 크기·흔들림·리드타임과 관계없이 정해 둔 일수가 그대로 MP에 들어갑니다' },
    { ic: 'gear', h: '정해 둔 가중치', big: WRANGE, cap: `품목그룹별 실적 가중치 (현재 ${D.weights.length}개 그룹)`,
      body: '실적과 FCST를 섞는 비중이 그룹마다 다르지만, 그 값이 최선인지 실제 출고로 확인하는 절차가 없습니다' },
    { ic: 'chart', h: '많은 경우의 수', big: man(SPACE), cap: `그룹 ${NG}개 × 후보 ${CANDS.length}개의 조합`,
      body: '가중치 조합마다 안전재고가 달라집니다. 모든 조합을 사람이 계산해 비교하기는 어렵습니다' },
  ];
  const cw = (CW - 0.6) / 3, y = 1.72, h = 3.7;
  cards.forEach((c, i) => {
    const x = MX + i * (cw + 0.3);
    RR(s, x, y, cw, h, C.light);
    icon(s, c.ic, x + 0.3, y + 0.3, 0.7, C.navy);
    T(s, c.h, { x: x + 1.2, y: y + 0.3, w: cw - 1.4, h: 0.7, fontSize: 17, bold: true, color: C.navy, valign: 'middle' });
    T(s, c.big, { x: x + 0.3, y: y + 1.2, w: cw - 0.6, h: 0.75, fontSize: 34, bold: true, color: C.amber, valign: 'middle' });
    T(s, c.cap, { x: x + 0.3, y: y + 1.95, w: cw - 0.6, h: 0.34, fontSize: 12, color: C.slate, valign: 'middle' });
    line(s, x + 0.3, y + 2.42, cw - 0.6, 0, C.line);
    T(s, c.body, { x: x + 0.3, y: y + 2.55, w: cw - 0.6, h: 1.0, fontSize: 12.5, color: C.text });
  });
  RR(s, MX, 5.62, CW, 1.15, C.navy);
  icon(s, 'target', MX + 0.3, 5.84, 0.7, C.teal);
  T(s, [
    { text: '이번 과제의 질문', options: { bold: true, color: '8FC9C9', fontSize: 13, breakLine: true } },
    { text: '여러 경우의 수 가운데 실제 출고를 가장 잘 맞히는 가중평균은 무엇이고, 그 값으로 만든 안전재고일수는 충분했나?', options: { bold: true, color: C.white, fontSize: 15 } },
  ], { x: MX + 1.25, y: 5.68, w: CW - 1.5, h: 1.03, valign: 'middle', paraSpaceAfter: 3 });
})();

// ======================================================================= Before / After
(function beforeAfter() {
  const s = slide({
    title: '가중치를 데이터로 고르고, 결과를 실제 출고로 채점합니다',
    sub: '같은 업무를 Before(현재 방식)와 After(AI 기반 산출)로 나란히 놓고, 이번 데이터의 근거 수치를 붙였습니다',
    tagText: '구현', tagKind: 'teal', source: `${SRC} › 산출결과·재고등급·검증요약${WS ? ' · ' + WS.source : ''}`,
  });
  const rows = [
    ['가중치 결정', '그룹별로 정해 둔 값을 계속 사용', `후보 ${CANDS.length}개를 모두 채점해 오차가 가장 작은 값 선택`,
      wsReady ? `${pct(wsCur.wape)} → ${pct(wsNew.wape)}` : PENDING, `검증 구간 ${HOLD} 수요 오차율`],
    ['일별 안전재고', `고정값 ${oldLevels.length}종 (${D.days.old_safety_n} SKU)`, 'SKU마다 수요 편차·리드타임·등급으로 계산', `${f1(SS.new_min)}~${f1(SS.new_max)}일`, `산식 안전재고일수 (${SS.n} SKU)`],
    ['서비스 수준', '품목그룹마다 같은 일수 (등급과 무관)', '재고등급(단가·회전율·리드타임·품질)별 Z값 적용', `${D.grades.match}/${D.grades.total}`, '등급 규칙이 기존 등급과 일치'],
    ['성능 확인', '맞았는지 확인하지 않음', '이후 9주 실제 출고로 오차율·결품 방어율 채점', pct(M['가중평균'].wape), `사후 검증 ${VAL} 오차율`],
    ['이상 값', '조용히 섞임', '점검메모 10종과 채점 제외 사유를 기록', `${D.memos.with_memo} SKU`, `메모가 붙은 SKU (${D.memos.total}개 중)`],
    ['운영', '매주 사람이 다시 계산', '명령 한 번으로 산출·탐색·검증·리포트 재현', `${int(TESTS)}개`, '자동 테스트로 규칙 고정'],
  ];
  const y0 = 1.72;
  const cx = { item: MX, before: MX + 1.75, arrow: MX + 5.6, after: MX + 6.05, ev: MX + 9.95 };
  const cwid = { item: 1.6, before: 3.75, after: 3.75, ev: CW - 9.95 };
  T(s, 'Before · 현재 방식', { x: cx.before, y: y0, w: cwid.before, h: 0.34, fontSize: 13, bold: true, color: C.red, valign: 'middle' });
  T(s, 'After · AI 기반 산출', { x: cx.after, y: y0, w: cwid.after, h: 0.34, fontSize: 13, bold: true, color: C.teal, valign: 'middle' });
  T(s, '근거 수치', { x: cx.ev, y: y0, w: cwid.ev, h: 0.34, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
  const rh = 0.7, step = 0.77;
  rows.forEach((r, i) => {
    const y = y0 + 0.45 + i * step;
    box(s, r[0], cx.item, y, cwid.item, rh, { fill: C.navy, color: C.white, bold: true, fontSize: 13, round: 0.06 });
    box(s, r[1], cx.before, y, cwid.before, rh, { fill: C.redL, fontSize: 12.5, align: 'left', margin: 10, round: 0.06 });
    arrow(s, cx.arrow + 0.02, y + rh / 2, 0.36, 0, C.teal, { width: 2 });
    box(s, r[2], cx.after, y, cwid.after, rh, { fill: C.tealL, fontSize: 12.5, align: 'left', margin: 10, round: 0.06, bold: true, color: C.navy });
    T(s, r[3], { x: cx.ev + 0.1, y: y + 0.02, w: cwid.ev - 0.1, h: 0.38, fontSize: 17, bold: true, color: C.teal, valign: 'middle' });
    T(s, r[4], { x: cx.ev + 0.1, y: y + 0.4, w: cwid.ev - 0.1, h: 0.3, fontSize: 10.5, color: C.gray, valign: 'middle' });
  });
})();

// ======================================================================= 계획 대비 재정의
(function replan() {
  const s = slide({
    title: '계획을 3단계로 나누고, 데이터 제약에 맞게 고쳤습니다',
    sub: '범위를 줄인 것이 아니라, 확인된 데이터와 업무에 맞게 구체화했습니다',
    tagText: '계획 대비', tagKind: 'light', source: 'docs/superpowers/specs 설계 문서 3종(1·2·3단계)',
  });
  const hl = (t) => cellOpt(t, { bold: true, color: C.teal });
  const rows = [
    ['항목', '원래 계획', '실제로 한 것', '바꾼 이유'],
    [cellOpt('1단계 범위', { bold: true }), '가중평균 기준 일수 산출', hl('세 기준 산출 + 사후 검증 잣대'), '개선 효과를 잴 잣대를 먼저 만들어야 전후를 비교할 수 있음'],
    [cellOpt('2단계 재고등급', { bold: true }), '회전율·리드타임·품질 등급까지 모두 자동', hl('단가·회전율 등급, 점수, 종합, 재고등급 자동 (리드타임·품질 등급은 입력값)'), '리드타임 일수만으로는 등급이 정해지지 않음 (같은 FLT 14일이 A와 C로 나뉨)'],
    [cellOpt('3단계 가중치', { bold: true }), '과거 W01~22로 고르고 W23~31로 확인', hl(`AI 탐색: 후보 ${CANDS.length}개 전체 채점, ${SEL}로 고르고 ${HOLD}로 확인`), 'W23~31 시점의 FCST 스냅샷이 없어 과거 구간을 쓸 수 없음'],
    [cellOpt('운영 방식', { bold: true }), '수동 실행', hl('명령 한 번으로 재현되는 파이프라인 + 자동 테스트 (MLOps 기반)'), '매주 같은 잣대로 다시 고르고 확인하려면 재현성이 먼저'],
    [cellOpt('범위 밖', { bold: true }), '—', '결품 확률·악성 재고 위험·MP DB 연동은 후속', '검증 잣대를 먼저 확보'],
  ];
  table(s, rows, { x: MX, y: 1.75, w: CW, colW: [1.75, 2.9, 3.85, CW - 1.75 - 2.9 - 3.85], rowH: [0.5, 0.82, 0.82, 0.82, 0.82, 0.62], fontSize: 12 });
})();

// ======================================================================= 간지 2
divider(1);

// ======================================================================= AI 기술 구성
(function aiTech() {
  const s = slide({
    title: 'AI·머신러닝 기법 여섯 가지를 조합했습니다',
    sub: '각 기법이 이 과제에서 맡은 일과 근거 · 오른쪽 아래 표시는 구현 여부',
    tagText: '기술 구성', tagKind: 'teal',
    source: `src/safety_stock (weight_search · inventory_days) · tests/ 자동 테스트 ${int(TESTS)}개${WS ? ' · ' + WS.source + ' › 적용기준' : ''}`,
  });
  const cards = [
    ['search', 'Grid Search', '하이퍼파라미터 탐색', `실적 가중치 w를 0.0~1.0(후보 ${CANDS.length}개)로 바꿔 가며 품목그룹 ${NG}개를 모두 채점합니다`, `탐색 공간 ${man(SPACE)} 조합`, '구현'],
    ['calendar', 'Time-series Holdout', '시계열 검증 (백테스트)', `고르는 구간(${SEL})과 확인 구간(${HOLD})을 나눠 미래 정보가 선택에 섞이지 않게 합니다`, '확인 구간은 선택에 미사용', '구현'],
    ['lock', 'Regularization · Pooling', '과적합 방지', `SKU ${MINSKU}개 미만 그룹은 전체를 합쳐 고른 w를 써서 작은 표본의 우연을 걸러 냅니다`, `${SMALL.join('·')} 적용`, '구현'],
    ['chart', 'Statistical Safety Stock', '통계적 안전재고 모형', '수요 편차(σ)·리드타임·서비스 수준(Z)으로 SKU별 안전재고와 안전재고일수를 계산합니다', `${D.sku_total} SKU 산출`, '구현'],
    ['flow', 'MLOps Pipeline', 'MLOps 파이프라인', '데이터 검사 → 산출 → 탐색 → 검증 → 리포트를 명령 한 번으로 재현하고 버전으로 관리합니다', `자동 테스트 ${int(TESTS)}개`, '구현 + 제안'],
    ['gear', 'Generative AI Agent', '생성형 AI 개발 에이전트', '설계·코드·테스트·독립 검토를 AI 에이전트(Claude)가 맡고, 사람은 판단과 승인을 맡습니다', '단계마다 별도 AI 검토', '구현'],
  ];
  const cw = (CW - 0.4) / 3, ch = 2.25;
  cards.forEach(([ic, en, ko, body, ev, st], i) => {
    const x = MX + (i % 3) * (cw + 0.2), y = 1.62 + Math.floor(i / 3) * (ch + 0.15);
    RR(s, x, y, cw, ch, C.light);
    icon(s, ic, x + 0.22, y + 0.22, 0.62, i === 0 ? C.teal : C.navy);
    T(s, en, { x: x + 1.0, y: y + 0.2, w: cw - 1.2, h: 0.26, fontSize: 10.5, color: C.slate, valign: 'middle' });
    T(s, ko, { x: x + 1.0, y: y + 0.46, w: cw - 1.2, h: 0.4, fontSize: 15, bold: true, color: C.navy, valign: 'middle' });
    T(s, body, { x: x + 0.22, y: y + 1.0, w: cw - 0.44, h: 0.8, fontSize: 11.5, color: C.text });
    line(s, x + 0.22, y + 1.83, cw - 0.44, 0, C.line);
    T(s, ev, { x: x + 0.22, y: y + 1.88, w: cw - 1.55, h: 0.34, fontSize: 11.5, bold: true, color: C.teal, valign: 'middle' });
    box(s, st, x + cw - 1.27, y + 1.92, 1.05, 0.27, { round: 0.13, fill: st === '구현' ? C.teal : C.amber, color: C.white, bold: true, fontSize: 9.5, margin: 0 });
  });
  box(s, [
    { text: '핵심  ', options: { bold: true, color: '8FC9C9' } },
    { text: '예측 공식은 현업이 쓰던 가중평균 그대로 두고, 그 가중치를 머신러닝의 모델 선택 방식으로 고릅니다. 결과를 설명할 수 있고 MP에 바로 넣을 수 있습니다.', options: { color: C.white } },
  ], MX, 6.36, CW, 0.6, { fill: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 구성과 데이터 경계
(function architecture() {
  const s = slide({
    title: '파이프라인 한 번으로 산출·탐색·검증이 끝납니다',
    sub: '기준 데이터 → 세 영역 + AI 탐색 → 결과 xlsx · 사후 검증용 실제 출고는 계산 입력에 섞지 않습니다',
    tagText: '구현', tagKind: 'teal', source: 'src/safety_stock/ 폴더 구조 · README.md (코드 이름은 부록 5)',
  });
  const ix = MX, iw = 2.95, iy = 1.72, ih = 3.2;
  RR(s, ix, iy, iw, ih, C.light);
  T(s, '기준 데이터 (8/28 점검 파일)', { x: ix + 0.2, y: iy + 0.15, w: iw - 0.4, h: 0.4, fontSize: 13.5, bold: true, color: C.navy, valign: 'middle' });
  const inputs = ['구매마스터', 'SCM FCST (W35~47, 13주)', `WMS 출고 (${HIST})`, '현재고 (창고 합산)', '등급 입력 (리드타임·품질)'];
  inputs.forEach((t, i) => T(s, '·  ' + t, { x: ix + 0.25, y: iy + 0.68 + i * 0.47, w: iw - 0.4, h: 0.42, fontSize: 12, color: C.text, valign: 'middle' }));
  const mx = 4.05, mw = 4.25;
  const mods = [
    ['① 구매마스터 영역', '구매마스터 읽기 · 재고등급(Z값) 계산'],
    ['② MP 시뮬레이션 영역', 'SCM FCST·기말재고 예상 읽기'],
    ['③ 재고일수 영역', '출고·현재고 · 세 기준 계산 · 사후 검증'],
  ];
  mods.forEach(([h, b], i) => {
    const y = 1.72 + i * 1.1;
    RR(s, mx, y, mw, 0.95, C.navy);
    T(s, h, { x: mx + 0.22, y: y + 0.1, w: mw - 0.4, h: 0.38, fontSize: 14, bold: true, color: C.white, valign: 'middle' });
    T(s, b, { x: mx + 0.22, y: y + 0.5, w: mw - 0.4, h: 0.36, fontSize: 12, color: 'C9D3DD', valign: 'middle' });
    arrow(s, ix + iw + 0.05, y + 0.47, mx - ix - iw - 0.12, 0, C.slate);
  });
  const ox = 8.95, ow = W - MX - ox, oy = 1.72;
  RR(s, ox, oy, ow, 3.2, C.tealL);
  T(s, '결과 xlsx (주요 시트)', { x: ox + 0.2, y: oy + 0.15, w: ow - 0.4, h: 0.4, fontSize: 13.5, bold: true, color: C.teal, valign: 'middle' });
  const sheets = [['산출결과', 'SKU별 세 기준 값'], ['재고등급', '등급 계산 과정'], ['검증_SKU', '실제 출고와 비교'], ['검증요약', '기준·그룹별 지표'], ['적용기준', '산식·규칙 기록'], ['가중치탐색', '후보 채점·선택·확인']];
  sheets.forEach(([a, b], i) => {
    T(s, [{ text: a, options: { bold: true, color: C.navy } }, { text: '  ' + b, options: { color: C.slate } }],
      { x: ox + 0.25, y: oy + 0.65 + i * 0.42, w: ow - 0.4, h: 0.38, fontSize: 12, valign: 'middle' });
  });
  arrow(s, mx + mw + 0.05, 3.3, ox - mx - mw - 0.12, 0, C.teal, { width: 2 });
  line(s, MX, 5.25, CW, 0, C.red, { width: 1.5, dashType: 'dash' });
  box(s, '데이터 경계: 사후 검증 구간의 실제 출고는 계산 입력에 넣지 않는다', MX, 5.08, 6.1, 0.34, { fill: C.white, lineColor: C.white, color: C.red, bold: true, fontSize: 12.5, align: 'left', margin: 2 });
  const by = 5.6, bh = 1.15;
  RR(s, MX, by, 3.3, bh, C.light);
  T(s, 'W44 업로드 파일 › 출고실적', { x: MX + 0.2, y: by + 0.12, w: 3.0, h: 0.38, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
  T(s, `${VAL} 실제 출고 (사후 검증 전용)`, { x: MX + 0.2, y: by + 0.55, w: 3.0, h: 0.45, fontSize: 12, color: C.text, valign: 'middle' });
  arrow(s, MX + 3.35, by + bh / 2, 0.6, 0, C.red);
  RR(s, 4.6, by, 3.7, bh, C.white, { line: { color: C.red, width: 1.25 } });
  T(s, '사후 검증', { x: 4.8, y: by + 0.12, w: 3.3, h: 0.38, fontSize: 13, bold: true, color: C.red, valign: 'middle' });
  T(s, '8/24 계산값과 비교 → 검증_SKU·검증요약', { x: 4.8, y: by + 0.55, w: 3.3, h: 0.45, fontSize: 12, color: C.text, valign: 'middle' });
  // 꺾은선: 사후 검증 → 결과 시트
  line(s, 8.3, by + bh / 2, 1.05, 0, C.red, { width: 1.5 });
  arrow(s, 9.35, by + bh / 2, 0, -(by + bh / 2 - 4.97), C.red);
  T(s, 'AI · 자동화', { x: 9.75, y: by - 0.02, w: 2.98, h: 0.3, fontSize: 11, bold: true, color: C.slate, valign: 'middle' });
  [['AI 가중치 탐색', '후보 전체 채점'], ['발표 자동 생성', '숫자 직접 추출']].forEach(([a, b], i) => {
    box(s, [{ text: a, options: { bold: true, color: C.navy, breakLine: true, fontSize: 11.5 } }, { text: b, options: { color: C.slate, fontSize: 10.5 } }],
      9.75 + i * 1.52, by + 0.32, 1.46, bh - 0.32, { fill: C.white, lineColor: C.line, round: 0.06, margin: 3 });
  });
})();

// ======================================================================= 산식 한 장
(function formulas() {
  const zs = String(D.criteria['Z값'] || '').split(',').map((t) => t.trim().split('='));
  const s = slide({
    title: 'MP에 올릴 값은 안전재고일수입니다',
    sub: '가중평균 수요 → 수요 편차 → 리드타임 수요 → 안전재고 → 안전재고일수 (중간 반올림 없음)',
    tagText: '구현', tagKind: 'teal', source: `${SRC} › 적용기준 · 사용자 답변(기존 값 = MP 일별 안전재고, 리드타임 수요 미포함)`,
  });
  const chain = [
    ['가중평균 수요 (주)', '실적평균 × 실적 가중치 + FCST평균 × FCST 가중치'],
    ['수요 편차 (주)', '실적 편차와 FCST 편차를 같은 가중치로 섞음'],
    ['리드타임 수요', '가중평균 수요 × 리드타임 ÷ 7'],
    ['리드타임 수요 편차', '수요 편차 × √(리드타임 ÷ 7)'],
    ['안전재고', 'Z값 × 리드타임 수요 편차'],
    ['안전재고일수', '안전재고 ÷ 일평균 수요', 'MP 일별 안전재고를 대체할 값'],
    ['발주점', '리드타임 수요 + 안전재고', '현재고가 이 아래면 발주'],
    ['적정 재고일수', '리드타임 + 안전재고일수', '발주점을 일수로 본 값 (비교용)'],
  ];
  const bw = 2.82, bh = 1.32, gx = 0.29;
  chain.forEach(([h, f, c], i) => {
    const row = i < 4 ? 0 : 1, col = i % 4;
    const x = MX + col * (bw + gx), y = row === 0 ? 1.72 : 3.42;
    const hl = h === '안전재고일수';
    RR(s, x, y, bw, bh, hl ? C.teal : C.light, hl ? {} : {});
    T(s, h, { x: x + 0.18, y: y + 0.1, w: bw - 0.36, h: 0.32, fontSize: 13, bold: true, color: hl ? C.white : C.teal, valign: 'middle' });
    T(s, f, { x: x + 0.18, y: y + 0.44, w: bw - 0.36, h: c ? 0.5 : 0.78, fontSize: 12.5, bold: true, color: hl ? C.white : C.navy, valign: 'top' });
    if (c) T(s, c, { x: x + 0.18, y: y + 0.96, w: bw - 0.36, h: 0.3, fontSize: 11, bold: hl, color: hl ? 'E6F4F4' : C.gray, valign: 'middle' });
    if (col < 3) arrow(s, x + bw + 0.02, y + bh / 2, gx - 0.04, 0, C.slate);
  });
  line(s, MX + 3 * (bw + gx) + bw / 2, 1.72 + bh, 0, 3.2 - 1.72 - bh, C.slate, { width: 1.5 });
  line(s, MX + bw / 2, 3.2, 3 * (bw + gx), 0, C.slate, { width: 1.5 });
  arrow(s, MX + bw / 2, 3.2, 0, 0.2, C.slate);
  // 아래: Z값, 세 기준, 용어
  const y = 5.0;
  RR(s, MX, y, 3.1, 1.75, C.navy);
  T(s, 'Z값 (재고등급별 서비스 수준)', { x: MX + 0.2, y: y + 0.12, w: 2.8, h: 0.34, fontSize: 12.5, bold: true, color: C.white, valign: 'middle' });
  zs.forEach(([g, v], i) => {
    box(s, [{ text: g, options: { bold: true, fontSize: 15, breakLine: true } }, { text: v, options: { fontSize: 12 } }],
      MX + 0.2 + i * 0.93, y + 0.6, 0.83, 0.95, { fill: C.navy2, color: C.white, round: 0.06, margin: 2 });
  });
  table(s, [
    ['기준', '평균 수요 (주)', '수요 편차 (주)'],
    [cellOpt('가중평균', { bold: true, color: C.teal }), '실적·FCST 평균을 가중치로', '실적·FCST 편차를 가중치로'],
    [cellOpt('실적기준', { bold: true }), '실적평균 × 보정계수', '실적 편차'],
    [cellOpt('FCST기준', { bold: true }), 'FCST평균 × 보정계수', 'FCST 편차'],
  ], { x: 3.95, y, w: 4.6, colW: [1.1, 1.75, 1.75], rowH: 0.42, fontSize: 11, headSize: 11.5 });
  RR(s, 8.85, y, W - MX - 8.85, 1.75, C.light);
  const notes = [
    ['실적평균', `출고 ${HIST} 주평균 (빈 칸 제외)`],
    ['FCST평균', '8/24 FCST 앞 13주 평균 (빈 칸은 0)'],
    ['일평균 수요', '가중평균 수요 ÷ 7'],
    ['보정계수', '과소예측 보정계수 (현업 입력)'],
  ];
  notes.forEach(([a, b], i) => T(s, [{ text: a + '  ', options: { bold: true, color: C.navy } }, { text: b, options: { color: C.text } }],
    { x: 9.05, y: y + 0.14 + i * 0.37, w: W - MX - 9.2, h: 0.35, fontSize: 11.5, valign: 'middle' }));
})();

// ======================================================================= 사례 장(대표·반대)
function caseSlide(c, { title, sub, tagText, footRule, favorable }) {
  const s = slide({ title, sub, tagText, tagKind: 'light', source: `${SRC} › 산출결과·검증_SKU (${c.sku})` });
  const ltw = c.lt / 7, daily = c.g / 7;
  const hl = (t, extra = {}) => cellOpt(t, Object.assign({ bold: true, color: C.white, fill: { color: C.teal } }, extra));
  const rows = [
    ['단계', '계산', '값'],
    [`실적평균 (${HIST})`, '출고 주평균', f1(c.hm)],
    ['FCST평균 (13주)', '8/24 FCST 주평균', f1(c.fm)],
    ['가중평균 수요', `${f1(c.hm)} × ${wfmt(c.wh)} + ${f1(c.fm)} × ${wfmt(c.wf)}`, `${f1(c.g)} /주`],
    ['수요 편차', '실적·FCST 편차를 같은 가중치로', f1(c.sigma)],
    ['리드타임 수요', `${f1(c.g)} × ${c.lt} ÷ 7`, qty(c.k)],
    ['리드타임 수요 편차', `${f1(c.sigma)} × √(${c.lt}÷7) = ${f1(c.sigma)} × ${Math.sqrt(ltw).toFixed(2)}`, f1(c.l)],
    ['안전재고', `Z(${c.grade}등급) ${c.z} × ${f1(c.l)}`, qty(c.t)],
    [hl('안전재고일수 (MP 값)'), hl(`${qty(c.t)} ÷ (${f1(c.g)} ÷ 7 = ${f1(daily)})`), hl(`${f1(c.s)}일`, { fontSize: 13 })],
    ['발주점', `${qty(c.k)} + ${qty(c.t)}`, qty(c.rop)],
    ['적정 재고일수 (비교용)', `리드타임 ${c.lt} + ${f1(c.s)}`, `${f1(c.u)}일`],
  ];
  table(s, rows, { x: MX, y: 1.72, w: 7.3, colW: [2.15, 3.6, 1.55], rowH: 0.385, fontSize: 11.5, headSize: 11.5 });
  const rx = 8.25, rw = W - MX - rx;
  RR(s, rx, 1.72, rw, 1.95, C.light);
  T(s, '기존 MP 일별 안전재고', { x: rx + 0.22, y: 1.82, w: rw / 2 - 0.3, h: 0.3, fontSize: 12, bold: true, color: C.slate });
  T(s, `${c.old_ss_days}일`, { x: rx + 0.22, y: 2.12, w: rw / 2 - 0.3, h: 0.7, fontSize: 32, bold: true, color: C.slate, valign: 'middle' });
  T(s, `적정 재고일수 ${c.old_days}일`, { x: rx + 0.22, y: 2.9, w: rw / 2 - 0.3, h: 0.5, fontSize: 11, color: C.gray });
  T(s, '산식 안전재고일수', { x: rx + rw / 2 + 0.1, y: 1.82, w: rw / 2 - 0.3, h: 0.3, fontSize: 12, bold: true, color: C.teal });
  T(s, `${f1(c.s)}일`, { x: rx + rw / 2 + 0.1, y: 2.12, w: rw / 2 - 0.3, h: 0.7, fontSize: 32, bold: true, color: C.teal, valign: 'middle' });
  T(s, `적정 재고일수 ${f1(c.u)}일`, { x: rx + rw / 2 + 0.1, y: 2.9, w: rw / 2 - 0.3, h: 0.5, fontSize: 11, color: C.gray });
  line(s, rx + rw / 2 - 0.05, 1.9, 0, 1.6, C.line, { width: 1 });
  const vy = 3.82;
  RR(s, rx, vy, rw, 2.93, C.white, { line: { color: C.line, width: 1 } });
  T(s, `사후 검증 · ${VAL} 실제 출고`, { x: rx + 0.22, y: vy + 0.1, w: rw - 0.4, h: 0.34, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
  const vr = [
    ['실제 리드타임 수요', qty(c.actual_lt), C.text],
    ['산식 목표재고량 (발주점)', `${qty(c.new_target)} → ${cov(c.new_cover)}`, c.new_cover === '방어' ? C.teal : C.red],
    ['기존 목표재고량', `${qty(c.old_target)} → ${cov(c.old_cover)}`, c.old_cover === '방어' ? C.teal : C.red],
    ['실제 기준 재고일수', `산식 ${f1(c.new_actual_days)}일 · 기존 ${f1(c.old_actual_days)}일`, C.text],
    ['예측 주수요 대 실제', `${f1(c.g)} 대 ${f1(c.actual_week)} (${c.error >= 0 ? '+' : ''}${pct(c.error / c.actual_week, 0)})`, C.amber],
  ];
  vr.forEach(([a, b, col], i) => {
    const y = vy + 0.52 + i * 0.44;
    T(s, a, { x: rx + 0.22, y, w: 2.1, h: 0.4, fontSize: 11.5, color: C.slate, valign: 'middle' });
    T(s, b, { x: rx + 2.25, y, w: rw - 2.45, h: 0.4, fontSize: 12, bold: true, color: col, valign: 'middle', align: 'right' });
  });
  T(s, `선정 규칙: ${footRule}${favorable ? ' (산식에 유리한 사례)' : ''}`, { x: MX, y: 6.79, w: CW, h: 0.24, fontSize: 10, color: C.gray, valign: 'middle' });
  return s;
}

(function example() {
  const c = D.case;
  const s = caseSlide(c, {
    title: `대표 사례: 일별 안전재고 ${c.old_ss_days}일 → ${f1(c.s)}일`,
    sub: `${c.sku} · ${c.group} ${c.name} · 재고등급 ${c.grade} · 리드타임 ${c.lt}일 · 다음 장은 반대 사례`,
    tagText: '사례', footRule: c.rule, favorable: true,
  });
  const both = c.new_cover === '방어' && c.old_cover === '방어';
  const msg = both
    ? `두 기준 모두 리드타임 수요를 덮었고, 산식 목표재고량은 기존의 ${pct(c.new_target / c.old_target, 0)}였습니다. 단, 예측 수요는 실제보다 높았습니다.`
    : `방어 결과: 산식 ${cov(c.new_cover)}, 기존 ${cov(c.old_cover)}. 목표재고량은 산식이 기존의 ${pct(c.new_target / c.old_target, 0)}입니다.`;
  box(s, msg, MX, 6.08, 7.3, 0.64, { fill: C.navy, color: C.white, fontSize: 12, bold: true, align: 'left', margin: 10, round: 0.06 });
})();

(function counterExample() {
  const c = D.counter_case;
  const both = c.new_cover === '결품' && c.old_cover === '결품'; // 검증 시트 값 '결품' = 리드타임 수요 미방어
  const s = caseSlide(c, {
    title: `반대 사례: ${c.old_ss_days}일 → ${f1(c.s)}일, ${both ? '두 기준 모두 방어 실패' : '산식 방어 실패'}`,
    sub: `${c.sku} · ${c.group} ${c.name} · 재고등급 ${c.grade} · 리드타임 ${c.lt}일 · 방어 실패 SKU`,
    tagText: '사례', footRule: c.rule, favorable: false,
  });
  const ratio = c.actual_week / c.g;
  box(s, `실제 출고가 예측의 ${ratio.toFixed(1)}배(9주 주평균)로, ${both ? '두 기준 모두' : '산식은'} 리드타임 수요를 덮지 못했습니다(방어 실패, 실제 결품 기록은 확인하지 않음).`,
    MX, 6.08, 7.3, 0.64, { fill: C.red, color: C.white, fontSize: 12, bold: true, align: 'left', margin: 10, round: 0.06 });
})();

// ======================================================================= 안전재고일수 분포
(function distribution() {
  const s = slide({
    title: '산식 안전재고일수는 고정값보다 짧고 SKU마다 다릅니다',
    sub: `기존 MP 일별 안전재고(고정값)와 산식 안전재고일수 분포 · 두 값이 모두 있는 ${SS.n} SKU`,
    tagText: '산출 결과', tagKind: 'teal', source: `${SRC} › 산출결과 (기존안전재고일수, 가중평균_안전재고일수)`,
  });
  const cx = MX, cwid = 8.2, top = 2.45, base = 5.75;
  RR(s, cx, 1.72, cwid, 5.05, C.light);
  T(s, '안전재고일수 구간별 SKU 수', { x: cx + 0.25, y: 1.82, w: 4.5, h: 0.32, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
  [['기존 고정값', C.mute], ['산식', C.teal]].forEach(([l, col], i) => {
    const x = cx + 5.2 + i * 1.5;
    R(s, x, 1.92, 0.2, 0.14, col);
    T(s, l, { x: x + 0.27, y: 1.82, w: 1.2, h: 0.32, fontSize: 11, color: C.text, valign: 'middle' });
  });
  const mx = Math.max(...SS.old, ...SS.new) * 1.15;
  const n = SS.labels.length, slot = (cwid - 0.5) / n, bw = Math.min(0.3, slot * 0.36);
  line(s, cx + 0.25, base, cwid - 0.5, 0, C.mute);
  SS.labels.forEach((lab, i) => {
    const sx = cx + 0.25 + i * slot + (slot - 2 * bw - 0.04) / 2;
    [[SS.old[i], C.mute], [SS.new[i], C.teal]].forEach(([v, col], k) => {
      const bh = (v / mx) * (base - top);
      if (v > 0) R(s, sx + k * (bw + 0.04), base - bh, bw, bh, col);
      T(s, String(v), { x: sx + k * (bw + 0.04) - 0.1, y: base - bh - 0.26, w: bw + 0.2, h: 0.24, fontSize: 10.5, bold: true, align: 'center', valign: 'bottom', color: k ? C.teal : C.slate });
    });
    T(s, lab, { x: cx + 0.25 + i * slot, y: base + 0.06, w: slot, h: 0.28, fontSize: 10.5, color: C.text, align: 'center' });
  });
  T(s, `기존 값은 ${oldLevels.join('·')}일 네 가지뿐입니다. 산식 값은 SKU마다 다릅니다. 구간은 아래 값 포함·위 값 미포함(예: 30일은 30~40).`,
    { x: cx + 0.25, y: 6.3, w: cwid - 0.5, h: 0.36, fontSize: 11, color: C.gray, valign: 'middle' });
  const rx = cx + cwid + 0.3, rw = W - MX - rx;
  const kp = [
    [`${f1(SS.old_median)}일 → ${f1(SS.new_median)}일`, `안전재고일수 중앙값 (${SS.n} SKU)`],
    [`${SS.shorter} / ${SS.n}`, '산식이 기존보다 짧은 SKU'],
    [`${f1(SS.new_min)}~${f1(SS.new_max)}일`, `산식 범위 (${SS.n} SKU) · 전체 ${SS.all_new_n} SKU 최대 ${f1(SS.all_new_max)}일`],
  ];
  kp.forEach(([v, l], i) => {
    const y = 1.72 + i * 1.32;
    RR(s, rx, y, rw, 1.2, C.white, { line: { color: C.line, width: 1 } });
    T(s, v, { x: rx + 0.2, y: y + 0.08, w: rw - 0.4, h: 0.6, fontSize: 22, bold: true, color: C.teal, valign: 'middle' });
    T(s, l, { x: rx + 0.2, y: y + 0.68, w: rw - 0.4, h: 0.46, fontSize: 11, color: C.slate, valign: 'middle' });
  });
  box(s, '짧다고 좋은 것은 아닙니다. 리드타임 수요를 덮었는지는 사후 검증 결과에서 따로 봅니다.',
    rx, 5.7, rw, 1.07, { fill: C.amberL, color: C.navy, fontSize: 12, bold: true, align: 'left', margin: 10, round: 0.06 });
})();

// ======================================================================= 재고등급 자동화
(function grades() {
  const G = D.grades;
  const s = slide({
    title: `재고등급 규칙이 기존 등급 ${G.total}개와 모두 일치합니다`,
    sub: '단가·회전율 백분위 → 등급 → 점수 → 종합 → 재고등급 → Z값 · 리드타임·품질 등급은 현업 입력값',
    tagText: '구현', tagKind: 'teal',
    source: `${SRC} › 재고등급`,
  });
  const cols = [MX, 2.5, 4.27, 6.04, 7.4]; const cws = [1.7, 1.55, 1.55, 1.1, 1.45];
  const ys = [1.75, 2.72, 3.69, 4.66]; const h = 0.78;
  const calc = (t, x, y, w) => box(s, t, x, y, w, h, { fill: C.light, fontSize: 11.5, round: 0.06, margin: 3 });
  const inp = (t, x, y, w) => box(s, t, x, y, w, h, { fill: C.amberL, fontSize: 11.5, round: 0.06, margin: 3, bold: true });
  const sc = (t, x, y, w) => box(s, t, x, y, w, h, { fill: C.white, lineColor: C.teal, fontSize: 11.5, round: 0.06, margin: 3, color: C.teal, bold: true });
  inp('구매단가', cols[0], ys[0], cws[0]);
  calc('단가 백분위', cols[1], ys[0], cws[1]);
  calc(lines(['단가 등급', '≥0.75 A · ≥0.26 B']), cols[2], ys[0], cws[2]);
  sc('A5 B3 C1', cols[3], ys[0], cws[3]);
  inp(lines(['FCST ÷ 평균재고', '= 회전율']), cols[0], ys[1], cws[0]);
  calc('회전율 백분위', cols[1], ys[1], cws[1]);
  calc(lines(['회전율 등급', '(분모 0이면 C)']), cols[2], ys[1], cws[2]);
  sc('A10 B5 C1', cols[3], ys[1], cws[3]);
  inp(lines(['리드타임 등급', '(입력값)']), cols[0], ys[2], cws[0]);
  inp(lines(['품질이슈 등급', '(입력값)']), cols[0], ys[3], cws[0]);
  sc('A5 B3 C1', cols[3], ys[2], cws[3]);
  sc('A5 B3 C0', cols[3], ys[3], cws[3]);
  [0, 1].forEach((r) => {
    arrow(s, cols[0] + cws[0] + 0.02, ys[r] + h / 2, cols[1] - cols[0] - cws[0] - 0.04, 0);
    arrow(s, cols[1] + cws[1] + 0.02, ys[r] + h / 2, cols[2] - cols[1] - cws[1] - 0.04, 0);
    arrow(s, cols[2] + cws[2] + 0.02, ys[r] + h / 2, cols[3] - cols[2] - cws[2] - 0.04, 0);
  });
  [2, 3].forEach((r) => arrow(s, cols[0] + cws[0] + 0.02, ys[r] + h / 2, cols[3] - cols[0] - cws[0] - 0.04, 0, C.mute, { dashType: 'dash' }));
  const ty = ys[0], th = ys[3] + h - ys[0];
  RR(s, cols[4], ty, cws[4], th, C.navy);
  const steps = ['종합', '= 네 점수 합', '↓', '종합 백분위', '↓', '재고등급', 'A · B · C', '↓', 'Z값'];
  T(s, steps.map((t, i) => ({ text: t, options: { breakLine: i < steps.length - 1, bold: ['종합', '재고등급', 'Z값'].includes(t), color: t === '↓' ? '8FC9C9' : C.white, fontSize: ['종합', '재고등급', 'Z값'].includes(t) ? 14 : 11.5 } })),
    { x: cols[4], y: ty + 0.15, w: cws[4], h: th - 0.3, align: 'center', valign: 'middle' });
  [0, 1, 2, 3].forEach((r) => arrow(s, cols[3] + cws[3] + 0.02, ys[r] + h / 2, cols[4] - cols[3] - cws[3] - 0.04, 0, C.teal));
  const ly = 5.62;
  R(s, MX, ly + 0.08, 0.22, 0.22, C.amberL); T(s, '현업 입력', { x: MX + 0.3, y: ly, w: 1.1, h: 0.38, fontSize: 11, valign: 'middle' });
  R(s, MX + 1.45, ly + 0.08, 0.22, 0.22, C.light); T(s, '계산', { x: MX + 1.75, y: ly, w: 0.7, h: 0.38, fontSize: 11, valign: 'middle' });
  R(s, MX + 2.5, ly + 0.08, 0.22, 0.22, C.white, { line: { color: C.teal, width: 1 } }); T(s, '점수', { x: MX + 2.8, y: ly, w: 0.7, h: 0.38, fontSize: 11, valign: 'middle' });
  box(s, `백분위 = x보다 작은 값의 수 ÷ (n−1), 소수 셋째 자리 내림 (Excel PERCENTRANK.INC와 같음) · 모집단은 구매마스터 ${G.total}행 전체`,
    MX, 6.06, 8.25, 0.66, { fill: C.light, fontSize: 11, align: 'left', margin: 8, round: 0.06 });
  const rx = 9.2, rw = W - MX - rx;
  RR(s, rx, 1.75, rw, 2.25, C.tealL);
  T(s, `${G.match} / ${G.total}`, { x: rx + 0.25, y: 1.85, w: rw - 0.5, h: 0.8, fontSize: 40, bold: true, color: C.teal, valign: 'middle' });
  T(s, '기존 등급과 일치한 SKU', { x: rx + 0.25, y: 2.66, w: rw - 0.5, h: 0.34, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  T(s, lines(['대조: 단가 등급 · 회전율 등급 · 종합 · 재고등급', '리드타임·품질 등급은 현업 입력값']),
    { x: rx + 0.25, y: 3.02, w: rw - 0.5, h: 0.9, fontSize: 11.5, color: C.slate, paraSpaceAfter: 2 });
  T(s, '계산 재고등급 분포', { x: rx, y: 4.15, w: rw, h: 0.32, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
  const tot = G.dist.A + G.dist.B + G.dist.C; let bx = rx;
  [['A', C.teal], ['B', C.slate], ['C', C.mute]].forEach(([k, col]) => {
    const bw = (G.dist[k] / tot) * rw;
    box(s, `${k} ${G.dist[k]}`, bx, 4.52, bw, 0.5, { fill: col, color: C.white, bold: true, fontSize: 12, margin: 0 });
    bx += bw;
  });
  RR(s, rx, 5.2, rw, 1.52, C.light);
  T(s, [
    { text: `회전율 분모가 0인 ${G.turnover_blank}행`, options: { breakLine: true, bold: true, color: C.navy } },
    { text: `→ 회전율 등급 ${Object.entries(G.turnover_blank_grade).map(([k, v]) => `${k} ${v}`).join(' · ')} (기존 규칙과 같음)`, options: { breakLine: true } },
    { text: `등급산출불가 ${G.not_computable}건` },
  ], { x: rx + 0.2, y: 5.32, w: rw - 0.4, h: 1.3, fontSize: 11.5, color: C.text, valign: 'middle', paraSpaceAfter: 3 });
})();

// ======================================================================= 적용 장치
(function guards() {
  const V = D.validation;
  const exW = Object.values(V.by_basis['가중평균'].excluded).reduce((a, b) => a + b, 0) + (V.memo['실적파일없음'] || 0);
  const topMemo = Object.entries(D.memos.counts).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([k, v]) => `${k} ${v}`).join(' · ');
  const s = slide({
    title: '틀린 값이 조용히 섞이지 않게 네 가지 장치를 두었습니다',
    sub: '계산할 수 없는 값은 비우고 이유를 남깁니다 — 임의 값으로 채우지 않습니다',
    tagText: '구현', tagKind: 'teal', source: `${SRC} › 산출결과·재고등급·검증_SKU`,
  });
  const cards = [
    ['doc', '점검메모', '입력이 빠지거나 이상하면 SKU 행에 사유를 적습니다 (10종, 부록 2)', `${D.memos.with_memo}/${D.memos.total}`, '메모가 붙은 SKU', topMemo],
    ['calendar', '미래 정보 차단', `가중치는 ${SEL}로만 고르고, ${HOLD} 실제 출고는 확인에만 씁니다`, HOLD, '선택에 쓰지 않은 확인 구간', '확인 구간 출고를 바꿔도 선택이 같음을 테스트로 확인'],
    ['lock', '등급산출불가', '단가·리드타임·품질 등급이 없으면 등급을 비우고 Z값을 쓰지 않습니다 (조용히 C로 두지 않음)', `${D.grades.not_computable}건`, '이번 데이터', `구매마스터 ${D.grades.total}행 모두 계산됨`],
    ['check', '채점 제외 규칙', '입력이 없거나 보정계수가 0 또는 3 초과면 채점에서 빼고 사유를 남깁니다', `${exW}개`, '가중평균 채점 제외 SKU', `${D.validation.total}개 중 ${D.validation.by_basis['가중평균'].demand_scored}개 채점`],
  ];
  const cw = (CW - 0.6) / 4, y = 1.75, h = 4.25;
  cards.forEach(([ic, a, b, big, bigl, foot], i) => {
    const x = MX + i * (cw + 0.2);
    RR(s, x, y, cw, h, C.light);
    icon(s, ic, x + 0.28, y + 0.28, 0.72, C.navy);
    T(s, a, { x: x + 0.28, y: y + 1.12, w: cw - 0.5, h: 0.4, fontSize: 16, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: x + 0.28, y: y + 1.58, w: cw - 0.5, h: 1.15, fontSize: 12, color: C.text });
    line(s, x + 0.28, y + 2.8, cw - 0.56, 0, C.line);
    T(s, big, { x: x + 0.28, y: y + 2.9, w: cw - 0.5, h: 0.6, fontSize: 26, bold: true, color: C.teal, valign: 'middle' });
    T(s, bigl, { x: x + 0.28, y: y + 3.48, w: cw - 0.5, h: 0.28, fontSize: 11, bold: true, color: C.slate, valign: 'middle' });
    T(s, foot, { x: x + 0.28, y: y + 3.76, w: cw - 0.5, h: 0.45, fontSize: 10.5, color: C.gray });
  });
  box(s, `보정계수 이상은 산출은 그대로, 채점만 제외 · 사후 검증 구간 실제 출고는 계산 입력에 넣지 않음 · 자동 테스트 ${int(TESTS)}개로 규칙 고정`,
    MX, 6.2, CW, 0.56, { fill: C.navy, color: C.white, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 간지 3
divider(2);

// ======================================================================= 가중치 탐색 방법
(function wsMethod() {
  const s = slide({
    title: `AI 가중치 탐색: 앞 4주로 고르고 뒤 5주로 확인합니다`,
    sub: `8/24 계산값은 그대로 두고, 사후 검증 구간을 선택 구간(${SEL})과 검증 구간(${HOLD})으로 나눴습니다`,
    tagText: '실험 방법', tagKind: 'light',
    source: `가중치 탐색 설계(2026-11-20) · ${SRC} › 산출결과(현재 가중치)${WS ? ' · ' + WS.source + ' › 적용기준' : ''}`,
  });
  const x0 = MX + 0.2, x1 = W - MX - 0.2, wk = (x1 - x0) / 47, X = (w) => x0 + (w - 1) * wk;
  const y = 2.05;
  T(s, '주차', { x: MX - 0.05, y: y - 0.35, w: 1, h: 0.28, fontSize: 10.5, color: C.gray });
  box(s, `출고 실적 ${HIST} (계산 입력)`, X(1), y, X(32) - X(1) - 0.03, 0.5, { fill: C.mute, color: C.white, bold: true, fontSize: 12, margin: 0 });
  box(s, '미사용', X(32), y, X(35) - X(32) - 0.03, 0.5, { fill: 'E6EBF0', color: C.gray, fontSize: 10, margin: 0 });
  box(s, lines(['선택 구간', SEL]), X(35), y, X(39) - X(35) - 0.03, 0.5, { fill: C.amber, color: C.white, bold: true, fontSize: 10.5, margin: 0 });
  box(s, lines(['검증 구간', HOLD]), X(39), y, X(44) - X(39) - 0.03, 0.5, { fill: C.teal, color: C.white, bold: true, fontSize: 10.5, margin: 0 });
  box(s, 'FCST 8/24 스냅샷 W35~47 (계산 입력)', X(35), y + 0.6, X(48) - X(35), 0.36, { fill: 'EFE3CB', color: '8A6420', bold: true, fontSize: 10.5, margin: 0 });
  line(s, X(35) - 0.015, y - 0.15, 0, 1.25, C.red, { width: 1.5, dashType: 'dash' });
  T(s, '8/24 계산 시점', { x: X(35) - 1.75, y: y + 0.62, w: 1.7, h: 0.32, fontSize: 10.5, bold: true, color: C.red, align: 'right', valign: 'middle' });
  [1, 10, 20, 31, 35, 39, 43, 47].forEach((w) => T(s, `W${String(w).padStart(2, '0')}`, { x: X(w) - 0.3 + wk / 2, y: y + 1.0, w: 0.6, h: 0.24, fontSize: 10, color: C.gray, align: 'center' }));
  const crit = WS ? WS.criteria : {};
  const rules = [
    ['후보', '실적 가중치(w) 0.0~1.0을 0.1 간격으로 11개, FCST 가중치는 1 − w'],
    ['단위', crit['단위'] || '품목그룹마다 하나의 w'],
    ['고르는 기준', '선택 구간 가중평균 오차율이 가장 낮은 w (같으면 현재 w 쪽)'],
    ['작은 그룹', '선택 구간 SKU 5개 미만이면 전체를 합쳐 고른 w (전체 합산 최적)'],
  ];
  const cw = 3.55, ch = 1.12;
  rules.forEach(([a, b], i) => {
    const x = MX + (i % 2) * (cw + 0.2), yy = 3.55 + Math.floor(i / 2) * (ch + 0.18);
    RR(s, x, yy, cw, ch, C.light);
    num(s, i + 1, x + 0.18, yy + 0.18, 0.42, C.navy, C.white, 13);
    T(s, a, { x: x + 0.72, y: yy + 0.14, w: cw - 0.9, h: 0.36, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: x + 0.72, y: yy + 0.5, w: cw - 0.9, h: 0.58, fontSize: 11, color: C.text });
  });
  box(s, '왜 앞뒤로 나눴나: W23~31 시점의 FCST가 없어 과거 구간으로는 FCST 쪽 값을 만들 수 없습니다. W38에 다시 고르는 상황을 흉내 낸 것이라 실제 운영보다 낙관적일 수 있습니다.',
    MX, 6.12, 2 * cw + 0.2, 0.66, { fill: C.amberL, fontSize: 11, align: 'left', margin: 8, round: 0.06 });
  const tx = MX + 2 * cw + 0.5, tw = W - MX - tx;
  label(s, '현재 가중치 (구매마스터)', tx, 3.5, tw);
  const rows = [['품목그룹', 'SKU(구매마스터)', '실적 가중치', 'FCST 가중치']];
  D.weights.forEach((w) => rows.push([cellOpt(w.group, { bold: true }), String(w.skus), wfmt(w.wh), wfmt(w.wf)]));
  table(s, rows, { x: tx, y: 3.88, w: tw, colW: [0.95, 1.35, (tw - 2.3) / 2, (tw - 2.3) / 2], rowH: 0.36, fontSize: 11, headSize: 10.5, align: 'center' });
})();

// ======================================================================= 후보 전체 채점(히트맵)
(function wsHeatmap() {
  if (!wsReady) return;
  const s = slide({
    title: `후보 ${CANDS.length}개를 모두 채점해 그룹별 최적 w를 골랐습니다`,
    sub: `선택 구간 ${SEL}의 가중평균 수요 오차율 · 행 안에서 오차가 작을수록 진한 청록 · ★ 선택 · ○ 현재`,
    tagText: '탐색 과정', tagKind: 'teal', source: `${WS.source} › 후보성능·선택가중치`,
  });
  const rowsDef = [{ key: '전체', label: '전체 합산', choice: null }, ...WS.choices.map((c) => ({ key: c.group, label: c.group, choice: c }))];
  const pooled = WS.candidates.filter((c) => c.scope === '전체' && isNum(c.wape));
  const pooledBest = pooled.reduce((a, b) => (b.wape < a.wape ? b : a)).w;
  const lx = MX, nx = MX + 1.2, x0 = MX + 1.9, rcw = 1.6;
  const cw = (CW - 1.9 - rcw - 0.1) / CANDS.length;
  const rx = x0 + cw * CANDS.length + 0.1;
  T(s, '실적 가중치 w  (FCST 가중치 = 1 − w)', { x: x0, y: 1.62, w: cw * CANDS.length, h: 0.3, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  T(s, '그룹', { x: lx, y: 1.95, w: 1.15, h: 0.32, fontSize: 11.5, bold: true, color: C.navy, valign: 'middle' });
  T(s, 'SKU', { x: nx, y: 1.95, w: 0.65, h: 0.32, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  CANDS.forEach((w, j) => T(s, wfmt(w), { x: x0 + j * cw, y: 1.95, w: cw, h: 0.32, fontSize: 11.5, bold: true, color: C.slate, align: 'center', valign: 'middle' }));
  T(s, '현재 → 선택', { x: rx, y: 1.95, w: rcw, h: 0.32, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  const PAL = [[C.teal, C.white], ['5FA8A8', C.white], ['A8D3D3', C.navy], [C.tealL, C.navy], ['F3F5F7', C.slate]];
  const rh = 0.4, rs = 0.44, y0 = 2.34;
  rowsDef.forEach((r, i) => {
    const y = y0 + i * rs;
    const cells = CANDS.map((w) => WS.candidates.find((c) => c.scope === r.key && c.w === w));
    const vals = cells.filter((c) => c && isNum(c.wape)).map((c) => c.wape);
    const mn = Math.min(...vals), mx = Math.max(...vals);
    const chosen = r.choice ? r.choice.chosen : pooledBest;
    const current = r.choice ? r.choice.current : null;
    const first = cells.find((c) => c);
    const small = r.choice && r.choice.method === '공통최적';
    if (i === 0) R(s, lx, y - 0.02, CW, rh + 0.04, 'EEF1F4');
    T(s, small ? `${r.label} *` : r.label, { x: lx, y, w: 1.15, h: rh, fontSize: 12.5, bold: true, color: small ? C.gray : C.navy, valign: 'middle' });
    T(s, first ? String(first.n) : '', { x: nx, y, w: 0.65, h: rh, fontSize: 11.5, color: C.text, align: 'center', valign: 'middle' });
    cells.forEach((c, j) => {
      const x = x0 + j * cw;
      if (!c || !isNum(c.wape)) { box(s, '–', x + 0.02, y, cw - 0.04, rh, { fill: 'F7F8FA', color: C.mute, fontSize: 10.5, margin: 0 }); return; }
      const t = mx > mn ? (c.wape - mn) / (mx - mn) : 0;
      const lv = t <= 0.001 ? 0 : t < 0.25 ? 1 : t < 0.5 ? 2 : t < 0.75 ? 3 : 4;
      const isC = c.w === chosen, isNow = c.w === current;
      const mark = `${isC ? '★' : ''}${isNow ? '○' : ''}`;
      box(s, `${mark}${mark ? ' ' : ''}${pct(c.wape, c.wape < 1 ? 1 : 0)}`, x + 0.02, y, cw - 0.04, rh, {
        fill: PAL[lv][0], color: PAL[lv][1], fontSize: 10, bold: isC, margin: 0,
        lineColor: isC ? C.amber : PAL[lv][0], lineWidth: isC ? 2.25 : 0.75 });
    });
    const how = r.choice ? (r.choice.method === '그룹최적' ? '그룹 최적' : r.choice.method === '공통최적' ? '전체 합산 값 사용' : '현재 유지') : '전체 합산 최적';
    T(s, [{ text: r.choice ? `${wfmt(current)} → ${wfmt(chosen)}` : wfmt(chosen), options: { bold: true, color: r.choice && chosen !== current ? C.teal : C.navy, breakLine: true, fontSize: 12 } },
      { text: how, options: { color: C.gray, fontSize: 9.5 } }],
    { x: rx, y: y - 0.02, w: rcw, h: rh + 0.04, align: 'center', valign: 'middle' });
  });
  const ny = y0 + rowsDef.length * rs + 0.04;
  T(s, `* SKU ${MINSKU}개 미만 그룹은 그룹 안의 최솟값 대신 전체 합산 최적 w(${wfmt(pooledBest)})를 씁니다 (과적합 방지) · 색은 행 안에서만 비교 · SKU = 선택 구간 채점 수`,
    { x: MX, y: ny, w: CW, h: 0.3, fontSize: 10.5, color: C.gray, valign: 'middle' });
  box(s, [
    { text: `왜 ${CANDS.length * NG}번 계산으로 충분한가  `, options: { bold: true, color: '8FC9C9' } },
    { text: `그룹의 오차는 그 그룹의 w에만 달려 있습니다. 그래서 후보 ${CANDS.length}개 × 그룹 ${NG}개 = ${CANDS.length * NG}번 계산으로 ${man(SPACE)} 조합 전체를 비교한 것과 같은 답을 얻습니다 (작은 그룹은 일부러 전체 합산 값을 씀).`, options: { color: C.white } },
  ], MX, ny + 0.38, CW, 0.66, { fill: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 가중치 탐색 결과
(function wsResult() {
  const ready = wsReady;
  const cv = wsCur, nv = wsNew;
  const coverDown = ready && isNum(nv.cover) && isNum(cv.cover) && nv.cover < cv.cover;
  const s = slide({
    title: ready ? `AI 탐색 결과: 검증 구간 오차율 ${pct(cv.wape)} → ${pct(nv.wape)}` : `가중치 탐색 결과: ${PENDING}`,
    sub: ready
      ? `개선은 ${topContribText}에 몰려 있고, 결품 방어율은 ${pct(cv.cover)} → ${pct(nv.cover)}${coverDown ? '로 낮아졌습니다' : '입니다'} · 검증 구간 ${HOLD}`
      : '가중치 탐색 결과 파일(out/가중치탐색.xlsx)이 생기면 두 명령을 다시 돌려 이 장을 채웁니다',
    tagText: '실험 결과', tagKind: 'dark',
    source: ready ? `${WS.source} › 선택가중치·검증비교·후보성능 · 그룹 몫은 extract.py가 SKU별로 다시 계산` : 'out/가중치탐색.xlsx (산출 예정)',
  });
  const lw = 7.15;
  label(s, `그룹별 선택 (선택 구간 ${SEL}로 고름)`, MX, 1.72, lw);
  const rows = [['그룹', 'SKU', '현재 w', '선택 w', '방식', '선택 구간 오차율', '검증 구간 오차율', '감소 몫']];
  if (ready) {
    WS.choices.forEach((c) => {
      const hc = wsRow('검증', c.group, '현재', '가중평균'), hn = wsRow('검증', c.group, '탐색', '가중평균');
      const hv = hc && hn && isNum(hc.wape) && isNum(hn.wape);
      const col = hv ? (hn.wape < hc.wape ? C.teal : hn.wape > hc.wape ? C.red : C.text) : C.gray;
      const sh = contrib && contrib[c.group] ? contrib[c.group].share : null;
      rows.push([cellOpt(c.group, { bold: true }), String(c.n), wfmt(c.current),
        cellOpt(wfmt(c.chosen), { bold: true, color: c.chosen !== c.current ? C.teal : C.text }),
        cellOpt(METHOD[c.method] || c.method, { color: c.method === '그룹최적' ? C.navy : C.gray }),
        `${pct(c.wape_current)} → ${pct(c.wape_chosen)}`,
        cellOpt(hv ? `${pct(hc.wape)} → ${pct(hn.wape)}` : '–', { bold: true, color: col }),
        cellOpt(isNum(sh) ? pct(sh, 0) : '–', { bold: isNum(sh) && sh >= 0.1, color: isNum(sh) && sh < 0 ? C.red : C.text })]);
    });
  } else {
    D.weights.forEach((w) => rows.push([cellOpt(w.group, { bold: true }), String(w.skus), wfmt(w.wh),
      cellOpt('예정', { color: C.gray }), cellOpt('–', { color: C.gray }), cellOpt(PENDING, { color: C.gray }), cellOpt(PENDING, { color: C.gray }), cellOpt('–', { color: C.gray })]));
  }
  table(s, rows, { x: MX, y: 2.1, w: lw, colW: [0.65, 0.5, 0.7, 0.7, 0.9, 1.4, 1.5, 0.8], rowH: 0.37, fontSize: 10.5, headSize: 10.5, align: 'center' });
  const noteY = 2.1 + 0.37 * rows.length + 0.06;
  const ends = ready ? WS.choices.filter((c) => isNum(c.chosen) && (c.chosen === 0 || c.chosen === 1) && c.method === '그룹최적') : [];
  const endText = ends.map((c) => `${c.group} ${wfmt(c.chosen)}(${c.chosen === 0 ? 'FCST만' : '실적만'})`).join(' · ');
  T(s, ready
    ? lines([
      'SKU = 선택 구간 채점 수 · 검증 구간은 공통 범위 SKU · 감소 몫 = 오차 감소분 중 그룹 몫',
      `${endText ? `후보의 끝값을 고른 그룹: ${endText} · ` : ''}SKU 1~4개 그룹의 큰 변화는 우연일 수 있음`,
    ])
    : 'SKU = 구매마스터 SKU 수 · 탐색 결과가 나오면 선택 w와 구간별 오차율이 채워집니다',
  { x: MX, y: noteY, w: lw, h: 0.55, fontSize: 10.5, color: C.gray });
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  label(s, `검증 구간 ${HOLD} · 전체`, rx, 1.72, rw);
  const kp = [
    ['가중평균 오차율', 'wape', '가중평균', true],
    ['결품 방어율 (가중평균)', 'cover', '가중평균', false],
    ['결품 방어율 (기존수기)', 'cover', '기존수기', false],
    ['재고일수 중앙값 (가중평균)', 'days_median', '가중평균', null],
  ];
  kp.forEach(([l, key, b, lowGood], i) => {
    const y = 2.1 + i * 0.6;
    RR(s, rx, y, rw, 0.53, C.light);
    T(s, l, { x: rx + 0.15, y, w: 2.3, h: 0.53, fontSize: 11.5, bold: true, color: C.navy, valign: 'middle' });
    if (!ready) { T(s, PENDING, { x: rx + 2.4, y, w: rw - 2.55, h: 0.53, fontSize: 13, bold: true, color: C.gray, align: 'right', valign: 'middle' }); return; }
    const a = wsRow('검증', '전체', '현재', b)[key], z = wsRow('검증', '전체', '탐색', b)[key];
    const f = key === 'days_median' ? (v) => f1(v) + '일' : (v) => pct(v);
    let col = C.text;
    if (lowGood !== null && isNum(a) && isNum(z) && a !== z) col = ((z < a) === lowGood) ? C.teal : C.red;
    T(s, [{ text: f(a), options: { color: C.slate } }, { text: ' → ', options: { color: C.slate } }, { text: f(z), options: { color: col } }],
      { x: rx + 2.35, y, w: rw - 2.5, h: 0.53, fontSize: 14, bold: true, align: 'right', valign: 'middle' });
  });
  T(s, '기존수기 목표재고량도 가중평균 수요로 리드타임 수요를 계산하므로 가중치를 바꾸면 함께 바뀝니다',
    { x: rx, y: 4.5, w: rw, h: 0.45, fontSize: 10.5, color: C.gray });
  const cy = 5.0;
  if (ready) {
    const cand = WS.candidates.filter((c) => c.scope === '전체' && isNum(c.wape));
    const mn = Math.min(...cand.map((c) => c.wape)), mxv = Math.max(...cand.map((c) => c.wape));
    T(s, `선택 구간 전체 합산 오차율 (w별): ${pct(mn)} ~ ${pct(mxv)}`, { x: rx, y: cy, w: rw, h: 0.3, fontSize: 11.5, bold: true, color: C.navy, valign: 'middle' });
    const bx = rx + 0.05, bw = (rw - 0.1) / cand.length, base = 6.45, lo = mn * 0.9, ph = 1.0;
    cand.forEach((c, i) => {
      const h = Math.max(0.04, ((c.wape - lo) / (mxv - lo)) * ph);
      R(s, bx + i * bw + bw * 0.15, base - h, bw * 0.7, h, c.wape === mn ? C.amber : C.mute);
      T(s, wfmt(c.w), { x: bx + i * bw, y: base + 0.03, w: bw, h: 0.22, fontSize: 9.5, color: C.gray, align: 'center' });
    });
    T(s, '막대 높이는 최솟값 근처를 확대한 것 · 주황 = 최소', { x: rx, y: 6.72, w: rw, h: 0.24, fontSize: 10, color: C.gray });
  } else {
    box(s, `후보 w별 오차율 곡선: ${PENDING}`, rx, cy, rw, 1.85, { fill: C.white, dash: true, lineColor: C.mute, color: C.gray, fontSize: 12 });
  }
  if (ready) {
    const up = [], down = [];
    WS.choices.forEach((c) => {
      const hc = wsRow('검증', c.group, '현재', '가중평균'), hn = wsRow('검증', c.group, '탐색', '가중평균');
      if (hc && hn && isNum(hc.wape) && isNum(hn.wape)) {
        const t = `${c.group} ${pct(hc.wape)}→${pct(hn.wape)}`;
        if (hn.wape < hc.wape) up.push(t); else if (hn.wape > hc.wape) down.push(`${c.group} 오차율 ${pct(hc.wape)}→${pct(hn.wape)}`);
      }
      if (hc && hn && isNum(hc.cover) && isNum(hn.cover) && hn.cover < hc.cover) down.push(`${c.group} 방어율 ${pct(hc.cover)}→${pct(hn.cover)}`);
    });
    if (coverDown) down.unshift(`전체 방어율 ${pct(cv.cover)}→${pct(nv.cover)}`);
    const pair = (arr) => arr.reduce((acc, t, i) => (i % 2 ? (acc[acc.length - 1] += ' · ' + t, acc) : (acc.push(t), acc)), []);
    const by = noteY + 0.62, bh = 6.85 - by, hw = (lw - 0.15) / 2;
    RR(s, MX, by, hw, bh, C.tealL);
    T(s, [{ text: '좋아진 것 (오차율)', options: { bold: true, color: C.teal, breakLine: true } }, ...lines(pair(up).length ? pair(up) : ['없음'], { color: C.text })],
      { x: MX + 0.15, y: by + 0.06, w: hw - 0.3, h: bh - 0.12, fontSize: 11, valign: 'top' });
    RR(s, MX + hw + 0.15, by, hw, bh, C.redL);
    T(s, [{ text: '나빠진 것', options: { bold: true, color: C.red, breakLine: true } }, ...lines(pair(down).length ? pair(down) : ['없음'], { color: C.text })],
      { x: MX + hw + 0.3, y: by + 0.06, w: hw - 0.3, h: bh - 0.12, fontSize: 11, valign: 'top' });
  }
})();

// ======================================================================= 간지 4
divider(3);


// ======================================================================= 실험 설계
(function design() {
  const V = D.validation, ex = V.by_basis['가중평균'].excluded;
  const s = slide({
    title: '실험: 8/24에 계산한 값이 이후 9주 출고를 맞혔나',
    sub: `같은 입력으로 네 기준을 만들고, 계산에 쓰지 않은 사후 검증 구간(${VAL}) 실제 출고로 같은 잣대를 적용했습니다`,
    tagText: '실험 방법', tagKind: 'light', source: `${SRC} › 검증_SKU·검증요약·적용기준`,
  });
  const pw = 4.25;
  RR(s, MX, 1.72, pw, 5.03, C.navy);
  icon(s, 'search', MX + 0.3, 1.92, 0.7, C.teal);
  T(s, '실험이 답하려는 질문', { x: MX + 1.15, y: 2.0, w: 2.9, h: 0.5, fontSize: 13, color: C.sky, bold: true, valign: 'middle' });
  T(s, `8/24 시점에 계산한 값이 그 뒤 ${VAL} 실제 출고를 얼마나 맞혔나?`,
    { x: MX + 0.3, y: 2.75, w: pw - 0.6, h: 1.2, fontSize: 19, bold: true, color: C.white });
  // 기준 × 지표 표(어떤 기준이 어떤 지표로 채점되나)
  const mx0 = MX + 0.3, my0 = 4.1, mc = [1.25, 0.88, 0.88, 0.88];
  ['기준', '오차율', '결품 방어', '재고일수'].forEach((t, i) => T(s, t, { x: mx0 + mc.slice(0, i).reduce((a, b) => a + b, 0), y: my0, w: mc[i], h: 0.3, fontSize: 10.5, color: C.sky, bold: true, align: i ? 'center' : 'left', valign: 'middle' }));
  BASES.forEach((b, r) => {
    const yy = my0 + 0.33 + r * 0.33;
    T(s, b, { x: mx0, y: yy, w: mc[0], h: 0.3, fontSize: 11, color: C.white, valign: 'middle' });
    [b === '기존수기' ? '–' : '●', '●', '●'].forEach((t, k) => T(s, t, { x: mx0 + mc.slice(0, k + 1).reduce((a, x) => a + x, 0), y: yy, w: mc[k + 1], h: 0.3, fontSize: 11, color: t === '●' ? '8FC9C9' : C.mute, align: 'center', valign: 'middle' }));
  });
  T(s, lines([`계산 입력: 출고 ${HIST}`, 'FCST 8/24 스냅샷 앞 13주 (W35~47)']),
    { x: MX + 0.3, y: 5.95, w: pw - 0.6, h: 0.65, fontSize: 11.5, color: C.sky, valign: 'bottom' });
  const rx = 5.15, rw = W - MX - rx;
  label(s, '비교한 기준 4가지', rx, 1.72, rw, C.amber);
  const bases = [['기존수기', 'MP 일별 안전재고'], ['가중평균', '실적·FCST 가중'], ['실적기준', '실적 × 보정계수'], ['FCST기준', 'FCST × 보정계수']];
  const bw = (rw - 0.45) / 4;
  bases.forEach(([a, b], i) => {
    const x = rx + i * (bw + 0.15);
    RR(s, x, 2.1, bw, 0.78, C.light);
    R(s, x, 2.1, 0.08, 0.78, BASIS_COLOR[a]);
    T(s, a, { x: x + 0.2, y: 2.16, w: bw - 0.3, h: 0.34, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: x + 0.2, y: 2.5, w: bw - 0.3, h: 0.3, fontSize: 11, color: C.slate, valign: 'middle' });
  });
  label(s, '지표 3가지', rx, 3.05, rw, C.amber);
  const mets = [
    ['수요 오차율 (WAPE)', 'Σ|예측 − 실제| ÷ Σ실제', '낮을수록 좋음 · 기존수기는 수요 예측이 없어 제외'],
    ['결품 방어율', '목표재고량이 실제 리드타임 수요 이상인 SKU 비율', '높을수록 좋음 · 실제 결품이 아닌 대리 지표'],
    ['실제 기준 재고일수', '목표재고량 ÷ 실제 일평균 출고', '방어율이 같다면 낮을수록 좋음'],
  ];
  const mw = (rw - 0.3) / 3;
  mets.forEach(([a, b, c], i) => {
    const x = rx + i * (mw + 0.15);
    RR(s, x, 3.43, mw, 1.45, C.white, { line: { color: C.line, width: 1 } });
    T(s, a, { x: x + 0.18, y: 3.5, w: mw - 0.36, h: 0.34, fontSize: 13, bold: true, color: C.teal, valign: 'middle' });
    T(s, b, { x: x + 0.18, y: 3.86, w: mw - 0.36, h: 0.55, fontSize: 11, color: C.text });
    T(s, c, { x: x + 0.18, y: 4.4, w: mw - 0.36, h: 0.45, fontSize: 10.5, color: C.gray });
  });
  label(s, '공정 비교 규칙', rx, 5.0, rw, C.amber);
  RR(s, rx, 5.36, rw, 1.39, C.light);
  const exText = Object.entries(ex).map(([k, v]) => `${k.replace(', ', '·')} ${v}`).join(' · ');
  const rl = [
    `공통 범위: 수요는 세 산식 기준 모두 채점된 SKU(${M['가중평균'].n_demand}개), 결품 방어는 네 기준 모두 채점된 SKU(${M['가중평균'].n_lt}개)`,
    `채점 제외(가중평균): ${exText} · 실적파일없음 ${V.memo['실적파일없음'] || 0}`,
    `보정계수 0 또는 3 초과 SKU는 실적·FCST기준만 채점 제외 (${D.memos.counts['보정계수이상'] || 0}개)`,
    `리드타임 63일(사후 검증 구간) 초과 SKU는 결품 방어 채점 불가 (${V.lt_over_window}개)`,
  ];
  T(s, rl.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < rl.length - 1 } })),
    { x: rx + 0.2, y: 5.42, w: rw - 0.4, h: 1.28, fontSize: 11, color: C.text, valign: 'middle', paraSpaceAfter: 2 });
})();

// ======================================================================= 결과 1 평균
(function result1() {
  const s = slide({
    title: `결과 1. 세 산식 중 ${bestWape}의 수요 오차가 가장 낮습니다`,
    sub: coverGap > 0
      ? `단, 결품 방어율은 기존수기(${pct(M['기존수기'].cover)})가 가중평균(${pct(M['가중평균'].cover)})보다 높습니다 · 사후 검증 구간 ${VAL} · 공통 범위`
      : `결품 방어율도 가중평균이 기존수기 이상입니다 · 사후 검증 구간 ${VAL} · 공통 범위`,
    tagText: '실험 결과', tagKind: 'dark', source: `${SRC} › 검증요약 (공통 범위)`,
  });
  const cw = (CW - 0.6) / 3, y = 1.72, h = 3.95;
  const mk = (key) => BASES.map((b) => ({ label: b, value: M[b][key], color: BASIS_COLOR[b], strong: b === '가중평균', empty: '없음' }));
  const maxW = Math.max(...formulaBases.map((b) => M[b].wape)) * 1.2;
  const maxD = Math.max(...BASES.map((b) => M[b].days_median)) * 1.2;
  [0, 1, 2].forEach((i) => RR(s, MX + i * (cw + 0.3), y, cw, h, C.light));
  bars(s, { x: MX + 0.2, y: y + 0.12, w: cw - 0.4, h: h - 0.2, title: '수요 오차율 (WAPE)', better: `낮을수록 좋음 · ${M['가중평균'].n_demand} SKU · 기존수기는 예측 없음`, items: mk('wape'), max: maxW });
  bars(s, { x: MX + cw + 0.5, y: y + 0.12, w: cw - 0.4, h: h - 0.2, title: '결품 방어율', better: `높을수록 좋음 · ${M['가중평균'].n_lt} SKU`, items: mk('cover'), max: 1.15 });
  bars(s, { x: MX + 2 * (cw + 0.3) + 0.2, y: y + 0.12, w: cw - 0.4, h: h - 0.2, title: '실제 기준 재고일수 중앙값', better: '방어율이 같다면 낮을수록 좋음', items: mk('days_median'), max: maxD, fmt: (v) => f1(v) + '일' });
  const daysCut = 1 - M['가중평균'].days_median / M['기존수기'].days_median;
  box(s, lines([
    `세 산식 기준 중 가중평균의 수요 오차가 가장 낮습니다 (기존수기는 수요 예측이 없어 오차율 비교 대상이 아닙니다).`,
    coverGap > 0
      ? `가중평균은 기존수기보다 재고일수가 ${pct(daysCut, 0)} 짧고, 결품 방어율은 ${pp(M['기존수기'].cover, M['가중평균'].cover)} 낮습니다.`
      : `가중평균은 기존수기보다 재고일수가 ${pct(daysCut, 0)} 짧습니다.`,
  ]), MX, 5.8, CW, 0.78, { fill: C.navy, color: C.white, fontSize: 12.5, bold: true, align: 'left', margin: 12, round: 0.06, paraSpaceAfter: 2 });
  T(s, `참고 · 전체 범위(기준마다 SKU 수 다름) 오차율: 가중평균 ${pct(MALL['가중평균'].wape)} · 실적기준 ${pct(MALL['실적기준'].wape)} · FCST기준 ${pct(MALL['FCST기준'].wape)} / 결품 방어율: 기존수기 ${pct(MALL['기존수기'].cover)} · 가중평균 ${pct(MALL['가중평균'].cover)} · 실적기준 ${pct(MALL['실적기준'].cover)} · FCST기준 ${pct(MALL['FCST기준'].cover)}`,
    { x: MX, y: 6.62, w: CW, h: 0.36, fontSize: 10.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 결과 2 품목그룹 히트맵
(function result2() {
  const gm = (g, b) => D.metrics[g][b];
  const n = (g) => gm(g, '가중평균').n_demand;
  const allBad = formulaBases.every((b) => gm(worstGroup, b).wape > 0.4);
  const s = slide({
    title: `결과 2. ${bestGroup}는 잘 맞고, ${worstGroup}는 ${allBad ? '세 기준 모두' : '크게'} 빗나갔습니다`,
    sub: `품목그룹 × 기준 히트맵 · 사후 검증 구간 ${VAL} · 공통 범위 SKU를 그룹별로 나눈 값`,
    tagText: '실험 결과', tagKind: 'dark', source: `${SRC} › 검증요약 (품목그룹 범위)`,
  });
  const lvW = (v) => (v <= 0.2 ? 0 : v <= 0.4 ? 1 : v <= 0.8 ? 2 : 3);
  const lvC = (v) => (v >= 0.9 ? 0 : v >= 0.7 ? 1 : v >= 0.5 ? 2 : 3);
  const PAL = [[C.teal, C.white], [C.tealL, C.navy], ['F3D9A8', C.navy], ['D9665A', C.white]];
  const cols = [
    ...formulaBases.map((b) => ({ grp: '수요 오차율 (낮을수록 좋음)', b, key: 'wape', lv: lvW, fmt: (v) => pct(v, 0) })),
    ...BASES.map((b) => ({ grp: '결품 방어율 (높을수록 좋음)', b, key: 'cover', lv: lvC, fmt: (v) => pct(v, 0) })),
    ...['기존수기', '가중평균'].map((b) => ({ grp: '재고일수 중앙값', b, key: 'days_median', lv: null, fmt: (v) => f1(v) })),
  ];
  const x0 = MX + 1.55, cw = 1.08, gap = 0.05, y0 = 1.75;
  T(s, '품목그룹', { x: MX, y: y0 + 0.36, w: 1.0, h: 0.34, fontSize: 11.5, bold: true, color: C.navy, valign: 'middle' });
  T(s, lines(['SKU', '(공통)']), { x: MX + 0.95, y: y0 + 0.33, w: 0.6, h: 0.42, fontSize: 10.5, bold: true, color: C.navy, valign: 'middle', align: 'center' });
  let gx = x0, prev = null, startX = x0;
  const flushGroup = (name, sx, ex) => box(s, name, sx, y0, ex - sx - gap, 0.32, { fill: C.navy, color: C.white, bold: true, fontSize: 11.5, margin: 0 });
  cols.forEach((c) => {
    if (c.grp !== prev) { if (prev) { flushGroup(prev, startX, gx); gx += 0.12; } startX = gx; prev = c.grp; }
    c.x = gx;
    T(s, c.b, { x: gx, y: y0 + 0.36, w: cw, h: 0.34, fontSize: 11, bold: true, color: BASIS_TEXT[c.b], align: 'center', valign: 'middle' });
    gx += cw + gap;
  });
  flushGroup(prev, startX, gx);
  const rh = 0.39, rs = 0.44;
  groupsOrdered.forEach((g, r) => {
    const y = y0 + 0.8 + r * rs, small = n(g) < 5;
    T(s, small ? `${g} *` : g, { x: MX, y, w: 1.0, h: rh, fontSize: 13, bold: true, color: small ? C.gray : C.navy, valign: 'middle' });
    T(s, String(n(g)), { x: MX + 1.0, y, w: 0.5, h: rh, fontSize: 12, color: C.text, valign: 'middle', align: 'center' });
    cols.forEach((c) => {
      const v = gm(g, c.b)[c.key];
      if (!isNum(v)) { box(s, '–', c.x, y, cw, rh, { fill: 'F7F8FA', color: C.mute, fontSize: 11, margin: 0 }); return; }
      const [f, tc] = c.lv ? PAL[c.lv(v)] : ['EEF1F4', C.text];
      box(s, c.fmt(v), c.x, y, cw, rh, { fill: f, color: tc, fontSize: 12, bold: c.b === '가중평균', margin: 0 });
    });
  });
  const ly = y0 + 0.8 + groupsOrdered.length * rs + 0.08;
  const legend = [[0, '오차율 ≤20% · 방어 ≥90%'], [1, '≤40% · ≥70%'], [2, '≤80% · ≥50%'], [3, '>80% · <50%']];
  legend.forEach(([lv, t], i) => {
    const x = MX + i * 2.35;
    R(s, x, ly + 0.07, 0.3, 0.22, PAL[lv][0]);
    T(s, t, { x: x + 0.38, y: ly, w: 1.95, h: 0.36, fontSize: 10.5, color: C.text, valign: 'middle' });
  });
  T(s, '* SKU 5개 미만: 해석 보류', { x: MX + 9.4, y: ly, w: CW - 9.4, h: 0.36, fontSize: 10.5, color: C.gray, valign: 'middle', align: 'right' });
  const cy = ly + 0.42, ch = 6.85 - cy;
  const smallList = groupsOrdered.filter((g) => n(g) < 5);
  const cards = [
    [C.teal, `${bestGroup} (${n(bestGroup)} SKU)`, `오차율 ${pct(gm(bestGroup, '가중평균').wape, 0)}, 결품 방어율 ${pct(gm(bestGroup, '가중평균').cover, 0)} — 가중평균이 잘 맞음`],
    [C.red, `${worstGroup} (${n(worstGroup)} SKU)`, `세 기준 오차율 ${formulaBases.map((b) => pct(gm(worstGroup, b).wape, 0)).join('·')}, 모두 과소예측`],
    [C.gray, smallList.join('·'), 'SKU 1~4개 그룹: 값이 극단적이어도 결론을 내리지 않음'],
  ];
  const ccw = (CW - 0.4) / 3;
  cards.forEach(([col, a, b], i) => {
    const x = MX + i * (ccw + 0.2);
    RR(s, x, cy, ccw, ch, C.light);
    s.addShape(pres.shapes.OVAL, { x: x + 0.16, y: cy + 0.12, w: 0.18, h: 0.18, fill: { color: col }, line: { color: col, width: 0 } });
    T(s, [{ text: a, options: { bold: true, color: C.navy, fontSize: 12, breakLine: true } }, { text: b, options: { color: C.text } }],
      { x: x + 0.42, y: cy + 0.06, w: ccw - 0.55, h: ch - 0.12, fontSize: 11, valign: 'top' });
  });
})();

// ======================================================================= 결과 3 원인 추적
(function result3() {
  const GS = D.group_sums, worst = worstGroup, best = bestGroup;
  const ratio = GS[worst].actual / GS[worst].hist;
  const s = slide({
    title: `결과 3. ${worst}는 실제 출고가 과거 평균의 ${ratio.toFixed(1)}배였습니다`,
    sub: '그룹별 주수요 합계와 오차 기여 · 원인은 데이터로 확인하지 않았으므로 단정하지 않습니다',
    tagText: '결과 분석', tagKind: 'dark', source: `${SRC} › 산출결과·검증_SKU (공통 범위 합계, 오차기여율)`,
  });
  const lx = MX, lw = 5.9, top = 2.4, base = 4.95;
  RR(s, lx, 1.72, lw, 4.2, C.light);
  T(s, '주수요 합계 (공통 범위, 그룹 SKU 합)', { x: lx + 0.25, y: 1.82, w: lw - 0.5, h: 0.32, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
  const series = [['hist', `과거 실적 평균 (${HIST})`, C.mute], ['fcst', 'FCST 평균 (8/24)', 'C8A25E'], ['wavg', '가중평균 예측', C.teal], ['actual', `실제 (${VAL})`, C.navy]];
  const grpList = [best, worst];
  const maxV = Math.max(...grpList.flatMap((g) => series.map(([k]) => GS[g][k]))) * 1.15;
  const slot = (lw - 0.5) / 2, bw = 0.5;
  line(s, lx + 0.25, base, lw - 0.5, 0, C.mute);
  grpList.forEach((g, gi) => {
    const sx = lx + 0.25 + gi * slot + (slot - series.length * (bw + 0.08)) / 2;
    series.forEach(([k, , col], i) => {
      const v = GS[g][k], bh = (v / maxV) * (base - top);
      R(s, sx + i * (bw + 0.08), base - bh, bw, bh, col);
      T(s, int(v), { x: sx + i * (bw + 0.08) - 0.15, y: base - bh - 0.28, w: bw + 0.3, h: 0.26, fontSize: 11, bold: true, align: 'center', valign: 'bottom', color: k === 'actual' ? C.navy : C.text });
    });
    T(s, `${g} (${GS[g].n} SKU)`, { x: lx + 0.25 + gi * slot, y: base + 0.05, w: slot, h: 0.32, fontSize: 12.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  });
  series.forEach(([, l, col], i) => {
    const x = lx + 0.25 + (i % 2) * 2.85, y = 5.38 + Math.floor(i / 2) * 0.26;
    R(s, x, y + 0.07, 0.18, 0.13, col);
    T(s, l, { x: x + 0.25, y, w: 2.55, h: 0.24, fontSize: 10.5, color: C.text, valign: 'middle' });
  });
  const rx = 6.8, rw = W - MX - rx;
  label(s, `오차 기여 상위 ${D.top.rows.length}개 SKU (가중평균)`, rx, 1.72, rw);
  const rows = [['SKU', '그룹', '예측(주)', '실제(주)', '방향', '기여율']];
  D.top.rows.forEach((r) => rows.push([r.sku, r.group, f1(r.forecast), f1(r.actual), r.error > 0 ? '과대' : '과소', cellOpt(pct(r.share), { bold: true })]));
  table(s, rows, { x: rx, y: 2.1, w: rw, colW: [1.85, 0.7, 0.95, 0.95, 0.65, rw - 5.1], rowH: 0.38, fontSize: 11, headSize: 11.5 });
  T(s, `기여율 = |수요오차| ÷ 전체 Σ|수요오차| (채점 ${D.top.n_scored} SKU) · 상위 ${D.top.rows.length}개 합 ${pct(D.top.share_sum, 0)}`,
    { x: rx, y: 4.45, w: rw, h: 0.32, fontSize: 10.5, color: C.gray, valign: 'middle' });
  const gShare = D.top.group_share[worst];
  RR(s, rx, 4.85, rw, 1.07, C.amberL);
  T(s, lines([
    `${worst} ${GS[worst].n}개 SKU를 합치면 가중평균 전체 오차의 ${pct(gShare, 0)}입니다.`,
    '개별 SKU는 작아 상위 표에 없지만 그룹 전체로는 가장 큽니다.',
  ]), { x: rx + 0.18, y: 4.9, w: rw - 0.36, h: 0.97, fontSize: 11.5, color: C.text, valign: 'middle', paraSpaceAfter: 2 });
  const bias = (b) => pct(D.metrics[worst][b].bias, 0);
  box(s, lines([
    `관찰: ${worst} 실제 출고(${int(GS[worst].actual)})가 과거 실적(${int(GS[worst].hist)})·FCST(${int(GS[worst].fcst)})보다 컸고, 세 기준 모두 낮게 예측했습니다.`,
    `편향(+ 과대 · − 과소): 가중평균 ${bias('가중평균')} · 실적기준 ${bias('실적기준')} · FCST기준 ${bias('FCST기준')}(두 기준은 보정계수 적용 후). 원인은 현업 확인이 필요합니다.`,
  ]), MX, 6.05, CW, 0.75, { fill: C.navy, color: C.white, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06, paraSpaceAfter: 2 });
})();

// ======================================================================= 해석과 한계
(function limits() {
  const SP = D.supplement;
  const s = slide({
    title: '해석과 한계: 말할 수 있는 것과 아직 말할 수 없는 것',
    sub: '개선 주장은 채점에 쓰지 않은 구간의 수치로만 하고, 나빠진 지표도 함께 적었습니다',
    tagText: '해석', tagKind: 'light', source: `${SRC}${WS ? ' · ' + WS.source : ''}`,
  });
  const lw = 5.3;
  RR(s, MX, 1.72, lw, 5.05, C.navy);
  T(s, '말할 수 있는 것', { x: MX + 0.3, y: 1.84, w: lw - 0.6, h: 0.4, fontSize: 16, bold: true, color: C.white, valign: 'middle' });
  const can = [
    `세 산식 기준 중 ${bestWape}의 수요 오차가 가장 낮았다 (오차율 ${pct(M[bestWape].wape)} · 사후 검증 구간 ${VAL} · 공통 ${M['가중평균'].n_demand} SKU)`,
    `산식은 기존수기보다 재고가 적고 (재고일수 ${f1(M['가중평균'].days_median)} 대 ${f1(M['기존수기'].days_median)}일) 결품 방어율도 낮다 (${pct(M['가중평균'].cover)} 대 ${pct(M['기존수기'].cover)} · 공통 ${M['가중평균'].n_lt} SKU)`,
    `재고등급 규칙은 기존 등급 ${D.grades.total}행과 같다 (리드타임·품질 등급은 현업 입력값)`,
    wsReady
      ? `AI 가중치 탐색: 검증 구간 ${HOLD}(${WS.contributions ? WS.contributions.n : ''} SKU)에서 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}, 개선은 ${topContribText}에 몰림, 방어율 ${pct(wsCur.cover)} → ${pct(wsNew.cover)} (①과 구간이 다름)`
      : `가중치 탐색 결과: ${PENDING}`,
  ];
  const ih = 1.1;
  can.forEach((t, i) => {
    const y = 2.32 + i * ih;
    num(s, i + 1, MX + 0.3, y + (ih - 0.42) / 2 - 0.05, 0.42, C.teal, C.white, 13);
    T(s, t, { x: MX + 0.9, y, w: lw - 1.15, h: ih - 0.1, fontSize: 11.5, color: C.white, valign: 'middle' });
  });
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  T(s, '아직 말할 수 없는 것 (한계)', { x: rx, y: 1.72, w: rw, h: 0.4, fontSize: 16, bold: true, color: C.amber, valign: 'middle' });
  const lim = [
    ['target', '결품 방어는 대리 지표', '실제 결품이 아니라 목표재고량이 실제 리드타임 수요를 덮었는지만 봅니다. 기존수기 목표재고량은 업무 가정으로 만든 값입니다.'],
    ['calendar', '선택 구간 4주', 'W38에 다시 고르는 상황을 흉내 낸 것이라 실제 운영보다 낙관적일 수 있습니다.'],
    ['chart', '유의성 검정 없음', `공통 ${M['가중평균'].n_demand} SKU, 그룹은 1~${Math.max(...groupsOrdered.map((g) => D.metrics[g]['가중평균'].n_demand))}개. 차이가 우연인지 검정하지 않았습니다.`],
    ['doc', 'FCST 스냅샷 1개', '8/24 한 시점만 썼습니다. 주마다 다른 FCST로 같은 비교를 반복해야 합니다.'],
    ['search', '채점 범위', `${D.validation.total} SKU 중 가중평균 수요 채점은 ${D.validation.by_basis['가중평균'].demand_scored}개입니다. 입력이 없는 SKU는 사유와 함께 뺐습니다.`],
    ['db', '입력 일부 역산', SP ? `확인되지 않은 입력 ${int(SP.total)}칸(기말재고 예상·FCST·현재고)은 산식과 다른 시트 값으로 역산했습니다. 원본 파일로 재확인합니다.` : '원본 파일로 입력을 재확인합니다.'],
  ];
  const step = 0.775;
  lim.forEach(([ic, a, b], i) => {
    const y = 2.18 + i * step;
    RR(s, rx, y, rw, step - 0.07, C.amberL);
    icon(s, ic, rx + 0.14, y + (step - 0.07 - 0.5) / 2, 0.5, C.amber);
    T(s, a, { x: rx + 0.78, y, w: 1.65, h: step - 0.07, fontSize: 12, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: rx + 2.45, y, w: rw - 2.6, h: step - 0.07, fontSize: 11, color: C.text, valign: 'middle' });
  });
})();

// ======================================================================= 간지 5
divider(4);

// ======================================================================= 운영 흐름 제안
(function ops() {
  const s = slide({
    title: '제안: MLOps로 매주 다시 고르고, 사람은 승인만 합니다',
    sub: '담당자별 할 일 · 매주 쌓이는 사후 검증이 다음 주 가중치 선택(재학습)의 근거가 됩니다',
    tagText: '제안', tagKind: 'amber', source: '발표자료 설계 §3 · 설계 문서 2026-11-19 §2(주 단위 구동)',
  });
  const lanes = [['SCM 담당', C.light], ['시스템 (자동)', C.tealL], ['구매 담당', C.light], ['MP 시스템 담당', C.light]];
  const ly0 = 1.72, lh = 1.0, lg = 0.1, lx = MX, lw = 1.55;
  lanes.forEach(([n, col], i) => {
    const y = ly0 + i * (lh + lg);
    R(s, lx + lw, y, CW - lw, lh, col);
    box(s, n, lx, y, lw - 0.06, lh, { fill: i === 1 ? C.teal : C.navy, color: C.white, bold: true, fontSize: 12.5 });
  });
  const steps = [
    [0, '① MP 시뮬레이션', '주 1회 FCST·기말재고'],
    [1, '② 자동 실행', '데이터 검사·등급·산출'],
    [1, '③ AI 재탐색', 'w 다시 고르기'],
    [2, '④ 검토·승인', '메모·미방어 SKU'],
    [3, '⑤ MP 배포', '일별 안전재고 교체'],
    [1, '⑥ 성능 모니터링', '오차율·방어율 추적'],
  ];
  const sx0 = lx + lw + 0.2, sw = 1.6, sg = (CW - lw - 0.4 - 6 * sw) / 5, bh = 0.7;
  const pos = steps.map(([lane], i) => ({ x: sx0 + i * (sw + sg), y: ly0 + lane * (lh + lg) + 0.1 }));
  steps.forEach(([lane, a, b], i) => {
    const { x, y } = pos[i];
    RR(s, x, y, sw, bh, C.white, { line: { color: lane === 1 ? C.teal : C.navy, width: 1.25 } });
    T(s, a, { x: x + 0.08, y: y + 0.03, w: sw - 0.16, h: 0.32, fontSize: 11, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: x + 0.08, y: y + 0.36, w: sw - 0.16, h: 0.32, fontSize: 10, color: C.slate, valign: 'top' });
    if (i < steps.length - 1) {
      const n = pos[i + 1];
      const ax = x + sw, ay = y + bh / 2, bxx = n.x, byy = n.y + bh / 2;
      if (Math.abs(ay - byy) < 0.01) arrow(s, ax + 0.02, ay, bxx - ax - 0.04, 0, C.slate);
      else {
        line(s, ax, ay, (bxx - ax) / 2, 0, C.slate, { width: 1.5 });
        line(s, ax + (bxx - ax) / 2, Math.min(ay, byy), 0, Math.abs(byy - ay), C.slate, { width: 1.5 });
        arrow(s, ax + (bxx - ax) / 2, byy, (bxx - ax) / 2 - 0.02, 0, C.slate);
      }
    }
  });
  // 되먹임: ⑥ → ② (다음 주 가중치·보정계수)
  const p2 = pos[1], p6 = pos[5], fy = p2.y + bh + 0.1;
  line(s, p6.x + sw / 2, p6.y + bh, 0, fy - p6.y - bh, C.amber, { width: 1.5, dashType: 'dash' });
  line(s, p2.x + sw / 2, fy, p6.x - p2.x, 0, C.amber, { width: 1.5, dashType: 'dash' });
  arrow(s, p2.x + sw / 2, fy, 0, -0.09, C.amber, { width: 1.5 });
  box(s, '다음 주 재학습: 가중치·보정계수 다시 고르기', pos[3].x + 0.15, fy - 0.12, 2.9, 0.24, { fill: C.tealL, color: '8A6420', fontSize: 10, bold: true, margin: 0 });
  const by = ly0 + 4 * (lh + lg) + 0.05;
  box(s, [
    { text: 'MLOps 구성  ', options: { bold: true, color: C.amber } },
    { text: '재현 파이프라인·자동 테스트·버전 관리는 구현했고, 주간 재탐색·성능 모니터링(오차율이 기준을 넘으면 알림)·승인 후 배포가 이번 제안입니다.', options: { color: C.white } },
  ], MX, by, CW, 0.6, { fill: C.navy, fontSize: 12, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 다음 단계
(function next() {
  const s = slide({
    title: '다음 단계: 원본 파일로 다시 확인한 뒤 MP에 반영합니다',
    sub: '순서대로 진행합니다 · 결과와 발표 숫자는 명령 한 번으로 다시 만들어집니다',
    tagText: '계획', tagKind: 'light', source: '설계 문서 비목표·한계 항목 · 이번 결과(결과 1~3, 가중치 탐색, 한계)',
  });
  const GSw = D.group_sums[worstGroup];
  const items = [
    ['원본 파일로 재확인', '역산해 채운 입력을 원본 파일 값으로 바꿔 같은 산출·탐색·검증을 다시 돌림', '가장 먼저',
      `역산해 채운 입력 ${D.supplement ? int(D.supplement.total) : '?'}칸 · 명령 한 번으로 재실행`],
    ['MP 시스템 반영', '산식 안전재고일수를 일별 안전재고 값으로 올리는 업로드 형식 확정', '재검증 후',
      `산식 안전재고일수 중앙값 ${f1(SS.new_median)}일 (기존 고정값 ${oldLevels.join('·')}일)`],
    ['서비스 수준(Z) 조정', '재고를 줄인 만큼 결품 방어율이 낮아진 것을 Z값으로 다시 맞춤', '현업 협의',
      `결품 방어율 가중평균 ${pct(M['가중평균'].cover)} 대 기존수기 ${pct(M['기존수기'].cover)}`],
    ['등급·보정계수 기준', '리드타임 등급 기준을 문서로 정하고, 보정계수 k도 같은 방법으로 탐색', '현업 협의',
      `보정계수 이상 ${D.memos.counts['보정계수이상'] || 0}개 · 없음 ${D.memos.counts['보정계수없음'] || 0}개`],
    ['주간 재탐색', 'FCST 스냅샷을 매주 쌓아 가중치를 다시 고르고 같은 비교를 반복 → 유의성 판단', '매주 자동',
      '지금은 8/24 스냅샷 1개, 선택 구간 4주'],
    [`${worstGroup} 수요 확인`, `${worstGroup} 출고가 과거보다 크게 늘어난 원인을 현업과 확인`, '현업 확인',
      `실제 ${int(GSw.actual)} 대 과거 ${int(GSw.hist)} (주수요 합) · 방어율 ${pct(D.metrics[worstGroup]['가중평균'].cover, 0)}`],
  ];
  const cw = (CW - 1.0) / 6, y = 1.8;
  line(s, MX + cw / 2, y + 0.3, CW - cw, 0, C.line, { width: 2 });
  items.forEach(([a, b, c, ev], i) => {
    const x = MX + i * (cw + 0.2);
    num(s, i + 1, x + cw / 2 - 0.3, y, 0.6, i === 0 ? C.teal : C.navy, C.white, 18);
    RR(s, x, y + 0.8, cw, 3.65, C.light);
    T(s, a, { x: x + 0.15, y: y + 0.9, w: cw - 0.3, h: 0.62, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: x + 0.15, y: y + 1.55, w: cw - 0.3, h: 1.25, fontSize: 11, color: C.text });
    RR(s, x + 0.1, y + 2.85, cw - 0.2, 0.98, C.white);
    T(s, [{ text: '근거', options: { bold: true, color: C.teal, breakLine: true } }, { text: ev, options: { color: C.text } }],
      { x: x + 0.18, y: y + 2.89, w: cw - 0.36, h: 0.9, fontSize: 10.5, valign: 'top' });
    box(s, c, x + 0.15, y + 3.95, cw - 0.3, 0.36, { fill: C.white, lineColor: C.line, color: C.slate, fontSize: 10.5, round: 0.18 });
  });
  box(s, '원본 파일로 같은 탐색·검증을 다시 돌려 결론이 유지되는지 확인한 뒤 MP에 반영합니다.',
    MX, 6.4, CW, 0.46, { fill: C.amberL, color: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 계획 목표 대비 달성도
(function achievement() {
  const s = slide({
    title: '계획 목표 대비 달성도',
    sub: '현재 근거로 말할 수 있는 수준을 판정했습니다',
    tagText: '계획 대비', tagKind: 'light', source: `${SRC} · ${WS ? WS.source : 'out/가중치탐색.xlsx(산출 예정)'}`,
  });
  const V = { ok: ['달성', C.teal], part: ['부분', C.amber], no: ['미검증', C.red], next: ['후속', C.slate], wip: ['진행 중', C.slate] };
  const ws = wsReady
    ? [`검증 구간 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)} (개선은 ${topContribText}), 결품 방어율 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}`, (wsNew.wape < wsCur.wape && wsNew.cover >= wsCur.cover) ? 'ok' : 'part']
    : [`${PENDING} (탐색 결과 파일 대기)`, 'wip'];
  const rows = [
    ['안전재고일수 산출 (세 기준)', `${D.sku_total} SKU 처리, 입력 누락은 점검메모로 표시`, 'ok'],
    ['재고등급 산출', `${D.grades.match}/${D.grades.total} 기존 등급과 일치. 리드타임·품질 등급은 입력값으로 남음`, 'part'],
    ['AI 최적 가중치 탐색', ws[0], ws[1]],
    ['더 맞는 값 찾기 (사후 검증)', `세 산식 중 ${bestWape} 오차율 ${pct(M[bestWape].wape)} 최저 · 기존수기보다 재고는 적고 결품 방어율도 낮음`, 'part'],
    ['MLOps 파이프라인', `재현 파이프라인·자동 테스트 ${int(TESTS)}개 구현, 주간 재탐색·모니터링은 제안`, 'part'],
    ['통계적 유의성', `검정하지 않음 (공통 ${M['가중평균'].n_demand} SKU, FCST 스냅샷 1개)`, 'no'],
    ['MP 시스템 반영·주간 운영', '운영안 제안까지 (원본 파일 재확인 후 반영)', 'next'],
  ];
  const tr = [['계획 목표', '확인한 결과', '판정']];
  rows.forEach(([a, b, v]) => tr.push([a, b, cellOpt(V[v][0], { bold: true, color: C.white, fill: { color: V[v][1] }, align: 'center' })]));
  table(s, tr, { x: MX, y: 1.72, w: CW, colW: [3.2, CW - 3.2 - 1.5, 1.5], rowH: 0.6, fontSize: 12, headSize: 12.5 });
  T(s, '판정 기준: 달성 = 근거 수치로 확인 · 부분 = 일부만 확인되거나 나빠진 지표가 있음 · 미검증 = 확인하지 않음 · 후속 = 이번 범위 밖',
    { x: MX, y: 6.55, w: CW, h: 0.3, fontSize: 10.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 정리
(function closing() {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.navy };
  T(s, '정리', { x: 0.9, y: 0.45, w: 8, h: 0.7, fontSize: 32, bold: true, color: C.white, valign: 'middle' });
  T(s, '왜 만들었고, 무엇을 비교했고, 결과는 어땠는가', { x: 0.9, y: 1.12, w: 9.5, h: 0.38, fontSize: 14, color: C.sky, valign: 'middle' });
  tag(s, '정리', 'outline');
  num(s, 1, 0.9, 1.78, 0.62, C.teal, C.white, 18);
  T(s, '왜 만들었나', { x: 1.8, y: 1.7, w: 10.6, h: 0.36, fontSize: 15, bold: true, color: C.sky, valign: 'middle' });
  T(s, `MP 일별 안전재고가 고정값(${oldLevels.join('·')}일)이고, 실적과 FCST를 섞는 가중치가 최선인지 확인할 방법이 없었습니다.`,
    { x: 1.8, y: 2.06, w: 10.6, h: 0.42, fontSize: 13.5, color: C.white, valign: 'middle' });
  num(s, 2, 0.9, 2.78, 0.62, C.teal, C.white, 18);
  T(s, '무엇끼리 비교했나', { x: 1.8, y: 2.7, w: 10.6, h: 0.36, fontSize: 15, bold: true, color: C.sky, valign: 'middle' });
  T(s, `가중치 ${man(SPACE)} 조합을 AI 방식으로 탐색해 고르고, 고를 때 쓰지 않은 실제 출고와 네 가지 기준을 같은 잣대로 비교했습니다.`,
    { x: 1.8, y: 3.06, w: 10.6, h: 0.42, fontSize: 13.5, color: C.white, valign: 'middle' });
  num(s, 3, 0.9, 3.78, 0.62, C.teal, C.white, 18);
  T(s, '결과는', { x: 1.8, y: 3.7, w: 10.6, h: 0.36, fontSize: 15, bold: true, color: C.sky, valign: 'middle' });
  const res = [
    wsReady ? `AI 탐색 가중치로 검증 구간 수요 오차율이 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}로 낮아졌습니다 (개선은 ${topContribText}에 몰림).` : `가중치 탐색 결과는 ${PENDING}입니다.`,
    `세 산식 기준 중 ${bestWape}의 수요 오차가 가장 낮았습니다 (오차율 ${pct(M[bestWape].wape)}, ${VAL}).`,
    `산식은 기존수기보다 재고가 적고(${f1(M['가중평균'].days_median)} 대 ${f1(M['기존수기'].days_median)}일) 결품 방어율도 낮습니다(${pct(M['가중평균'].cover)} 대 ${pct(M['기존수기'].cover)}). 서비스 수준(Z) 조정이 다음 과제입니다.`,
  ];
  T(s, res.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < res.length - 1 } })),
    { x: 1.8, y: 4.08, w: 10.8, h: 2.15, fontSize: 13, color: C.white, valign: 'top', paraSpaceAfter: 5 });
  T(s, '감사합니다 · 질의응답', { x: 0.9, y: 6.4, w: 8, h: 0.45, fontSize: 17, bold: true, color: C.sky, valign: 'middle' });
  footer(s, `${SRC} › 검증요약(공통)${WS ? ' · ' + WS.source : ''}`, true);
})();

// ======================================================================= 부록 1. 용어 풀이
(function glossary() {
  const s = slide({
    title: '부록 1. 용어와 기호 풀이',
    sub: '본문에서 쓴 약어·지표·구간 이름과 산식 기호',
    tagText: '부록', tagKind: 'light', source: `${SRC} › 적용기준 · README.md`,
  });
  const left = [
    ['안전재고', '수요가 예측보다 많을 때를 대비해 더 들고 있는 재고'],
    ['안전재고일수', '안전재고 ÷ 일평균 수요. MP 일별 안전재고를 대체할 값'],
    ['적정 재고일수', '리드타임 + 안전재고일수. 발주점을 일수로 본 값(비교용)'],
    ['리드타임 (LT)', '발주에서 입고까지 걸리는 일수'],
    ['발주점 (ROP)', '리드타임 수요 + 안전재고. 현재고가 이 아래면 발주'],
    ['Z값', `재고등급별 서비스 수준 계수 (${String(D.criteria['Z값'] || '').replace(/,\s*/g, ' · ')}). 클수록 안전재고가 큼`],
    ['수요 편차 (σ)', '주간 수요가 평균에서 흔들리는 정도 (모표준편차) · 보정계수 k: 실적·FCST기준에 곱하는 현업 입력값'],
    ['수요 오차율 (WAPE)', 'Σ|예측 − 실제| ÷ Σ실제. 낮을수록 잘 맞음'],
    ['편향', 'Σ(예측 − 실제) ÷ Σ실제. +는 과대예측, −는 과소예측'],
    ['기존수기', `지금 MP의 고정 일별 안전재고(${oldLevels.join('·')}일)로 만든 비교 기준`],
  ];
  const right = [
    ['목표재고량', '산식 기준은 발주점, 기존수기는 리드타임 수요 + 기존 안전재고일수 × 일평균 수요'],
    ['결품 방어율', '목표재고량이 실제 리드타임 수요 이상인 SKU 비율 (실제 결품이 아닌 대리 지표, 못 덮으면 "미방어")'],
    ['실제 기준 재고일수', '목표재고량 ÷ 실제 일평균 출고'],
    ['공통 범위', '비교하는 모든 기준이 채점된 SKU만 모은 범위'],
    ['사후 검증 구간', `8/24 계산값을 채점한 ${VAL} (9주)`],
    ['선택 구간 · 검증 구간', `가중치 탐색에서 w를 고른 ${SEL} · 확인한 ${HOLD} (시계열 검증)`],
    ['실적 가중치 (w)', '가중평균 수요에서 실적평균의 비중 (FCST는 1 − w)'],
    ['그룹 · 전체 합산 최적', `그룹 안에서 고른 w · SKU ${MINSKU}개 미만 그룹은 전체를 합쳐 고른 w (과적합 방지)`],
    ['그리드 탐색', '후보 값을 빠짐없이 계산해 가장 좋은 값을 고르는 머신러닝 하이퍼파라미터 탐색 방법'],
    ['MLOps', '데이터 → 산출 → 검증 → 배포 → 모니터링을 자동·재현 가능하게 운영하는 방식'],
    ['산식 기호', 'Hm 실적평균 · Fm FCST평균 · wh/wf 가중치 · G 가중평균 수요 · K 리드타임 수요 · L 리드타임 수요 편차 · T 안전재고 · S 안전재고일수 · U 적정 재고일수'],
  ];
  const tw = (CW - 0.3) / 2;
  [left, right].forEach((list, k) => {
    const rows = [['용어', '뜻']];
    list.forEach(([a, b]) => rows.push([cellOpt(a, { bold: true, color: C.navy }), b]));
    table(s, rows, { x: MX + k * (tw + 0.3), y: 1.55, w: tw, colW: [1.75, tw - 1.75], rowH: 0.4, fontSize: 10, headSize: 11 });
  });
})();

// ======================================================================= 부록 2. 점검메모
(function appendixMemo() {
  const s = slide({
    title: '부록 2. 점검메모 10종과 이번 데이터의 SKU 수',
    sub: `산출결과 시트의 점검메모 열 · 여러 개면 쉼표로 함께 붙음 · ${D.memos.total} SKU 중 ${D.memos.with_memo}개에 메모`,
    tagText: '부록', tagKind: 'light', source: `${SRC} › 산출결과(점검메모) · README.md`,
  });
  const meaning = [
    ['구매마스터없음', '출고·FCST에만 있고 구매마스터에 없는 SKU'],
    ['등급없음', '계산 재고등급이 없음 (Z 관련 값 비움)'],
    ['수요없음', '가중평균 주수요가 0 (일수 비움)'],
    ['실적없음', '출고 시트에 SKU가 없거나 실적 주차 값이 모두 빈 칸'],
    ['FCST없음', 'SCM FCST 시트에 SKU가 없음'],
    ['가중치없음', '실적 가중치나 FCST 가중치가 비어 있음'],
    ['가중치합오류', '두 가중치의 합이 1이 아님 (값은 그대로 사용)'],
    ['리드타임없음', '리드타임이 없거나 0 이하'],
    ['보정계수없음', '과소예측 보정계수가 없음 (실적·FCST기준 비움)'],
    ['보정계수이상', '보정계수가 0이거나 3보다 큼 (계산은 그대로, 실적·FCST기준 채점 제외)'],
  ];
  const rows = [['메모', '뜻', 'SKU 수']];
  meaning.forEach(([k, v]) => rows.push([cellOpt(k, { bold: true, color: C.navy }), v, cellOpt(String(D.memos.counts[k] || 0), { align: 'center', bold: true })]));
  table(s, rows, { x: MX, y: 1.72, w: CW, colW: [2.3, CW - 2.3 - 1.4, 1.4], rowH: 0.43, fontSize: 12, headSize: 12 });
  T(s, `참고: 가중평균 발주 판정 — 발주필요 ${D.memos.order_needed} · 정상 ${D.memos.order_ok} · 판정 없음 ${D.memos.total - D.memos.order_needed - D.memos.order_ok}`,
    { x: MX, y: 6.6, w: CW, h: 0.3, fontSize: 10.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 부록 3. 산식·규칙 상세
(function appendixRules() {
  const s = slide({
    title: '부록 3. 산식과 사후 검증 규칙 상세',
    sub: '적용기준 시트 문구 그대로 · WK는 본문의 W와 같은 표기 · 이 표의 "검증 구간"은 본문의 사후 검증 구간 W35~43',
    tagText: '부록', tagKind: 'light', source: `${SRC} › 적용기준`,
  });
  const keys = ['실적 주차', 'FCST 주 수', '검증 주차', '평균·σ 정의', '재고등급', '기존수기 목표재고량', '결품 방어', '범위 정의'];
  const rows = [['항목', '내용']];
  keys.filter((k) => D.criteria[k]).forEach((k) => rows.push([cellOpt(k === '검증 주차' ? '사후 검증 주차' : k, { bold: true, color: C.navy }), String(D.criteria[k])]));
  table(s, rows, { x: MX, y: 1.65, w: CW, colW: [2.2, CW - 2.2], fontSize: 10.5, headSize: 11, rowH: [0.36, 0.34, 0.34, 0.34, 0.5, 0.72, 0.5, 0.36, 0.72] });
})();

// ======================================================================= 부록 4. 가중치 탐색 규칙
(function appendixWs() {
  const s = slide({
    title: '부록 4. 가중치 탐색 규칙',
    sub: '가중치탐색.xlsx 적용기준 시트 문구를 그대로 옮겼습니다',
    tagText: '부록', tagKind: 'light', source: WS ? `${WS.source} › 적용기준` : 'out/가중치탐색.xlsx (산출 예정)',
  });
  const keys = ['선택 구간', '검증 구간', '참고 전체 구간', '후보', '단위', '고르는 기준', '동점 규칙', '작은 그룹 규칙', '현재유지', '한계'];
  const rows = [['항목', '내용']];
  if (WS) keys.filter((k) => WS.criteria[k]).forEach((k) => rows.push([cellOpt(k, { bold: true, color: C.navy }), String(WS.criteria[k])]));
  else rows.push(['전체', PENDING]);
  if (WS && WS.contributions) rows.push([cellOpt('그룹별 감소 몫', { bold: true, color: C.navy }), `extract.py 계산: ${WS.contributions.note}. 다시 계산한 전체 오차율 ${pct(WS.contributions.wape_current_recomputed)} → ${pct(WS.contributions.wape_search_recomputed)} (파일 값과 같음)`]);
  table(s, rows, { x: MX, y: 1.65, w: CW, colW: [2.2, CW - 2.2], fontSize: 11, headSize: 11, rowH: 0.42 });
})();

// ======================================================================= 부록 5. 실행 명령
(function appendixCmd() {
  const s = slide({
    title: '부록 5. 다시 돌리는 명령과 코드 구성',
    sub: '입력이 바뀌면 아래 순서로 다시 돌리면 발표 숫자가 함께 바뀝니다 (숫자는 손으로 적지 않음) · PowerShell 기준',
    tagText: '부록', tagKind: 'light', source: 'README.md · presentation/extract.py · presentation/build.js',
  });
  const cmds = [
    ["$BASE = '<기준 파일: 8/28 재고등급·안전재고 점검 xlsx>'", C.white],
    ["$ACT  = 'data\\W44_가중평균수요_반영_상품_발주점검_업로드.xlsx'", C.white],
    ['# 1. 산출·사후 검증 / 2. AI 가중치 탐색', C.mute],
    [`uv run python -m safety_stock $BASE --actual $ACT -o ${SRC.replace(/\//g, '\\')}`, C.white],
    ['uv run python -m safety_stock.weight_search $BASE --actual $ACT -o out\\가중치탐색.xlsx', C.white],
    ['# 3. 발표 숫자 추출 → 발표자료 생성 / 4. 자동 테스트', C.mute],
    ['uv run python presentation\\extract.py ; node presentation\\build.js', C.white],
    ['uv run pytest', C.white],
  ];
  RR(s, MX, 1.72, CW, 2.75, '16263A');
  T(s, cmds.map(([t, col], i) => ({ text: t, options: { color: col, breakLine: i < cmds.length - 1 } })),
    { x: MX + 0.3, y: 1.8, w: CW - 0.6, h: 2.6, fontFace: 'Consolas', fontSize: 11, valign: 'middle', paraSpaceAfter: 3 });
  const files = [
    ['입력', '기준 파일(8/28 재고등급·안전재고 점검) · W44 업로드 파일(사후 실제 출고, 검증 전용)'],
    ['결과', `${SRC} · ${D.sources.search}`],
    ['발표', 'presentation/data.json → out/발표_적정재고일수.pptx'],
    ['코드 영역', 'purchase_master(구매마스터·재고등급) · mp_simulation(MP 시뮬레이션) · inventory_days(재고일수·사후 검증) · weight_search(AI 가중치 탐색) · presentation(발표 자동 생성)'],
    ['검증', `자동 테스트 ${int(TESTS)}개 (uv run pytest)`],
  ];
  const rows = [['구분', '파일']];
  files.forEach(([a, b]) => rows.push([cellOpt(a, { bold: true, color: C.navy }), b]));
  table(s, rows, { x: MX, y: 4.62, w: CW, colW: [1.3, CW - 1.3], rowH: 0.38, fontSize: 10.5, headSize: 11 });
})();

// ---------------------------------------------------------------- 저장
// 글자 run의 언어를 한국어로 표시한다. 그래야 PowerPoint가 한글을 어절 단위로 줄바꿈한다(단어 중간에서 끊기지 않음).
async function save() {
  const buf = await pres.write({ outputType: 'nodebuffer' });
  const zip = await JSZip.loadAsync(buf);
  const files = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  for (const f of files) {
    const xml = (await zip.file(f).async('string'))
      .replace(/lang="en-US"/g, 'lang="ko-KR" altLang="en-US"')
      .replace(/(<a:ea typeface="[^"]*" pitchFamily="\d+" charset=")-122"/g, '$1-127"');
    zip.file(f, xml);
  }
  fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
  fs.writeFileSync(OUT_PATH, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`${path.relative(ROOT, OUT_PATH)} 작성: ${page}장, 가중치 탐색 ${wsReady ? '반영' : PENDING}`);
}
save();
