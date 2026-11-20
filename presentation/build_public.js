// 발표자료 공개본 생성: presentation/data.json → out/발표_적정재고일수_공개본.pptx (15장, 비식별)
// 대학원 배포용이다. 품목그룹은 '그룹 A~G'로, 사례 품목은 '품목 X'로 쓰고 절대 수량은 쓰지 않는다.
// 숫자는 모두 data.json(= extract.py가 결과 파일에서 읽은 값)에서 가져온다. 손으로 적은 숫자는 없다.
// 사내본(build.js)은 건드리지 않는다. 양식 도우미는 build.js에서 그대로 복사했다.
// 실행: node presentation/build_public.js [--data presentation/data.json] [--out out/발표_적정재고일수_공개본.pptx]
'use strict';
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

// ---------------------------------------------------------------- 인자·데이터
const ROOT = path.resolve(__dirname, '..');
function arg(name, def) {
  const i = process.argv.indexOf(name);
  if (i < 0) return def;
  const v = process.argv[i + 1];
  if (v === undefined || v.startsWith('--')) throw new Error(`${name} 옵션에 경로 값이 없습니다. 예: ${name} <경로>`);
  return path.resolve(v);
}
const DATA_PATH = arg('--data', path.join(__dirname, 'data.json'));
const OUT_PATH = arg('--out', path.join(ROOT, 'out', '발표_적정재고일수_공개본.pptx'));
const D = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));

// ---------------------------------------------------------------- 양식 상수
const FONT = '맑은 고딕';
const FOOT = 'AI 기반 적정 재고일수 산출 · 테스트 데이터 기반';
const W = 13.333, MX = 0.6, CW = W - 2 * MX; // LAYOUT_WIDE 13.333 x 7.5 in
const C = {
  navy: '1B2E44', navy2: '27405C', teal: '0B7F80', tealL: 'DCEEEE', slate: '52657A', gray: '5F6F80',
  light: 'F0F3F6', line: 'D3DBE3', text: '1F2D3D', amber: 'C8891A', amberL: 'F6ECDA', red: 'C0392B',
  redL: 'F9E6E3', white: 'FFFFFF', sky: 'A9B8C8', mute: '9AA8B6',
};

// ---------------------------------------------------------------- 숫자·주차 표기
const isNum = (v) => typeof v === 'number' && !Number.isNaN(v);
const pct = (v, d = 1) => (isNum(v) ? (v * 100).toFixed(d) + '%' : '–');
// 차이는 화면에 보이는 반올림 값끼리 뺀다(예: 69.0% − 60.6% = 8.4%p)
const pp = (a, b) => ((Math.round(a * 1000) - Math.round(b * 1000)) / 10).toFixed(1) + '%p';
const f1 = (v) => (isNum(v) ? v.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : '–');
const int = (v) => (isNum(v) ? Math.round(v).toLocaleString('en-US') : '–');
const wfmt = (v) => (isNum(v) ? v.toFixed(1) : '–');
// 'WK35~WK43' 같은 주차 범위를 'W35~43'으로 통일한다
function wr(s) {
  const m = String(s || '').match(/W[K]?(\d+)\D+W?[K]?(\d+)/);
  return m ? `W${m[1]}~${m[2]}` : String(s || '');
}
// 주차 범위의 처음·끝 주 번호와 주 수
function wspan(s) {
  const m = String(s || '').match(/W[K]?(\d+)\D+W?[K]?(\d+)/);
  if (!m) throw new Error(`주차 범위를 읽지 못했습니다: ${s}`);
  const a = Number(m[1]), b = Number(m[2]);
  return { a, b, n: b - a + 1 };
}
const man = (v) => (isNum(v) ? Math.round(v / 1e4).toLocaleString('en-US') + '만' : '–');
// 부호를 붙인 비율(+12% / −3%)
const spct0 = (v) => (isNum(v) ? (v < 0 ? '−' : '') + pct(Math.abs(v), 0) : '–');
const spct = (v, d = 1) => (isNum(v) ? (v >= 0 ? '+' : '−') + pct(Math.abs(v), d) : '–');

// ---------------------------------------------------------------- 파생 값
const WS = D.weight_search;
const SRCH = WS ? WS.criteria : {};
function wsRow(period, scope, weights, basis) {
  if (!WS) return null;
  return WS.compare.find((r) => r.period === period && r.scope === scope && r.weights === weights && r.basis === basis) || null;
}
const wsCur = wsRow('검증', '전체', '현재', '가중평균');
const wsNew = wsRow('검증', '전체', '탐색', '가중평균');
const wsReady = !!(wsCur && wsNew && isNum(wsCur.wape) && isNum(wsNew.wape));
if (!WS || !wsReady) throw new Error('가중치 탐색 결과가 없습니다: extract.py를 먼저 실행');

const HIST = wr(D.criteria['실적 주차']); // 출고 실적 W01~31
const VAL = wr(D.criteria['검증 주차']); // 사후 검증 W35~43
const SEL = wr(SRCH['선택 구간']); // 선택 구간 W35~38
const HOLD = wr(SRCH['검증 구간']); // 검증 구간 W39~43
const HS = wspan(HIST), VS = wspan(VAL), SS_ = wspan(SEL), HO = wspan(HOLD);
const FW = (() => {
  const m = String(D.criteria['FCST 주 수'] || '').match(/(\d+)\s*주/);
  if (!m) throw new Error('예측 주 수를 읽지 못했습니다');
  return Number(m[1]);
})();
const FCST_END = SS_.a + FW - 1; // 예측이 덮는 마지막 주

// 문자열 코드포인트 비교(로캘에 따라 순서가 바뀌지 않게)
const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
// 품목그룹 별칭: 가중평균 채점 SKU 수 내림차순, 같으면 코드 오름차순으로 '그룹 A'부터 붙인다(코드에 대응표를 적지 않음)
const groupsOrdered = Object.keys(D.metrics)
  .filter((g) => !['전체', '공통', '미분류'].includes(g))
  .sort((a, b) => (D.metrics[b]['가중평균'].n_demand - D.metrics[a]['가중평균'].n_demand) || cmp(a, b));
const GA = {};
groupsOrdered.forEach((code, i) => { GA[code] = '그룹 ' + String.fromCharCode(65 + i); });
const ga = (code) => {
  if (!(code in GA)) throw new Error('별칭이 없는 품목그룹이 있습니다');
  return GA[code];
};
const byAlias = (a, b) => cmp(ga(a), ga(b));
const aliasLetters = (codes) => '그룹 ' + codes.map((c) => ga(c).slice(-1)).sort().join('·');

const SS = D.safety; // 안전재고일수(기존 고정값 대 산식)
const oldLevels = Object.keys(D.days.old_safety_levels);
const contrib = WS.contributions ? WS.contributions.groups : null;
const topContrib = contrib ? Object.entries(contrib).filter(([, v]) => isNum(v.share)).sort((a, b) => b[1].share - a[1].share).slice(0, 2) : [];
const topContribText = topContrib.map(([g, v]) => `${ga(g)} ${pct(v.share, 0)}`).join(' · ');
const METHOD = { '그룹최적': '그룹 최적', '공통최적': '전체 합산', '현재유지': '현재 유지' };
// 그룹 오차는 그 그룹의 w에만 달려 있어서 그룹 × 후보 계산으로 전체 조합에서 고른 것과 같은 답을 얻는다
const CANDS = [...new Set(WS.candidates.map((c) => c.w))].sort((a, b) => a - b);
const NG = WS.choices.length;
const SPACE = Math.pow(CANDS.length, NG);
const CSTEP = CANDS.length > 1 ? CANDS[1] - CANDS[0] : null;
const TESTS = isNum(D.tests) ? D.tests : null;
const SMALL = WS.choices.filter((c) => c.method === '공통최적').map((c) => c.group);
const MINSKU = (() => {
  const m = String(SRCH['작은 그룹 규칙'] || '').match(/(\d+)\s*개?\s*미만/);
  if (!m) throw new Error('작은 그룹 규칙의 SKU 기준을 읽지 못했습니다');
  return Number(m[1]);
})();
const WRANGE = D.weights.length
  ? `${wfmt(Math.min(...D.weights.map((w) => w.wh)))}~${wfmt(Math.max(...D.weights.map((w) => w.wh)))}` : '–';
const choicesSorted = [...WS.choices].sort((a, b) => byAlias(a.group, b.group));
const NHOLD = wsNew.n_demand; // 검증 구간 수요 채점 SKU 수
const coverDown = isNum(wsNew.cover) && isNum(wsCur.cover) && wsNew.cover < wsCur.cover;
const maxGroupN = Math.max(...groupsOrdered.map((g) => D.metrics[g]['가중평균'].n_demand));
const minGroupN = Math.min(...groupsOrdered.map((g) => D.metrics[g]['가중평균'].n_demand));

// ---------------------------------------------------------------- 프레젠테이션
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'AI 기반 적정 재고일수 산출';
pres.subject = '테스트 데이터로 개발·검증한 가중평균 수요 탐색';
pres.author = '';
pres.company = '';
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

// 도형으로 그린 아이콘(원 안). 외부 그림 파일 없이도 깨지지 않는다.
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
// 꼬리말: 왼쪽은 발표 이름과 데이터 표기, 오른쪽은 쪽 번호(근거 파일 경로·작성자는 넣지 않음)
function footer(s, dark = false) {
  const col = dark ? C.sky : C.gray;
  T(s, FOOT, { x: MX, y: 7.05, w: 8.0, h: 0.26, fontSize: 9, color: col, valign: 'middle' });
  T(s, String(page), { x: 12.25, y: 7.05, w: 0.48, h: 0.26, fontSize: 10, color: col, align: 'right', valign: 'middle', bold: true });
}
// 제목이 한 줄에 들어가도록 글자 크기를 줄인다(한글 1em, 영문·숫자 약 0.58em으로 어림)
function fitSize(text, w, max, min) {
  const em = [...String(text)].reduce((a, ch) => a + (ch === ' ' ? 0.32 : /[\x21-\x7e]/.test(ch) ? 0.58 : 1.0), 0);
  let size = max;
  while (size > min && (em * size) / 72 > w) size -= 0.5;
  return size;
}
function slide({ title, sub, tagText, tagKind }) {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.white };
  T(s, title, { x: MX, y: 0.4, w: 10.0, h: 0.62, fontSize: fitSize(title, 9.8, 24, 20), bold: true, color: C.navy, valign: 'middle' });
  if (sub) T(s, sub, { x: MX, y: 1.04, w: 10.0, h: 0.36, fontSize: fitSize(sub, 9.9, 13.5, 11.5), color: C.slate, valign: 'middle' });
  tag(s, tagText, tagKind);
  footer(s);
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
// 검증 값 '결품'은 실제 결품이 아니라 '목표재고 < 실제 리드타임 수요'(대리 지표)라서 화면에는 '미방어'로 쓴다
const cov = (v) => (v === '결품' ? '미방어' : v);

// ======================================================================= 1. 표지
(function cover() {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.navy };
  T(s, '대학원 프로젝트 발표 · 테스트 데이터셋(비식별 처리)', { x: 0.9, y: 1.05, w: 9.5, h: 0.4, fontSize: 15, color: C.sky });
  // 사용자 지시: 처음부터 '어떤 AI 기술을 활용한 무엇'인지, 최종 방향(MLOps)과 지금 단계를 밝힌다
  T(s, [
    { text: 'MLOps 기반', options: { fontSize: 36, color: '8FC9C9', breakLine: true } },
    { text: '적정 안전재고일수 산출 시스템', options: { fontSize: 42, color: C.white } },
  ], { x: 0.9, y: 1.6, w: 11.5, h: 1.95, bold: true, valign: 'middle', lineSpacingMultiple: 1.1 });
  T(s, [
    { text: '1단계: ', options: { bold: true, color: '8FC9C9' } },
    { text: 'AutoML(그리드 탐색 · 시계열 검증)로 실적·예측 최적 가중치를 자동 선택', options: { color: 'D5DEE7' } },
  ], { x: 0.9, y: 3.75, w: 11.5, h: 0.6, fontSize: 20, valign: 'middle' });
  R(s, 0.9, 4.75, 2.6, 0.06, C.teal);
  T(s, D.run_date, { x: 0.9, y: 5.1, w: 6, h: 0.4, fontSize: 15, bold: true, color: C.white });
  T(s, `가중치 ${man(SPACE)} 조합 탐색 · 시계열 분할 검증`, { x: 0.9, y: 5.55, w: 11.5, h: 0.36, fontSize: 13, color: C.sky });
  tag(s, '프로젝트 발표', 'outline');
  footer(s, true);
})();

// ======================================================================= 2. 한 장 요약
(function summary() {
  const s = slide({
    title: `AutoML 탐색으로 최적 가중평균을 찾아 오차율을 ${pp(wsCur.wape, wsNew.wape)} 낮췄습니다`,
    sub: '어떤 AI 기술로 무엇을 만들었고, 무엇이 문제였고, 무엇을 얻었는가 · 테스트 데이터셋(비식별 처리)',
    tagText: '요약', tagKind: 'dark',
  });
  // 사용자 지시: AI 기술(AutoML·HPO)을 먼저 말하고, 그다음 문제·결과
  const cards = [
    ['search', 'AI 기술 · AutoML(HPO)', ['그리드 탐색 + 시계열 분할 검증', '+ 과적합 방지(작은 그룹 합산)', `→ ${man(SPACE)} 조합을 ${CANDS.length * NG}번 계산으로 탐색해 선택`], C.teal],
    ['calendar', '문제', [`안전재고일수가 고정값 ${oldLevels.length}종(${oldLevels.join('·')}일)`, `실적·예측 가중치(${WRANGE})를 정한 근거 없음`], C.navy],
    ['chart', '결과', [`검증 구간 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}`, `고를 때 쓰지 않은 ${HOLD}(${HO.n}주)로 확인`], C.navy],
  ];
  const cw = (CW - 0.4) / 3, cy = 1.6, chh = 1.8;
  cards.forEach(([ic, head, body, col], i) => {
    const x = MX + i * (cw + 0.2);
    RR(s, x, cy, cw, chh, i === 0 ? C.tealL : C.light);
    icon(s, ic, x + 0.25, cy + 0.2, 0.6, col);
    T(s, head, { x: x + 1.0, y: cy + 0.25, w: cw - 1.15, h: 0.5, fontSize: 16.5, bold: true, color: col, valign: 'middle' });
    T(s, lines(body), { x: x + 0.28, y: cy + 0.92, w: cw - 0.45, h: 0.8, fontSize: 12, color: C.text, paraSpaceAfter: 2 });
  });
  label(s, `핵심 수치 · 검증 구간 ${HOLD} · 안전재고일수는 기존값·산식 값이 모두 있는 ${SS.n} SKU`, MX, 3.5, 10, C.navy, 13);
  const kw = (CW - 0.6) / 4;
  const kpis = [
    { v: pct(wsNew.wape), l: '검증 오차율 (AI 탐색 w)', n: `현재 w ${pct(wsCur.wape)}보다 ${pp(wsCur.wape, wsNew.wape)} 낮음`, c: C.teal },
    { v: man(SPACE), l: '탐색한 가중치 조합', n: `그룹 ${NG}개 × 후보 ${CANDS.length}개`, c: C.teal },
    { v: `${f1(SS.old_median)} → ${f1(SS.new_median)}일`, l: '안전재고일수 중앙값', n: `고정값 → 산식 · ${SS.shorter}/${SS.n} SKU 짧아짐`, c: C.teal },
    { v: `${pct(wsCur.cover)} → ${pct(wsNew.cover)}`, l: '결품 방어율 (대리 지표)', n: '서비스 수준(Z) 조정이 다음 과제', c: coverDown ? C.red : C.teal, warn: coverDown },
  ];
  kpis.forEach((k, i) => {
    const x = MX + i * (kw + 0.2), y = 3.88, h = 1.5;
    RR(s, x, y, kw, h, k.warn ? C.redL : C.white, { line: { color: k.warn ? 'E8B4AD' : C.line, width: 1 } });
    if (k.warn) box(s, '주의', x + kw - 0.78, y + 0.76, 0.6, 0.28, { round: 0.14, fill: C.red, color: C.white, fontSize: 10, bold: true });
    T(s, k.v, { x: x + 0.22, y: y + 0.1, w: kw - 0.44, h: 0.62, fontSize: fitSize(k.v, kw - 0.5, 28, 18), bold: true, color: k.c, valign: 'middle' });
    T(s, k.l, { x: x + 0.22, y: y + 0.74, w: kw - 0.44, h: 0.32, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
    T(s, k.n, { x: x + 0.22, y: y + 1.06, w: kw - 0.44, h: 0.36, fontSize: 10.5, color: C.slate, valign: 'middle' });
  });
  // MLOps 진행 띠: 최종 방향과 현재 단계(사용자 지시)
  const by = 5.52, bh = 0.92;
  RR(s, MX, by, CW, bh, C.light);
  T(s, lines(['MLOps', '진행 단계']), { x: MX + 0.2, y: by, w: 1.3, h: bh, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
  const lv = [
    ['Level 0', '수동 실험', 'done'],
    ['Level 1', '파이프라인 자동화·지속 학습', 'now'],
    ['Level 2', 'CI/CD 자동화', 'plan'],
  ];
  const sx = MX + 1.55, sw = (CW - 1.75 - 0.5) / 3, sy = by + 0.36, shh = 0.44;
  lv.forEach(([a, b, st], i) => {
    const x = sx + i * (sw + 0.25);
    const o = st === 'done' ? { fill: C.teal, color: C.white }
      : st === 'now' ? { fill: C.tealL, color: C.navy, lineColor: C.teal, lineWidth: 1.5 }
        : { fill: C.white, color: C.gray, dash: true, lineColor: C.mute };
    box(s, [{ text: a + '  ', options: { bold: true } }, { text: b }], x, sy, sw, shh, Object.assign({ fontSize: 12, round: 0.06, margin: 2 }, o));
    if (i < lv.length - 1) arrow(s, x + sw + 0.03, sy + shh / 2, 0.19, 0, C.slate, { width: 1.75 });
  });
  // 현재 위치: Level 0 → 1 경계
  const mx = sx + sw + 0.125;
  box(s, '▼ 현재: Level 0→1 진입', mx - 1.2, by + 0.05, 2.4, 0.28, { round: 0.14, fill: C.amber, color: C.white, bold: true, fontSize: 11, margin: 0 });
  T(s, lines([
    `오차율(WAPE): 실제 출고 합 대비 절대 오차 합의 비율, 낮을수록 좋음 · 개선은 ${topContribText}에 몰려 있음`,
    '결품 방어율: 목표재고가 실제 리드타임 수요를 덮은 SKU 비율(실제 결품 기록이 아닌 대리 지표), 높을수록 좋음',
  ]), { x: MX, y: 6.5, w: CW, h: 0.48, fontSize: 10, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 3. 문제
(function problem() {
  const s = slide({
    title: '안전재고일수가 고정값이라 수요가 바뀌어도 그대로입니다',
    sub: `지금 방식의 한계 세 가지 · 기존 안전재고일수가 있는 ${D.days.old_safety_n} SKU`,
    tagText: '문제 정의', tagKind: 'light',
  });
  const lw = 6.3, y = 1.72, h = 4.3;
  RR(s, MX, y, lw, h, C.light);
  T(s, '기존 안전재고일수별 SKU 수', { x: MX + 0.3, y: y + 0.15, w: lw - 0.6, h: 0.36, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  T(s, `품목마다 계산하지 않고 ${oldLevels.length}가지 값 중 하나를 정해 씁니다`, { x: MX + 0.3, y: y + 0.5, w: lw - 0.6, h: 0.3, fontSize: 11.5, color: C.gray, valign: 'middle' });
  const vals = oldLevels.map((k) => D.days.old_safety_levels[k]);
  const mxv = Math.max(...vals) * 1.15;
  const top = y + 1.2, base = y + h - 0.65, slot = (lw - 0.8) / oldLevels.length, bw = Math.min(0.9, slot * 0.55);
  line(s, MX + 0.4, base, lw - 0.8, 0, C.mute);
  oldLevels.forEach((k, i) => {
    const v = D.days.old_safety_levels[k], bh = (v / mxv) * (base - top);
    const bx = MX + 0.4 + i * slot + (slot - bw) / 2;
    R(s, bx, base - bh, bw, bh, C.mute);
    T(s, `${v} SKU`, { x: MX + 0.4 + i * slot, y: base - bh - 0.34, w: slot, h: 0.3, fontSize: 13, bold: true, color: C.navy, align: 'center', valign: 'bottom' });
    T(s, `${k}일`, { x: MX + 0.4 + i * slot, y: base + 0.08, w: slot, h: 0.36, fontSize: 15, bold: true, color: C.slate, align: 'center', valign: 'middle' });
  });
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  const cards = [
    ['calendar', `고정 일수 ${oldLevels.length}종`, '수요 크기·흔들림·리드타임과 관계없이 정해 둔 일수를 그대로 씁니다'],
    ['gear', `근거 없는 가중치 ${WRANGE}`, '그룹마다 실적·예측 비율을 다르게 쓰지만, 그 값이 최선인지 확인한 적이 없습니다'],
    ['check', '사후 채점 없음', '계산한 값이 이후 실제 출고를 맞혔는지 확인하지 않습니다'],
  ];
  const ch = (h - 0.3) / 3;
  cards.forEach(([ic, a, b], i) => {
    const cy = y + i * (ch + 0.15);
    RR(s, rx, cy, rw, ch, C.white, { line: { color: C.line, width: 1 } });
    icon(s, ic, rx + 0.25, cy + (ch - 0.7) / 2, 0.7, C.amber);
    T(s, a, { x: rx + 1.15, y: cy + 0.2, w: rw - 1.35, h: 0.42, fontSize: 17, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: rx + 1.15, y: cy + 0.66, w: rw - 1.35, h: ch - 0.8, fontSize: 12.5, color: C.text });
  });
  box(s, [
    { text: '그래서 → AI(AutoML) 탐색  ', options: { bold: true, color: '8FC9C9' } },
    { text: '가중치를 사람이 정하지 않고 AI 탐색으로 데이터에서 고르고, 고를 때 쓰지 않은 기간으로 채점합니다', options: { color: C.white } },
  ], MX, y + h + 0.2, CW, 0.65, { fill: C.navy, fontSize: 13.5, bold: true, align: 'left', margin: 14, round: 0.06 });
})();

// ======================================================================= 4. 목표·데이터
(function goal() {
  const s = slide({
    title: '질문: AutoML이 고른 가중치가 처음 보는 기간에도 더 맞는가',
    sub: '연구 질문 두 가지와, 개발·검증에 쓴 테스트 데이터셋',
    tagText: '목표', tagKind: 'light',
  });
  const lw = 5.6, y = 1.72, h = 5.05;
  RR(s, MX, y, lw, h, C.navy);
  icon(s, 'target', MX + 0.3, y + 0.28, 0.7, C.teal);
  T(s, '연구 질문', { x: MX + 1.15, y: y + 0.33, w: lw - 1.4, h: 0.6, fontSize: 18, bold: true, color: C.white, valign: 'middle' });
  const qs = [
    ['AutoML이 고르는 최적 w는?', `AutoML(HPO: 그리드 탐색 + 시계열 분할 검증 + 과적합 방지)이 그룹마다 후보 ${CANDS.length}개를 모두 채점해 고른 실적·예측 비율 w는 얼마인가`],
    ['AI가 고른 w가 새 기간에도 맞나?', `고를 때 쓰지 않은 기간(${HOLD})에도 현재 w보다 수요 오차가 줄어드는가`],
  ];
  qs.forEach(([a, b], i) => {
    const qy = y + 1.35 + i * 1.75;
    num(s, `Q${i + 1}`, MX + 0.3, qy, 0.62, C.teal, C.white, 14);
    T(s, a, { x: MX + 1.15, y: qy - 0.02, w: lw - 1.4, h: 0.5, fontSize: 17, bold: true, color: C.white, valign: 'middle' });
    T(s, b, { x: MX + 1.15, y: qy + 0.5, w: lw - 1.4, h: 1.0, fontSize: 13, color: 'D5DEE7' });
  });
  const rx = MX + lw + 0.35, rw = W - MX - rx;
  label(s, '테스트 데이터셋 (비식별 처리)', rx, y, rw);
  const k = (t) => cellOpt(t, { bold: true, color: C.navy });
  const rows = [
    ['항목', '값'],
    [k('품목'), `${D.sku_total} SKU`],
    [k('그룹'), `${groupsOrdered.length}개 (그룹 A~${String.fromCharCode(64 + groupsOrdered.length)})`],
    [k('출고 실적'), `${HIST} (${HS.n}주, 주 단위) · 계산 입력`],
    [k('예측'), `${FW}주 (${wr(`W${SS_.a}~W${FCST_END}`)}) · 계산 입력`],
    [k('선택 구간'), cellOpt(`${SEL} (${SS_.n}주) · 가중치를 고름`, { bold: true, color: C.amber })],
    [k('검증 구간'), cellOpt(`${HOLD} (${HO.n}주) · 선택에 쓰지 않음`, { bold: true, color: C.teal })],
    [k('비식별'), '품목·그룹 코드 치환, 수량 비공개'],
  ];
  table(s, rows, { x: rx, y: y + 0.45, w: rw, colW: [1.6, rw - 1.6], rowH: 0.5, fontSize: 12.5, headSize: 12.5 });
  T(s, '이 발표의 숫자는 비율(%), 일수, 배수, SKU 수, 주차 번호만 씁니다.',
    { x: rx, y: y + 0.45 + 0.5 * rows.length + 0.15, w: rw, h: 0.4, fontSize: 11.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 5. 시스템 구성
(function system() {
  const s = slide({
    title: 'MLOps 파이프라인 중 검사부터 리포트까지 구현했습니다',
    sub: '한 줄 흐름 · 청록 = 구현, 회색 점선 = 계획 · 구현 범위는 명령 한 번으로 재현됩니다',
    tagText: '시스템', tagKind: 'teal',
  });
  T(s, '입력: 출고 실적 · 수요 예측 · 품목 기준정보(리드타임 등) · 재고', { x: MX, y: 1.6, w: CW, h: 0.32, fontSize: 12.5, bold: true, color: C.slate, valign: 'middle' });
  const steps = [
    ['db', '데이터 검사', '빈 값·이상값에 점검메모', true],
    ['chart', '산출', '안전재고일수와 재고등급 계산', true],
    ['search', 'AutoML 탐색', `w 후보 ${CANDS.length}개를 그룹별로 채점`, true],
    ['calendar', '시계열 검증', '선택에 안 쓴 구간으로 확인', true],
    ['doc', '리포트', '결과표·발표 자료 자동 생성', true],
    ['check', '승인·배포', '사람이 승인한 뒤 값 교체', false],
    ['target', '모니터링', '오차율·방어율 추적, 경보', false],
    ['gear', '재학습', '매주 가중치 다시 고르기', false],
  ];
  const n = steps.length, gap = 0.2, bw = (CW - gap * (n - 1)) / n, y = 2.45, bh = 2.25;
  const xs = steps.map((_, i) => MX + i * (bw + gap));
  steps.forEach(([ic, a, b, done], i) => {
    const x = xs[i];
    if (done) RR(s, x, y, bw, bh, C.tealL, { line: { color: C.teal, width: 1.25 } });
    else box(s, '', x, y, bw, bh, { fill: C.white, dash: true, lineColor: C.mute, round: 0.08 });
    icon(s, ic, x + (bw - 0.66) / 2, y + 0.22, 0.66, done ? C.teal : C.mute);
    T(s, a, { x: x + 0.05, y: y + 0.98, w: bw - 0.1, h: 0.4, fontSize: 12.5, bold: true, color: done ? C.navy : C.gray, align: 'center', valign: 'middle' });
    T(s, b, { x: x + 0.1, y: y + 1.4, w: bw - 0.2, h: 0.75, fontSize: 11, color: done ? C.text : C.gray, align: 'center' });
    if (i < n - 1) arrow(s, x + bw + 0.02, y + bh / 2, gap - 0.04, 0, done && steps[i + 1][3] ? C.teal : C.mute, { width: 1.5 });
  });
  // 구현과 계획의 경계: 현재 여기
  const k = steps.findIndex((st) => !st[3]);
  const bx = xs[k] - gap / 2;
  line(s, bx, y - 0.12, 0, bh + 0.24, C.red, { width: 2, dashType: 'dash' });
  box(s, '현재 여기 ▼', bx - 0.7, y - 0.45, 1.4, 0.3, { round: 0.15, fill: C.red, color: C.white, bold: true, fontSize: 11.5, margin: 0 });
  // 되먹임: 재학습 → AutoML 탐색
  const fy = y + bh + 0.28, from = xs[n - 1] + bw / 2, to = xs[2] + bw / 2;
  line(s, from, y + bh, 0, fy - y - bh, C.mute, { width: 1.5, dashType: 'dash' });
  line(s, to, fy, from - to, 0, C.mute, { width: 1.5, dashType: 'dash' });
  arrow(s, to, fy, 0, -(fy - y - bh) + 0.02, C.mute, { width: 1.5 });
  box(s, '되먹임(계획): 다음 주 데이터로 AutoML 탐색을 다시 돌림', to + 1.2, fy - 0.15, 4.6, 0.3, { fill: C.white, lineColor: C.white, color: C.gray, fontSize: 11, bold: true, margin: 0 });
  // 구현 / 계획 범위
  const ry = fy + 0.38;
  box(s, [
    { text: '구현  ', options: { bold: true, color: '8FC9C9' } },
    { text: `데이터 검사 · 산출 · AutoML 탐색 · 시계열 검증 · 리포트 + 자동 테스트 ${int(TESTS)}개 · 버전 관리(git) · 검증 구간 실제값은 계산 입력에 쓰지 않음`, options: { color: C.white } },
  ], MX, ry, CW, 0.62, { fill: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
  box(s, [
    { text: '계획  ', options: { bold: true, color: C.slate } },
    { text: '승인 후 배포 · 성능 모니터링(기준을 넘으면 경보) · 주간 재학습', options: { color: C.gray } },
  ], MX, ry + 0.72, CW, 0.5, { fill: C.white, dash: true, lineColor: C.mute, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 6. 산출 방법
(function method() {
  const G = D.grades;
  const s = slide({
    title: '가중평균 수요로 품목별 안전재고일수를 계산합니다',
    sub: '일반식 두 개와 결과 값 · 세부 산식과 등급별 계수 값은 생략합니다',
    tagText: '방법', tagKind: 'teal',
  });
  const fw = 4.45, gap = 0.35, fh = 2.6, y = 1.72;
  const formulas = [
    ['① 가중평균 수요', ['가중평균 수요', '= w·실적 + (1−w)·예측'], ['w: 실적 가중치 (그룹마다 하나, AutoML이 고름)', `실적: 과거 출고 주평균 · 예측: 앞 ${FW}주 예측 평균`]],
    ['② 안전재고', ['안전재고', '= Z·σ·√LT'], ['Z: 서비스 수준 계수 (재고등급별)', 'σ: 수요 편차 · LT: 리드타임']],
  ];
  formulas.forEach(([h, f, notes], i) => {
    const x = MX + i * (fw + gap);
    RR(s, x, y, fw, fh, C.light);
    T(s, h, { x: x + 0.3, y: y + 0.18, w: fw - 0.6, h: 0.36, fontSize: 14, bold: true, color: C.teal, valign: 'middle' });
    box(s, [{ text: f[0], options: { fontSize: 13, color: C.slate, breakLine: true } }, { text: f[1], options: { fontSize: 21, color: C.navy } }],
      x + 0.3, y + 0.62, fw - 0.6, 1.05, { fill: C.white, lineColor: C.line, bold: true, round: 0.06 });
    T(s, lines(notes), { x: x + 0.3, y: y + 1.8, w: fw - 0.6, h: 0.7, fontSize: 11.5, color: C.slate, paraSpaceAfter: 3 });
    arrow(s, x + fw + 0.04, y + fh / 2, gap - 0.08, 0, i ? C.teal : C.slate, { width: 2 });
  });
  const ox = MX + 2 * (fw + gap), ow = W - MX - ox;
  RR(s, ox, y, ow, fh, C.teal);
  T(s, '결과 값', { x: ox + 0.15, y: y + 0.18, w: ow - 0.3, h: 0.36, fontSize: 14, bold: true, color: 'E6F4F4', valign: 'middle' });
  T(s, [{ text: '안전재고일수', options: { fontSize: 20, breakLine: true } }, { text: '= 안전재고', options: { fontSize: 15, breakLine: true } }, { text: '÷ 일평균 수요', options: { fontSize: 15 } }],
    { x: ox + 0.15, y: y + 0.62, w: ow - 0.3, h: 1.05, bold: true, color: C.white, valign: 'middle', paraSpaceAfter: 4 });
  T(s, '품목마다 다른 일수가 나옵니다 (고정값 대신)', { x: ox + 0.15, y: y + 1.8, w: ow - 0.3, h: 0.6, fontSize: 11.5, color: 'E6F4F4' });
  // 등급 카드
  const gy = y + fh + 0.35, gh = 6.77 - gy;
  RR(s, MX, gy, CW, gh, C.white, { line: { color: C.line, width: 1 } });
  icon(s, 'table', MX + 0.3, gy + 0.3, 0.7, C.navy);
  T(s, '재고등급(A·B·C)별 서비스 수준 Z 적용', { x: MX + 1.2, y: gy + 0.25, w: 6.4, h: 0.42, fontSize: 16, bold: true, color: C.navy, valign: 'middle' });
  T(s, `규칙 기반 등급이 기존 등급 ${G.match}/${G.total} 일치`, { x: MX + 1.2, y: gy + 0.7, w: 6.4, h: 0.4, fontSize: 14, bold: true, color: C.teal, valign: 'middle' });
  T(s, '단가·회전율·리드타임·품질 이슈를 종합해 등급을 매기고, 등급이 높을수록 더 높은 서비스 수준을 씁니다',
    { x: MX + 1.2, y: gy + 1.15, w: 6.4, h: 0.65, fontSize: 12, color: C.text });
  const bx0 = MX + 8.0, bwid = CW - 8.3;
  T(s, `계산 재고등급 분포 (${G.total} SKU)`, { x: bx0, y: gy + 0.3, w: bwid, h: 0.32, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
  const tot = G.dist.A + G.dist.B + G.dist.C; let bx = bx0;
  [['A', C.teal], ['B', C.slate], ['C', C.mute]].forEach(([k, col]) => {
    const bw = (G.dist[k] / tot) * bwid;
    box(s, `${k} ${G.dist[k]}`, bx, gy + 0.75, bw, 0.6, { fill: col, color: C.white, bold: true, fontSize: 13, margin: 0 });
    bx += bw;
  });
  T(s, '숫자는 SKU 수', { x: bx0, y: gy + 1.42, w: bwid, h: 0.28, fontSize: 10.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 7. AI 기법
(function aiTech() {
  const s = slide({
    title: '머신러닝의 모델 선택 방식으로 가중치를 고릅니다',
    sub: '핵심 기법 세 가지와 받쳐 주는 기법 세 가지 · 각 기법이 왜 필요한가',
    tagText: 'AI 기법', tagKind: 'teal',
  });
  const big = [
    ['search', 'Grid Search', '하이퍼파라미터 탐색',
      `실적 가중치 w를 ${wfmt(CANDS[0])}~${wfmt(CANDS[CANDS.length - 1])} 후보 ${CANDS.length}개로 바꿔 가며 그룹 ${NG}개를 모두 채점합니다`,
      '사람이 정해 둔 값 대신, 데이터가 가장 잘 맞힌 값을 고르려고', `탐색 공간 ${man(SPACE)} 조합`],
    ['calendar', 'Time-series Holdout', '시계열 분할 검증',
      `고르는 구간(${SEL})과 확인 구간(${HOLD})을 시간 순서로 나눠, 확인 구간은 선택에 쓰지 않습니다`,
      '미래 정보가 선택에 섞이면 성능이 부풀려지므로', '확인 구간은 선택에 미사용'],
    ['lock', 'Regularization · Pooling', '과적합 방지',
      `SKU ${MINSKU}개 미만 그룹은 그룹 안의 최솟값 대신 전체를 합쳐 고른 w를 씁니다`,
      'SKU가 적은 그룹의 우연한 오차에 w가 끌려가지 않게', `작은 그룹 ${SMALL.length}개에 적용`],
  ];
  const cw = (CW - 0.4) / 3, ch = 2.95, y = 1.62;
  big.forEach(([ic, en, ko, body, why, ev], i) => {
    const x = MX + i * (cw + 0.2);
    RR(s, x, y, cw, ch, C.light);
    icon(s, ic, x + 0.22, y + 0.22, 0.66, C.teal);
    T(s, en, { x: x + 1.02, y: y + 0.2, w: cw - 1.2, h: 0.28, fontSize: 11, color: C.slate, valign: 'middle' });
    T(s, ko, { x: x + 1.02, y: y + 0.48, w: cw - 1.2, h: 0.42, fontSize: 16.5, bold: true, color: C.navy, valign: 'middle' });
    T(s, body, { x: x + 0.22, y: y + 1.05, w: cw - 0.44, h: 0.8, fontSize: 12, color: C.text });
    box(s, [{ text: '왜  ', options: { bold: true, color: C.teal } }, { text: why, options: { color: C.navy, bold: true } }],
      x + 0.22, y + 1.9, cw - 0.44, 0.55, { fill: C.white, lineColor: C.white, fontSize: 11.5, align: 'left', margin: 6, round: 0.06 });
    T(s, ev, { x: x + 0.22, y: y + 2.52, w: cw - 0.44, h: 0.32, fontSize: 12, bold: true, color: C.teal, valign: 'middle' });
  });
  const small = [
    ['chart', '통계적 안전재고 모형', `수요 편차·리드타임·서비스 수준으로 품목별 안전재고를 계산 · ${D.sku_total} SKU`],
    ['flow', 'MLOps 파이프라인', `검사부터 리포트까지 명령 한 번으로 재현 · 자동 테스트 ${int(TESTS)}개`],
    ['gear', '생성형 AI 개발 에이전트', '설계·코드·테스트·독립 검토는 AI 에이전트가, 판단과 승인은 사람이 맡음'],
  ];
  const sy = y + ch + 0.17, sh = 1.2;
  small.forEach(([ic, a, b], i) => {
    const x = MX + i * (cw + 0.2);
    RR(s, x, sy, cw, sh, C.white, { line: { color: C.line, width: 1 } });
    icon(s, ic, x + 0.2, sy + 0.2, 0.5, C.navy);
    T(s, a, { x: x + 0.85, y: sy + 0.18, w: cw - 1.0, h: 0.36, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: x + 0.85, y: sy + 0.55, w: cw - 1.0, h: 0.6, fontSize: 10.5, color: C.text });
  });
  box(s, [
    { text: '핵심  ', options: { bold: true, color: '8FC9C9' } },
    { text: '학습된 예측 모델이 아니라, 머신러닝의 모델 선택(하이퍼파라미터 탐색) 방식을 가중치 w에 적용했습니다. 공식은 가중평균 그대로라 결과를 설명할 수 있습니다.', options: { color: C.white } },
  ], MX, sy + sh + 0.15, CW, 0.56, { fill: C.navy, fontSize: 12.5, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 8. 탐색 설계
(function wsMethod() {
  const s = slide({
    title: `탐색 설계: 앞 ${SS_.n}주로 고르고 뒤 ${HO.n}주로 확인합니다`,
    sub: `계산 시점의 값은 그대로 두고, 그 뒤 실제 출고를 선택 구간(${SEL})과 검증 구간(${HOLD})으로 나눴습니다`,
    tagText: '탐색 설계', tagKind: 'light',
  });
  const LAST = Math.max(FCST_END, HO.b);
  const x0 = MX + 0.2, x1 = W - MX - 0.2, wk = (x1 - x0) / LAST, X = (w) => x0 + (w - 1) * wk;
  const y = 2.05;
  T(s, '주차', { x: MX - 0.05, y: y - 0.35, w: 1, h: 0.28, fontSize: 10.5, color: C.gray });
  box(s, `출고 실적 ${HIST} (계산 입력)`, X(HS.a), y, X(HS.b + 1) - X(HS.a) - 0.03, 0.5, { fill: C.mute, color: C.white, bold: true, fontSize: 12, margin: 0 });
  if (SS_.a > HS.b + 1) box(s, '미사용', X(HS.b + 1), y, X(SS_.a) - X(HS.b + 1) - 0.03, 0.5, { fill: 'E6EBF0', color: C.gray, fontSize: 10, margin: 0 });
  box(s, lines(['선택 구간', SEL]), X(SS_.a), y, X(SS_.b + 1) - X(SS_.a) - 0.03, 0.5, { fill: C.amber, color: C.white, bold: true, fontSize: 10.5, margin: 0 });
  box(s, lines(['검증 구간', HOLD]), X(HO.a), y, X(HO.b + 1) - X(HO.a) - 0.03, 0.5, { fill: C.teal, color: C.white, bold: true, fontSize: 10.5, margin: 0 });
  box(s, `예측(${FW}주) ${wr(`W${SS_.a}~W${FCST_END}`)} (계산 입력)`, X(SS_.a), y + 0.6, X(FCST_END + 1) - X(SS_.a), 0.36, { fill: 'EFE3CB', color: '8A6420', bold: true, fontSize: 10.5, margin: 0 });
  line(s, X(SS_.a) - 0.015, y - 0.15, 0, 1.25, C.red, { width: 1.5, dashType: 'dash' });
  T(s, '계산 시점', { x: X(SS_.a) - 1.75, y: y + 0.62, w: 1.7, h: 0.32, fontSize: 10.5, bold: true, color: C.red, align: 'right', valign: 'middle' });
  [...new Set([HS.a, 10, 20, HS.b, SS_.a, HO.a, HO.b, FCST_END])].forEach((w) => T(s, `W${String(w).padStart(2, '0')}`, { x: X(w) - 0.3 + wk / 2, y: y + 1.0, w: 0.6, h: 0.24, fontSize: 10, color: C.gray, align: 'center' }));
  const rules = [
    ['후보', `실적 가중치 w ${wfmt(CANDS[0])}~${wfmt(CANDS[CANDS.length - 1])}을 ${wfmt(CSTEP)} 간격으로 ${CANDS.length}개, 예측 가중치는 1 − w`],
    ['단위', '품목그룹마다 w 하나 (그룹 안 모든 품목에 같은 값)'],
    ['고르는 기준', '선택 구간 가중평균 오차율이 가장 낮은 w (같으면 현재 w 쪽)'],
    ['작은 그룹', `선택 구간 SKU ${MINSKU}개 미만이면 전체를 합쳐 고른 w (과적합 방지)`],
  ];
  const cw = 3.55, ch = 1.12;
  rules.forEach(([a, b], i) => {
    const x = MX + (i % 2) * (cw + 0.2), yy = 3.55 + Math.floor(i / 2) * (ch + 0.18);
    RR(s, x, yy, cw, ch, C.light);
    num(s, i + 1, x + 0.18, yy + 0.18, 0.42, C.navy, C.white, 13);
    T(s, a, { x: x + 0.72, y: yy + 0.14, w: cw - 0.9, h: 0.36, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: x + 0.72, y: yy + 0.5, w: cw - 0.9, h: 0.58, fontSize: 11, color: C.text });
  });
  box(s, `왜 앞뒤로 나눴나: 과거 시점의 예측 기록이 없어 계산 시점 이후 출고만 쓸 수 있습니다. W${SS_.b}에 다시 고르는 상황을 흉내 낸 것이라 실제 운영보다 낙관적일 수 있습니다.`,
    MX, 6.12, 2 * cw + 0.2, 0.66, { fill: C.amberL, fontSize: 11, align: 'left', margin: 8, round: 0.06 });
  // 오른쪽: 분해 가능성 설명
  const tx = MX + 2 * cw + 0.5, tw = W - MX - tx, ty = 3.55, th = 6.78 - ty;
  RR(s, tx, ty, tw, th, C.navy);
  T(s, `왜 ${CANDS.length * NG}번 계산으로 ${man(SPACE)} 조합의 최적을 얻나`, { x: tx + 0.25, y: ty + 0.1, w: tw - 0.5, h: 0.62, fontSize: 13.5, bold: true, color: C.white, valign: 'middle' });
  const cx = tx + 0.25, cwid = (tw - 0.75) / 2;
  box(s, [{ text: '모든 조합', options: { fontSize: 11, color: C.sky, breakLine: true } }, { text: `${CANDS.length}의 ${NG}제곱`, options: { fontSize: 13, bold: true, color: C.white, breakLine: true } }, { text: `= ${man(SPACE)}`, options: { fontSize: 15, bold: true, color: C.white } }],
    cx, ty + 0.82, cwid, 1.0, { fill: C.navy2, round: 0.06, margin: 3 });
  box(s, [{ text: '실제 계산', options: { fontSize: 11, color: '8FC9C9', breakLine: true } }, { text: `${CANDS.length} × ${NG}`, options: { fontSize: 13, bold: true, color: C.white, breakLine: true } }, { text: `= ${CANDS.length * NG}번`, options: { fontSize: 15, bold: true, color: C.white } }],
    cx + cwid + 0.25, ty + 0.82, cwid, 1.0, { fill: C.teal, round: 0.06, margin: 3 });
  T(s, lines([
    '전체 오차(절대 오차의 합)는 그룹별 오차의 합이고, 각 그룹 오차는 그 그룹의 w에만 달려 있습니다.',
    '그래서 그룹마다 따로 최적 w를 고르면 모든 조합을 비교한 답과 같습니다.',
    `단, SKU ${MINSKU}개 미만 그룹은 과적합 방지를 위해 이 값 대신 전체 합산 w를 일부러 씁니다.`,
  ]), { x: tx + 0.25, y: ty + 1.88, w: tw - 0.5, h: th - 2.0, fontSize: 10.5, color: C.white, paraSpaceAfter: 3, valign: 'top' });
})();

// ======================================================================= 9. 후보 전체 채점(히트맵)
(function wsHeatmap() {
  const s = slide({
    title: `후보 ${CANDS.length}개를 모두 채점해 그룹별 최적 w를 골랐습니다`,
    sub: `선택 구간 ${SEL}의 가중평균 수요 오차율 · 행 안에서 오차가 작을수록 진한 청록 · ★ 선택 · ○ 현재`,
    tagText: '탐색 과정', tagKind: 'teal',
  });
  const rowsDef = [{ key: '전체', label: '전체 합산', choice: null }, ...choicesSorted.map((c) => ({ key: c.group, label: ga(c.group), choice: c }))];
  const pooled = WS.candidates.filter((c) => c.scope === '전체' && isNum(c.wape));
  const pooledBest = pooled.reduce((a, b) => (b.wape < a.wape ? b : a)).w;
  const lx = MX, nx = MX + 1.2, x0 = MX + 1.9, rcw = 1.6;
  const cw = (CW - 1.9 - rcw - 0.1) / CANDS.length;
  const rx = x0 + cw * CANDS.length + 0.1;
  T(s, '실적 가중치 w  (예측 가중치 = 1 − w)', { x: x0, y: 1.62, w: cw * CANDS.length, h: 0.3, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
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
    const how = r.choice ? (METHOD[r.choice.method] || r.choice.method) + (r.choice.method === '공통최적' ? ' 값 사용' : '') : '전체 합산 최적';
    T(s, [{ text: r.choice ? `${wfmt(current)} → ${wfmt(chosen)}` : wfmt(chosen), options: { bold: true, color: r.choice && chosen !== current ? C.teal : C.navy, breakLine: true, fontSize: 12 } },
      { text: how, options: { color: C.gray, fontSize: 10 } }],
    { x: rx, y: y - 0.02, w: rcw, h: rh + 0.04, align: 'center', valign: 'middle' });
  });
  const ny = y0 + rowsDef.length * rs + 0.04;
  T(s, `* SKU ${MINSKU}개 미만 그룹은 그룹 안의 최솟값 대신 전체 합산 최적 w(${wfmt(pooledBest)})를 씁니다 (과적합 방지) · 색은 행 안에서만 비교 · SKU = 선택 구간 채점 수`,
    { x: MX, y: ny, w: CW, h: 0.3, fontSize: 10.5, color: C.gray, valign: 'middle' });
  // 읽는 법: 끝값(0.0 예측만 / 1.0 실적만)을 고른 그룹
  const ends = choicesSorted.filter((c) => c.method === '그룹최적' && (c.chosen === 0 || c.chosen === 1));
  const toHist = ends.filter((c) => c.chosen === 1).map((c) => c.group);
  const toFcst = ends.filter((c) => c.chosen === 0).map((c) => c.group);
  const parts = [];
  if (toHist.length) parts.push(`${aliasLetters(toHist)}는 ${wfmt(1)}(실적만)`);
  if (toFcst.length) parts.push(`${aliasLetters(toFcst)}는 ${wfmt(0)}(예측만)`);
  box(s, [
    { text: '읽는 법  ', options: { bold: true, color: '8FC9C9' } },
    // 그룹 최적을 쓴 그룹이 모두 끝값을 골랐을 때만 '모두 한쪽만'이라고 쓴다
    { text: parts.length && ends.length === choicesSorted.filter((c) => c.method === '그룹최적').length
      ? `SKU ${MINSKU}개 이상 그룹은 섞는 것보다 한쪽만 쓰는 편이 선택 구간에서 더 맞았습니다: ${parts.join(', ')}. 단, ${SS_.n}주로 고른 값이라 다음 장의 검증 구간에서 다시 확인합니다.`
      : `그룹마다 오차가 가장 낮은 w가 다릅니다. ${SS_.n}주로 고른 값이라 다음 장의 검증 구간에서 다시 확인합니다.`, options: { color: C.white } },
  ], MX, ny + 0.4, CW, 0.66, { fill: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 10. 결과 1
(function result1() {
  const s = slide({
    title: `검증 구간 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}`,
    sub: `고를 때 쓰지 않은 ${HOLD}(${HO.n}주) · 전체 ${NHOLD} SKU · 현재 w → AI 탐색 w`,
    tagText: '결과', tagKind: 'dark',
  });
  const kp = [
    ['가중평균 수요 오차율', 'wape', true, (v) => pct(v), '낮을수록 좋음'],
    ['결품 방어율', 'cover', false, (v) => pct(v), '높을수록 좋음 · 대리 지표'],
    ['실제 기준 재고일수 중앙값', 'days_median', null, (v) => f1(v) + '일', '목표재고 ÷ 실제 출고'],
  ];
  const cw = (CW - 0.4) / 3, y = 1.72, h = 2.2;
  kp.forEach(([l, key, lowGood, f, note], i) => {
    const x = MX + i * (cw + 0.2);
    const a = wsCur[key], z = wsNew[key];
    let col = C.navy;
    if (lowGood !== null && isNum(a) && isNum(z) && a !== z) col = ((z < a) === lowGood) ? C.teal : C.red;
    const bad = col === C.red;
    RR(s, x, y, cw, h, bad ? C.redL : i === 0 ? C.tealL : C.light);
    T(s, l, { x: x + 0.3, y: y + 0.18, w: cw - 0.6, h: 0.38, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
    T(s, [{ text: f(a), options: { color: C.slate } }, { text: ' → ', options: { color: C.slate } }, { text: f(z), options: { color: col } }],
      { x: x + 0.3, y: y + 0.62, w: cw - 0.6, h: 0.8, fontSize: 30, bold: true, valign: 'middle' });
    const diff = key === 'days_median'
      ? `${z - a >= 0 ? '+' : '−'}${f1(Math.abs(z - a))}일`
      : `${z >= a ? '+' : '−'}${pp(Math.max(a, z), Math.min(a, z))}`;
    T(s, diff, { x: x + 0.3, y: y + 1.45, w: 1.5, h: 0.36, fontSize: 15, bold: true, color: col, valign: 'middle' });
    T(s, note, { x: x + 1.7, y: y + 1.45, w: cw - 2.0, h: 0.36, fontSize: 11, color: C.gray, valign: 'middle', align: 'right' });
    if (bad) box(s, '나빠짐', x + cw - 1.05, y + 0.2, 0.8, 0.3, { round: 0.15, fill: C.red, color: C.white, fontSize: 10.5, bold: true, margin: 0 });
  });
  // 왼쪽 아래: 선택 구간 전체 합산 오차율 w별 막대
  const by = y + h + 0.3, lw = 6.6, bh = 6.82 - by;
  RR(s, MX, by, lw, bh, C.light);
  const cand = WS.candidates.filter((c) => c.scope === '전체' && isNum(c.wape)).sort((a, b) => a.w - b.w);
  const mn = Math.min(...cand.map((c) => c.wape)), mxv = Math.max(...cand.map((c) => c.wape));
  const best = cand.find((c) => c.wape === mn);
  T(s, `참고: 모든 그룹에 같은 w를 쓸 때 (선택 구간 ${SEL})`, { x: MX + 0.25, y: by + 0.12, w: lw - 0.5, h: 0.32, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
  T(s, `전체 합산 오차율 ${pct(mn)} ~ ${pct(mxv)} · 가장 낮은 w = ${wfmt(best.w)} · 막대는 최솟값 근처를 확대`, { x: MX + 0.25, y: by + 0.44, w: lw - 0.5, h: 0.28, fontSize: 10.5, color: C.gray, valign: 'middle' });
  const bx = MX + 0.3, bwid = (lw - 0.6) / cand.length, base = by + bh - 0.4, lo = mn * 0.9, ph = bh - 1.45;
  line(s, bx, base, lw - 0.6, 0, C.mute);
  cand.forEach((c, i) => {
    const hh = Math.max(0.04, ((c.wape - lo) / (mxv - lo)) * ph);
    const isMin = c.wape === mn;
    R(s, bx + i * bwid + bwid * 0.18, base - hh, bwid * 0.64, hh, isMin ? C.amber : C.mute);
    T(s, pct(c.wape), { x: bx + i * bwid - 0.05, y: base - hh - 0.26, w: bwid + 0.1, h: 0.24, fontSize: 9.5, bold: isMin, color: isMin ? C.amber : C.slate, align: 'center', valign: 'bottom' });
    T(s, wfmt(c.w), { x: bx + i * bwid, y: base + 0.04, w: bwid, h: 0.26, fontSize: 10.5, color: C.text, align: 'center', valign: 'middle' });
  });
  // 오른쪽 아래: 해석
  const rx = MX + lw + 0.25, rw = W - MX - rx;
  RR(s, rx, by, rw, bh, C.navy);
  T(s, '해석', { x: rx + 0.25, y: by + 0.12, w: rw - 0.5, h: 0.36, fontSize: 14, bold: true, color: '8FC9C9', valign: 'middle' });
  const full = { c: wsRow('전체구간', '전체', '현재', '가중평균'), n: wsRow('전체구간', '전체', '탐색', '가중평균') };
  const notes = [
    `선택에 쓰지 않은 ${HO.n}주로 확인한 개선입니다. 모든 그룹에 같은 w를 쓰면 w에 따른 차이가 작습니다(${pct(mn)}~${pct(mxv)}, 왼쪽).`,
    coverDown
      ? `결품 방어율은 ${pp(wsCur.cover, wsNew.cover)} 낮아졌습니다. 가중평균 수요 편향이 ${spct(wsCur.bias)} → ${spct(wsNew.bias)}로 ${wsNew.bias < wsCur.bias ? '과소' : '과대'} 쪽으로 움직였습니다 → 서비스 수준(Z) 조정 과제입니다.`
      : `결품 방어율은 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}입니다.`,
  ];
  if (full.c && full.n && isNum(full.c.wape) && isNum(full.n.wape)) notes.push(`참고: 선택 구간을 포함한 ${VAL} 전체는 ${pct(full.c.wape)} → ${pct(full.n.wape)}`);
  T(s, notes.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < notes.length - 1, color: i === 2 ? C.sky : C.white, fontSize: i === 2 ? 10.5 : 12 } })),
    { x: rx + 0.25, y: by + 0.55, w: rw - 0.45, h: bh - 0.65, valign: 'top', paraSpaceAfter: 6 });
})();

// ======================================================================= 11. 결과 2
(function result2() {
  const [t1, t2] = topContrib;
  const s = slide({
    title: t2 ? `개선은 ${ga(t1[0])} ${pct(t1[1].share, 0)} · ${ga(t2[0])} ${pct(t2[1].share, 0)}에 몰렸습니다` : `개선은 ${ga(t1[0])}에 몰렸습니다`,
    sub: `그룹별 선택과 검증 구간 ${HOLD} 오차율 · 감소 몫 = 전체 오차 감소분 중 그 그룹의 몫`,
    tagText: '결과', tagKind: 'dark',
  });
  const lw = 7.6, y = 1.72, rh = 0.42;
  const rows = [['그룹', 'SKU', '현재 w', '선택 w', '방식', '검증 구간 오차율 (전→후)', '감소 몫']];
  const up = [], down = [];
  choicesSorted.forEach((c) => {
    const hc = wsRow('검증', c.group, '현재', '가중평균'), hn = wsRow('검증', c.group, '탐색', '가중평균');
    const hv = hc && hn && isNum(hc.wape) && isNum(hn.wape);
    const col = hv ? (hn.wape < hc.wape ? C.teal : hn.wape > hc.wape ? C.red : C.text) : C.gray;
    const sh = contrib && contrib[c.group] ? contrib[c.group].share : null;
    rows.push([cellOpt(ga(c.group), { bold: true }), String(c.n), wfmt(c.current),
      cellOpt(wfmt(c.chosen), { bold: true, color: c.chosen !== c.current ? C.teal : C.text }),
      cellOpt(METHOD[c.method] || c.method, { color: c.method === '그룹최적' ? C.navy : C.gray }),
      cellOpt(hv ? `${pct(hc.wape)} → ${pct(hn.wape)}` : '–', { bold: true, color: col }),
      cellOpt(isNum(sh) ? spct0(sh) : '–', { bold: isNum(sh) && sh >= 0.1, color: isNum(sh) && sh < 0 ? C.red : C.text })]);
    if (hv) {
      if (hn.wape < hc.wape) up.push(`${ga(c.group)} ${pct(hc.wape)}→${pct(hn.wape)}`);
      else if (hn.wape > hc.wape) down.push(`${ga(c.group)} 오차율 ${pct(hc.wape)}→${pct(hn.wape)}`);
    }
    if (hc && hn && isNum(hc.cover) && isNum(hn.cover) && hn.cover < hc.cover) down.push(`${ga(c.group)} 방어율 ${pct(hc.cover)}→${pct(hn.cover)}`);
  });
  if (coverDown) down.unshift(`전체 방어율 ${pct(wsCur.cover)}→${pct(wsNew.cover)}`);
  table(s, rows, { x: MX, y, w: lw, colW: [0.95, 0.6, 0.8, 0.8, 1.1, 2.4, 0.95], rowH: rh, fontSize: 11.5, headSize: 11, align: 'center' });
  const tEnd = y + rh * rows.length;
  T(s, `SKU = 선택 구간 채점 수 · 방식 '전체 합산' = SKU ${MINSKU}개 미만이라 전체를 합쳐 고른 w · 검증 구간 전체 ${NHOLD} SKU`,
    { x: MX, y: tEnd + 0.06, w: lw, h: 0.3, fontSize: 10.5, color: C.gray, valign: 'middle' });
  // 오른쪽: 감소 몫 가로 막대(표의 행과 높이를 맞춤)
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  T(s, '오차 감소 몫', { x: rx, y, w: rw, h: rh, fontSize: 12, bold: true, color: C.navy, valign: 'middle' });
  const shares = choicesSorted.map((c) => (contrib && contrib[c.group] ? contrib[c.group].share : null));
  const negMax = Math.max(0, ...shares.filter(isNum).map((v) => -v)), posMax = Math.max(0, ...shares.filter(isNum));
  const lab = 0.6, usable = rw - lab - 0.1;
  const scale = usable / (negMax + posMax || 1);
  const zx = rx + negMax * scale;
  line(s, zx, y + rh, 0, rh * choicesSorted.length, C.mute, { width: 1 });
  shares.forEach((v, i) => {
    if (!isNum(v)) return;
    const ry = y + rh * (i + 1) + 0.09, bhh = rh - 0.18;
    const len = Math.max(0.03, Math.abs(v) * scale);
    if (v >= 0) R(s, zx, ry, len, bhh, v >= 0.1 ? C.teal : '7FBFBF');
    else R(s, zx - len, ry, len, bhh, C.red);
    const tx = v >= 0 ? zx + len + 0.06 : zx + 0.06;
    T(s, spct0(v), { x: tx, y: ry - 0.03, w: lab, h: bhh + 0.06, fontSize: 11, bold: v >= 0.1, color: v < 0 ? C.red : C.navy, valign: 'middle' });
  });
  // 아래: 좋아진 것 / 나빠진 것
  const pair = (arr) => arr.reduce((acc, t, i) => (i % 2 ? (acc[acc.length - 1] += ' · ' + t, acc) : (acc.push(t), acc)), []);
  const by = tEnd + 0.5, bh = 6.85 - by, hw = (CW - 0.2) / 2;
  RR(s, MX, by, hw, bh, C.tealL);
  T(s, [{ text: '좋아진 것 (검증 구간 오차율)', options: { bold: true, color: C.teal, breakLine: true, fontSize: 12.5 } }, ...lines(pair(up).length ? pair(up) : ['없음'], { color: C.text })],
    { x: MX + 0.2, y: by + 0.1, w: hw - 0.4, h: bh - 0.2, fontSize: 11.5, valign: 'top', paraSpaceAfter: 3 });
  RR(s, MX + hw + 0.2, by, hw, bh, C.redL);
  T(s, [{ text: '나빠진 것', options: { bold: true, color: C.red, breakLine: true, fontSize: 12.5 } }, ...lines(pair(down).length ? pair(down) : ['없음'], { color: C.text })],
    { x: MX + hw + 0.4, y: by + 0.1, w: hw - 0.4, h: bh - 0.2, fontSize: 11.5, valign: 'top', paraSpaceAfter: 3 });
})();

// ======================================================================= 12. 사례
(function example() {
  const c = D.case, cc = D.counter_case;
  const s = slide({
    title: `사례: 품목 X의 안전재고일수 ${c.old_ss_days}일 → ${f1(c.s)}일`,
    sub: `${ga(c.group)} · 재고등급 ${c.grade} · 리드타임 ${c.lt}일 · 그룹의 현재 w 기준(AI 탐색 w 아님) · 수량 없이 일수와 비율로`,
    tagText: '사례', tagKind: 'light',
  });
  const lw = 6.0, y = 1.72;
  // 왼쪽: 일수 비교
  RR(s, MX, y, lw, 2.75, C.light);
  T(s, '안전재고일수 비교', { x: MX + 0.25, y: y + 0.12, w: lw - 0.5, h: 0.34, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  const maxD = Math.max(c.old_ss_days, c.s);
  const barX = MX + 1.55, barW = lw - 1.55 - 1.35;
  [['기존 고정', c.old_ss_days, `${c.old_ss_days}일`, C.mute, C.slate], ['산식', c.s, `${f1(c.s)}일`, C.teal, C.teal]].forEach(([l, v, t, col, tc], i) => {
    const ry = y + 0.65 + i * 0.85;
    T(s, l, { x: MX + 0.25, y: ry, w: 1.25, h: 0.6, fontSize: 13, bold: true, color: tc, valign: 'middle' });
    R(s, barX, ry + 0.12, Math.max(0.05, (v / maxD) * barW), 0.36, col);
    T(s, t, { x: barX + Math.max(0.05, (v / maxD) * barW) + 0.1, y: ry, w: 1.3, h: 0.6, fontSize: 24, bold: true, color: tc, valign: 'middle' });
  });
  T(s, `리드타임을 더한 적정 재고일수: ${c.old_days}일 → ${f1(c.u)}일`, { x: MX + 0.25, y: y + 2.3, w: lw - 0.5, h: 0.32, fontSize: 11.5, color: C.slate, valign: 'middle' });
  // 왼쪽 아래: 이 품목에 적용한 식
  const fy = y + 2.9, fh = 1.3;
  RR(s, MX, fy, lw, fh, C.white, { line: { color: C.line, width: 1 } });
  T(s, '이 품목에 적용한 식', { x: MX + 0.25, y: fy + 0.1, w: lw - 0.5, h: 0.32, fontSize: 12.5, bold: true, color: C.teal, valign: 'middle' });
  T(s, lines([
    `가중평균 수요 = ${wfmt(c.wh)}·실적 + ${wfmt(c.wf)}·예측  (그룹의 현재 w)`,
    `안전재고 = Z(등급 ${c.grade})·σ·√${c.lt}  → 일평균 수요로 나눠 ${f1(c.s)}일`,
  ]), { x: MX + 0.25, y: fy + 0.46, w: lw - 0.5, h: 0.78, fontSize: 12, color: C.navy, bold: true, paraSpaceAfter: 4 });
  // 오른쪽: 사후 검증
  const rx = MX + lw + 0.3, rw = W - MX - rx, vh = fy + fh - y;
  RR(s, rx, y, rw, vh, C.white, { line: { color: C.line, width: 1 } });
  T(s, `사후 검증 · ${VAL}(${VS.n}주) 실제 출고`, { x: rx + 0.25, y: y + 0.12, w: rw - 0.5, h: 0.34, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  const ok = (v) => (v === '방어' ? C.teal : C.red);
  const fdiff = c.g / c.actual_week - 1;
  const vr = [
    ['리드타임 수요 방어', [[`산식 ${cov(c.new_cover)}`, ok(c.new_cover)], ['  ·  ', C.gray], [`기존 ${cov(c.old_cover)}`, ok(c.old_cover)]]],
    ['실제 기준 재고일수', [[`산식 ${f1(c.new_actual_days)}일`, C.teal], ['  ·  ', C.gray], [`기존 ${f1(c.old_actual_days)}일`, C.slate]]],
    ['목표재고 비율 (산식 ÷ 기존)', [[pct(c.new_target / c.old_target, 0), C.teal]]],
    ['실제 대비 가중평균 수요 차이', [[`${fdiff >= 0 ? '+' : '−'}${pct(Math.abs(fdiff), 0)}`, C.amber, true], [fdiff >= 0 ? '(가중평균 수요가 실제보다 높음)' : '(가중평균 수요가 실제보다 낮음)', C.gray]]],
  ];
  const step = 0.72;
  vr.forEach(([a, parts], i) => {
    const ry = y + 0.6 + i * step;
    if (i) line(s, rx + 0.25, ry - 0.04, rw - 0.5, 0, C.line);
    T(s, a, { x: rx + 0.25, y: ry, w: 2.5, h: step - 0.08, fontSize: 12, color: C.slate, valign: 'middle' });
    T(s, parts.map(([t, col, br]) => ({ text: t, options: { color: col, bold: col !== C.gray, fontSize: col === C.gray ? 10.5 : 15, breakLine: !!br } })),
      { x: rx + 2.7, y: ry, w: rw - 2.95, h: step - 0.08, align: 'right', valign: 'middle' });
  });
  T(s, `선정 규칙: ${c.rule} (산식에 유리한 사례)`, { x: rx + 0.25, y: y + vh - 0.62, w: rw - 0.5, h: 0.52, fontSize: 10, color: C.gray, valign: 'bottom' });
  // 아래: 요약과 반대 사례
  const both = c.new_cover === '방어' && c.old_cover === '방어';
  const msg = both
    ? `두 기준 모두 리드타임 수요를 덮었고, 산식 목표재고는 기존의 ${pct(c.new_target / c.old_target, 0)}였습니다. 단, 가중평균 수요는 실제보다 ${pct(Math.abs(fdiff), 0)} ${fdiff >= 0 ? '높았' : '낮았'}습니다.`
    : `방어 결과: 산식 ${cov(c.new_cover)}, 기존 ${cov(c.old_cover)} · 산식 목표재고는 기존의 ${pct(c.new_target / c.old_target, 0)}입니다.`;
  box(s, msg, MX, 6.03, CW, 0.42, { fill: C.navy, color: C.white, fontSize: 12, bold: true, align: 'left', margin: 10, round: 0.06 });
  const ccBoth = cc.new_cover === '결품' && cc.old_cover === '결품';
  box(s, [
    { text: '반대 사례  ', options: { bold: true, color: C.white } },
    { text: `${ga(cc.group)}의 한 품목은 ${cc.old_ss_days}일 → ${f1(cc.s)}일이었고, 실제 출고가 가중평균 수요의 ${(cc.actual_week / cc.g).toFixed(1)}배여서 ${ccBoth ? '두 기준 모두' : '산식은'} 방어에 실패했습니다.`, options: { color: C.white } },
  ], MX, 6.52, CW, 0.42, { fill: C.red, fontSize: 12, align: 'left', margin: 10, round: 0.06 });
})();

// ======================================================================= 13. 해석과 한계
(function limits() {
  const s = slide({
    title: `검증 오차는 줄었지만, 선택 ${SS_.n}주·유의성 검정 없음의 한계가 있습니다`,
    sub: '개선 주장은 채점에 쓰지 않은 구간의 수치로만 하고, 나빠진 지표도 함께 적었습니다',
    tagText: '해석', tagKind: 'light',
  });
  const lw = 5.3, y = 1.72, h = 5.05;
  RR(s, MX, y, lw, h, C.navy);
  T(s, '말할 수 있는 것', { x: MX + 0.3, y: y + 0.14, w: lw - 0.6, h: 0.4, fontSize: 16, bold: true, color: C.white, valign: 'middle' });
  const can = [
    [`검증 오차율 개선`, `선택에 안 쓴 ${HOLD}: ${pct(wsCur.wape)} → ${pct(wsNew.wape)} (${NHOLD} SKU)`],
    ['고정값보다 짧고 품목마다 다른 일수', `중앙값 ${f1(SS.old_median)} → ${f1(SS.new_median)}일 · ${SS.shorter}/${SS.n} SKU가 짧아짐`],
    ['등급 규칙 일치', `규칙 기반 재고등급이 기존 등급 ${D.grades.match}/${D.grades.total}과 같음`],
  ];
  const ih = 1.45;
  can.forEach(([a, b], i) => {
    const iy = y + 0.75 + i * ih;
    num(s, i + 1, MX + 0.3, iy + 0.05, 0.46, C.teal, C.white, 14);
    T(s, a, { x: MX + 0.95, y: iy, w: lw - 1.2, h: 0.5, fontSize: 14.5, bold: true, color: C.white, valign: 'middle' });
    T(s, b, { x: MX + 0.95, y: iy + 0.52, w: lw - 1.2, h: 0.8, fontSize: 12, color: 'D5DEE7' });
  });
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  T(s, '아직 말할 수 없는 것 (한계)', { x: rx, y, w: rw, h: 0.4, fontSize: 16, bold: true, color: C.amber, valign: 'middle' });
  const lim = [
    ['db', '테스트 데이터셋 결과', '비식별 처리한 테스트 데이터셋에서 얻은 결과입니다. 다른 데이터에서도 같다고 보장하지 않습니다.'],
    ['calendar', `선택 구간 ${SS_.n}주`, `W${SS_.b}에 다시 고르는 상황을 흉내 낸 것이라 실제 운영보다 낙관적일 수 있습니다.`],
    ['chart', '유의성 검정 없음', `검증 ${NHOLD} SKU, 그룹은 ${minGroupN}~${maxGroupN}개. 차이가 우연인지 검정하지 않았습니다.`],
    ['target', '결품 방어는 대리 지표', lines(['목표재고가 실제 리드타임 수요를 덮었는지만 봅니다(실제 결품 기록 아님).', `방어율 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}`])],
    ['doc', '예측 스냅샷 1개', '한 시점의 예측만 썼습니다. 주마다 쌓아 같은 비교를 반복해야 합니다.'],
  ];
  const step = (h - 0.5) / lim.length;
  lim.forEach(([ic, a, b], i) => {
    const ly = y + 0.5 + i * step;
    RR(s, rx, ly, rw, step - 0.08, C.amberL);
    icon(s, ic, rx + 0.16, ly + (step - 0.08 - 0.55) / 2, 0.55, C.amber);
    T(s, a, { x: rx + 0.85, y: ly, w: 2.05, h: step - 0.08, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: rx + 2.95, y: ly, w: rw - 3.1, h: step - 0.08, fontSize: 11.5, color: C.text, valign: 'middle' });
  });
})();

// ======================================================================= 14. 향후 과제와 MLOps
(function roadmap() {
  const s = slide({
    title: 'MLOps 로드맵: 지금은 Level 0에서 1로 넘어가는 단계입니다',
    sub: 'Google MLOps 성숙도 기준으로 본 이 프로젝트의 위치와, 목표로 하는 모니터링 화면 구상',
    tagText: 'MLOps 로드맵', tagKind: 'amber',
  });
  // 왼쪽: 성숙도 Level 0/1/2
  const lw = 5.75, y0 = 1.62, ch = 1.48, cg = 0.1;
  const levels = [
    ['Level 0', '수동 실험', '스크립트로 실험·학습, 사람이 결과를 확인하고 반영', '달성', C.teal,
      'AutoML 탐색·시계열 검증을 스크립트로 실행하고 사람이 확인'],
    ['Level 1', '파이프라인 자동화·지속 학습', '파이프라인 자동 실행, 새 데이터로 재학습, 데이터·성능 검증', '진행 중', C.amber,
      `구현: 파이프라인 · 자동 테스트 ${int(TESTS)}개 · 버전 관리 · 다음: 주간 재탐색 · 모니터링`],
    ['Level 2', 'CI/CD 자동화', '파이프라인 자체를 CI/CD로 시험·배포, 자동 배포와 되돌리기', '계획', C.mute,
      '승인 후 자동 배포, 경보가 뜨면 재학습을 자동으로 시작'],
  ];
  levels.forEach(([a, b, core, st, col, mine], i) => {
    const y = y0 + i * (ch + cg), now = st === '진행 중';
    RR(s, MX, y, lw, ch, now ? C.amberL : C.light, now ? { line: { color: C.amber, width: 2.25 } } : {});
    box(s, a, MX + 0.18, y + 0.15, 0.95, 0.32, { round: 0.16, fill: i === 2 ? C.slate : C.navy, color: C.white, bold: true, fontSize: 11.5, margin: 0 });
    T(s, b, { x: MX + 1.25, y: y + 0.12, w: lw - 2.6, h: 0.38, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
    box(s, now ? '진행 중 · 현재' : st, MX + lw - 1.38, y + 0.15, 1.2, 0.32, { round: 0.16, fill: col, color: C.white, bold: true, fontSize: 11, margin: 0 });
    T(s, [{ text: '핵심  ', options: { bold: true, color: C.slate } }, { text: core, options: { color: C.text } }],
      { x: MX + 0.2, y: y + 0.55, w: lw - 0.4, h: 0.4, fontSize: 10.5, valign: 'middle' });
    T(s, [{ text: '이 프로젝트  ', options: { bold: true, color: C.teal } }, { text: mine, options: { color: C.navy } }],
      { x: MX + 0.2, y: y + 0.95, w: lw - 0.4, h: 0.46, fontSize: 10.5, valign: 'middle' });
  });
  // 오른쪽: 목표 대시보드 목업(도형). 숫자는 data.json 값을 예시로 쓴다
  const wx = MX + lw + 0.3, ww = W - MX - wx;
  box(s, '목표 화면(구상) · 예시 값', wx, y0, ww, 0.46, { fill: C.amber, color: C.white, bold: true, fontSize: 16, round: 0.06 });
  const wy = y0 + 0.56, wh = y0 + 3 * ch + 2 * cg - wy;
  RR(s, wx, wy, ww, wh, C.white, { line: { color: C.mute, width: 1.25 } });
  R(s, wx + 0.02, wy + 0.02, ww - 0.04, 0.32, 'E6EBF0');
  [C.red, C.amber, C.teal].forEach((col, i) => s.addShape(pres.shapes.OVAL, { x: wx + 0.14 + i * 0.2, y: wy + 0.1, w: 0.13, h: 0.13, fill: { color: col }, line: { color: col, width: 0 } }));
  T(s, '주간 가중치 점검 화면 (구상)', { x: wx + 0.85, y: wy + 0.02, w: ww - 1.0, h: 0.32, fontSize: 10.5, bold: true, color: C.slate, valign: 'middle' });
  // 주간 검증 오차율 막대: 탐색 전·후 값(예시) + 앞으로 쌓일 주(빈 막대)
  const cx = wx + 0.15, cwid = 3.25, cy = wy + 0.45, chh = 2.55;
  T(s, '주간 검증 오차율', { x: cx, y: cy, w: cwid, h: 0.28, fontSize: 11, bold: true, color: C.navy, valign: 'middle' });
  const vmax = wsCur.wape * 1.3, base = cy + chh - 0.3, top = cy + 0.55, ph = base - top;
  const bars5 = [['탐색 전', wsCur.wape, C.mute], ['탐색 후', wsNew.wape, C.teal], ['다음 주', null], ['+2주', null], ['+3주', null]];
  const slot = cwid / bars5.length, bwid = slot * 0.56;
  line(s, cx, base, cwid, 0, C.mute);
  bars5.forEach(([l, v, col], i) => {
    const bx = cx + i * slot + (slot - bwid) / 2;
    if (isNum(v)) {
      const hh = (v / vmax) * ph;
      R(s, bx, base - hh, bwid, hh, col);
      T(s, pct(v), { x: cx + i * slot - 0.05, y: base - hh - 0.24, w: slot + 0.1, h: 0.22, fontSize: 9.5, bold: true, color: i === 1 ? C.teal : C.slate, align: 'center', valign: 'bottom' });
    } else {
      box(s, '', bx, base - ph * 0.5, bwid, ph * 0.5, { fill: C.white, dash: true, lineColor: C.line });
    }
    T(s, l, { x: cx + i * slot - 0.05, y: base + 0.03, w: slot + 0.1, h: 0.24, fontSize: 9.5, color: C.text, align: 'center', valign: 'middle' });
  });
  const ty = base - (wsCur.wape / vmax) * ph;
  line(s, cx, ty, cwid, 0, C.red, { width: 1, dashType: 'dash' });
  T(s, '경보 기준(예: 탐색 전 수준)', { x: cx + cwid - 2.0, y: ty - 0.24, w: 2.0, h: 0.22, fontSize: 9, color: C.red, align: 'right', valign: 'middle' });
  // 그룹별 가중치 변경 표(예시: 이번 탐색 결과)
  const tx = cx + cwid + 0.2, tw = wx + ww - 0.15 - tx;
  T(s, '그룹별 가중치 변경', { x: tx, y: cy, w: tw, h: 0.28, fontSize: 11, bold: true, color: C.navy, valign: 'middle' });
  const rows = [['그룹', '현재 w', '선택 w']];
  choicesSorted.forEach((c) => rows.push([ga(c.group), wfmt(c.current), cellOpt(wfmt(c.chosen), { bold: true, color: c.chosen !== c.current ? C.teal : C.text })]));
  table(s, rows, { x: tx, y: cy + 0.32, w: tw, colW: [tw - 1.3, 0.65, 0.65], rowH: 0.27, fontSize: 9.5, headSize: 9.5, align: 'center', margin: [0.02, 0.04, 0.02, 0.04] });
  // 경보 예시 두 개
  const worse = choicesSorted.map((c) => ({ c, a: wsRow('검증', c.group, '현재', '가중평균'), z: wsRow('검증', c.group, '탐색', '가중평균') }))
    .filter((r) => r.a && r.z && isNum(r.a.wape) && isNum(r.z.wape) && r.z.wape > r.a.wape)
    .sort((p, q) => (q.z.wape - q.a.wape) - (p.z.wape - p.a.wape))[0];
  const alerts = [];
  if (coverDown) alerts.push(`결품 방어율 ${pct(wsCur.cover)} → ${pct(wsNew.cover)} 하락 · 서비스 수준 검토`);
  if (worse) alerts.push(`${ga(worse.c.group)} 오차율 ${pct(worse.a.wape)} → ${pct(worse.z.wape)} 상승 · 승인 전 확인`);
  const ay0 = wy + wh - 0.12 - alerts.length * 0.42;
  alerts.forEach((t, i) => {
    box(s, [{ text: '경보  ', options: { bold: true, color: C.red } }, { text: t, options: { color: C.text } }],
      wx + 0.15, ay0 + i * 0.42, ww - 0.3, 0.36, { fill: C.redL, lineColor: 'E8B4AD', fontSize: 10.5, align: 'left', margin: 6, round: 0.05 });
  });
  // 아래: 남은 과제 한 줄
  box(s, [
    { text: '다음 과제  ', options: { bold: true, color: '8FC9C9' } },
    { text: '서비스 수준(Z)도 같은 탐색·시계열 검증으로 조정 · 주간 재탐색으로 유의성 판단 · 실운영 데이터 적용', options: { color: C.white } },
  ], MX, 6.4, CW, 0.5, { fill: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 15. 정리
(function closing() {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.navy };
  T(s, '정리', { x: 0.9, y: 0.45, w: 8, h: 0.7, fontSize: 32, bold: true, color: C.white, valign: 'middle' });
  T(s, '무엇이 문제였고, 어떻게 풀었고, 무엇을 얻었는가', { x: 0.9, y: 1.12, w: 9.5, h: 0.38, fontSize: 14, color: C.sky, valign: 'middle' });
  tag(s, '정리', 'outline');
  const items = [
    ['무엇이 문제였나', `안전재고일수가 고정값(${oldLevels.join('·')}일)이었고, 실적과 예측을 섞는 가중치에 근거가 없었습니다.`],
    ['어떻게 풀었나', `AutoML(그리드 탐색 + 시계열 분할 검증)으로 가중치 ${man(SPACE)} 조합을 ${CANDS.length * NG}번 계산으로 탐색하고 과적합 방지 규칙(작은 그룹 합산)을 적용해 고른 뒤, 고를 때 쓰지 않은 ${HOLD}로 검증했습니다.`],
    ['무엇을 얻었나', `검증 구간 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}(${pp(wsCur.wape, wsNew.wape)} 개선), 안전재고일수 중앙값 ${f1(SS.old_median)} → ${f1(SS.new_median)}일. 결품 방어율 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}는 서비스 수준 조정 과제로 남았습니다.`],
  ];
  items.forEach(([a, b], i) => {
    const y = 1.75 + i * 1.3;
    num(s, i + 1, 0.9, y + 0.05, 0.62, C.teal, C.white, 18);
    T(s, a, { x: 1.8, y, w: 10.6, h: 0.4, fontSize: 16, bold: true, color: C.sky, valign: 'middle' });
    T(s, b, { x: 1.8, y: y + 0.42, w: 10.8, h: 0.8, fontSize: 14, color: C.white, valign: 'top' });
  });
  // 3번 항목에 이어 MLOps 다음 단계 한 줄(사용자 지시)
  box(s, [
    { text: 'MLOps 다음 단계  ', options: { bold: true, color: C.white } },
    { text: '지금은 Level 0→1 진입 · 다음은 주간 재학습·성능 모니터링을 자동화하고 승인 후 배포(Level 1 완성)', options: { color: C.white } },
  ], 1.8, 1.75 + 2 * 1.3 + 1.25, 10.8, 0.5, { fill: C.teal, fontSize: 13, align: 'left', margin: 10, round: 0.06 });
  T(s, '감사합니다 · 질의응답', { x: 0.9, y: 6.3, w: 8, h: 0.5, fontSize: 18, bold: true, color: C.sky, valign: 'middle' });
  footer(s, true);
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
  console.log(`${path.relative(ROOT, OUT_PATH)} 작성: ${page}장`);
}
save();
