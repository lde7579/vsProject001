// 발표자료 공개본 생성: presentation/data.json → out/발표_적정재고일수_공개본.pptx (15장, 비식별)
// 핵심 메시지: 현업이 품목마다 계산할 수 없어 비슷한 값(4종)으로 넣던 일별 안전재고(일)를 품목별로 자동 산출한다.
// 대표 대상은 가중평균 기준, 비교군은 실적기준·FCST기준이다. AI 탐색은 현업 산정 가중치를 출발점으로 확인·보정하는 단계다.
// 표현 원칙: 사내 시스템명 대신 '기존 수요예측 엔진', 'FCST(예측)'은 처음에 풀어 쓴다, 구현·계획 단계는 색으로만 구분한다.
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
const FOOT = 'AI 기반 품목별 일별 안전재고 자동 산출 · 테스트 데이터 기반';
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

// 세 기준(대표: 가중평균, 비교: 실적기준·FCST기준)과 화면 이름·색. 9~15장도 이 상수를 쓴다.
const BASIS3 = ['가중평균', '실적기준', 'FCST기준'];
const BLABEL = { 가중평균: '가중평균 기준', 실적기준: '실적기준', FCST기준: 'FCST기준', 기존수기: '기존 수기' };
const BCOL = { 기존수기: C.mute, 가중평균: C.teal, 실적기준: C.navy2, FCST기준: '6F8FAF' };
const BSHORT = { 가중평균: '가중평균', 실적기준: '실적', FCST기준: 'FCST' }; // 좁은 칸에 쓰는 짧은 이름
const BTINT = { 기존수기: 'E4E9EE', 가중평균: 'C9E4E4', 실적기준: 'D5DDE6', FCST기준: 'DCE6F0' };
const M = D.metrics['공통']; // 세 기준 비교 범위: 검증 주차 전체, 공통 SKU
const SB = SS.bases; // 네 기준 안전재고일수 분포(네 값이 모두 있는 SKU)
if (!SB || !BASIS3.every((b) => SB[b]) || !SB['기존수기']) throw new Error('네 기준 안전재고일수 분포가 없습니다: extract.py를 먼저 실행');
if (!M || !BASIS3.every((b) => M[b] && isNum(M[b].wape))) throw new Error('공통 범위 세 기준 지표가 없습니다: extract.py를 먼저 실행');
// 오차율이 낮은 기준부터
const BYWAPE = [...BASIS3].sort((a, b) => M[a].wape - M[b].wape);
// 화면 표기: 결측은 '–', 방어 값은 cov()를 거친다
const cv = (v) => (v === null || v === undefined || v === '' ? '–' : cov(v));

// ---------------------------------------------------------------- 프레젠테이션
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'AI 기반 품목별 일별 안전재고 자동 산출';
pres.subject = '가중평균 기준 산출, 실적·FCST기준 비교, 그리드 탐색·시계열 검증 보정 (테스트 데이터)';
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
  // 사용자 지시: 처음부터 'AI 기반 무엇'인지 밝힌다. 단계·현재 위치 같은 표현은 쓰지 않는다.
  T(s, [
    { text: 'AI 기반 품목별', options: { fontSize: 36, color: '8FC9C9', breakLine: true } },
    { text: '일별 안전재고 자동 산출 시스템', options: { fontSize: 42, color: C.white } },
  ], { x: 0.9, y: 1.6, w: 11.5, h: 1.95, bold: true, valign: 'middle', lineSpacingMultiple: 1.1 });
  // 부제: 산출 · 비교 · 보정 세 줄기
  const chips = [
    ['산출', '가중평균 기준 안전재고'],
    ['비교', '실적기준 · FCST(예측)기준'],
    ['보정', '그리드 탐색 · 시계열 검증'],
  ];
  const cw = 3.7, cg = 0.2;
  chips.forEach(([a, b], i) => {
    box(s, [{ text: a + '  ', options: { bold: true, color: '8FC9C9' } }, { text: b, options: { color: 'D5DEE7', bold: true } }],
      0.9 + i * (cw + cg), 3.8, cw, 0.52, { fill: C.navy2, lineColor: '3A5675', fontSize: 15, round: 0.08, margin: 4 });
  });
  R(s, 0.9, 4.75, 2.6, 0.06, C.teal);
  T(s, D.run_date, { x: 0.9, y: 5.1, w: 6, h: 0.4, fontSize: 15, bold: true, color: C.white });
  T(s, `${D.sku_total} SKU · 품목그룹 ${groupsOrdered.length}개 · 세 기준 비교와 가중치 ${man(SPACE)} 조합 탐색`, { x: 0.9, y: 5.55, w: 11.5, h: 0.36, fontSize: 13, color: C.sky });
  tag(s, '프로젝트 발표', 'outline');
  footer(s, true);
})();

// ======================================================================= 2. 한 장 요약
(function summary() {
  const OLDN = SB['기존수기'].distinct;
  const s = slide({
    title: `현업 수기 ${OLDN}종 대신 품목별 일별 안전재고를 자동 산출합니다`,
    sub: '목표 · 세 기준 비교 · AI 보정 · 테스트 데이터셋(비식별 처리)',
    tagText: '요약', tagKind: 'dark',
  });
  const wBest = BYWAPE[0] === '가중평균';
  const cards = [
    ['target', '목표', ['품목마다 자동 산출되는 일별 안전재고값', `현업 수기 ${oldLevels.length}종(${oldLevels.join('·')}일)을 대체`, '대표 대상: 가중평균 기준'], C.teal],
    ['chart', '세 기준 비교', ['가중평균 기준 vs 실적기준 vs FCST기준', `같은 SKU의 ${VAL} 실제 출고 오차율`, BYWAPE.map((b) => `${BSHORT[b]} ${pct(M[b].wape)}`).join(' · ')], C.navy],
    ['search', 'AI 보정', ['현업 산정 가중치를 출발점으로', '그리드 탐색 + 시계열 분할 검증', `검증 구간 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}`], C.navy],
  ];
  const cw = (CW - 0.4) / 3, cy = 1.6, chh = 1.85;
  cards.forEach(([ic, head, body, col], i) => {
    const x = MX + i * (cw + 0.2);
    RR(s, x, cy, cw, chh, i === 0 ? C.tealL : C.light);
    icon(s, ic, x + 0.25, cy + 0.2, 0.6, col);
    T(s, head, { x: x + 1.0, y: cy + 0.25, w: cw - 1.15, h: 0.5, fontSize: 16.5, bold: true, color: col, valign: 'middle' });
    T(s, lines(body), { x: x + 0.28, y: cy + 0.92, w: cw - 0.4, h: 0.85, fontSize: 12, color: C.text, paraSpaceAfter: 2 });
  });
  // 핵심 수치: 범위가 다른 두 묶음을 머리띠로 나눈다(세 기준 비교 / AI 보정)
  const kw = (CW - 0.6) / 4, ky = 4.0, kh = 1.5;
  const kx = (i) => MX + i * (kw + 0.2);
  const scopes = [
    [0, 2, `세 기준 비교 · ${VAL} · 공통 ${M['가중평균'].n_demand} SKU`, C.navy],
    [3, 3, `AI 보정 · 검증 구간 ${HOLD}`, C.teal],
  ];
  scopes.forEach(([a, b, t, col]) => {
    const x = kx(a), w = kx(b) + kw - x;
    T(s, t, { x, y: ky - 0.4, w, h: 0.28, fontSize: 12, bold: true, color: col, valign: 'middle' });
    line(s, x, ky - 0.09, w, 0, col, { width: 1.5 });
  });
  const wDown = M['가중평균'].cover < M['기존수기'].cover;
  const kpis = [
    { v: `${OLDN}종 → ${SB['가중평균'].distinct}개`, l: '안전재고일수 값의 종류', n: `기존 수기 → 가중평균 기준 · ${SB.n} SKU`, c: C.teal },
    { v: pct(M['가중평균'].wape), l: '가중평균 기준 수요 오차율', n: wBest ? `세 기준 중 가장 낮음 · ${M['가중평균'].n_demand} SKU` : `${M['가중평균'].n_demand} SKU`, c: C.teal },
    { v: pct(M['가중평균'].cover), l: '결품 방어율 (가중평균 기준)', n: wDown ? `기존 수기 ${pct(M['기존수기'].cover)}보다 낮음 · ${M['가중평균'].n_lt} SKU` : `기존 수기 ${pct(M['기존수기'].cover)} · ${M['가중평균'].n_lt} SKU`, c: wDown ? C.red : C.teal, warn: wDown },
    { v: `${pct(wsCur.wape)} → ${pct(wsNew.wape)}`, l: 'AI 보정 후 수요 오차율', n: `현업 w → 보정 w · ${NHOLD} SKU`, c: C.teal },
  ];
  kpis.forEach((k, i) => {
    const x = kx(i);
    RR(s, x, ky, kw, kh, k.warn ? C.redL : C.white, { line: { color: k.warn ? 'E8B4AD' : C.line, width: 1 } });
    if (k.warn) box(s, '주의', x + kw - 0.78, ky + 0.2, 0.6, 0.28, { round: 0.14, fill: C.red, color: C.white, fontSize: 10, bold: true });
    T(s, k.v, { x: x + 0.22, y: ky + 0.1, w: kw - (k.warn ? 1.05 : 0.44), h: 0.62, fontSize: fitSize(k.v, kw - (k.warn ? 1.1 : 0.5), 28, 18), bold: true, color: k.c, valign: 'middle' });
    T(s, k.l, { x: x + 0.22, y: ky + 0.74, w: kw - 0.3, h: 0.32, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
    T(s, k.n, { x: x + 0.22, y: ky + 1.06, w: kw - 0.3, h: 0.36, fontSize: 10.5, color: C.slate, valign: 'middle' });
  });
  box(s, [
    { text: '핵심  ', options: { bold: true, color: '8FC9C9' } },
    { text: '현업이 품목마다 계산할 수 없던 일별 안전재고를 같은 공식으로 품목마다 산출하고, 실제 출고로 채점했습니다', options: { color: C.white } },
  ], MX, 5.72, CW, 0.5, { fill: C.navy, fontSize: 13, bold: true, align: 'left', margin: 12, round: 0.06 });
  const coverNote = coverDown ? ` · AI 보정 후 방어율도 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}로 낮아져 서비스 수준(Z) 조정이 다음 과제` : '';
  T(s, lines([
    '오차율(WAPE): 실제 출고 합 대비 절대 오차 합의 비율, 낮을수록 좋음',
    `결품 방어율: 목표재고가 실제 리드타임 수요를 덮은 SKU 비율(실제 결품 기록이 아닌 대리 지표), 높을수록 좋음${coverNote}`,
  ]), { x: MX, y: 6.36, w: CW, h: 0.6, fontSize: 10, color: C.gray, valign: 'middle', paraSpaceAfter: 2 });
})();

// ======================================================================= 3. 문제
(function problem() {
  const s = slide({
    title: `현업 산정 안전재고일수가 ${oldLevels.length}가지 값에 몰려 있습니다`,
    sub: `방식이 틀렸다기보다 품목마다 계산할 수단이 없습니다 · 기존 안전재고일수가 있는 ${D.days.old_safety_n} SKU`,
    tagText: '문제 정의', tagKind: 'light',
  });
  const lw = 6.3, y = 1.72, h = 4.3;
  RR(s, MX, y, lw, h, C.light);
  T(s, '현업 산정 안전재고일수별 SKU 수', { x: MX + 0.3, y: y + 0.15, w: lw - 0.6, h: 0.36, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  T(s, `${D.days.old_safety_n} SKU가 ${oldLevels.join('·')}일 ${oldLevels.length}가지 값 중 하나를 씁니다`, { x: MX + 0.3, y: y + 0.5, w: lw - 0.6, h: 0.3, fontSize: 11.5, color: C.gray, valign: 'middle' });
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
    ['calendar', '비슷한 값 일괄 입력', `품목마다 계산할 수단이 없어 ${oldLevels.join('·')}일 중 하나를 골라 넣습니다`],
    ['gear', '수요·리드타임 변화 미반영', '수요 크기·흔들림·리드타임이 바뀌어도 넣어 둔 일수가 그대로 남습니다'],
    ['check', '사후 채점 없음', '넣은 값이 이후 실제 출고를 덮었는지 확인하지 않습니다'],
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
    { text: '그래서 → 목표  ', options: { bold: true, color: '8FC9C9' } },
    { text: '품목마다 같은 공식으로 일별 안전재고(일)를 자동 산출하고, 실제 출고로 채점합니다', options: { color: C.white } },
  ], MX, y + h + 0.2, CW, 0.65, { fill: C.navy, fontSize: 13.5, bold: true, align: 'left', margin: 14, round: 0.06 });
})();

// ======================================================================= 4. 목표와 비교 설계
(function goal() {
  const s = slide({
    title: '목표: 품목별 일별 안전재고를 자동 산출하고 세 기준으로 비교합니다',
    sub: '대표 대상은 가중평균 기준 · 비교군은 실적기준과 FCST(예측)기준 · 테스트 데이터셋(비식별 처리)',
    tagText: '목표', tagKind: 'light',
  });
  const lw = 7.75, y = 1.62;
  // 목표 한 줄
  box(s, [
    { text: '목표  ', options: { bold: true, color: '8FC9C9' } },
    { text: '품목마다 자동 산출되는 일별 안전재고값 (안전재고 ÷ 일평균 수요 = 일)', options: { color: C.white } },
  ], MX, y, lw, 0.56, { fill: C.navy, fontSize: 13.5, bold: true, align: 'left', margin: 12, round: 0.06 });
  // 세 기준 카드
  const bases = [
    ['가중평균', '대표', '현업 산정 가중치 w로 실적과 예측을 섞음'],
    ['실적기준', '비교', '실제 출고 실적만으로 수요를 잡음'],
    ['FCST기준', '비교', '기존 수요예측 엔진의 예측값만으로 수요를 잡음'],
  ];
  const bw = (lw - 0.3) / 3, by = y + 0.72, bh = 1.3;
  bases.forEach(([k, role, body], i) => {
    const x = MX + i * (bw + 0.15), main = i === 0;
    RR(s, x, by, bw, bh, main ? C.tealL : C.light, main ? { line: { color: C.teal, width: 1.5 } } : {});
    R(s, x, by, 0.09, bh, BCOL[k]);
    T(s, BLABEL[k], { x: x + 0.25, y: by + 0.14, w: bw - 1.1, h: 0.4, fontSize: 15, bold: true, color: main ? C.teal : C.navy, valign: 'middle' });
    box(s, role, x + bw - 0.78, by + 0.18, 0.62, 0.3, { round: 0.15, fill: main ? C.teal : C.white, color: main ? C.white : C.slate, lineColor: main ? C.teal : C.line, bold: true, fontSize: 10.5, margin: 0 });
    T(s, body, { x: x + 0.25, y: by + 0.6, w: bw - 0.4, h: bh - 0.68, fontSize: 12, color: C.text });
  });
  // 현업 가중치 설명
  const wy = by + bh + 0.14, wh = 0.78;
  RR(s, MX, wy, lw, wh, C.white, { line: { color: C.line, width: 1 } });
  icon(s, 'gear', MX + 0.18, wy + 0.13, 0.52, C.slate);
  T(s, [
    { text: `현업 산정 가중치 w (${WRANGE}) `, options: { bold: true, color: C.navy, breakLine: true } },
    { text: '그룹마다 현업이 자체 공식으로 계산한 값 · AI 탐색은 이 값에서 출발해 확인·보정', options: { color: C.text, fontSize: 11.5 } },
  ], { x: MX + 0.85, y: wy + 0.06, w: lw - 1.0, h: wh - 0.12, fontSize: 13, valign: 'middle', paraSpaceAfter: 2 });
  // 연구 질문 세 개
  const qy0 = wy + wh + 0.12;
  label(s, '연구 질문', MX, qy0, lw, C.navy, 14);
  const qs = [
    ['세 기준 중 실제 출고에 가장 가까운 것은?', `같은 SKU를 ${VAL}(${VS.n}주) 실제 출고로 채점`],
    ['품목마다 다른 값이 나오는가?', `현업 수기 ${oldLevels.length}종과 같은 SKU에서 분포 비교`],
    ['AI 보정으로 더 나아지는가?', `현업 w에서 출발해 탐색, 고를 때 안 쓴 ${HOLD}로 확인`],
  ];
  const qh = 0.56;
  qs.forEach(([a, b], i) => {
    const qy = qy0 + 0.38 + i * (qh + 0.07);
    RR(s, MX, qy, lw, qh, C.light);
    num(s, `Q${i + 1}`, MX + 0.12, qy + (qh - 0.4) / 2, 0.4, C.teal, C.white, 11);
    T(s, a, { x: MX + 0.62, y: qy, w: 3.75, h: qh, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: MX + 4.4, y: qy, w: lw - 4.5, h: qh, fontSize: 11, color: C.slate, valign: 'middle' });
  });
  // 오른쪽: 테스트 데이터셋
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  label(s, '테스트 데이터셋 (비식별 처리)', rx, y, rw, C.navy, 14);
  const k = (t) => cellOpt(t, { bold: true, color: C.navy });
  const rows = [
    ['항목', '값'],
    [k('품목'), `${D.sku_total} SKU`],
    [k('그룹'), `${groupsOrdered.length}개 (그룹 A~${String.fromCharCode(64 + groupsOrdered.length)})`],
    [k('출고 실적'), `${HIST} (${HS.n}주) · 계산 입력`],
    [k('예측'), `${FW}주 (${wr(`W${SS_.a}~W${FCST_END}`)}) · 계산 입력`],
    [k('세 기준 채점'), cellOpt(`${VAL} (${VS.n}주) 실제 출고`, { bold: true, color: C.navy })],
    [k('AI 선택 구간'), cellOpt(`${SEL} (${SS_.n}주) · w를 고름`, { bold: true, color: C.amber })],
    [k('AI 검증 구간'), cellOpt(`${HOLD} (${HO.n}주) · 선택에 미사용`, { bold: true, color: C.teal })],
    [k('비식별'), '품목·그룹 코드 치환, 수량 비공개'],
  ];
  const rh = 0.46;
  table(s, rows, { x: rx, y: y + 0.42, w: rw, colW: [1.3, rw - 1.3], rowH: rh, fontSize: 11.5, headSize: 12 });
  T(s, '실제 출고는 채점에만 쓰고 계산 입력에는 쓰지 않습니다. 숫자는 비율(%), 일수, SKU 수, 주차 번호만 씁니다.',
    { x: rx, y: y + 0.42 + rh * rows.length + 0.12, w: rw, h: 0.6, fontSize: 10.5, color: C.gray, valign: 'top' });
})();

// ======================================================================= 5. 시스템 구성
(function system() {
  const s = slide({
    title: 'MLOps 파이프라인으로 세 기준 산출부터 리포트까지 자동화했습니다',
    sub: '한 줄 흐름 · 구현 범위는 명령 한 번으로 재현됩니다',
    tagText: '시스템', tagKind: 'teal',
  });
  T(s, '입력: 출고 실적 · 수요 예측 · 품목 기준정보(리드타임 등) · 재고', { x: MX, y: 1.6, w: 8.5, h: 0.32, fontSize: 12.5, bold: true, color: C.slate, valign: 'middle' });
  // 범례: 구현은 청록, 계획은 회색 점선(위치 표시 없이 색으로만 구분)
  const lgx = W - MX - 2.9;
  RR(s, lgx, 1.67, 0.34, 0.2, C.tealL, { line: { color: C.teal, width: 1.25 } });
  T(s, '구현', { x: lgx + 0.42, y: 1.6, w: 0.8, h: 0.32, fontSize: 11, color: C.text, valign: 'middle' });
  box(s, '', lgx + 1.35, 1.67, 0.34, 0.2, { fill: C.white, dash: true, lineColor: C.mute, round: 0.04 });
  T(s, '계획', { x: lgx + 1.77, y: 1.6, w: 0.8, h: 0.32, fontSize: 11, color: C.text, valign: 'middle' });
  const steps = [
    ['db', '데이터 검사', '빈 값·이상값에 점검메모', true],
    ['chart', '세 기준 산출', '기준별 안전재고일수 계산', true],
    ['table', '재고등급', '규칙 기반 A·B·C → 서비스 수준', true],
    ['search', 'AI 보정', `현업 w에서 출발, 후보 ${CANDS.length}개 채점`, true],
    ['calendar', '시계열 검증', '선택에 안 쓴 구간으로 확인', true],
    ['doc', '리포트', '결과표·발표 자료 자동 생성', true],
    ['check', '승인·배포', '사람이 승인한 뒤 값 교체', false],
    ['target', '모니터링', '오차율·방어율 추적, 경보', false],
    ['gear', '재탐색', '매주 가중치 다시 고르기', false],
  ];
  const n = steps.length, gap = 0.17, bw = (CW - gap * (n - 1)) / n, y = 2.2, bh = 2.3;
  const xs = steps.map((_, i) => MX + i * (bw + gap));
  steps.forEach(([ic, a, b, done], i) => {
    const x = xs[i];
    if (done) RR(s, x, y, bw, bh, C.tealL, { line: { color: C.teal, width: 1.25 } });
    else box(s, '', x, y, bw, bh, { fill: C.white, dash: true, lineColor: C.mute, round: 0.08 });
    icon(s, ic, x + (bw - 0.62) / 2, y + 0.2, 0.62, done ? C.teal : C.mute);
    T(s, a, { x: x + 0.03, y: y + 0.92, w: bw - 0.06, h: 0.4, fontSize: 12.5, bold: true, color: done ? C.navy : C.gray, align: 'center', valign: 'middle' });
    T(s, b, { x: x + 0.08, y: y + 1.36, w: bw - 0.16, h: 0.85, fontSize: 10.5, color: done ? C.text : C.gray, align: 'center' });
    if (i < n - 1) arrow(s, x + bw + 0.01, y + bh / 2, gap - 0.02, 0, done && steps[i + 1][3] ? C.teal : C.mute, { width: 1.5 });
  });
  // 되먹임(계획): 재탐색 → AI 보정
  const fy = y + bh + 0.3, from = xs[n - 1] + bw / 2, to = xs[3] + bw / 2;
  line(s, from, y + bh, 0, fy - y - bh, C.mute, { width: 1.5, dashType: 'dash' });
  line(s, to, fy, from - to, 0, C.mute, { width: 1.5, dashType: 'dash' });
  arrow(s, to, fy, 0, -(fy - y - bh) + 0.02, C.mute, { width: 1.5 });
  box(s, '되먹임(계획): 다음 주 데이터로 AI 보정을 다시 돌림', to + 1.0, fy - 0.15, 4.6, 0.3, { fill: C.white, lineColor: C.white, color: C.gray, fontSize: 11, bold: true, margin: 0 });
  // 구현 / 계획 범위
  const ry = fy + 0.42;
  box(s, [
    { text: '구현  ', options: { bold: true, color: '8FC9C9' } },
    { text: `데이터 검사 · 세 기준 산출 · 재고등급 · AI 보정 · 시계열 검증 · 리포트 + 자동 테스트 ${int(TESTS)}개 · 버전 관리(git)`, options: { color: C.white } },
  ], MX, ry, CW, 0.56, { fill: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
  box(s, [
    { text: '계획  ', options: { bold: true, color: C.slate } },
    { text: '승인 후 배포 · 성능 모니터링(기준을 넘으면 경보) · 주간 재탐색', options: { color: C.gray } },
  ], MX, ry + 0.66, CW, 0.5, { fill: C.white, dash: true, lineColor: C.mute, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
  T(s, '검증 구간의 실제 출고는 채점에만 쓰고 계산 입력에는 쓰지 않습니다.', { x: MX, y: ry + 1.24, w: CW, h: 0.28, fontSize: 10.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 6. 산출 방법
(function method() {
  const G = D.grades;
  const s = slide({
    title: '세 기준은 수요를 잡는 방식만 다르고 안전재고 식은 같습니다',
    sub: '일반식 두 개와 결과 값 · 세부 산식과 등급별 계수 값은 생략합니다',
    tagText: '방법', tagKind: 'teal',
  });
  const y = 1.66, fh = 3.0;
  // ① 수요: 기준별
  const dw = 5.0;
  RR(s, MX, y, dw, fh, C.light);
  T(s, '① 수요 (기준마다 다름)', { x: MX + 0.25, y: y + 0.12, w: dw - 0.5, h: 0.36, fontSize: 14, bold: true, color: C.teal, valign: 'middle' });
  const dem = [
    ['가중평균', 'w·실적 + (1−w)·예측', 'w: 현업 산정 가중치(그룹마다 하나)'],
    ['실적기준', '실적', '과거 출고 실적의 주평균'],
    ['FCST기준', '예측', `기존 수요예측 엔진의 앞 ${FW}주 예측 평균`],
  ];
  const rh = 0.74;
  dem.forEach(([k, f, note], i) => {
    const ry = y + 0.58 + i * (rh + 0.06), main = i === 0;
    RR(s, MX + 0.2, ry, dw - 0.4, rh, C.white, { line: { color: main ? C.teal : C.line, width: main ? 1.5 : 1 } });
    R(s, MX + 0.2, ry, 0.08, rh, BCOL[k]);
    T(s, BLABEL[k], { x: MX + 0.4, y: ry + 0.04, w: 1.5, h: rh - 0.08, fontSize: 12.5, bold: true, color: main ? C.teal : C.navy, valign: 'middle' });
    T(s, [{ text: `수요 = ${f}`, options: { bold: true, color: C.navy, fontSize: main ? 14 : 13, breakLine: true } }, { text: note, options: { color: C.slate, fontSize: 10.5 } }],
      { x: MX + 1.95, y: ry + 0.04, w: dw - 2.25, h: rh - 0.08, valign: 'middle' });
  });
  arrow(s, MX + dw + 0.04, y + fh / 2, 0.27, 0, C.slate, { width: 2 });
  // ② 안전재고
  const sx = MX + dw + 0.35, sw = 3.55;
  RR(s, sx, y, sw, fh, C.light);
  T(s, '② 안전재고 (세 기준 공통)', { x: sx + 0.25, y: y + 0.12, w: sw - 0.5, h: 0.36, fontSize: 14, bold: true, color: C.teal, valign: 'middle' });
  box(s, [{ text: '안전재고', options: { fontSize: 13, color: C.slate, breakLine: true } }, { text: '= Z·σ·√LT', options: { fontSize: 22, color: C.navy } }],
    sx + 0.25, y + 0.62, sw - 0.5, 1.15, { fill: C.white, lineColor: C.line, bold: true, round: 0.06 });
  T(s, lines(['Z: 서비스 수준 계수 (재고등급별)', 'σ: 그 기준 수요의 흔들림', 'LT: 리드타임']), { x: sx + 0.25, y: y + 1.92, w: sw - 0.5, h: 0.95, fontSize: 11.5, color: C.slate, paraSpaceAfter: 3 });
  arrow(s, sx + sw + 0.04, y + fh / 2, 0.27, 0, C.teal, { width: 2 });
  // 결과 값
  const ox = sx + sw + 0.35, ow = W - MX - ox;
  RR(s, ox, y, ow, fh, C.teal);
  T(s, '결과 값', { x: ox + 0.2, y: y + 0.12, w: ow - 0.4, h: 0.36, fontSize: 14, bold: true, color: 'E6F4F4', valign: 'middle' });
  T(s, [{ text: '안전재고일수', options: { fontSize: 20, breakLine: true } }, { text: '= 안전재고', options: { fontSize: 15, breakLine: true } }, { text: '÷ 일평균 수요', options: { fontSize: 15 } }],
    { x: ox + 0.2, y: y + 0.62, w: ow - 0.4, h: 1.2, bold: true, color: C.white, valign: 'middle', paraSpaceAfter: 4 });
  T(s, '품목마다, 기준마다 다른 일수가 나옵니다 (현업 수기 값 대신)', { x: ox + 0.2, y: y + 1.95, w: ow - 0.4, h: 0.9, fontSize: 11.5, color: 'E6F4F4' });
  // 등급 카드(짧게)
  const gy = y + fh + 0.3, gh = 6.85 - gy;
  RR(s, MX, gy, CW, gh, C.white, { line: { color: C.line, width: 1 } });
  icon(s, 'table', MX + 0.3, gy + 0.3, 0.7, C.navy);
  T(s, '재고등급(A·B·C)별 서비스 수준 Z 적용', { x: MX + 1.2, y: gy + 0.22, w: 6.4, h: 0.42, fontSize: 16, bold: true, color: C.navy, valign: 'middle' });
  T(s, `규칙 기반 등급이 기존 등급 ${G.match}/${G.total} 일치`, { x: MX + 1.2, y: gy + 0.66, w: 6.4, h: 0.4, fontSize: 14, bold: true, color: C.teal, valign: 'middle' });
  T(s, '단가·회전율·리드타임·품질 이슈를 종합해 등급을 매기고, 등급이 높을수록 더 높은 서비스 수준을 씁니다',
    { x: MX + 1.2, y: gy + 1.1, w: 6.4, h: gh - 1.2, fontSize: 11.5, color: C.text });
  const bx0 = MX + 8.0, bwid = CW - 8.3;
  T(s, `계산 재고등급 분포 (${G.total} SKU)`, { x: bx0, y: gy + 0.3, w: bwid, h: 0.32, fontSize: 12.5, bold: true, color: C.navy, valign: 'middle' });
  const tot = G.dist.A + G.dist.B + G.dist.C; let bx = bx0;
  [['A', C.teal], ['B', C.slate], ['C', C.mute]].forEach(([k, col]) => {
    const bw = (G.dist[k] / tot) * bwid;
    box(s, `${k} ${G.dist[k]}`, bx, gy + 0.72, bw, 0.55, { fill: col, color: C.white, bold: true, fontSize: 13, margin: 0 });
    bx += bw;
  });
  T(s, '숫자는 SKU 수', { x: bx0, y: gy + 1.32, w: bwid, h: 0.28, fontSize: 10.5, color: C.gray, valign: 'middle' });
})();

// ======================================================================= 7. 결과 1: 품목별 값
(function spread() {
  const ALL = ['기존수기', ...BASIS3];
  const OLD = SB['기존수기'];
  const s = slide({
    title: `자동 산출하면 ${OLD.distinct}종이던 안전재고일수가 품목마다 달라집니다`,
    sub: `같은 ${SB.n} SKU의 안전재고일수 · 막대 = 중앙값, 띠 = P10~P90, 최댓값은 글자로 · 고유값은 소수 1자리 기준`,
    tagText: '결과 1', tagKind: 'dark',
  });
  // 축: P90 중 가장 큰 값까지(10일 단위로 올림). 그보다 큰 최댓값은 글자로만 적는다.
  const AX = Math.ceil((Math.max(...ALL.map((b) => SB[b].p90)) * 1.08) / 10) * 10;
  const lx = MX, labW = 1.75, cx = MX + labW, cwid = 4.75;
  const X = (v) => cx + (Math.min(v, AX) / AX) * cwid;
  // 표 열
  const cols = [['중앙값', 0.95], ['P10~P90', 1.5], ['최댓값', 1.05], ['고유값', 0.85], ['P90÷P10', 0.95]];
  const tx0 = cx + cwid + 0.25;
  const colX = []; cols.reduce((acc, [, w]) => (colX.push(acc), acc + w), tx0);
  const y0 = 1.72, hh = 0.36, rh = 0.8;
  // 머리 줄
  T(s, '기준', { x: lx, y: y0, w: labW, h: hh, fontSize: 11.5, bold: true, color: C.navy, valign: 'middle' });
  T(s, '안전재고일수 (일)', { x: cx, y: y0, w: cwid, h: hh, fontSize: 11.5, bold: true, color: C.navy, valign: 'middle' });
  cols.forEach(([t, w], j) => T(s, t, { x: colX[j], y: y0, w, h: hh, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' }));
  line(s, lx, y0 + hh + 0.02, CW, 0, C.navy, { width: 1.25 });
  const ry0 = y0 + hh + 0.1, chartH = rh * ALL.length;
  // 눈금
  for (let v = 0; v <= AX; v += 10) {
    line(s, X(v), ry0, 0, chartH, v === 0 ? C.mute : 'E6EBF0', { width: v === 0 ? 1 : 0.75 });
    T(s, String(v), { x: X(v) - 0.3, y: ry0 + chartH + 0.02, w: 0.6, h: 0.24, fontSize: 9.5, color: C.gray, align: 'center', valign: 'middle' });
  }
  ALL.forEach((b, i) => {
    const d = SB[b], ry = ry0 + i * rh, mid = ry + rh / 2, main = b === '가중평균';
    if (main) R(s, lx, ry + 0.04, CW, rh - 0.08, 'F1F8F8');
    if (i) line(s, lx, ry, CW, 0, C.line, { width: 0.75 });
    T(s, BLABEL[b], { x: lx + 0.1, y: ry, w: labW - 0.1, h: rh, fontSize: 13, bold: true, color: main ? C.teal : b === '기존수기' ? C.slate : C.navy, valign: 'middle' });
    // P10~P90 띠
    R(s, X(d.p10), mid - 0.22, Math.max(0.04, X(d.p90) - X(d.p10)), 0.44, BTINT[b]);
    // 중앙값 막대(0 → 중앙값)와 표시
    R(s, cx, mid - 0.07, Math.max(0.03, X(d.median) - cx), 0.14, BCOL[b]);
    R(s, X(d.median) - 0.03, mid - 0.26, 0.06, 0.52, BCOL[b]);
    if (d.max > AX) T(s, `최대 ${f1(d.max)}일 →`, { x: cx + cwid - 1.6, y: mid - 0.38, w: 1.58, h: 0.22, fontSize: 9, color: C.gray, align: 'right', valign: 'middle' });
    // 표 값
    const vals = [
      [`${f1(d.median)}일`, true],
      [`${f1(d.p10)} ~ ${f1(d.p90)}일`, false],
      [`${f1(d.max)}일`, false],
      [`${d.distinct}개`, true],
      [isNum(d.p10) && d.p10 > 0 ? `${(d.p90 / d.p10).toFixed(1)}배` : '–', false],
    ];
    vals.forEach(([t, bold], j) => T(s, t, { x: colX[j], y: ry, w: cols[j][1], h: rh, fontSize: bold ? 14 : 12, bold, color: main ? C.teal : C.text, align: 'center', valign: 'middle' }));
  });
  line(s, lx, ry0 + chartH, CW, 0, C.line, { width: 0.75 });
  T(s, 'P90÷P10 = 퍼짐 배수(작을수록 고르게 분포)', { x: tx0, y: ry0 + chartH + 0.02, w: W - MX - tx0, h: 0.24, fontSize: 9.5, color: C.gray, align: 'right', valign: 'middle' });
  // 해석(데이터로 조건부 작성)
  // 퍼짐은 크기가 다른 기준끼리 견주도록 P90÷P10(배수)으로 본다
  const ratio = (b) => (SB[b].p10 > 0 ? SB[b].p90 / SB[b].p10 : Infinity);
  const narrow = [...BASIS3].sort((a, b) => ratio(a) - ratio(b))[0];
  const extreme = BASIS3.filter((b) => b !== '가중평균' && SB[b].max > SB['가중평균'].max);
  const notes = [
    `기존 수기는 ${OLD.distinct}종(${int(OLD.min)}~${int(OLD.max)}일)뿐이지만, 세 기준은 ${BASIS3.map((b) => SB[b].distinct).join('·')}개 값으로 품목마다 달라집니다.`,
  ];
  const others = BASIS3.filter((b) => b !== narrow);
  if (narrow === '가중평균') notes.push(`가중평균 기준은 P90이 P10의 ${ratio(narrow).toFixed(1)}배로, ${others.map((b) => `${BLABEL[b]} ${ratio(b).toFixed(1)}배`).join('·')}보다 고르게 퍼져 안정적입니다.`);
  else notes.push(`P90÷P10이 가장 작은 기준은 ${BLABEL[narrow]}(${ratio(narrow).toFixed(1)}배)입니다.`);
  if (extreme.length) notes.push(`${extreme.map((b) => BLABEL[b]).join('·')}은 최댓값이 ${extreme.map((b) => f1(SB[b].max) + '일').join('·')}로 극단값이 나옵니다(가중평균 기준 최대 ${f1(SB['가중평균'].max)}일).`);
  notes.push('짧을수록 좋은 것이 아니라, 실제 출고를 덮는지가 기준입니다 → 실제 출고 채점에서 확인');
  const by = ry0 + chartH + 0.36, bh = 6.9 - by;
  RR(s, MX, by, CW, bh, C.navy);
  T(s, notes.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < notes.length - 1, color: i === notes.length - 1 ? C.sky : C.white, fontSize: i === notes.length - 1 ? 11 : 12 } })),
    { x: MX + 0.25, y: by + 0.08, w: CW - 0.5, h: bh - 0.16, valign: 'middle', paraSpaceAfter: 3 });
})();

// ======================================================================= 8. 사례: 품목 X
(function example() {
  const c = D.case, cc = D.counter_case;
  const s = slide({
    title: `사례: 품목 X의 안전재고일수 기존 ${c.old_ss_days}일 → 가중평균 기준 ${f1(c.s)}일`,
    sub: `${ga(c.group)} · 재고등급 ${c.grade} · 리드타임 ${c.lt}일 · 현업 w 기준(AI 보정 전) · 수량 없이 일수로`,
    tagText: '사례', tagKind: 'light',
  });
  const rowsDef = [
    ['기존수기', c.old_ss_days, c.old_days, c.old_cover, c.old_actual_days],
    ['가중평균', c.s, c.u, c.new_cover, c.new_actual_days],
    ['실적기준', c.act_s, c.act_u, c.act_cover, c.act_actual_days],
    ['FCST기준', c.fc_s, c.fc_u, c.fc_cover, c.fc_actual_days],
  ];
  const fmtD = (v) => (isNum(v) ? `${Number.isInteger(v) ? v : f1(v)}일` : '–');
  const lw = 5.6, y = 1.66, ph = 3.55;
  // 왼쪽: 네 기준 안전재고일수 막대
  RR(s, MX, y, lw, ph, C.light);
  T(s, '안전재고일수 비교', { x: MX + 0.25, y: y + 0.12, w: lw - 0.5, h: 0.34, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  const maxD = Math.max(...rowsDef.map((r) => (isNum(r[1]) ? r[1] : 0)));
  const barX = MX + 1.75, barW = lw - 1.75 - 1.15;
  rowsDef.forEach(([b, v], i) => {
    const ry = y + 0.6 + i * 0.6, main = b === '가중평균';
    const tc = main ? C.teal : b === '기존수기' ? C.slate : C.navy;
    T(s, BLABEL[b], { x: MX + 0.25, y: ry, w: 1.45, h: 0.5, fontSize: 12.5, bold: true, color: tc, valign: 'middle' });
    const bl = isNum(v) ? Math.max(0.05, (v / maxD) * barW) : 0;
    if (isNum(v)) R(s, barX, ry + 0.1, bl, 0.3, BCOL[b]);
    T(s, fmtD(v), { x: barX + bl + 0.08, y: ry, w: 1.1, h: 0.5, fontSize: main ? 20 : 16, bold: true, color: tc, valign: 'middle' });
  });
  const fy = y + 0.6 + 4 * 0.6 + 0.05;
  line(s, MX + 0.25, fy, lw - 0.5, 0, C.line);
  T(s, lines([
    `가중평균 수요 = ${wfmt(c.wh)}·실적 + ${wfmt(c.wf)}·예측 (그룹의 현업 w)`,
    `안전재고 = Z(등급 ${c.grade})·σ·√LT → 일평균 수요로 나눔`,
  ]), { x: MX + 0.25, y: fy + 0.08, w: lw - 0.5, h: y + ph - fy - 0.16, fontSize: 11, color: C.navy, valign: 'middle', paraSpaceAfter: 2 });
  // 오른쪽: 실제 출고 채점 표
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  RR(s, rx, y, rw, ph, C.white, { line: { color: C.line, width: 1 } });
  T(s, `실제 출고로 채점 · ${VAL}(${VS.n}주)`, { x: rx + 0.25, y: y + 0.12, w: rw - 0.5, h: 0.34, fontSize: 14, bold: true, color: C.navy, valign: 'middle' });
  const ok = (v) => (v === '방어' ? C.teal : v === '결품' ? C.red : C.gray);
  const trows = [['기준', '안전재고일수', '적정재고일수', '리드타임 수요 방어', '실제 기준 재고일수']];
  rowsDef.forEach(([b, sd, ud, cvr, ad]) => {
    const main = b === '가중평균';
    trows.push([
      cellOpt(BLABEL[b], { bold: true, color: main ? C.teal : C.navy, align: 'left', fill: { color: main ? 'F1F8F8' : C.white } }),
      cellOpt(fmtD(sd), { bold: main, fill: { color: main ? 'F1F8F8' : C.white } }),
      cellOpt(fmtD(ud), { fill: { color: main ? 'F1F8F8' : C.white } }),
      cellOpt(cv(cvr), { bold: true, color: ok(cvr), fill: { color: main ? 'F1F8F8' : C.white } }),
      cellOpt(fmtD(ad), { bold: main, fill: { color: main ? 'F1F8F8' : C.white } }),
    ]);
  });
  const tw = rw - 0.5;
  table(s, trows, { x: rx + 0.25, y: y + 0.58, w: tw, colW: [1.4, 1.15, tw - 1.4 - 1.15 - 1.05 - 1.0, 1.05, 1.0], rowH: [0.55, 0.44, 0.44, 0.44, 0.44], fontSize: 12, headSize: 10.5, align: 'center' });
  T(s, lines([
    '적정재고일수 = 리드타임 + 안전재고일수',
    '방어 = 목표재고가 실제 리드타임 수요를 덮음(대리 지표)',
    '실제 기준 재고일수 = 목표재고 ÷ 실제 일평균 출고',
  ]), { x: rx + 0.25, y: y + 0.58 + 0.55 + 0.44 * 4 + 0.08, w: rw - 0.5, h: 0.55, fontSize: 10, color: C.gray, valign: 'top', paraSpaceAfter: 1 });
  T(s, `선정 규칙: ${c.rule} (가중평균에 유리한 사례)`, { x: MX, y: y + ph + 0.06, w: CW, h: 0.3, fontSize: 10, color: C.gray, valign: 'middle' });
  // 아래: 요약과 반대 사례
  const covs = rowsDef.map((r) => r[3]);
  const allOk = covs.every((v) => v === '방어');
  const fdiff = c.g / c.actual_week - 1;
  const msg = allOk
    ? `네 기준 모두 방어했고, 실제 기준 재고일수는 기존 수기 ${f1(c.old_actual_days)}일 → 가중평균 기준 ${f1(c.new_actual_days)}일입니다. 단, 가중평균 수요는 실제보다 ${pct(Math.abs(fdiff), 0)} ${fdiff >= 0 ? '높았' : '낮았'}습니다.`
    : `방어 결과: ${rowsDef.map((r) => `${BLABEL[r[0]]} ${cv(r[3])}`).join(' · ')}`;
  box(s, msg, MX, 5.62, CW, 0.58, { fill: C.navy, color: C.white, fontSize: 12, bold: true, align: 'left', valign: 'middle', margin: 10, round: 0.06 });
  const ccCovs = [cc.old_cover, cc.new_cover, cc.act_cover, cc.fc_cover];
  const ccAll = ccCovs.every((v) => v === '결품');
  box(s, [
    { text: '반대 사례  ', options: { bold: true, color: C.white } },
    { text: `${ga(cc.group)}의 한 품목은 기존 ${cc.old_ss_days}일 → 가중평균 기준 ${f1(cc.s)}일이었고, 실제 출고가 가중평균 수요의 ${(cc.actual_week / cc.g).toFixed(1)}배여서 ${ccAll ? '네 기준 모두' : '가중평균 기준은'} 방어에 실패했습니다.`, options: { color: C.white } },
  ], MX, 6.3, CW, 0.58, { fill: C.red, fontSize: 12, align: 'left', valign: 'middle', margin: 10, round: 0.06 });
})();

// ======================================================================= 9. 결과 2: 실제 출고 채점
(function scoring() {
  const ALL = ['기존수기', ...BASIS3];
  const best = BYWAPE[0];
  const title = best === '가중평균'
    ? `가중평균 기준이 실제 출고에 가장 가깝습니다 (오차율 ${pct(M[best].wape)})`
    : `실제 출고에 가장 가까운 기준은 ${BLABEL[best]}입니다 (오차율 ${pct(M[best].wape)})`;
  // 방어율·재고일수 채점 SKU 수(네 기준이 같으면 한 값)
  const NLT = [...new Set(ALL.map((b) => M[b].n_lt))].sort((a, b) => a - b);
  const nlt = NLT.length === 1 ? `${NLT[0]}` : `${NLT[0]}~${NLT[NLT.length - 1]}`;
  const s = slide({
    title,
    sub: `실제 출고로 채점 · ${VAL}(${VS.n}주) · 공통 ${M['가중평균'].n_demand} SKU (방어율·재고일수는 ${nlt} SKU) · 청록 = 지표별 가장 좋은 값`,
    tagText: '결과 2', tagKind: 'dark',
  });
  const days = (v) => (isNum(v) ? `${f1(v)}일` : '–');
  // best: 'min' 낮을수록, 'abs' 0에 가까울수록, 'max' 높을수록 좋음, null 은 우열 없이 참고
  const rowsDef = [
    { key: 'wape', name: '수요 오차율 (WAPE)', how: '낮을수록 좋음 · |예측 − 실제| 합 ÷ 실제 합', fmt: (v) => pct(v), best: 'min' },
    { key: 'bias', name: '수요 편향', how: '0에 가까울수록 좋음 · + 많게, − 적게 잡음', fmt: (v) => spct(v), best: 'abs' },
    { key: 'cover', name: '결품 방어율', how: '높을수록 좋음 · 리드타임 수요를 덮은 비율(대리 지표)', fmt: (v) => pct(v), best: 'max' },
    { key: 'days_median', name: '실제 기준 재고일수 중앙값', how: '참고 · 길수록 재고를 많이 듦', fmt: days, best: null },
  ];
  const score = (r, v) => (r.best === 'min' ? v : r.best === 'abs' ? Math.abs(v) : -v);
  const bestOf = (r) => {
    if (!r.best) return null;
    const c = ALL.filter((b) => isNum(M[b][r.key]));
    return c.length ? c.reduce((a, b) => (score(r, M[b][r.key]) < score(r, M[a][r.key]) ? b : a)) : null;
  };
  const labW = 3.75, bw = (CW - labW) / ALL.length;
  const bx = (j) => MX + labW + j * bw;
  const y0 = 1.62, hh = 0.46, rh = 0.74;
  const ry0 = y0 + hh + 0.06, tableH = rh * rowsDef.length;
  // 가중평균 열 강조 띠
  const mj = ALL.indexOf('가중평균');
  R(s, bx(mj), y0, bw, hh + 0.06 + tableH, 'F1F8F8');
  // 머리 줄
  T(s, '지표', { x: MX + 0.1, y: y0, w: labW - 0.1, h: hh, fontSize: 12, bold: true, color: C.navy, valign: 'middle' });
  ALL.forEach((b, j) => T(s, BLABEL[b], {
    x: bx(j), y: y0, w: bw, h: hh, fontSize: 13.5, bold: true, align: 'center', valign: 'middle',
    color: b === '가중평균' ? C.teal : b === '기존수기' ? C.slate : C.navy,
  }));
  line(s, MX, y0 + hh + 0.02, CW, 0, C.navy, { width: 1.25 });
  rowsDef.forEach((r, i) => {
    const ry = ry0 + i * rh, win = bestOf(r);
    if (i) line(s, MX, ry, CW, 0, C.line, { width: 0.75 });
    T(s, [
      { text: r.name, options: { fontSize: 14, bold: true, color: C.navy, breakLine: true } },
      { text: r.how, options: { fontSize: 10, color: C.gray } },
    ], { x: MX + 0.1, y: ry, w: labW - 0.2, h: rh, valign: 'middle', paraSpaceAfter: 2 });
    ALL.forEach((b, j) => {
      const v = M[b][r.key];
      if (b === win) {
        box(s, r.fmt(v), bx(j) + 0.22, ry + 0.11, bw - 0.44, rh - 0.22, { fill: C.teal, color: C.white, bold: true, fontSize: 19, round: 0.08, margin: 0 });
      } else if (!isNum(v)) {
        T(s, [
          { text: '–', options: { fontSize: 17, color: C.mute, breakLine: true } },
          { text: '자체 수요예측 없음', options: { fontSize: 9.5, color: C.gray } },
        ], { x: bx(j), y: ry, w: bw, h: rh, align: 'center', valign: 'middle' });
      } else {
        T(s, r.fmt(v), { x: bx(j), y: ry, w: bw, h: rh, fontSize: 18, bold: b === '가중평균', color: b === '가중평균' ? C.teal : C.text, align: 'center', valign: 'middle' });
      }
    });
  });
  line(s, MX, ry0 + tableH, CW, 0, C.line, { width: 0.75 });
  T(s, '기존 수기는 자체 수요예측이 없어(수요는 가중평균 기준과 같고 안전재고일수만 고정 값) 오차율·편향을 매기지 않습니다 · 방어 = 목표재고가 실제 리드타임 수요를 덮음',
    { x: MX, y: ry0 + tableH + 0.04, w: CW, h: 0.28, fontSize: 10, color: C.gray, valign: 'middle' });
  // 해석(데이터로 조건부 작성)
  const W0 = M['가중평균'], OLD = M['기존수기'];
  const notes = [];
  const others = BASIS3.filter((b) => b !== best);
  const biasBest = bestOf(rowsDef[1]);
  notes.push(`수요 오차율은 ${BLABEL[best]}이 ${pct(M[best].wape)}로 ${others.map((b) => `${BLABEL[b]} ${pct(M[b].wape)}`).join('·')}보다 낮${best === '가중평균' && biasBest === '가중평균' ? `고, 편향도 ${spct(W0.bias)}로 0에 가장 가깝습니다.` : '습니다.'}`);
  const coverTop = bestOf(rowsDef[2]);
  const daysTop = ALL.filter((b) => isNum(M[b].days_median)).reduce((a, b) => (M[b].days_median > M[a].days_median ? b : a));
  if (coverTop === '기존수기' && daysTop === '기존수기') {
    notes.push(`기존 수기는 방어율이 ${pct(OLD.cover)}로 가장 높지만 재고일수 중앙값도 ${days(OLD.days_median)}로 가장 깁니다 → 더 많은 재고로 막은 것입니다.`);
  }
  const higher = ALL.filter((b) => b !== '가중평균' && isNum(M[b].cover) && M[b].cover > W0.cover);
  if (higher.length) {
    const dd = (Math.round(OLD.days_median * 10) - Math.round(W0.days_median * 10)) / 10;
    notes.push(`가중평균 기준의 방어율 ${pct(W0.cover)}는 ${higher.map((b) => `${BLABEL[b]}(${pct(M[b].cover)})`).join('·')}보다 낮습니다${dd > 0 ? `. 대신 재고일수 중앙값은 기존 수기보다 ${dd.toFixed(1)}일 짧습니다` : ''} → 서비스 수준(Z) 조정 과제`);
  }
  notes.push('다음: 품목그룹별로 나눠 보면, 가장 맞는 기준이 그룹마다 같을까?');
  const by = ry0 + tableH + 0.42, bh = 6.9 - by;
  RR(s, MX, by, CW, bh, C.navy);
  T(s, notes.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < notes.length - 1, color: i === notes.length - 1 ? C.sky : C.white, fontSize: i === notes.length - 1 ? 11 : 12.5 } })),
    { x: MX + 0.25, y: by + 0.08, w: CW - 0.5, h: bh - 0.16, valign: 'middle', paraSpaceAfter: 3 });
})();

// ======================================================================= 10. 결과 3: 그룹별 비교
(function groupHeat() {
  // 그룹마다 오차율이 가장 낮은 기준(값이 없으면 null)
  const rowBest = (m) => {
    const c = BASIS3.filter((b) => m[b] && isNum(m[b].wape));
    return c.length ? c.reduce((a, b) => (m[b].wape < m[a].wape ? b : a)) : null;
  };
  // 그룹 A~G 중 채점 값이 있는 그룹만('미분류'는 groupsOrdered에서 이미 빠짐)
  const groups = groupsOrdered.filter((g) => BASIS3.some((b) => isNum(D.metrics[g][b].wape)));
  const winners = groups.map((g) => rowBest(D.metrics[g]));
  const kinds = [...new Set(winners.filter(Boolean))];
  const title = kinds.length > 1
    ? '품목그룹마다 실제 출고에 가장 맞는 기준이 다릅니다'
    : `모든 품목그룹에서 ${BLABEL[kinds[0]]}이 실제 출고에 가장 맞았습니다`;
  const s = slide({
    title,
    sub: `그룹별 수요 오차율 · ${VAL}, 공통 ${M['가중평균'].n_demand} SKU · 행 안에서 오차가 작을수록 진한 청록 · ★ 행 최솟값`,
    tagText: '결과 3', tagKind: 'dark',
  });
  const rowsDef = [{ label: '공통 합산', m: M, all: true }, ...groups.map((g) => ({ label: ga(g), m: D.metrics[g], all: false }))];
  const lx = MX, nx = MX + 1.45, x0 = MX + 2.35, rcw = 2.3;
  const cw = (CW - 2.35 - rcw - 0.15) / BASIS3.length;
  const rx = x0 + cw * BASIS3.length + 0.15;
  const hy = 1.66;
  T(s, '그룹', { x: lx, y: hy, w: 1.4, h: 0.34, fontSize: 12, bold: true, color: C.navy, valign: 'middle' });
  T(s, 'SKU', { x: nx, y: hy, w: 0.8, h: 0.34, fontSize: 12, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  BASIS3.forEach((b, j) => T(s, BLABEL[b], { x: x0 + j * cw, y: hy, w: cw, h: 0.34, fontSize: 12.5, bold: true, color: b === '가중평균' ? C.teal : C.navy, align: 'center', valign: 'middle' }));
  T(s, '가장 맞는 기준', { x: rx, y: hy, w: rcw, h: 0.34, fontSize: 12, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  const PAL = [[C.teal, C.white], ['5FA8A8', C.white], ['A8D3D3', C.navy], [C.tealL, C.navy], ['F3F5F7', C.slate]];
  const rh = 0.4, rs = 0.45, y0 = 2.06;
  rowsDef.forEach((r, i) => {
    const y = y0 + i * rs;
    const vals = BASIS3.map((b) => (r.m[b] && isNum(r.m[b].wape) ? r.m[b].wape : null));
    const ok = vals.filter((v) => v !== null);
    const mn = Math.min(...ok), mx = Math.max(...ok);
    const win = rowBest(r.m);
    const n = r.m['가중평균'].n_demand;
    const small = !r.all && n < MINSKU;
    if (r.all) R(s, lx, y - 0.02, CW, rh + 0.04, 'EEF1F4');
    T(s, small ? `${r.label} *` : r.label, { x: lx + 0.05, y, w: 1.4, h: rh, fontSize: 12.5, bold: true, color: small ? C.gray : C.navy, valign: 'middle' });
    T(s, String(n), { x: nx, y, w: 0.8, h: rh, fontSize: 12, color: C.text, align: 'center', valign: 'middle' });
    vals.forEach((v, j) => {
      const x = x0 + j * cw;
      if (v === null) { box(s, '–', x + 0.03, y, cw - 0.06, rh, { fill: 'F7F8FA', color: C.mute, fontSize: 11, margin: 0 }); return; }
      const t = mx > mn ? (v - mn) / (mx - mn) : 0;
      const lv = t <= 0.001 ? 0 : t < 0.25 ? 1 : t < 0.5 ? 2 : t < 0.75 ? 3 : 4;
      const isW = BASIS3[j] === win;
      box(s, `${isW ? '★ ' : ''}${pct(v, v < 1 ? 1 : 0)}`, x + 0.03, y, cw - 0.06, rh, {
        fill: PAL[lv][0], color: PAL[lv][1], fontSize: 12, bold: isW, margin: 0 });
    });
    T(s, win ? BLABEL[win] : '–', { x: rx, y, w: rcw, h: rh, fontSize: 12.5, bold: true, color: win === '가중평균' ? C.teal : C.navy, align: 'center', valign: 'middle' });
  });
  const ny = y0 + rowsDef.length * rs + 0.02;
  T(s, `* SKU ${MINSKU}개 미만 그룹: 품목이 적어 우연의 영향이 큽니다 · 색은 행 안에서만 비교 · SKU = 가중평균 기준 수요 채점 수 · 공통 합산은 그룹을 모두 합친 값`,
    { x: MX, y: ny, w: CW, h: 0.28, fontSize: 10, color: C.gray, valign: 'middle' });
  // 결론 줄: 가중평균이 가장 맞은 그룹과 다른 기준이 더 맞은 그룹
  const wGroups = groups.filter((g, i) => winners[i] === '가중평균');
  const byOther = BASIS3.filter((b) => b !== '가중평균')
    .map((b) => [b, groups.filter((g, i) => winners[i] === b)])
    .filter(([, gs]) => gs.length);
  const parts = [];
  if (wGroups.length) parts.push(`가중평균 기준이 가장 맞은 그룹은 ${aliasLetters(wGroups)}입니다`);
  if (byOther.length) parts.push(`${byOther.map(([b, gs]) => `${aliasLetters(gs)}는 ${BLABEL[b]}`).join(', ')}이 더 맞았습니다`);
  const msg = byOther.length
    ? `${parts.join('. ')}. 그룹마다 실적과 예측을 섞는 비율이 달라야 한다는 뜻입니다 → 그룹별로 가중치를 보정할 여지가 있습니다.`
    : `${parts.join('. ')}. 현업 산정 가중치가 그룹마다 잘 맞습니다 → AI 보정으로 더 나아지는지 확인합니다.`;
  box(s, [
    { text: '읽는 법  ', options: { bold: true, color: '8FC9C9' } },
    { text: msg, options: { color: C.white, breakLine: true } },
    { text: '다음: AI 보정 — 현업 가중치를 출발점으로 그룹별 w를 데이터로 확인하고 보정합니다', options: { color: C.sky, fontSize: 11, bold: false } },
  ], MX, ny + 0.36, CW, 6.9 - (ny + 0.36), { fill: C.navy, fontSize: 12.5, bold: true, align: 'left', valign: 'middle', margin: 12, round: 0.06 });
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
