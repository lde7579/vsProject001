// 발표자료 공개본 생성: presentation/data.json → out/발표_적정재고일수_공개본_<YYMMDD_HH>.pptx (15장, 비식별)
// 핵심 메시지: 현업이 품목마다 계산할 수 없어 비슷한 값(4종)으로 넣던 일별 안전재고(일)를 품목별로 자동 산출한다.
// 대표 대상은 가중평균 기준, 비교군은 실적기준·FCST기준이다. Grid Search 가중치 최적화는 현업 산정 가중치를 출발점으로 확인·보정하는 단계다.
// 표현 원칙: 사내 시스템명 대신 '기존 수요예측 엔진', 'FCST(예측)'은 처음에 풀어 쓴다, 구현·계획 단계는 색으로만 구분한다.
// 대학원 배포용이다. 품목그룹은 '그룹 A~G'로, 사례 품목은 '품목 X'로 쓰고 절대 수량은 쓰지 않는다.
// 숫자는 모두 data.json(= extract.py가 결과 파일에서 읽은 값)에서 가져온다. 손으로 적은 숫자는 없다.
// 사내본(build.js)은 건드리지 않는다. 양식 도우미는 build.js에서 그대로 복사했다.
// 실행: node presentation/build_public.js [--data presentation/data.json] [--out <경로>] (기본 출력 파일명 끝에 만든 시각 YYMMDD_HH)
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
// 파일명 끝에 만든 시각(년월일_시, 예: 261121_01)을 붙인다(사용자 지시)
const STAMP = (() => {
  const d = new Date(), p = (n) => String(n).padStart(2, '0');
  return `${p(d.getFullYear() % 100)}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}`;
})();
const OUT_PATH = arg('--out', path.join(ROOT, 'out', `발표_적정재고일수_공개본_${STAMP}.pptx`));
const D = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));

// ---------------------------------------------------------------- 양식 상수
const FONT = '맑은 고딕';
const FOOT = '품목별 일별 안전재고 자동 산출 · Grid Search 가중치 최적화 · 테스트 데이터 기반';
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
const METHOD = { '그룹최적': '그룹 최적', '공통최적': '전체 합산', '현재유지': '현업 w 유지' };
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
const wapeDown = wsNew.wape < wsCur.wape; // 검증 구간 오차율이 실제로 줄었을 때만 '개선/감소'라고 쓴다
const coverDown = isNum(wsNew.cover) && isNum(wsCur.cover) && wsNew.cover < wsCur.cover;
// 가중치 최적화 범위의 그룹당 SKU 수(가중치 탐색 choices의 그룹별 채점 SKU 수. 13장 표의 SKU 열과 같은 값)
const maxChoiceN = Math.max(...WS.choices.map((c) => c.n));
const minChoiceN = Math.min(...WS.choices.map((c) => c.n));

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
pres.title = '품목별 일별 안전재고 자동 산출 · Grid Search 가중치 최적화';
pres.subject = '가중평균 기준 산출, 실적·FCST기준 비교, Grid Search·Time-Series Holdout 가중치 최적화 (테스트 데이터)';
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
  // 사용자 지시: 'AI 기반' 같은 막연한 말 대신 누구나 아는 기술 명칭(Grid Search 등)을 쓴다. 단계·현재 위치 같은 표현은 쓰지 않는다.
  T(s, [
    { text: 'Grid Search 가중치 최적화로', options: { fontSize: 32, color: '8FC9C9', breakLine: true } },
    { text: '품목별 일별 안전재고 자동 산출', options: { fontSize: 42, color: C.white } },
  ], { x: 0.9, y: 1.6, w: 11.5, h: 1.95, bold: true, valign: 'middle', lineSpacingMultiple: 1.1 });
  // 부제: 산출 · 비교 · 보정 세 줄기
  const chips = [
    ['산출', '가중평균 기준 안전재고'],
    ['비교', '실적기준 · FCST(예측)기준'],
    ['최적화', 'Grid Search · Holdout 검증'],
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
    sub: '목표 · 세 기준 비교 · 가중치 최적화(Grid Search) · 테스트 데이터셋(비식별 처리)',
    tagText: '요약', tagKind: 'dark',
  });
  const wBest = BYWAPE[0] === '가중평균';
  const cards = [
    ['target', '목표', ['품목마다 자동 산출되는 일별 안전재고값', `현업 수기 ${oldLevels.length}종(${oldLevels.join('·')}일)을 대체`, '대표 대상: 가중평균 기준'], C.teal],
    ['chart', '세 기준 비교', ['가중평균 기준 vs 실적기준 vs FCST기준', `같은 SKU의 ${VAL} 실제 출고 오차율`, BYWAPE.map((b) => `${BSHORT[b]} ${pct(M[b].wape)}`).join(' · ')], C.navy],
    ['search', '가중치 최적화', ['현업 산정 가중치를 출발점으로', 'Grid Search + Time-Series Holdout', `검증 구간 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}`], C.navy],
  ];
  const cw = (CW - 0.4) / 3, cy = 1.6, chh = 1.85;
  cards.forEach(([ic, head, body, col], i) => {
    const x = MX + i * (cw + 0.2);
    RR(s, x, cy, cw, chh, i === 0 ? C.tealL : C.light);
    icon(s, ic, x + 0.25, cy + 0.2, 0.6, col);
    T(s, head, { x: x + 1.0, y: cy + 0.25, w: cw - 1.15, h: 0.5, fontSize: 16.5, bold: true, color: col, valign: 'middle' });
    T(s, lines(body), { x: x + 0.28, y: cy + 0.92, w: cw - 0.4, h: 0.85, fontSize: 12, color: C.text, paraSpaceAfter: 2 });
  });
  // 핵심 수치: 범위가 다른 두 묶음을 머리띠로 나눈다(세 기준 비교 / 가중치 최적화)
  const kw = (CW - 0.6) / 4, ky = 4.0, kh = 1.5;
  const kx = (i) => MX + i * (kw + 0.2);
  const scopes = [
    [0, 2, `세 기준 비교 · ${VAL} · 공통 ${M['가중평균'].n_demand} SKU`, C.navy],
    [3, 3, `가중치 최적화 · 검증 구간 ${HOLD}`, C.teal],
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
    { v: `${pct(wsCur.wape)} → ${pct(wsNew.wape)}`, l: '최적화 후 수요 오차율', n: `현업 w → 보정 w · ${NHOLD} SKU`, c: C.teal },
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
  // 기간이 다른 숫자를 섞지 않도록 가중치 최적화 수치에는 검증 구간을 붙인다
  const coverNote = coverDown ? [`가중치 최적화 후(검증 구간 ${HOLD}) 결품 방어율도 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}로 낮아져 서비스 수준(Z) 조정이 다음 과제`] : [];
  T(s, lines([
    // 가중치 최적화의 출발값과 세 기준 비교 값이 이어진 숫자로 읽히지 않게 범위를 밝힌다
    `오차율(WAPE): 실제 출고 합 대비 절대 오차 합의 비율, 낮을수록 좋음 · 최적화 전 ${pct(wsCur.wape)}는 같은 현업 w를 ${HOLD}만으로 다시 채점한 값(${pct(M['가중평균'].wape)}와 범위가 다름)`,
    '결품 방어율: 목표재고가 실제 리드타임 수요를 덮은 SKU 비율(실제 결품 기록이 아닌 대리 지표), 높을수록 좋음',
    ...coverNote,
  ]), { x: MX, y: 6.3, w: CW, h: 0.68, fontSize: 10, color: C.gray, valign: 'middle', paraSpaceAfter: 2 });
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
    { text: '그룹마다 현업이 자체 공식으로 계산한 값 · Grid Search는 이 값에서 출발해 확인·보정', options: { color: C.text, fontSize: 11.5 } },
  ], { x: MX + 0.85, y: wy + 0.06, w: lw - 1.0, h: wh - 0.12, fontSize: 13, valign: 'middle', paraSpaceAfter: 2 });
  // 연구 질문 세 개
  const qy0 = wy + wh + 0.12;
  label(s, '연구 질문', MX, qy0, lw, C.navy, 14);
  const qs = [
    ['세 기준 중 실제 출고에 가장 가까운 것은?', `같은 SKU를 ${VAL}(${VS.n}주) 실제 출고로 채점`],
    ['품목마다 다른 값이 나오는가?', `현업 수기 ${oldLevels.length}종과 같은 SKU에서 분포 비교`],
    ['가중치 최적화로 더 나아지는가?', `현업 w에서 출발해 탐색, 고를 때 안 쓴 ${HOLD}로 확인`],
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
    [k('선택 구간'), cellOpt(`${SEL} (${SS_.n}주) · w를 고름`, { bold: true, color: C.amber })],
    [k('검증 구간'), cellOpt(`${HOLD} (${HO.n}주) · 선택에 미사용`, { bold: true, color: C.teal })],
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
    ['search', 'Grid Search', `현업 w에서 출발, 후보 ${CANDS.length}개 채점`, true],
    ['calendar', '시계열 검증', '선택에 안 쓴 구간으로 확인', true],
    ['doc', '리포트', '결과표·발표 자료 자동 생성', true],
    ['check', '승인·배포', '사람이 승인한 뒤 값 교체', false],
    ['target', '모니터링', '오차율·방어율 추적, 경보', false],
    ['gear', '재탐색', '매주 가중치 다시 고르기', false],
  ];
  const n = steps.length, gap = 0.17, bw = (CW - gap * (n - 1)) / n, y = 2.25, bh = 2.55;
  const xs = steps.map((_, i) => MX + i * (bw + gap));
  steps.forEach(([ic, a, b, done], i) => {
    const x = xs[i];
    if (done) RR(s, x, y, bw, bh, C.tealL, { line: { color: C.teal, width: 1.25 } });
    else box(s, '', x, y, bw, bh, { fill: C.white, dash: true, lineColor: C.mute, round: 0.08 });
    icon(s, ic, x + (bw - 0.62) / 2, y + 0.28, 0.62, done ? C.teal : C.mute);
    T(s, a, { x: x + 0.03, y: y + 1.02, w: bw - 0.06, h: 0.4, fontSize: 12.5, bold: true, color: done ? C.navy : C.gray, align: 'center', valign: 'middle' });
    T(s, b, { x: x + 0.08, y: y + 1.48, w: bw - 0.16, h: 0.9, fontSize: 10.5, color: done ? C.text : C.gray, align: 'center' });
    if (i < n - 1) arrow(s, x + bw + 0.01, y + bh / 2, gap - 0.02, 0, done && steps[i + 1][3] ? C.teal : C.mute, { width: 1.5 });
  });
  // 되먹임(계획): 재탐색 → Grid Search
  const fy = y + bh + 0.3, from = xs[n - 1] + bw / 2, to = xs[3] + bw / 2;
  line(s, from, y + bh, 0, fy - y - bh, C.mute, { width: 1.5, dashType: 'dash' });
  line(s, to, fy, from - to, 0, C.mute, { width: 1.5, dashType: 'dash' });
  arrow(s, to, fy, 0, -(fy - y - bh) + 0.02, C.mute, { width: 1.5 });
  box(s, '되먹임(계획): 다음 주 데이터로 Grid Search를 다시 돌림', to + 1.0, fy - 0.15, 4.6, 0.3, { fill: C.white, lineColor: C.white, color: C.gray, fontSize: 11, bold: true, margin: 0 });
  // 구현 범위(계획 단계는 위 점선 상자와 범례로만 조용히 구분한다)
  const ry = fy + 0.5;
  box(s, [
    { text: '구현  ', options: { bold: true, color: '8FC9C9' } },
    { text: `데이터 검사 · 세 기준 산출 · 재고등급 · Grid Search · Holdout 검증 · 리포트 + 자동 테스트 ${int(TESTS)}개 · 버전 관리(git)`, options: { color: C.white } },
  ], MX, ry, CW, 0.56, { fill: C.navy, fontSize: 12, bold: true, align: 'left', margin: 12, round: 0.06 });
  T(s, '검증 구간의 실제 출고는 채점에만 쓰고 계산 입력에는 쓰지 않습니다.', { x: MX, y: ry + 0.7, w: CW, h: 0.28, fontSize: 10.5, color: C.gray, valign: 'middle' });
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
  T(s, `규칙 기반 재고등급이 기존 등급 ${G.grade_match}/${G.total} 일치`, { x: MX + 1.2, y: gy + 0.66, w: 6.4, h: 0.4, fontSize: 14, bold: true, color: C.teal, valign: 'middle' });
  T(s, '단가·회전율·리드타임·품절 이력을 점수로 합산해 등급을 매기고, 등급이 높을수록 더 높은 서비스 수준을 씁니다',
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
  // 비교 범위는 세 기준뿐이다(기존 수기는 값이 4종이라 배수가 작아도 품목별 값이 아니다)
  if (narrow === '가중평균') notes.push(`세 기준 중에서는 가중평균 기준이 P90÷P10 ${ratio(narrow).toFixed(1)}배로, ${others.map((b) => `${BLABEL[b]} ${ratio(b).toFixed(1)}배`).join('·')}보다 고르게 퍼져 안정적입니다.`);
  else notes.push(`세 기준 중 P90÷P10이 가장 작은 기준은 ${BLABEL[narrow]}(${ratio(narrow).toFixed(1)}배)입니다.`);
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
    sub: `${ga(c.group)} · 재고등급 ${c.grade} · 리드타임 ${c.lt}일 · 현업 w 기준(최적화 전) · 수량 없이 일수로`,
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
  // '유리한 사례' 표시는 가중평균 기준이 실제로 방어한 경우에만 붙인다
  const favor = c.new_cover === '방어' ? ' (가중평균에 유리한 사례)' : '';
  T(s, `선정 규칙: ${c.rule}${favor}`,{ x: MX, y: y + ph + 0.06, w: CW, h: 0.3, fontSize: 10, color: C.gray, valign: 'middle' });
  // 아래: 요약과 반대 사례
  const covs = rowsDef.map((r) => r[3]);
  const allOk = covs.every((v) => v === '방어');
  // 실제 출고가 0이면 차이 비율을 쓰지 않는다
  const fdiff = isNum(c.actual_week) && c.actual_week > 0 ? c.g / c.actual_week - 1 : null;
  const fnote = isNum(fdiff) ? ` 단, 가중평균 수요는 실제보다 ${pct(Math.abs(fdiff), 0)} ${fdiff >= 0 ? '높았' : '낮았'}습니다.` : '';
  const msg = allOk
    ? `네 기준 모두 방어했고, 실제 기준 재고일수는 기존 수기 ${f1(c.old_actual_days)}일 → 가중평균 기준 ${f1(c.new_actual_days)}일입니다.${fnote}`
    : `방어 결과: ${rowsDef.map((r) => `${BLABEL[r[0]]} ${cv(r[3])}`).join(' · ')}`;
  box(s, msg, MX, 5.62, CW, 0.58, { fill: C.navy, color: C.white, fontSize: 12, bold: true, align: 'left', valign: 'middle', margin: 10, round: 0.06 });
  const ccCovs = [cc.old_cover, cc.new_cover, cc.act_cover, cc.fc_cover];
  const ccAll = ccCovs.every((v) => v === '결품');
  // 가중평균 수요가 0이면 배수를 쓰지 않는다. '방어 실패'는 가중평균 기준이 실제로 미방어일 때만 쓴다.
  const ccRatio = isNum(cc.g) && cc.g > 0 && isNum(cc.actual_week) ? `실제 출고가 가중평균 수요의 ${(cc.actual_week / cc.g).toFixed(1)}배여서 ` : '';
  const ccOut = ccAll ? `${ccRatio}네 기준 모두 방어에 실패했습니다.`
    : cc.new_cover === '결품' ? `${ccRatio}가중평균 기준은 방어에 실패했습니다.`
      : `방어 결과는 ${[['기존수기', cc.old_cover], ['가중평균', cc.new_cover], ['실적기준', cc.act_cover], ['FCST기준', cc.fc_cover]].map(([b, v]) => `${BLABEL[b]} ${cv(v)}`).join(' · ')}였습니다.`;
  box(s, [
    { text: '반대 사례  ', options: { bold: true, color: C.white } },
    { text: `${ga(cc.group)}의 한 품목은 기존 ${cc.old_ss_days}일 → 가중평균 기준 ${f1(cc.s)}일이었고, ${ccOut}`, options: { color: C.white } },
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
    // 방어율이 더 높은 기준마다 재고일수를 따로 견준다. '짧다'는 화면 반올림 값으로 실제 짧을 때만 쓴다.
    const dshort = (b) => (isNum(M[b].days_median) && isNum(W0.days_median)
      ? (Math.round(M[b].days_median * 10) - Math.round(W0.days_median * 10)) / 10 : null);
    const shorter = higher.filter((b) => isNum(dshort(b)) && dshort(b) > 0);
    let t = `가중평균 기준의 방어율 ${pct(W0.cover)}는 ${higher.map((b) => `${BLABEL[b]}(${pct(M[b].cover)})`).join('·')}보다 낮습니다`;
    if (shorter.length) t += `. 대신 재고일수 중앙값은 ${shorter.map((b) => `${BLABEL[b]}보다 ${dshort(b).toFixed(1)}일`).join('·')} 짧습니다`;
    // 실적기준이 방어율이 더 높을 때는, 그래도 가중평균 기준을 대표로 두는 까닭을 data로 적는다(편향이 더 크고 극단값이 나옴)
    if (higher.includes('실적기준')) {
      const A = M['실적기준'];
      const biasMore = isNum(A.bias) && isNum(W0.bias) && Math.abs(A.bias) > Math.abs(W0.bias);
      const extremeMore = isNum(SB['실적기준'].max) && isNum(SB['가중평균'].max) && SB['실적기준'].max > SB['가중평균'].max;
      const biasT = `수요를 ${spct(A.bias)} ${A.bias >= 0 ? '크게' : '작게'} 잡았`;
      const extT = `극단값(최대 ${f1(SB['실적기준'].max)}일)이 나옵니다`;
      const why = biasMore && extremeMore ? `${biasT}고(편향) ${extT}` : biasMore ? `${biasT}습니다(편향)` : extremeMore ? extT : '';
      if (why) t += `. 실적기준은 방어율이 높지만 ${why}`;
    }
    notes.push(`${t} → 서비스 수준(Z) 조정 과제`);
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
  // 결론과 제목은 SKU가 충분한(MINSKU 이상) 그룹만으로 판단한다. 소표본 그룹은 표에만 두고 참고로 쓴다.
  const big = groupsOrdered.filter((g) => D.metrics[g]['가중평균'].n_demand >= MINSKU && rowBest(D.metrics[g]));
  const bigWin = big.map((g) => rowBest(D.metrics[g]));
  const kinds = [...new Set(bigWin)];
  const hasSmall = groups.some((g) => D.metrics[g]['가중평균'].n_demand < MINSKU);
  const title = kinds.length > 1
    ? '품목그룹마다 실제 출고에 가장 맞는 기준이 다릅니다'
    : kinds.length === 1
      ? `SKU가 충분한 그룹은 모두 ${BLABEL[kinds[0]]}이 실제 출고에 가장 맞았습니다`
      : '품목그룹별 세 기준의 수요 오차율';
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
    const mn = ok.length ? Math.min(...ok) : 0, mx = ok.length ? Math.max(...ok) : 0;
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
  // 결론 줄: SKU가 충분한 그룹의 승자만으로 조건부 작성
  const SIDE = { 실적기준: '실적', FCST기준: '예측' };
  let msg;
  if (kinds.length > 1) {
    msg = `SKU가 충분한 ${aliasLetters(big)}에서도 가장 맞는 기준이 다릅니다(${big.map((g, i) => `${ga(g).slice(-1)} ${BLABEL[bigWin[i]]}`).join(' · ')}) → 실적과 예측 중 어느 쪽에 무게를 둘지가 그룹마다 달라, 그룹별로 가중치를 보정할 여지가 있습니다.`;
  } else if (kinds.length === 1 && kinds[0] === '가중평균') {
    msg = `SKU가 충분한 ${aliasLetters(big)}는 모두 가중평균 기준이 가장 맞았습니다 → 현업 산정 가중치가 잘 맞습니다. Grid Search로 더 나아지는지 확인합니다.`;
  } else if (kinds.length === 1) {
    msg = `SKU가 충분한 ${aliasLetters(big)}는 모두 ${BLABEL[kinds[0]]}이 가장 맞았습니다 → 현업 가중치보다 ${SIDE[kinds[0]]} 쪽에 무게를 더 둘 여지가 있습니다.`;
  } else {
    msg = `SKU ${MINSKU}개 이상인 그룹이 없어 그룹별 결론은 내리지 않습니다(표는 참고만).`;
  }
  if (kinds.length && hasSmall) msg += ' (* 소표본 그룹은 참고만)';
  box(s, [
    { text: '읽는 법  ', options: { bold: true, color: '8FC9C9' } },
    { text: msg, options: { color: C.white, breakLine: true } },
    { text: '다음: Grid Search — 현업 가중치를 출발점으로 그룹별 w를 데이터로 확인하고 보정합니다', options: { color: C.sky, fontSize: 11, bold: false } },
  ], MX, ny + 0.36, CW, 6.9 - (ny + 0.36), { fill: C.navy, fontSize: 12.5, bold: true, align: 'left', valign: 'middle', margin: 12, round: 0.06 });
})();

// ======================================================================= 11. 가중치 최적화: 기법과 탐색 설계
// 기존 '기법'과 '탐색 설계' 두 장을 합쳤다. 출발점은 현업이 자체 공식으로 산정한 w다.
(function aiMethod() {
  const s = slide({
    title: `현업 w에서 출발해 앞 ${SS_.n}주로 고르고, 뒤 ${HO.n}주로 확인합니다`,
    sub: 'Hyperparameter Tuning 방식(Grid Search · Time-Series Holdout · Pooling)을 가중치 w에 적용',
    tagText: 'Grid Search', tagKind: 'teal',
  });
  // 위: 주차 타임라인
  const LAST = Math.max(FCST_END, HO.b);
  const x0 = MX + 0.2, x1 = W - MX - 0.2, wk = (x1 - x0) / LAST, X = (w) => x0 + (w - 1) * wk;
  const y = 1.92;
  T(s, '주차', { x: MX - 0.05, y: y - 0.3, w: 1, h: 0.26, fontSize: 10.5, color: C.gray });
  box(s, `출고 실적 ${HIST} (계산 입력)`, X(HS.a), y, X(HS.b + 1) - X(HS.a) - 0.03, 0.46, { fill: C.mute, color: C.white, bold: true, fontSize: 12, margin: 0 });
  if (SS_.a > HS.b + 1) box(s, '미사용', X(HS.b + 1), y, X(SS_.a) - X(HS.b + 1) - 0.03, 0.46, { fill: 'E6EBF0', color: C.gray, fontSize: 10, margin: 0 });
  box(s, lines(['선택 구간', SEL]), X(SS_.a), y, X(SS_.b + 1) - X(SS_.a) - 0.03, 0.46, { fill: C.amber, color: C.white, bold: true, fontSize: 10, margin: 0 });
  box(s, lines(['검증 구간', HOLD]), X(HO.a), y, X(HO.b + 1) - X(HO.a) - 0.03, 0.46, { fill: C.teal, color: C.white, bold: true, fontSize: 10, margin: 0 });
  box(s, `예측(${FW}주) ${wr(`W${SS_.a}~W${FCST_END}`)} (계산 입력)`, X(SS_.a), y + 0.54, X(FCST_END + 1) - X(SS_.a), 0.32, { fill: 'EFE3CB', color: '8A6420', bold: true, fontSize: 10, margin: 0 });
  line(s, X(SS_.a) - 0.015, y - 0.12, 0, 1.0, C.red, { width: 1.5, dashType: 'dash' });
  T(s, '계산 시점', { x: X(SS_.a) - 1.75, y: y + 0.54, w: 1.7, h: 0.32, fontSize: 10.5, bold: true, color: C.red, align: 'right', valign: 'middle' });
  [...new Set([HS.a, 10, 20, HS.b, SS_.a, HO.a, HO.b, FCST_END])].forEach((w) => T(s, `W${String(w).padStart(2, '0')}`, { x: X(w) - 0.3 + wk / 2, y: y + 0.9, w: 0.6, h: 0.22, fontSize: 9.5, color: C.gray, align: 'center' }));
  // 가운데: 핵심 기법 카드 3개
  const cards = [
    ['search', 'Grid Search', '전수 후보 탐색',
      `실적 가중치 w ${wfmt(CANDS[0])}~${wfmt(CANDS[CANDS.length - 1])}, ${wfmt(CSTEP)} 간격 후보 ${CANDS.length}개를 품목그룹마다 모두 채점`,
      `탐색 공간 ${man(SPACE)} 조합`],
    ['calendar', 'Time-Series Holdout', '시계열 분할 검증',
      `선택 구간(${SEL}) 오차율이 가장 낮은 w를 고르고, 검증 구간(${HOLD})은 확인에만 씀`,
      '미래 정보가 선택에 안 섞임'],
    ['lock', 'Regularization · Pooling', '과적합 방지',
      `선택 구간 SKU ${MINSKU}개 미만 그룹은 전체를 합쳐 고른 w를 씀`,
      `작은 그룹 ${SMALL.length}개에 적용`],
  ];
  const cy = 3.22, ch = 2.28, rbw = 3.55;
  const cw = (CW - rbw - 0.3 - 0.4) / 3;
  cards.forEach(([ic, en, ko, body, ev], i) => {
    const x = MX + i * (cw + 0.2);
    RR(s, x, cy, cw, ch, C.light);
    icon(s, ic, x + 0.18, cy + 0.18, 0.56, C.teal);
    T(s, en, { x: x + 0.86, y: cy + 0.16, w: cw - 0.96, h: 0.26, fontSize: 10, color: C.slate, valign: 'middle' });
    T(s, ko, { x: x + 0.86, y: cy + 0.42, w: cw - 0.96, h: 0.36, fontSize: 15, bold: true, color: C.navy, valign: 'middle' });
    T(s, body, { x: x + 0.2, y: cy + 0.9, w: cw - 0.4, h: 0.95, fontSize: 11.5, color: C.text });
    T(s, ev, { x: x + 0.2, y: cy + ch - 0.42, w: cw - 0.4, h: 0.3, fontSize: 11.5, bold: true, color: C.teal, valign: 'middle' });
  });
  // 오른쪽: 77번 계산 = 1,949만 조합
  const tx = MX + 3 * cw + 0.4 + 0.3, tw = W - MX - tx;
  RR(s, tx, cy, tw, ch, C.navy);
  T(s, `왜 ${CANDS.length * NG}번 계산으로 전체 최적을 얻나`, { x: tx + 0.2, y: cy + 0.08, w: tw - 0.4, h: 0.5, fontSize: 13, bold: true, color: C.white, valign: 'middle' });
  const bw2 = (tw - 0.55) / 2;
  box(s, [{ text: `모든 조합 ${CANDS.length}의 ${NG}제곱`, options: { fontSize: 10, color: C.sky, breakLine: true } }, { text: `= ${man(SPACE)}`, options: { fontSize: 14, bold: true, color: C.white } }],
    tx + 0.2, cy + 0.64, bw2, 0.62, { fill: C.navy2, round: 0.06, margin: 2 });
  box(s, [{ text: `실제 계산 ${CANDS.length} × ${NG}`, options: { fontSize: 10, color: '8FC9C9', breakLine: true } }, { text: `= ${CANDS.length * NG}번`, options: { fontSize: 14, bold: true, color: C.white } }],
    tx + 0.35 + bw2, cy + 0.64, bw2, 0.62, { fill: C.teal, round: 0.06, margin: 2 });
  T(s, '전체 오차는 그룹별 오차의 합이고, 그룹 오차는 그 그룹의 w에만 달려 있어 그룹마다 따로 고르면 모든 조합을 비교한 답과 같습니다.',
    { x: tx + 0.2, y: cy + 1.34, w: tw - 0.4, h: ch - 1.42, fontSize: 10.5, color: C.white, valign: 'middle' });
  // 아래: 출발점 = 현업 w, 받쳐 주는 기법 한 줄
  const tieCur = /현재 가중치에 가까운/.test(String(SRCH['동점 규칙'] || ''));
  box(s, [
    { text: '출발점 = 현업 w  ', options: { bold: true, color: C.teal } },
    { text: `현업이 그룹별 자체 공식으로 산정한 w(${WRANGE})를 기준선으로 두고, 후보를 모두 채점해 더 맞는 w가 있는지 데이터로 확인·보정합니다${tieCur ? '(오차가 같으면 현업 w 쪽)' : ''}.`, options: { color: C.navy } },
  ], MX, 5.68, CW, 0.62, { fill: C.tealL, fontSize: 12.5, bold: true, align: 'left', margin: 12, round: 0.06 });
  T(s, [
    { text: '받쳐 주는 기법  ', options: { bold: true, color: C.slate } },
    { text: `통계적 안전재고 모형(Z·σ·√LT) · 검사부터 리포트까지 한 번에 재현하는 파이프라인(자동 테스트 ${int(TESTS)}개)`, options: { color: C.gray } },
  ], { x: MX, y: 6.42, w: CW, h: 0.42, fontSize: 10.5, valign: 'middle' });
})();

// ======================================================================= 12. 탐색 히트맵
(function wsHeatmap() {
  const s = slide({
    title: `후보 ${CANDS.length}개를 모두 채점해 그룹별 최적 w를 골랐습니다`,
    sub: `선택 구간 ${SEL}의 가중평균 수요 오차율 · 행 안에서 오차가 작을수록 진한 청록 · ★ 선택 · ○ 현업 w`,
    tagText: 'Grid Search', tagKind: 'teal',
  });
  const rowsDef = [{ key: '전체', label: '전체 합산', choice: null }, ...choicesSorted.map((c) => ({ key: c.group, label: ga(c.group), choice: c }))];
  const pooled = WS.candidates.filter((c) => c.scope === '전체' && isNum(c.wape));
  const pooledBest = pooled.reduce((a, b) => (b.wape < a.wape ? b : a)).w;
  const cand = (scope, w) => WS.candidates.find((c) => c.scope === scope && c.w === w);
  const lx = MX, nx = MX + 1.2, x0 = MX + 1.9, rcw = 1.6;
  const cw = (CW - 1.9 - rcw - 0.1) / CANDS.length;
  const rx = x0 + cw * CANDS.length + 0.1;
  T(s, '실적 가중치 w  (예측 가중치 = 1 − w)', { x: x0, y: 1.62, w: cw * CANDS.length, h: 0.3, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  T(s, '그룹', { x: lx, y: 1.95, w: 1.15, h: 0.32, fontSize: 11.5, bold: true, color: C.navy, valign: 'middle' });
  T(s, 'SKU', { x: nx, y: 1.95, w: 0.65, h: 0.32, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  CANDS.forEach((w, j) => T(s, wfmt(w), { x: x0 + j * cw, y: 1.95, w: cw, h: 0.32, fontSize: 11.5, bold: true, color: C.slate, align: 'center', valign: 'middle' }));
  T(s, '현업 → 선택', { x: rx, y: 1.95, w: rcw, h: 0.32, fontSize: 11.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  const PAL = [[C.teal, C.white], ['5FA8A8', C.white], ['A8D3D3', C.navy], [C.tealL, C.navy], ['F3F5F7', C.slate]];
  const rh = 0.4, rs = 0.44, y0 = 2.34;
  rowsDef.forEach((r, i) => {
    const y = y0 + i * rs;
    const cells = CANDS.map((w) => cand(r.key, w));
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
  // 읽는 법: SKU가 충분한 그룹마다 현업 w와 선택 w의 선택 구간 오차율 차이를 data로 계산해 적는다
  // NEAR_PP: '현업 w가 최적 근처'로 볼 선택 구간 오차율 절대 차이(%p) 기준. 화면에도 기준을 밝힌다.
  // 근거: 통계 검정이 아닌 발표용 판단 기준이다. 선택 구간 오차율(WAPE)의 절대 차이가 1%p 이하이면
  //       그룹 오차율 수준(이번 데이터 12~68%) 대비 작은 차이라 현업 w를 '최적 근처'로 본다.
  // 차이는 화면에 보이는 반올림 값끼리 뺀다.
  const NEAR_PP = 1.0;
  const gaps = choicesSorted.filter((c) => c.method === '그룹최적').map((c) => {
    const a = cand(c.group, c.current), z = cand(c.group, c.chosen);
    const wa = a && isNum(a.wape) ? a.wape : c.wape_current, wz = z && isNum(z.wape) ? z.wape : c.wape_chosen;
    const d = isNum(wa) && isNum(wz) ? (Math.round(wa * 1000) - Math.round(wz * 1000)) / 10 : null;
    return { g: c.group, d };
  }).filter((r) => isNum(r.d));
  const near = gaps.filter((r) => r.d <= NEAR_PP), far = gaps.filter((r) => r.d > NEAR_PP);
  const gtxt = (arr) => arr.map((r) => `${ga(r.g)} ${r.d.toFixed(1)}%p`).join(' · ');
  let msg;
  if (near.length && far.length) msg = `SKU ${MINSKU}개 이상 그룹 중 ${gtxt(near)} 차이로 현업 w가 이미 최적 근처(차이 ${NEAR_PP.toFixed(1)}%p 이하)였고, ${gtxt(far)} 차이로 보정 여지가 컸습니다.`;
  else if (near.length) msg = `SKU ${MINSKU}개 이상 그룹은 ${gtxt(near)} 차이로 모두 현업 w가 최적 근처(차이 ${NEAR_PP.toFixed(1)}%p 이하)였습니다.`;
  else if (far.length) msg = `SKU ${MINSKU}개 이상 그룹은 ${gtxt(far)} 차이로 보정 여지가 있었습니다(최적 근처 기준: 차이 ${NEAR_PP.toFixed(1)}%p 이하).`;
  else msg = '그룹마다 오차가 가장 낮은 w가 다릅니다.';
  msg += ` 단, ${SS_.n}주로 고른 값이라 다음 장의 검증 구간에서 다시 확인합니다.`;
  box(s, [
    { text: '읽는 법  ', options: { bold: true, color: '8FC9C9' } },
    { text: msg, options: { color: C.white } },
  ], MX, ny + 0.4, CW, 6.9 - (ny + 0.4), { fill: C.navy, fontSize: 12.5, bold: true, align: 'left', margin: 12, round: 0.06 });
})();

// ======================================================================= 13. 가중치 최적화 결과(검증 구간)
// 기존 '결과 KPI'와 '그룹별 감소 몫' 두 장을 합쳤다. 나빠진 지표(결품 방어율)도 함께 보인다.
(function wsResult() {
  const s = slide({
    title: `Grid Search 최적화 후 검증 구간 오차율 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}`,
    sub: `검증 구간 ${HOLD}(${HO.n}주, 고를 때 미사용) · ${NHOLD} SKU · 현업 w → 보정 w · 감소 몫 = 전체 오차 감소분 중 그 그룹의 몫`,
    tagText: '최적화 결과', tagKind: 'dark',
  });
  // 왼쪽: 핵심 지표 3개(세로)
  const lw = 4.25, y = 1.62, kh = 0.95, kg = 0.1;
  const kp = [
    ['가중평균 수요 오차율', '낮을수록 좋음', 'wape', true, (v) => pct(v)],
    ['결품 방어율', '높을수록 좋음 · 대리 지표', 'cover', false, (v) => pct(v)],
    ['실제 기준 재고일수 중앙값', '참고', 'days_median', null, (v) => f1(v) + '일'],
  ];
  kp.forEach(([l, how, key, lowGood, f], i) => {
    const ky = y + i * (kh + kg);
    const a = wsCur[key], z = wsNew[key];
    let col = C.navy;
    if (lowGood !== null && isNum(a) && isNum(z) && a !== z) col = ((z < a) === lowGood) ? C.teal : C.red;
    const bad = col === C.red;
    RR(s, MX, ky, lw, kh, bad ? C.redL : i === 0 ? C.tealL : C.light);
    T(s, [{ text: l, options: { bold: true, color: C.navy, fontSize: 12.5 } }, { text: `  ${how}`, options: { color: C.gray, fontSize: 9.5 } }],
      { x: MX + 0.2, y: ky + 0.08, w: lw - (bad ? 1.1 : 0.4), h: 0.32, valign: 'middle' });
    if (bad) box(s, '나빠짐', MX + lw - 0.92, ky + 0.1, 0.74, 0.28, { round: 0.14, fill: C.red, color: C.white, fontSize: 10, bold: true, margin: 0 });
    T(s, [{ text: `${f(a)} → `, options: { color: C.slate } }, { text: f(z), options: { color: col } }],
      { x: MX + 0.2, y: ky + 0.44, w: lw - 1.5, h: 0.5, fontSize: 22, bold: true, valign: 'middle' });
    const diff = key === 'days_median'
      ? `${z - a >= 0 ? '+' : '−'}${f1(Math.abs(z - a))}일`
      : `${z >= a ? '+' : '−'}${pp(Math.max(a, z), Math.min(a, z))}`;
    T(s, diff, { x: MX + lw - 1.3, y: ky + 0.44, w: 1.1, h: 0.5, fontSize: 15, bold: true, color: col, align: 'right', valign: 'middle' });
  });
  // 오른쪽: 그룹별 선택과 검증 구간 오차율, 감소 몫
  const tx = MX + lw + 0.25, tw = W - MX - tx, rh = 0.37;
  const rows = [['그룹', 'SKU', '현업 w', '선택 w', '방식', '검증 구간 오차율 (전→후)', '감소 몫']];
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
      cellOpt(isNum(sh) ? spct0(sh) : '–', { bold: isNum(sh) && sh >= 0.1, color: isNum(sh) && sh < 0 ? C.red : isNum(sh) && sh >= 0.1 ? C.teal : C.text })]);
    if (hv) {
      if (hn.wape < hc.wape) up.push(`${ga(c.group)} ${pct(hc.wape)}→${pct(hn.wape)}`);
      else if (hn.wape > hc.wape) down.push(`${ga(c.group)} 오차율 ${pct(hc.wape)}→${pct(hn.wape)}`);
    }
    if (hc && hn && isNum(hc.cover) && isNum(hn.cover) && hn.cover < hc.cover) down.push(`${ga(c.group)} 방어율 ${pct(hc.cover)}→${pct(hn.cover)}`);
  });
  if (coverDown) down.unshift(`전체 방어율 ${pct(wsCur.cover)}→${pct(wsNew.cover)}`);
  table(s, rows, { x: tx, y, w: tw, colW: [0.95, 0.6, 0.85, 0.85, 1.1, tw - 0.95 - 0.6 - 0.85 - 0.85 - 1.1 - 0.95, 0.95], rowH: rh, fontSize: 11.5, headSize: 11, align: 'center' });
  const tEnd = y + rh * rows.length;
  T(s, `SKU = 선택 구간 채점 수 · '전체 합산' = SKU ${MINSKU}개 미만이라 전체를 합쳐 고른 w`,
    { x: tx, y: tEnd + 0.04, w: tw, h: 0.26, fontSize: 10, color: C.gray, valign: 'middle' });
  // 오차율이 100%를 넘는 그룹: 실제 출고 합이 아주 작아 비율이 크게 나온다(수량은 쓰지 않음)
  const huge = choicesSorted.filter((c) => [wsRow('검증', c.group, '현재', '가중평균'), wsRow('검증', c.group, '탐색', '가중평균')]
    .some((r) => r && isNum(r.wape) && r.wape >= 1)).map((c) => c.group);
  let fnEnd = tEnd + 0.3;
  if (huge.length) {
    T(s, `${aliasLetters(huge)}: 실제 출고 합이 아주 작으면 작은 오차도 오차율(WAPE)이 매우 크게 나옵니다`,
      { x: tx, y: fnEnd, w: tw, h: 0.26, fontSize: 10, color: C.gray, valign: 'middle' });
    fnEnd += 0.26;
  }
  // 왼쪽 아래: 출발점 오차율이 세 기준 비교(9장)의 가중평균 오차율과 다른 까닭(범위가 다름)
  const kEnd = y + 3 * kh + 2 * kg;
  T(s, lines([
    `${pct(wsCur.wape)} = 같은 현업 w 가중평균 기준을 다시 채점한 값`,
    `채점 범위: 검증 구간 ${HOLD} · ${NHOLD} SKU`,
    `세 기준 비교 ${pct(M['가중평균'].wape)}는 ${VAL} · 공통 ${M['가중평균'].n_demand} SKU (범위가 다름)`,
  ]), { x: MX, y: kEnd + 0.06, w: lw, h: 0.56, fontSize: 10, color: C.gray, valign: 'top' });
  // 아래: 좋아진 것 / 나빠진 것(숨기지 않음)
  const pair = (arr) => arr.reduce((acc, t, i) => (i % 2 ? (acc[acc.length - 1] += ' · ' + t, acc) : (acc.push(t), acc)), []);
  const by = Math.max(fnEnd + 0.1, kEnd + 0.7), bh = 6.9 - by, hw = (CW - 0.2) / 2;
  RR(s, MX, by, hw, bh, C.tealL);
  const gain = [];
  if (topContribText) gain.push(`감소 몫: ${topContribText}에 몰림`);
  gain.push(...pair(up));
  T(s, [{ text: '좋아진 것 (검증 구간 오차율)', options: { bold: true, color: C.teal, breakLine: true, fontSize: 12.5 } }, ...lines(gain.length ? gain : ['없음'], { color: C.text })],
    { x: MX + 0.2, y: by + 0.1, w: hw - 0.4, h: bh - 0.2, fontSize: 11.5, valign: 'top', paraSpaceAfter: 3 });
  RR(s, MX + hw + 0.2, by, hw, bh, C.redL);
  const loss = pair(down);
  if (coverDown && isNum(wsCur.bias) && isNum(wsNew.bias)) loss.push(`수요 편향 ${spct(wsCur.bias)} → ${spct(wsNew.bias)}(${wsNew.bias < wsCur.bias ? '과소' : '과대'} 쪽) → 서비스 수준(Z) 조정 과제`);
  T(s, [{ text: '나빠진 것', options: { bold: true, color: C.red, breakLine: true, fontSize: 12.5 } }, ...lines(loss.length ? loss : ['없음'], { color: C.text })],
    { x: MX + hw + 0.4, y: by + 0.1, w: hw - 0.4, h: bh - 0.2, fontSize: 11.5, valign: 'top', paraSpaceAfter: 3 });
})();

// ======================================================================= 14. 해석과 한계
(function limits() {
  const W0 = M['가중평균'], OLDM = M['기존수기'];
  const wBest = BYWAPE[0] === '가중평균';
  const s = slide({
    title: wapeDown ? '품목별 값과 오차율 개선은 확인했지만, 유의성 검정은 아직 없습니다' : '품목별 값은 확인했지만, 유의성 검정은 아직 없습니다',
    sub: '주장은 범위(기간·SKU)를 밝힌 수치로만 하고, 나빠진 지표도 함께 적었습니다',
    tagText: '해석과 한계', tagKind: 'light',
  });
  const lw = 5.5, y = 1.62, h = 5.2;
  RR(s, MX, y, lw, h, C.navy);
  T(s, '말할 수 있는 것', { x: MX + 0.3, y: y + 0.12, w: lw - 0.6, h: 0.4, fontSize: 16, bold: true, color: C.white, valign: 'middle' });
  const OLD = SB['기존수기'], WB = SB['가중평균'];
  const can = [
    // 설명은 줄 단위로 나눠 숫자(예: 30.0 → 10.6일) 중간에서 줄이 바뀌지 않게 한다
    ['품목별 일별 안전재고 자동 산출', [`현업 수기 ${OLD.distinct}종 → 가중평균 기준 ${WB.distinct}개 값 (${SB.n} SKU)`, `중앙값 ${f1(OLD.median)} → ${f1(WB.median)}일`]],
    [wBest ? '세 기준 중 가중평균 기준 오차율 최저' : `세 기준 중 ${BLABEL[BYWAPE[0]]} 오차율 최저`,
      [`${VAL} 공통 ${W0.n_demand} SKU`, BYWAPE.map((b) => `${BSHORT[b]} ${pct(M[b].wape)}`).join(' · ')]],
    [wapeDown ? 'Grid Search로 검증 오차 감소' : 'Grid Search 후 검증 오차율', [`현업 w → 보정 w · 선택에 안 쓴 ${HOLD} (${NHOLD} SKU)`, `${pct(wsCur.wape)} → ${pct(wsNew.wape)}`]],
    ['등급 규칙 일치', [`규칙 기반 재고등급이 기존 등급 ${D.grades.grade_match}/${D.grades.total}과 같음`]],
  ];
  const ih = (h - 0.7) / can.length;
  can.forEach(([a, b], i) => {
    const iy = y + 0.66 + i * ih;
    num(s, i + 1, MX + 0.3, iy + 0.04, 0.42, i === can.length - 1 ? C.slate : C.teal, C.white, 13);
    T(s, a, { x: MX + 0.9, y: iy, w: lw - 1.1, h: 0.46, fontSize: 14, bold: true, color: C.white, valign: 'middle' });
    T(s, lines(b), { x: MX + 0.9, y: iy + 0.48, w: lw - 1.1, h: ih - 0.56, fontSize: 11.5, color: 'D5DEE7', paraSpaceAfter: 2 });
  });
  const rx = MX + lw + 0.3, rw = W - MX - rx;
  T(s, '아직 말할 수 없는 것 (한계)', { x: rx, y, w: rw, h: 0.4, fontSize: 16, bold: true, color: C.amber, valign: 'middle' });
  const lowCover = BASIS3.filter((b) => b !== '가중평균' && isNum(M[b].cover) && M[b].cover > W0.cover);
  const higherOld = isNum(OLDM.cover) && OLDM.cover > W0.cover;
  const lim = [
    ['chart', '같은 기간 채점 · 유의성 검정 없음', `세 기준은 같은 ${VAL}에서 채점했고, Grid Search 결과는 검증 ${NHOLD} SKU(그룹당 ${minChoiceN}~${maxChoiceN} SKU)입니다. 차이가 우연인지 검정하지 않았습니다.`],
    ['target', '결품 방어는 대리 지표', `실제 결품 기록이 아닙니다. 가중평균 기준 방어율 ${pct(W0.cover)}(${VAL} · ${W0.n_lt} SKU)${higherOld || lowCover.length ? `는 ${[...(higherOld ? ['기존 수기'] : []), ...lowCover.map((b) => BLABEL[b])].join('·')}보다 낮고` : ''}, 가중치 최적화 후(검증 구간 ${HOLD}) ${pct(wsCur.cover)} → ${pct(wsNew.cover)}입니다.`],
    ['db', '테스트 데이터셋 결과', '비식별 처리한 테스트 데이터셋의 결과입니다. 다른 데이터에서도 같다고 보장하지 않습니다.'],
    ['calendar', `선택 구간 ${SS_.n}주`, `W${SS_.b}에 다시 고르는 상황을 흉내 낸 것이라 실제 운영보다 낙관적일 수 있습니다.`],
    ['doc', '예측 스냅샷 1개', '한 시점의 예측만 썼습니다. 주마다 쌓아 같은 비교를 반복해야 합니다.'],
  ];
  const step = (h - 0.5) / lim.length;
  lim.forEach(([ic, a, b], i) => {
    const ly = y + 0.5 + i * step;
    RR(s, rx, ly, rw, step - 0.08, C.amberL);
    icon(s, ic, rx + 0.14, ly + (step - 0.08 - 0.5) / 2, 0.5, C.amber);
    T(s, a, { x: rx + 0.75, y: ly, w: 1.85, h: step - 0.08, fontSize: 12, bold: true, color: C.navy, valign: 'middle' });
    T(s, b, { x: rx + 2.7, y: ly, w: rw - 2.85, h: step - 0.08, fontSize: 11, color: C.text, valign: 'middle' });
  });
})();

// ======================================================================= 15. 정리
(function closing() {
  page += 1;
  const s = pres.addSlide();
  s.background = { color: C.navy };
  T(s, '정리', { x: 0.9, y: 0.45, w: 8, h: 0.7, fontSize: 32, bold: true, color: C.white, valign: 'middle' });
  T(s, '목표 · 세 기준 비교 · 가중치 최적화', { x: 0.9, y: 1.12, w: 9.5, h: 0.38, fontSize: 14, color: C.sky, valign: 'middle' });
  tag(s, '정리', 'outline');
  const OLD = SB['기존수기'], WB = SB['가중평균'], W0 = M['가중평균'];
  const best = BYWAPE[0];
  const coverLow = isNum(M['기존수기'].cover) && M['기존수기'].cover > W0.cover;
  const items = [
    [`목표: 현업 수기 ${OLD.distinct}종 → 품목별 자동 산출`,
      `현업이 품목마다 계산할 수 없어 ${oldLevels.join('·')}일에 몰려 있던 일별 안전재고를, 같은 공식으로 품목마다 산출했습니다(가중평균 기준 ${WB.distinct}개 값, 중앙값 ${f1(OLD.median)} → ${f1(WB.median)}일).`],
    [best === '가중평균' ? '세 기준 비교: 가중평균 기준 오차율 최저' : `세 기준 비교: ${BLABEL[best]} 오차율 최저`,
      `${VAL} 공통 ${W0.n_demand} SKU 실제 출고 오차율 ${BYWAPE.map((b) => `${BSHORT[b]} ${pct(M[b].wape)}`).join(' · ')}.${coverLow ? ` 가중평균 기준의 결품 방어율 ${pct(W0.cover)}(${W0.n_lt} SKU)는 기존 수기 ${pct(M['기존수기'].cover)}보다 낮아 과제로 남았습니다.` : ''}`],
    ['Grid Search: 현업 w에서 출발해 데이터로 보정',
      `Grid Search(${man(SPACE)} 조합을 ${CANDS.length * NG}번 계산)와 Time-Series Holdout 검증으로 고른 w가 검증 구간 ${HOLD}에서 ${pct(wsCur.wape)} → ${pct(wsNew.wape)}(${pp(Math.max(wsCur.wape, wsNew.wape), Math.min(wsCur.wape, wsNew.wape))} ${wapeDown ? '개선' : '악화'}).${coverDown ? ` 방어율 ${pct(wsCur.cover)} → ${pct(wsNew.cover)}는 함께 볼 과제입니다.` : ''}`],
  ];
  items.forEach(([a, b], i) => {
    const y = 1.75 + i * 1.3;
    num(s, i + 1, 0.9, y + 0.05, 0.62, C.teal, C.white, 18);
    T(s, a, { x: 1.8, y, w: 10.6, h: 0.4, fontSize: 16, bold: true, color: C.sky, valign: 'middle' });
    T(s, b, { x: 1.8, y: y + 0.42, w: 10.8, h: 0.8, fontSize: 14, color: C.white, valign: 'top' });
  });
  box(s, [
    { text: '다음 과제  ', options: { bold: true, color: C.white } },
    { text: '서비스 수준(Z)도 같은 탐색·시계열 검증으로 조정 · 주간 재탐색과 성능 모니터링 · 승인 후 운영 적용', options: { color: C.white } },
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
