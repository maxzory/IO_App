const $ = id => document.getElementById(id);
const fmt = (x, d = 3) => Number(x).toLocaleString('ko-KR', {maximumFractionDigits: d});
const SVGNS = 'http://www.w3.org/2000/svg';

const LEVELS = {
  dae:   {file: './data/io-2023-dae.xlsx',  label: '대분류'},
  jung:  {file: './data/io-2023-jung.xlsx', label: '중분류'},
  so:    {file: './data/io-2023-so.xlsx',   label: '소분류'},
  basic: {file: './data/io-2023.xlsx',      label: '기본부문'},
};
const TYPES = [
  {name: '이중 위험형', color: 'var(--t-double)', desc: '수입·수출 충격 모두에 약함'},
  {name: '수입 위험형', color: 'var(--t-import)', desc: '원자재·공급망 충격에 약함'},
  {name: '수출 위험형', color: 'var(--t-export)', desc: '관세·해외 경기에 약함'},
  {name: '안정형',      color: 'var(--t-stable)', desc: '내수 위주, 외부 충격에 강함'},
];
const typeColor = t => (TYPES.find(x => x.name === t) || TYPES[3]).color;

let worker, nextId = 0, requests = new Map();
let level = null, sectors = [], rowChoices = [], cells = [], basket = [], result = null;
let diag = [], cuts = null, selected = null, highlight = '', parents = {};

function text(tag, value, cls) { const e = document.createElement(tag); e.textContent = value; if (cls) e.className = cls; return e; }
function svg(tag, attrs) { const e = document.createElementNS(SVGNS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); return e; }
function call(type, payload) {
  return new Promise((resolve, reject) => { const id = ++nextId; requests.set(id, {resolve, reject}); worker.postMessage({id, type, payload}); });
}
function startWorker() {
  worker = new Worker('./worker.js');
  worker.onmessage = ({data: d}) => {
    if (d.type === 'status') { $('status').textContent = d.text; return; }
    const p = requests.get(d.id); if (!p) return; requests.delete(d.id);
    d.ok ? p.resolve(d.result) : p.reject(Error(d.error));
  };
  worker.onerror = e => { for (const p of requests.values()) p.reject(Error(e.message || 'Worker 실행 오류')); requests.clear(); worker = null; };
}
function options(el, list) {
  el.replaceChildren(...list.map(x => { const o = document.createElement('option'); o.value = x.code; o.textContent = x.code + ' ' + x.name; return o; }));
}
const INPUTS = ['sheet', 'row', 'sector', 'add', 'run', 'row-search', 'col-search', 'sector-search', 'amount', 'diag-download'];
function setInputs(on) { for (const id of INPUTS) $(id).disabled = !on; }
function invalidate() { result = null; $('results').hidden = true; }

// ---------------- 자료 불러오기 ----------------
fetch('./data/classification.json').then(r => r.ok ? r.json() : {}).then(j => { parents = j; if (diag.length) renderTable(); }).catch(() => {});

async function loadLevel(key) {
  const buttons = document.querySelectorAll('[data-level]');
  buttons.forEach(b => b.disabled = true);
  invalidate(); basket = []; renderBasket(); setInputs(false);
  try {
    $('status').textContent = LEVELS[key].label + ' Excel 내려받는 중…';
    const response = await fetch(LEVELS[key].file);
    if (!response.ok) throw Error(LEVELS[key].label + ' Excel 다운로드 실패 (HTTP ' + response.status + ') — data 폴더에 파일이 있는지 확인하세요');
    const buffer = await response.arrayBuffer();
    if (!worker) startWorker();
    const info = await call('init', {buffer, level: key});
    level = key; sectors = info.sectors; diag = info.diagnosis; cuts = info.cuts; selected = null; highlight = '';
    options($('sector'), sectors);
    for (const id of ['sector-search', 'row-search', 'col-search', 'diag-search']) $(id).value = '';
    $('type-filter').value = '';
    $('sheet').replaceChildren(...info.sheets.map(s => { const o = text('option', s); o.value = s; return o; }));
    $('validation').textContent = JSON.stringify(info.residuals, null, 2) + '\n최대 절대 잔차가 허용값 1e-8 미만이면 통과';
    setInputs(true);
    buttons.forEach(b => b.classList.toggle('active', b.dataset.level === key));
    renderDiagnosis();
    await explore();
    $('status').textContent = '준비 완료 · ' + info.level + ' ' + sectors.length + '개 산업 · ' + info.sheets.length + '개 시트 · 계수 검증 통과';
  } catch (e) {
    $('status').textContent = '준비 실패: ' + e.message;
  } finally {
    buttons.forEach(b => b.disabled = false);
  }
}
document.querySelectorAll('[data-level]').forEach(b => b.onclick = () => loadLevel(b.dataset.level));

// ---------------- 1. 충격 유형 진단 ----------------
function renderDiagnosis() {
  $('level-tag').textContent = LEVELS[level].label + ' · ' + diag.length + '개 산업';
  $('detail').hidden = true;
  $('parent-head').hidden = level === 'dae';
  renderTypeCards(); renderLegend(); renderScatter(); renderTable();
}

function renderTypeCards() {
  $('type-cards').replaceChildren(...TYPES.map(t => {
    const n = diag.filter(d => d.type === t.name).length;
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'card type-card' + (highlight === t.name ? ' on' : '');
    b.style.setProperty('--c', t.color);
    b.setAttribute('aria-pressed', highlight === t.name);
    b.append(text('span', t.name, 'type-name'), text('strong', n + '개'), text('span', t.desc, 'type-desc'));
    b.onclick = () => { highlight = highlight === t.name ? '' : t.name; $('type-filter').value = highlight; renderTypeCards(); renderScatter(); renderTable(); };
    return b;
  }));
}

function renderLegend() {
  $('legend').replaceChildren(...TYPES.map(t => {
    const s = text('span', t.name, 'legend-item'); s.style.setProperty('--c', t.color); return s;
  }));
}

function niceMax(v) { const p = Math.pow(10, Math.floor(Math.log10(v))); for (const m of [1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) if (m * p >= v) return m * p; return 10 * p; }

function renderScatter() {
  const W = 760, H = 480, m = {l: 64, r: 18, t: 30, b: 52};
  const xMax = 100, yMax = niceMax(Math.max(...diag.map(d => d.importVuln)) * 1.05);
  const X = v => m.l + v / xMax * (W - m.l - m.r), Y = v => H - m.b - v / yMax * (H - m.t - m.b);
  const root = svg('svg', {viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': '수출 의존 비중과 수입 충격 취약도 산점도'});
  // 사분면 배경과 이름
  const cx = X(cuts.exportShare), cy = Y(cuts.importVuln);
  // 사분면 이름은 그래프의 네 모서리에 표시 (점과 겹치지 않도록)
  const corners = [['수입 위험형', m.l + 8, m.t + 16, 'start'], ['이중 위험형', W - m.r - 8, m.t + 16, 'end'],
                   ['안정형', m.l + 8, H - m.b - 8, 'start'], ['수출 위험형', W - m.r - 8, H - m.b - 8, 'end']];
  // 격자와 눈금
  for (let i = 0; i <= 5; i++) {
    const xv = i * 20, yv = yMax * i / 5;
    root.append(svg('line', {x1: X(xv), x2: X(xv), y1: m.t, y2: H - m.b, class: 'grid'}));
    root.append(svg('line', {x1: m.l, x2: W - m.r, y1: Y(yv), y2: Y(yv), class: 'grid'}));
    const tx = svg('text', {x: X(xv), y: H - m.b + 18, 'text-anchor': 'middle', class: 'tick'}); tx.textContent = xv; root.append(tx);
    const ty = svg('text', {x: m.l - 8, y: Y(yv) + 4, 'text-anchor': 'end', class: 'tick'}); ty.textContent = fmt(yv, 2); root.append(ty);
  }
  const ax = svg('text', {x: (m.l + W - m.r) / 2, y: H - 10, 'text-anchor': 'middle', class: 'axis-title'}); ax.textContent = '수출 의존 비중 (%) → 수출 충격에 약함'; root.append(ax);
  const ay = svg('text', {x: -(m.t + H - m.b) / 2, y: 16, transform: 'rotate(-90)', 'text-anchor': 'middle', class: 'axis-title'}); ay.textContent = '수입 충격 취약도 (지수) →'; root.append(ay);
  // 중앙값 기준선
  root.append(svg('line', {x1: cx, x2: cx, y1: m.t, y2: H - m.b, class: 'cut'}));
  root.append(svg('line', {x1: m.l, x2: W - m.r, y1: cy, y2: cy, class: 'cut'}));
  const cl = svg('text', {x: cx, y: m.t - 8, 'text-anchor': 'middle', class: 'cut-label'}); cl.textContent = '중앙값 ' + fmt(cuts.exportShare, 1) + '%'; root.append(cl);
  const cl2 = svg('text', {x: W - m.r - 4, y: cy - 6, 'text-anchor': 'end', class: 'cut-label'}); cl2.textContent = '중앙값 ' + fmt(cuts.importVuln, 3); root.append(cl2);
  for (const [name, x, y, anchor] of corners) {
    const t = svg('text', {x, y, 'text-anchor': anchor, class: 'quad-label'}); t.textContent = name; root.append(t);
  }
  // 점
  const r = diag.length > 200 ? 4.5 : diag.length > 80 ? 5.5 : 7;
  const tip = $('tooltip');
  const ordered = [...diag].sort((a, b) => (a.code === selected) - (b.code === selected));
  for (const d of ordered) {
    const dim = highlight && d.type !== highlight;
    const c = svg('circle', {cx: X(d.exportShare), cy: Y(d.importVuln), r: d.code === selected ? r + 3 : r,
      class: 'dot' + (dim ? ' dim' : '') + (d.code === selected ? ' sel' : ''), tabindex: dim ? -1 : 0});
    c.style.fill = typeColor(d.type);
    const show = ev => {
      tip.hidden = false;
      tip.replaceChildren(text('b', d.name), text('span', d.type),
        text('span', '수출 의존 ' + fmt(d.exportShare, 1) + '% · 취약도 ' + fmt(d.importVuln, 3)));
      const box = $('scatter').getBoundingClientRect(), pt = ev.touches ? ev.touches[0] : ev;
      const x = (pt.clientX ?? box.left + box.width / 2) - box.left, y = (pt.clientY ?? box.top) - box.top;
      tip.style.left = Math.min(x + 14, box.width - 220) + 'px'; tip.style.top = (y + 14) + 'px';
    };
    c.addEventListener('mousemove', show);
    c.addEventListener('mouseleave', () => tip.hidden = true);
    c.addEventListener('focus', () => { tip.hidden = true; });
    c.addEventListener('click', () => selectSector(d.code));
    c.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); selectSector(d.code); } });
    root.append(c);
  }
  // 선택한 산업 이름표
  const s = diag.find(d => d.code === selected);
  if (s) {
    const lx = X(s.exportShare), ly = Y(s.importVuln) - r - 8;
    const t = svg('text', {x: lx, y: ly, 'text-anchor': lx > W * 0.7 ? 'end' : lx < W * 0.3 ? 'start' : 'middle', class: 'sel-label'}); t.textContent = s.name; root.append(t);
  }
  $('scatter').replaceChildren(root);
}

function describe(d) {
  const msgs = {
    '이중 위험형': '수입 원료 의존과 수출 의존이 모두 중앙값보다 높습니다. 원자재 가격 급등과 해외 수요 감소 양쪽에서 타격을 받을 수 있어 가장 먼저 관리해야 할 산업입니다.',
    '수입 위험형': '수입 원료 의존이 높고 그 충격이 다른 산업으로 퍼지기 쉽습니다. 유가·원자재 가격이나 공급망 차질에 대비한 수입처 다변화와 국산화가 중요합니다.',
    '수출 위험형': '생산의 상당 부분이 수출에 기대고 있습니다. 관세, 해외 경기 침체, 환율 변화에 민감하므로 수출 시장 다변화와 내수 기반 강화가 중요합니다.',
    '안정형': '수입 원료 의존과 수출 의존이 모두 중앙값보다 낮은 내수 중심 산업입니다. 외부 충격의 직접 영향은 상대적으로 작습니다.',
  };
  return msgs[d.type];
}

function selectSector(code) {
  selected = code;
  const d = diag.find(x => x.code === code); if (!d) return;
  renderScatter(); renderTable();
  const p = parents[d.code];
  const box = $('detail'); box.hidden = false; box.style.setProperty('--c', typeColor(d.type));
  const metrics = [['수출 의존 비중', fmt(d.exportShare, 1) + ' %'], ['수입 원료 비중', fmt(d.importShare, 1) + ' %'],
    ['영향력계수', fmt(d.influence, 3) + ' 배'], ['수입 충격 취약도', fmt(d.importVuln, 3)],
    ['수출이 유발한 생산액', fmt(d.exportInduced, 0) + ' 백만원'], ['수출액(국산)', fmt(d.exports, 0) + ' 백만원']];
  const grid = document.createElement('div'); grid.className = 'metrics';
  for (const [k, v] of metrics) { const m = document.createElement('div'); m.append(text('span', k), text('strong', v)); grid.append(m); }
  const title = document.createElement('h3');
  title.append(text('span', d.type, 'badge'), document.createTextNode(' ' + d.code + ' ' + d.name));
  const children = [title];
  if (p && level !== 'dae') children.push(text('p', '소속 대분류: ' + p.daeCode + ' ' + p.daeName, 'note'));
  children.push(text('p', describe(d)), grid);
  if (d.exports > 0) {
    const shock = -Math.round(d.exports * 0.1);
    const btn = text('button', '이 산업 수출 10% 감소(' + fmt(shock, 0) + '백만원)를 시뮬레이터에 넣기');
    btn.type = 'button';
    btn.onclick = () => { basket.push({code: d.code, name: d.name, amount: shock}); invalidate(); renderBasket(); $('simulation').scrollIntoView(); };
    children.push(btn);
  }
  box.replaceChildren(...children);
}

function filteredDiag() {
  const t = $('type-filter').value, q = $('diag-search').value.trim(), key = $('diag-sort').value;
  const rows = diag.filter(d => (!t || d.type === t) && (d.code + ' ' + d.name).includes(q));
  rows.sort((a, b) => key === 'code' ? a.code.localeCompare(b.code) : b[key] - a[key]);
  return rows;
}

function renderTable() {
  if (!diag.length) return;
  const rows = filteredDiag(), showParent = level !== 'dae';
  $('diag-caption').textContent = LEVELS[level].label + ' 진단 결과 · ' + rows.length + '개 산업 표시 (행을 누르면 그래프에서 강조)';
  $('diag-rows').replaceChildren(...rows.map(d => {
    const tr = document.createElement('tr'); if (d.code === selected) tr.className = 'sel';
    const p = parents[d.code];
    const vals = [d.code, d.name, showParent ? (p ? p.daeName : '') : null, null, fmt(d.exportShare, 1), fmt(d.importShare, 1), fmt(d.influence, 3), fmt(d.importVuln, 3), fmt(d.exports, 0)];
    vals.forEach((v, i) => {
      if (i === 2 && !showParent) return;
      if (i === 3) { const td = document.createElement('td'); const s = text('span', d.type, 'type-chip'); s.style.setProperty('--c', typeColor(d.type)); td.append(s); tr.append(td); return; }
      tr.append(text('td', v));
    });
    tr.onclick = () => { selectSector(d.code); $('diagnosis').querySelector('.chart-wrap').scrollIntoView({block: 'center'}); };
    return tr;
  }));
}
$('type-filter').onchange = () => { highlight = $('type-filter').value; renderTypeCards(); renderScatter(); renderTable(); };
$('diag-search').oninput = renderTable;
$('diag-sort').onchange = renderTable;

function csvCell(v) { let s = String(v); if (/^[=+@-]/.test(s) && typeof v !== 'number') s = "'" + s; return '"' + s.replaceAll('"', '""') + '"'; }
function downloadCsv(rows, name) {
  const url = URL.createObjectURL(new Blob(['﻿' + rows.map(r => r.map(csvCell).join(',')).join('\r\n')], {type: 'text/csv;charset=utf-8'}));
  const a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}
$('diag-download').onclick = () => {
  const rows = [['산업코드', '산업명', '소속대분류', '유형', '수출의존비중_%', '수입원료비중_%', '영향력계수', '수입충격취약도', '수출유발생산액_백만원', '수출액_백만원'],
    ...filteredDiag().map(d => [d.code, d.name, (parents[d.code] || {}).daeName || '', d.type, d.exportShare, d.importShare, d.influence, d.importVuln, d.exportInduced, d.exports])];
  downloadCsv(rows, 'shock-diagnosis-' + level + '.csv');
};

// ---------------- 2. 원자료 탐색 ----------------
let exploreSeq = 0;
async function explore() {
  const seq = ++exploreSeq;
  try {
    const d = await call('explore', {sheet: $('sheet').value, row: $('row').value});
    if (seq !== exploreSeq) return;
    rowChoices = d.rows; filterRows(d.selected); cells = d.cells;
    $('unit').textContent = d.unit;
    $('explore-caption').textContent = $('sheet').value + ' / 행: ' + d.selected + ' ' + d.name;
    renderCells();
  } catch (e) { $('status').textContent = e.message; }
}
function filterRows(sel) { const q = $('row-search').value.trim(); options($('row'), rowChoices.filter(r => (r.code + ' ' + r.name).includes(q))); if ([...$('row').options].some(o => o.value === sel)) $('row').value = sel; }
function renderCells() {
  const q = $('col-search').value.trim();
  $('cells').replaceChildren(...cells.filter(c => (c.code + ' ' + c.name).includes(q)).map(c => {
    const tr = document.createElement('tr');
    for (const v of [c.code, c.name, c.value === null ? '빈 셀' : typeof c.value === 'number' ? fmt(c.value) : c.value]) tr.append(text('td', v));
    return tr;
  }));
}
$('sheet').onchange = () => { $('row-search').value = ''; $('row').replaceChildren(); explore(); };
$('row').onchange = explore;
$('row-search').oninput = () => { filterRows(); if ($('row').value) explore(); else { $('cells').replaceChildren(); $('explore-caption').textContent = '일치하는 행이 없습니다.'; } };
$('col-search').oninput = renderCells;

// ---------------- 3. 시뮬레이터 ----------------
$('sector-search').oninput = () => { const q = $('sector-search').value.trim(); options($('sector'), sectors.filter(s => (s.code + ' ' + s.name).includes(q))); };
function renderBasket() {
  $('basket').replaceChildren(...basket.map((s, i) => {
    const li = text('li', s.code + ' ' + s.name + ' / ' + (s.amount > 0 ? '+' : '') + fmt(s.amount) + '백만원 ');
    const b = text('button', '삭제'); b.type = 'button'; b.onclick = () => { basket.splice(i, 1); invalidate(); renderBasket(); };
    li.append(b); return li;
  }));
  $('sim-status').textContent = basket.length ? basket.length + '개 입력 준비. 계산 버튼을 누르세요.' : '시나리오를 추가하세요.';
}
$('scenario').onsubmit = e => {
  e.preventDefault();
  const amount = Number($('amount').value), s = sectors.find(s => s.code === $('sector').value);
  if (!s || !Number.isFinite(amount) || amount === 0) { $('sim-status').textContent = '산업과 0이 아닌 변화액을 확인하세요.'; return; }
  basket.push({...s, amount}); invalidate(); renderBasket();
};
$('clear').onclick = () => { basket = []; invalidate(); renderBasket(); };
$('run').onclick = async () => {
  if (!basket.length) { $('sim-status').textContent = '시나리오를 먼저 추가하세요.'; return; }
  const snapshot = basket.map(s => ({...s}));
  $('run').disabled = true; $('sim-status').textContent = '계산 중…';
  try {
    const r = await call('simulate', snapshot);
    if (JSON.stringify(snapshot) !== JSON.stringify(basket)) { $('sim-status').textContent = '입력이 변경되었습니다. 다시 계산하세요.'; return; }
    result = r; renderResult(); $('sim-status').textContent = '계산 완료. 아래 결과는 현재 시나리오에 대응합니다.';
  } catch (e) { invalidate(); $('sim-status').textContent = '계산 실패: ' + e.message; }
  finally { $('run').disabled = false; }
};
function signed(v) { return (v > 0 ? '+' : '') + fmt(v, 1); }
function renderResult() {
  $('results').hidden = false;
  $('cards').replaceChildren(...[['수요 변화액', result.demand], ['생산 변화', result.production], ['수입 변화', result.imports], ['부가가치 변화', result.valueAdded]].map(([title, v]) => {
    const d = text('div', title, 'card'); d.append(text('strong', signed(v)), text('span', '백만원')); return d;
  }));
  const max = Math.max(...result.rows.map(r => Math.abs(r.production)), 1);
  $('bars').replaceChildren(...result.rows.slice(0, 15).map(r => {
    const d = text('div', r.name + ' · ' + signed(r.production) + '백만원', 'bar');
    const t = document.createElement('div'); t.className = 'track';
    const f = document.createElement('div'); f.className = 'fill' + (r.production < 0 ? ' neg' : '');
    f.style.width = Math.abs(r.production) / max * 100 + '%'; t.append(f); d.append(t); return d;
  }));
  $('result-rows').replaceChildren(...result.rows.map(r => {
    const tr = document.createElement('tr'); [r.code, r.name, signed(r.production), signed(r.imports), signed(r.valueAdded)].forEach(v => tr.append(text('td', v))); return tr;
  }));
}
$('download').onclick = () => {
  if (!result) return;
  downloadCsv([['산업코드', '산업명', '생산변화_백만원', '수입변화_백만원', '부가가치변화_백만원'], ...result.rows.map(r => [r.code, r.name, r.production, r.imports, r.valueAdded])], 'simulation-' + level + '.csv');
};
