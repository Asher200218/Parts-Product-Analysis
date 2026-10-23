/* Parts-Analysis local input tool.
   100% local: data lives in localStorage, Excel handled in-browser via vendored SheetJS.
   Importers also accept the older Chinese headers / sheet names / values, so existing files keep working. */
'use strict';

/* ---------- domain constants ---------- */
const DEFAULT_CATEGORY_HOURS = { 'High': 4.0, 'Low': 2.0, 'None': 0.25 };
// Live, user-customizable hours-per-difficulty. Mutated in place (never reassigned)
// so analysis.js, which reads it at call-time, always sees the current values.
const CATEGORY_HOURS = { 'High': 4.0, 'Low': 2.0, 'None': 0.25 };
const CATEGORIES = ['High', 'Low', 'None'];          // hardest -> easiest
const ABILITY_SKILLS = { A: ['High', 'Low', 'None'], B: ['Low', 'None'], C: ['None'] };
// legacy Chinese difficulty values (older saved data / Excel files) -> current values
const LEGACY_CATEGORY = { '高': 'High', '低': 'Low', '无': 'None' };
// hour-report task names (the hour-report export writes them in Chinese) -> canonical English task names
const TASK_ALIASES = { '订单设计': 'Order design', '订单校对': 'Order check', '加班': 'Overtime' };
const GRADES = ['', 'A', 'B', 'C'];
// Worker ability is split into two roles: design ('d') and check/verify ('v').
// Each part needs BOTH a design pass and a verify pass; w.grades[product] = { d, v }.
const ROLES = ['d', 'v'];
const ROLE_LABEL = { d: 'Design', v: 'Check' };
function emptyGrade(){ return { d: '', v: '' }; }
/* migrate legacy {product:'A'} -> {product:{d:'A',v:'A'}}; keep new {d,v} shape as-is */
function migrateGrades(grades){
  const out = {};
  for (const p in (grades || {})){
    const g = grades[p];
    if (g && typeof g === 'object') out[p] = { d: normGrade(g.d), v: normGrade(g.v) };
    else { const s = normGrade(g); if (s) out[p] = { d: s, v: s }; else out[p] = emptyGrade(); }
  }
  return out;
}
/* parse a worker-matrix column header into {product, role}; role '' = legacy (applies to both) */
function parseWorkerColRole(headerText){
  let t = clean(headerText), role = '', m;
  if ((m = t.match(/^(.*?)[\s·•\-_:：/|]*(设计|design)$/i))) { role = 'd'; t = m[1]; }
  else if ((m = t.match(/^(.*?)[\s·•\-_:：/|]*(校验|校对|verify|check)$/i))) { role = 'v'; t = m[1]; }
  return { product: t.trim(), role };
}
const DEFAULT_HOURS_PER_DAY = 8;   // workdays are fixed Monday–Friday (see analysis.js WORKDAYS_PER_WEEK)
// live, user-customizable shift length (mutated, never re-declared) so analysis.js sees it
let HOURS_PER_DAY = 8;
/* Check ratio: check hours = design hours × VERIFY_RATIO. Default 0.3;
   can be calibrated in one click on the Hour Log tab from actual logged hours (= total order-check hours ÷ total order-design hours). */
const DEFAULT_VERIFY_RATIO = 0.3;
let VERIFY_RATIO = DEFAULT_VERIFY_RATIO;
function applyVerifyRatio(v){
  v = (v == null || v === '' || isNaN(v) || v <= 0 || v > 3) ? DEFAULT_VERIFY_RATIO : Number(v);
  VERIFY_RATIO = v; state.settings = { ...(state.settings || {}), verifyRatio: v };
}

// header text -> internal field (Chinese aliases kept so existing Excel files still import)
const COLUMN_ALIASES = {
  '零件名称':'part','part':'part','partname':'part','part name':'part','零件':'part',
  '耗时等级':'category','硬度':'category','难度':'category','category':'category','level':'category','hardness':'category','等级':'category','difficulty':'category',
  '单件工时':'hours','hours':'hours','工时':'hours','hours per piece':'hours',
  '产品类型':'product','类型':'product','产品名称':'product','产品':'product','product':'product','producttype':'product','product type':'product',
  '所需零件':'parts','零件清单':'parts','所需零件清单':'parts','parts':'parts','partsneeded':'parts','partlist':'parts','parts needed':'parts',
  '提交时间':'due','提交日期':'due','交期':'due','due':'due','duedate':'due','submission':'due','submissiontime':'due','submitted':'due',
  '工人姓名':'worker','worker':'worker','name':'worker','姓名':'worker','worker name':'worker',
};
const CATEGORY_ALIASES = {
  '高':'High','high':'High','h':'High','难':'High',
  '低':'Low','low':'Low','l':'Low','中':'Low',
  '无':'None','none':'None','no':'None','n':'None','易':'None',
};
const PART_SPLIT_RE = /[;；、\n]+/;
const QTY_TAIL_RE = /[×✕✖]\s*([0-9]+(?:\.[0-9]+)?)\s*$/;
const CODE_PREFIX_RE = /^\s*\d+\s*[:：]\s*/;

// exact sheet names / headers used for exported .xlsx files
const SHEET = { parts:'Parts', workers:'Worker abilities', orders:'Orders' };
const LEGACY_SHEET = { '零件清单':'parts', '工人能力':'workers', '订单明细':'orders' };   // sheet names in older files
const HDR = {
  parts: ['Part name', 'Difficulty (High/Low/None)'],
  workersName: 'Worker name',
  orders: ['Product type', 'Parts needed (separate with ; — quantity ×N)', 'Submitted (optional, YYYY-MM-DD)'],
};

/* ---------- state ---------- */
const STORE_KEY = 'parts-analysis-web-v1';
let state = { products: [], parts: [], workers: [], orders: [], leaves: [], settings: { categoryHours: { ...DEFAULT_CATEGORY_HOURS }, hoursPerDay: DEFAULT_HOURS_PER_DAY, startDate: '' } };

function load() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (raw) state = Object.assign(state, JSON.parse(raw));
  } catch (e) { console.warn('load failed', e); }
  // normalise shape
  state.products = (state.products || []).map(String);
  state.parts = (state.parts || []).map(p => ({ name: p.name || '', category: LEGACY_CATEGORY[p.category] || p.category || '', t7: p.t7 ?? '', t8: p.t8 ?? '' }));
  if (state.settings && state.settings.categoryHours){   // migrate legacy Chinese difficulty keys
    const ch = state.settings.categoryHours;
    for (const k in LEGACY_CATEGORY) if (k in ch){ if (!(LEGACY_CATEGORY[k] in ch)) ch[LEGACY_CATEGORY[k]] = ch[k]; delete ch[k]; }
  }
  state.workers = (state.workers || []).map(w => ({ name: w.name || '', grades: migrateGrades(w.grades) }))
    .filter(w => !isSkipWorker(w.name));   // drop null / special order / order canceled placeholder names

  state.orders = (state.orders || []).map(o => ({ product: o.product || '', parts: String(o.parts || '').split(PART_SPLIT_RE).map(s => s.trim()).filter(Boolean).join('; '), time: o.time || '', actualFinish: o.actualFinish || '', designer: o.designer || '', checker: o.checker || '', required: o.required || '', so: o.so || '' }));
  state.leaves = (state.leaves || []).map(L => ({ worker: L.worker || '', start: L.start || '', end: L.end || '',
    hours: (L.hours != null && L.hours !== '') ? L.hours : (L.half ? HOURS_PER_DAY / 2 : '') }));   // half → half a shift (legacy data)
  if (!state.hourlog || typeof state.hourlog !== 'object' || !Array.isArray(state.hourlog.workers)) state.hourlog = null;   // hour log (hour-report import, stored aggregated)
  if (state.hourlog){   // migrate legacy Chinese task names
    const mapTask = t => TASK_ALIASES[t] || t;
    state.hourlog.workers.forEach(w => { const bt = {}; for (const k in (w.byTask || {})) bt[mapTask(k)] = (bt[mapTask(k)] || 0) + w.byTask[k]; w.byTask = bt; });
    (state.hourlog.tasks || []).forEach(t => { t.name = mapTask(t.name); });
  }
  applyCategoryHours(state.settings && state.settings.categoryHours);
  applyHoursPerDay(state.settings && state.settings.hoursPerDay);
  applyVerifyRatio(state.settings && state.settings.verifyRatio);
  if (typeof state.settings.startDate !== 'string') state.settings.startDate = '';
  if (!(state.settings.schedBuffer >= 0)) state.settings.schedBuffer = 2;   // schedule buffer for finishing early (workdays)
  if (!(state.settings.annualLeave >= 0)) state.settings.annualLeave = 15;  // annual leave (days/person·year), spread over the schedule's daily capacity
}

/* Custom hours per difficulty. Source of truth is state.settings.categoryHours;
   mirror it into the live CATEGORY_HOURS object so analysis.js picks it up. */
function applyCategoryHours(src) {
  const ch = {};
  CATEGORIES.forEach(c => {
    let v = src && src[c]; v = (v == null || v === '' || isNaN(v) || v < 0) ? DEFAULT_CATEGORY_HOURS[c] : Number(v);
    ch[c] = v; CATEGORY_HOURS[c] = v;
  });
  state.settings = { ...(state.settings || {}), categoryHours: ch };
}
/* Custom shift length (hours/worker/day). Mirrored into the live HOURS_PER_DAY. */
function applyHoursPerDay(v) {
  v = (v == null || v === '' || isNaN(v) || v <= 0) ? DEFAULT_HOURS_PER_DAY : Number(v);
  HOURS_PER_DAY = v; state.settings = { ...(state.settings || {}), hoursPerDay: v };
}
function hoursDefault() { return CATEGORIES.every(c => CATEGORY_HOURS[c] === DEFAULT_CATEGORY_HOURS[c]); }
function fmtHrs(h) { return Number.isInteger(h) ? String(h) : String(Math.round(h * 1000) / 1000); }
function catLabel(c) { return `${c} (${fmtHrs(CATEGORY_HOURS[c])}h)`; }
let saveTimer = null;
let quotaWarned = false;
function save() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
    catch (e) {
      console.warn('save failed', e);
      if (!quotaWarned){ quotaWarned = true; toast('Data too large: local save failed (browser storage is full). Export a backup to keep working.', 'err'); }
    }
  }, 150);
  updateCounts();
}
function updateCounts() {
  byId('cnt-products').textContent = state.products.length;
  byId('cnt-parts').textContent = state.parts.length;
  byId('cnt-workers').textContent = state.workers.length;
  byId('cnt-orders').textContent = state.orders.length;
  ['parts','workers','orders'].forEach(k => {
    const el = byId('dl-' + k); if (el) el.textContent = state[k].length + ' rows';
  });
  const dp = byId('dl-products'); if (dp) dp.textContent = state.products.length + ' products';
}

/* ---------- tiny helpers ---------- */
function byId(id){ return document.getElementById(id); }
function el(tag, attrs={}, kids=[]) {
  const n = document.createElement(tag);
  for (const k in attrs) {
    if (k === 'class') n.className = attrs[k];
    else if (k === 'html') n.innerHTML = attrs[k];
    else if (k.startsWith('on')) n.addEventListener(k.slice(2), attrs[k]);
    else if (attrs[k] === true) n.setAttribute(k, '');
    else if (attrs[k] !== false && attrs[k] != null) n.setAttribute(k, attrs[k]);
  }
  (Array.isArray(kids) ? kids : [kids]).forEach(c => c != null && n.append(c.nodeType ? c : document.createTextNode(c)));
  return n;
}
function clean(v){ return v == null ? '' : String(v).trim(); }
function normHeader(text){
  let t = clean(text).toLowerCase();
  for (const opener of ['(', '（']) if (t.includes(opener)) t = t.split(opener)[0];
  t = t.trim();
  return COLUMN_ALIASES[t] || t;
}
function normCategory(v){
  const t = clean(v).toLowerCase();
  return CATEGORY_ALIASES[t] || (CATEGORIES.includes(clean(v)) ? clean(v) : '');
}
function normGrade(v){
  const t = clean(v).toUpperCase();
  return ['A','B','C'].includes(t) ? t : '';
}
let toastTimer;
function toast(msg, kind){
  const t = byId('toast'); t.textContent = msg; t.className = 'toast' + (kind ? ' '+kind : '');
  t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(()=>t.hidden=true, 2600);
}

/* ---------- order-part parsing (mirror analyze.py) ---------- */
function parseOrderParts(text){
  const out = [];
  for (let chunk of String(text || '').split(PART_SPLIT_RE)){
    chunk = chunk.trim().replace(/^[\s　]+|[\s　]+$/g,'');
    if (!chunk) continue;
    let qty = 1;
    const m = chunk.match(QTY_TAIL_RE);
    if (m){ qty = parseFloat(m[1]); chunk = chunk.slice(0, m.index).trim(); }
    const name = chunk.replace(CODE_PREFIX_RE, '').trim();
    if (name) out.push({ name, qty });
  }
  return out;
}

/* =========================================================
   TABS
   ========================================================= */
byId('tabs').addEventListener('click', e => {
  const b = e.target.closest('.tab'); if (!b) return;
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t === b));
  const tab = b.dataset.tab;
  document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.id === 'panel-'+tab));
  if (tab === 'workers') renderWorkers();
  if (tab === 'retro') openRetro();
  if (tab === 'sched') openSchedule();
  if (tab === 'hourlog') renderHourlog();
});

/* =========================================================
   PRODUCTS
   ========================================================= */
function addProduct(name){
  name = clean(name);
  if (!name) return false;
  if (state.products.some(p => p === name)){ toast('Product already exists: '+name, 'err'); return false; }
  state.products.push(name);
  state.workers.forEach(w => { if (!(name in w.grades)) w.grades[name] = emptyGrade(); });
  save(); renderProducts();
  return true;
}
function renameProduct(idx, next){
  next = clean(next);
  const prev = state.products[idx];
  if (!next || next === prev) { renderProducts(); return; }
  if (state.products.some((p,i) => p === next && i !== idx)){ toast('A product with that name already exists', 'err'); renderProducts(); return; }
  state.products[idx] = next;
  state.workers.forEach(w => { w.grades[next] = w.grades[prev] || emptyGrade(); delete w.grades[prev]; });
  state.orders.forEach(o => { if (o.product === prev) o.product = next; });
  save(); renderProducts();
}
function deleteProduct(idx){
  const name = state.products[idx];
  const used = state.orders.filter(o => o.product === name).length;
  let msg = `Delete product "${name}"? Its column will be removed from worker abilities`;
  if (used) msg += `; ${used} orders reference it (the orders are kept, but their product field will be blank)`;
  if (!confirm(msg + '.')) return;
  state.products.splice(idx, 1);
  state.workers.forEach(w => delete w.grades[name]);
  save(); renderProducts();
}
function renderProducts(){
  const ul = byId('product-list'); ul.innerHTML = '';
  if (!state.products.length) ul.append(el('li', { class:'empty-note', style:'list-style:none' }, 'No products yet. Type a name above to add one, or import on the Import/Export tab.'));
  state.products.forEach((p, i) => {
    const input = el('input', { class:'pname', value:p, spellcheck:'false' });
    input.addEventListener('change', () => renameProduct(i, input.value));
    input.addEventListener('keydown', e => { if (e.key==='Enter') input.blur(); });
    const li = el('li', { class:'chip' }, [
      input,
      el('button', { class:'mini mv', title:'Move up', onclick:()=>moveProduct(i,-1) }, '↑'),
      el('button', { class:'mini mv', title:'Move down', onclick:()=>moveProduct(i, 1) }, '↓'),
      el('button', { class:'mini', title:'Delete', onclick:()=>deleteProduct(i) }, '✕'),
    ]);
    ul.append(li);
  });
  updateCounts();
}
function moveProduct(i, d){
  const j = i + d; if (j < 0 || j >= state.products.length) return;
  [state.products[i], state.products[j]] = [state.products[j], state.products[i]];
  save(); renderProducts();
}
byId('product-add').onclick = () => { if (addProduct(byId('product-input').value)) byId('product-input').value=''; };
byId('product-input').addEventListener('keydown', e => { if (e.key==='Enter') byId('product-add').click(); });
byId('product-bulk').onclick = () => openModal('Bulk add products', 'One product name per line. Existing ones are skipped.', '', text => {
  let n = 0; text.split('\n').forEach(line => { if (addProduct(line)) n++; });
  toast(`Added ${n} products`, 'ok');
});

/* =========================================================
   PARTS  (+ custom hours-per-difficulty settings)
   ========================================================= */
function renderHoursSettings(){
  const bar = byId('hours-settings'); if (!bar) return;
  bar.innerHTML = '';
  bar.append(el('span', { class:'settings-label' }, 'Hours per piece:'));
  CATEGORIES.forEach(c => {
    const inp = el('input', { type:'number', min:'0', step:'0.25', value:fmtHrs(CATEGORY_HOURS[c]) });
    inp.addEventListener('change', () => {
      let v = parseFloat(inp.value);
      if (isNaN(v) || v < 0){ v = CATEGORY_HOURS[c]; inp.value = fmtHrs(v); toast('Enter a number ≥ 0', 'err'); return; }
      CATEGORY_HOURS[c] = v; state.settings.categoryHours[c] = v; save();
      renderParts();              // refresh dropdown labels + hours column
      byId('analysis-out').innerHTML = '';  // analysis is now stale
      renderHoursSettings();      // refresh the "customized" marker
      toast(`Set ${c} = ${fmtHrs(v)}h`, 'ok');
    });
    bar.append(el('label', { class:'h-'+c }, [c, inp, el('span', { class:'unit' }, 'h')]));
  });
  bar.append(el('button', { class:'btn ghost', onclick:resetHours }, 'Reset to defaults'));
  bar.append(el('span', { class:'settings-hint' },
    hoursDefault()
      ? 'Changing these affects dropdown labels, hour totals and analysis.'
      : el('span', {}, [el('span', { class:'dirty' }, 'Customized'), ': exported parts.xlsx will include an "Hours per piece" column.'])));
}
function resetHours(){
  CATEGORIES.forEach(c => { CATEGORY_HOURS[c] = DEFAULT_CATEGORY_HOURS[c]; state.settings.categoryHours[c] = DEFAULT_CATEGORY_HOURS[c]; });
  save(); renderHoursSettings(); renderParts(); byId('analysis-out').innerHTML = '';
  toast('Reset to defaults (High 4 / Low 2 / None 0.25)', 'ok');
}

/* Check-ratio settings bar (Parts tab): check hours = design × ratio */
function renderVerifySettings(){
  const bar = byId('verify-settings'); if (!bar) return;
  bar.innerHTML = '';
  bar.append(el('span', { class:'settings-label' }, 'Check ratio:'));
  const inp = el('input', { type:'number', min:'0.01', max:'3', step:'0.01', value:String(VERIFY_RATIO), style:'width:90px' });
  inp.addEventListener('change', () => {
    const v = parseFloat(inp.value);
    if (isNaN(v) || v <= 0 || v > 3){ inp.value = String(VERIFY_RATIO); toast('Enter a number between 0 and 3', 'err'); return; }
    applyVerifyRatio(v); save();
    ['analysis-out','sched-out','retro-out'].forEach(id => { const n = byId(id); if (n) n.innerHTML = ''; });   // basis changed, old results are stale
    renderVerifySettings();
    toast(`Check ratio set to ${v} (check hours = design × ${v})`, 'ok');
  });
  bar.append(el('label', {}, [inp]));
  bar.append(el('span', { class:'settings-hint' },
    el('span', {}, [
      `Check hours = design hours × ratio (currently ${VERIFY_RATIO}${VERIFY_RATIO === DEFAULT_VERIFY_RATIO ? ', default' : ''}). `,
      'After importing an hour report on the Hour Log tab you can calibrate it from actual hours in one click (= total order-check hours ÷ total order-design hours).' ])));
}

let partFilter = '';
function renderParts(){
  renderHoursSettings();
  renderVerifySettings();
  const tb = byId('part-table').querySelector('tbody'); tb.innerHTML = '';
  const f = partFilter.toLowerCase();
  let shown = 0, matched = 0;
  state.parts.forEach((p, i) => {
    if (f && !p.name.toLowerCase().includes(f)) return;
    matched++;
    if (shown >= ROW_CAP) return;     // cap rendered rows; use the filter to narrow
    shown++;
    const nameI = el('input', { value:p.name, placeholder:'VC code' });
    nameI.addEventListener('input', () => { p.name = nameI.value; save(); });
    const sel = el('select', { class:'cell cat-'+p.category });
    [['','—'],['High','High'],['Low','Low'],['None','None']].forEach(([v,t]) =>
      sel.append(el('option', { value:v, selected:v===p.category }, t)));
    sel.addEventListener('change', () => { p.category = sel.value; sel.className = 'cell cat-'+p.category; save(); });
    const t7I = el('input', { type:'number', min:'0', step:'0.01', value:p.t7, placeholder:'col8', title:'Design hours for repeat orders (from column 8)' });
    t7I.addEventListener('input', () => { p.t7 = t7I.value; save(); });
    const t8I = el('input', { type:'number', min:'0', step:'0.01', value:p.t8, placeholder:'col9', title:'Design hours for the first order (from column 9)' });
    t8I.addEventListener('input', () => { p.t8 = t8I.value; save(); });
    tb.append(el('tr', {}, [
      el('td', { class:'rownum' }, String(i+1)),
      el('td', {}, nameI),
      el('td', {}, sel),
      el('td', {}, t7I),
      el('td', {}, t8I),
      el('td', { class:'act' }, el('button', { class:'row-del', title:'Delete', onclick:()=>{ state.parts.splice(i,1); save(); renderParts(); } }, '✕')),
    ]));
  });
  if (matched > shown)
    byId('part-table').querySelector('tbody').append(el('tr', {}, el('td', { colspan:'6', style:'padding:12px;color:var(--muted);text-align:center' },
      `Showing the first ${shown} (of ${matched}${f?' matching':''}). Use the filter above to narrow down.`)));
  updateCounts();
}
byId('part-add').onclick = () => { state.parts.push({ name:'', category:'', t7:'', t8:'' }); save(); partFilter=''; byId('part-filter').value=''; renderParts(); };
byId('part-filter').addEventListener('input', e => { partFilter = e.target.value; renderParts(); });
byId('part-clear').onclick = () => { if (state.parts.length && confirm('Clear all parts?')){ state.parts=[]; save(); renderParts(); } };
byId('part-paste').onclick = () => openModal('Paste parts (copy two columns from Excel)',
  'Each line: part name + difficulty, separated by Tab or comma. Difficulty can be High/Low/None (高/低/无 also accepted). Leave it blank if there is none.',
  'Custom housing\tHigh\nStandard bracket\tLow\nStandard screw pack\tNone', text => {
    let n=0;
    text.split('\n').forEach(line => {
      if (!line.trim()) return;
      const cols = line.split(/\t|,/).map(c=>c.trim());
      const name = cols[0]; if (!name) return;
      const cat = normCategory(cols[1] || '');
      const ex = state.parts.find(p => p.name === name);
      if (ex) ex.category = cat || ex.category; else state.parts.push({ name, category:cat });
      n++;
    });
    save(); renderParts(); toast(`Imported ${n} parts`, 'ok');
  });

/* =========================================================
   WORKERS (matrix)
   ========================================================= */
function renderWorkers(){
  renderLeaves();   // leave records (rendered independently, not blocked by "add products first")
  const noprod = byId('worker-noproducts');
  const table = byId('worker-table');
  if (!state.products.length){ noprod.hidden = false; table.style.display='none'; updateCounts(); return; }
  noprod.hidden = true; table.style.display='';
  const thead = table.querySelector('thead'); const tb = table.querySelector('tbody');
  thead.innerHTML=''; tb.innerHTML='';
  const prods = state.products;                          // show ALL products
  // two-row header: product name spans Design|Check; second row = the sub-columns
  const r1 = el('tr', {}, [ el('th', { class:'rownum', rowspan:'2' }, '#'), el('th', { class:'wname', rowspan:'2' }, 'Worker name') ]);
  prods.forEach(p => r1.append(el('th', { class:'role-head grp-start', colspan:'2' }, p)));
  r1.append(el('th', { class:'act', rowspan:'2' }, ''));
  const r2 = el('tr', {});
  prods.forEach(() => { r2.append(el('th', { class:'sub-d grp-start' }, 'Design')); r2.append(el('th', { class:'sub-v' }, 'Check')); });
  thead.append(r1, r2);

  const mkSel = (g, role) => {
    const sel = el('select', { class:'mini-grade g-'+(g[role]||''), title: ROLE_LABEL[role] });
    GRADES.forEach(v => sel.append(el('option', { value:v, selected:v===(g[role]||'') }, v || '—')));
    sel.addEventListener('change', () => { g[role] = sel.value; sel.className='mini-grade g-'+sel.value; save(); });
    return sel;
  };
  state.workers.slice(0, ROW_CAP).forEach((w, i) => {
    const nameI = el('input', { value:w.name, placeholder:'Name' });
    nameI.addEventListener('input', () => { w.name = nameI.value; save(); });
    const tr = el('tr', {}, [ el('td', { class:'rownum' }, String(i+1)), el('td', { class:'matrix-name' }, nameI) ]);
    prods.forEach(p => {
      if (!w.grades[p] || typeof w.grades[p] !== 'object') w.grades[p] = emptyGrade();
      const g = w.grades[p];
      tr.append(el('td', { class:'gradecell grp-start' }, mkSel(g,'d')));
      tr.append(el('td', { class:'gradecell' }, mkSel(g,'v')));
    });
    tr.append(el('td', { class:'act' }, el('button', { class:'row-del', title:'Delete', onclick:()=>{ state.workers.splice(i,1); save(); renderWorkers(); } }, '✕')));
    tb.append(tr);
  });
  if (state.workers.length > ROW_CAP)
    tb.append(el('tr', {}, el('td', { colspan: String(2 + prods.length*2 + 1), style:'padding:12px;color:var(--muted);text-align:center' },
      `Showing the first ${ROW_CAP} (of ${state.workers.length}).`)));
  updateCounts();
}
function addWorker(name){
  const g = {}; state.products.forEach(p => g[p]=emptyGrade());
  state.workers.push({ name: clean(name), grades:g });
}
byId('worker-add').onclick = () => { if (!state.products.length){ toast('Add a product first', 'err'); return; } addWorker(''); save(); renderWorkers(); };
byId('worker-clear').onclick = () => { if (state.workers.length && confirm('Clear all workers?')){ state.workers=[]; save(); renderWorkers(); } };

/* ---------- Leave records (only affect the Schedule's daily team capacity) ---------- */
function leaveDailyHours(L){                       // leave hours per day: blank/invalid = full shift H
  const v = +L.hours;
  return (L.hours !== '' && L.hours != null && isFinite(v) && v > 0) ? Math.min(HOURS_PER_DAY, v) : HOURS_PER_DAY;
}
function leaveSpanDays(L){                         // calendar days, start and end inclusive
  const s = parseISO(L.start); if (!s) return 0;
  let e = parseISO(L.end) || s; if (e.getTime() < s.getTime()) e = s;
  return Math.round((e.getTime() - s.getTime()) / 86400000) + 1;
}
function leaveTotalText(L){
  const d = leaveSpanDays(L); if (!d) return '—';
  const h = leaveDailyHours(L);
  return `${round(d * h, 1)}h (${d} days × ${fmtHrs(h)}h)`;
}
function workerNameOptions(selected){
  const names = [...new Set(state.workers.map(w => w.name).filter(Boolean))];
  if (selected && !names.includes(selected)) names.unshift(selected);   // keep names of deleted workers so records aren't lost
  return names;
}
function renderLeaves(){
  const wrap = byId('leave-wrap'); if (!wrap) return;
  const empty = byId('leave-empty');
  const leaves = state.leaves || (state.leaves = []);
  wrap.innerHTML = '';
  if (!leaves.length){ if (empty) empty.hidden = false; return; }
  if (empty) empty.hidden = true;
  const t = el('table', { class:'grid' });
  t.append(el('thead', {}, el('tr', {}, [
    el('th',{class:'rownum'},'#'), el('th',{style:'width:200px'},'Worker'),
    el('th',{style:'width:150px'},'From (incl.)'), el('th',{style:'width:150px'},'To (incl.)'),
    el('th',{style:'width:150px'},`Leave per day (h)`), el('th',{class:'num',style:'width:150px'},'Total hours'),
    el('th',{class:'act'},'') ])));
  const tb = el('tbody', {});
  leaves.forEach((L, i) => {
    const sel = el('select', {});
    sel.append(el('option', { value:'', selected:!L.worker }, '(choose worker)'));
    workerNameOptions(L.worker).forEach(n => sel.append(el('option', { value:n, selected:n===L.worker }, n)));
    sel.addEventListener('change', () => { L.worker = sel.value; save(); });
    const totCell = el('td', { class:'num' }, leaveTotalText(L));
    const si = el('input', { type:'date', value:L.start || '' });
    si.addEventListener('change', () => { L.start = si.value; if (!L.end) L.end = si.value; save(); renderLeaves(); });
    const ei = el('input', { type:'date', value:L.end || '' });
    ei.addEventListener('change', () => { L.end = ei.value; save(); renderLeaves(); });
    const hi = el('input', { type:'number', step:'0.5', min:'0', max:String(HOURS_PER_DAY),
      placeholder:`Full day (${fmtHrs(HOURS_PER_DAY)})`, value:(L.hours===''||L.hours==null) ? '' : L.hours, style:'width:110px' });
    hi.addEventListener('input', () => {                        // update the total while typing, without re-rendering the table (keeps focus)
      L.hours = hi.value === '' ? '' : (isNaN(parseFloat(hi.value)) ? '' : parseFloat(hi.value));
      totCell.textContent = leaveTotalText(L); save();
    });
    hi.addEventListener('change', () => {                       // on blur, normalize invalid / over-shift values
      if (hi.value !== ''){ let n = parseFloat(hi.value);
        if (isNaN(n) || n <= 0){ hi.value = ''; L.hours = ''; }
        else { if (n > HOURS_PER_DAY) n = HOURS_PER_DAY; hi.value = String(n); L.hours = n; } }
      else L.hours = '';
      totCell.textContent = leaveTotalText(L); save();
    });
    tb.append(el('tr', {}, [
      el('td',{class:'rownum'}, String(i+1)),
      el('td',{}, sel), el('td',{}, si), el('td',{}, ei),
      el('td',{}, hi), totCell,
      el('td',{class:'act'}, el('button',{class:'row-del',title:'Delete',onclick:()=>{ state.leaves.splice(i,1); save(); renderLeaves(); }}, '✕')) ]));
  });
  t.append(tb); wrap.append(t);
}
byId('leave-add') && (byId('leave-add').onclick = () => {
  (state.leaves || (state.leaves = [])).push({ worker:'', start:'', end:'', half:false });
  save(); renderLeaves();
});
byId('worker-rand').onclick = () => {
  if (!state.products.length){ toast('Add a product first', 'err'); return; }
  if (!state.workers.length){ toast('Add or import workers first', 'err'); return; }
  const ri = (a,b) => a + Math.floor(Math.random()*(b-a+1));
  const grade = () => ['A','B','C'][ri(0,2)];
  state.workers.forEach(w => {
    const role = ri(0,4);                         // 0-1 design only, 2-3 check only, 4 both
    const doD = role <= 1 || role === 4, doV = role >= 2;
    state.products.forEach(p => {
      const g = { d:'', v:'' };
      if (doD && Math.random() < 0.85) g.d = grade();   // occasionally blank = not involved in this product
      if (doV && Math.random() < 0.85) g.v = grade();
      w.grades[p] = g;
    });
  });
  save(); renderWorkers();
  if (byId('analysis-out')) byId('analysis-out').innerHTML = '';
  toast('Generated random worker abilities (test)', 'ok');
};
byId('worker-paste').onclick = () => {
  const head = 'Worker name\t' + (state.products.length ? state.products.map(p=>`${p}·Design\t${p}·Check`).join('\t') : 'Product A·Design\tProduct A·Check');
  openModal('Paste worker ability matrix (copy from Excel)',
    'First row is the header: <b>Worker name</b> + two columns per product, <b>Product·Design</b> / <b>Product·Check</b> (设计/校验 also accepted). Fill cells with A/B/C; blank = not involved in that role. A column without a role suffix sets the same grade for both design and check.',
    head + '\nAlice\tA\tB', text => {
      const rows = text.split('\n').map(r => r.split('\t')).filter(r => r.some(c => c.trim()));
      if (rows.length < 2){ toast('Need a header row + at least one data row', 'err'); return; }
      const cols = rows[0].slice(1).map(parseWorkerColRole);     // [{product, role}]
      cols.forEach(c => { if (c.product && !state.products.includes(c.product)) addProduct(c.product); });
      let n=0;
      for (let r = 1; r < rows.length; r++){
        const cells = rows[r]; const name = clean(cells[0]); if (!name || isSkipWorker(name)) continue;
        let w = state.workers.find(x => x.name === name);
        if (!w){ w = { name, grades:{} }; state.products.forEach(p=>w.grades[p]=emptyGrade()); state.workers.push(w); }
        cols.forEach((c, ci) => {
          if (!c.product) return;
          if (!w.grades[c.product] || typeof w.grades[c.product] !== 'object') w.grades[c.product] = emptyGrade();
          const g = normGrade(cells[ci+1] || '');
          if (c.role === 'd' || c.role === '') w.grades[c.product].d = g;
          if (c.role === 'v' || c.role === '') w.grades[c.product].v = g;
        });
        n++;
      }
      save(); renderWorkers(); renderProducts(); toast(`Imported ${n} workers`, 'ok');
    });
};

/* =========================================================
   ORDERS
   ========================================================= */
const ROW_CAP = 200;   // cap rendered rows so large imports stay responsive (data is all kept)
/* a <select> that lists at most OPT options + always the currently-stored value
   (so huge worker/product lists don't build tens of thousands of <option>s per row) */
function cappedSelect(current, names, OPT = 300){
  const s = el('select', { class:'cell' });
  s.append(el('option', { value:'', selected:!current }, '—'));
  let shown = 0, hasCur = false;
  for (const nm of names){
    if (!nm) continue;
    const isCur = nm === current; if (isCur) hasCur = true;
    if (shown < OPT || isCur){ s.append(el('option', { value:nm, selected:isCur }, nm)); if (shown < OPT) shown++; }
  }
  if (current && !hasCur) s.append(el('option', { value:current, selected:true }, current + ' (not on roster)'));
  return s;
}
let orderFilter = '';
function orderMatches(o, f){
  if (!f) return true;
  return (`${o.so||''} ${o.product||''} ${o.parts||''} ${o.designer||''} ${o.checker||''} ${o.required||''} ${o.time||''} ${o.actualFinish||''}`).toLowerCase().includes(f);
}
function renderOrders(){
  const tb = byId('order-table').querySelector('tbody'); tb.innerHTML='';
  const knownNames = new Set(state.parts.map(p => p.name).filter(Boolean));
  const workerNames = state.workers.map(w => w.name).filter(Boolean);
  const f = orderFilter.toLowerCase();
  let shown = 0, matched = 0;
  state.orders.forEach((o, i) => {
    if (!orderMatches(o, f)) return;
    matched++;
    if (shown >= ROW_CAP) return;     // cap rendered rows; use search to narrow
    shown++;
    // product select (capped options + the stored value)
    const sel = cappedSelect(o.product, state.products);
    sel.addEventListener('change', () => { o.product = sel.value; save(); });

    // worker select (designer / checker) — capped options + the stored value
    const workerSel = (field) => {
      const s = cappedSelect(o[field], workerNames);
      s.addEventListener('change', () => { o[field] = s.value; save(); });
      return s;
    };
    const reqI = el('input', { type:'date', value:o.required||'' });
    reqI.addEventListener('input', () => { o.required = reqI.value; save(); });
    const dateI = el('input', { type:'date', value:o.time });
    dateI.addEventListener('input', () => { o.time = dateI.value; save(); });
    const soI = el('input', { value:o.so||'', placeholder:'e.g. 123456789-1200', style:'width:150px' });
    soI.addEventListener('input', () => { o.so = soI.value.trim(); save(); });

    tb.append(el('tr', {}, [
      el('td', { class:'rownum' }, String(i+1)),
      el('td', {}, soI),
      el('td', {}, sel),
      el('td', { class:'parts-cell' }, buildPartsPicker(o, knownNames)),
      el('td', {}, workerSel('designer')),
      el('td', {}, workerSel('checker')),
      el('td', {}, reqI),
      el('td', {}, dateI),
      el('td', { class:'act' }, el('button', { class:'row-del', title:'Delete', onclick:()=>{ state.orders.splice(i,1); save(); renderOrders(); } }, '✕')),
    ]));
  });
  if (matched > shown)
    tb.append(el('tr', {}, el('td', { colspan:'9', style:'padding:12px;color:var(--muted);text-align:center' },
      `Showing the first ${shown} (of ${matched}${f?' matching':''}${f?'':', all imported'}). ${f?'':'Use the search above to find a specific order. '}Analysis includes all ${state.orders.length} orders.`)));
  updateCounts();
}
/* parse o.parts string -> merged [{name, qty}] (display order preserved, dups summed) */
function orderPartsItems(o){
  const map = new Map();
  parseOrderParts(o.parts).forEach(({ name, qty }) => map.set(name, (map.get(name) || 0) + (qty > 0 ? qty : 1)));
  return [...map.entries()].map(([name, qty]) => ({ name, qty }));
}
function fmtQty(q){ return Number.isInteger(q) ? String(q) : String(Math.round(q * 100) / 100); }
function serializeOrderParts(items){
  return items.map(it => it.qty === 1 ? it.name : `${it.name}×${fmtQty(it.qty)}`).join('; ');
}
/* pick parts that already exist in the Parts tab; quantities via steppers.
   knownNames is built once per render (passed in) to avoid O(orders×parts) work. */
function buildPartsPicker(o, knownNames){
  if (!knownNames) knownNames = new Set(state.parts.map(p => p.name).filter(Boolean));
  const items = orderPartsItems(o);
  const wrap = el('div', { class:'parts-picker' });

  items.forEach((it, idx) => {
    const unknown = !knownNames.has(it.name);
    const qty = el('input', { class:'pp-qty', type:'number', min:'0', step:'1', value:fmtQty(it.qty), title:'Quantity' });
    qty.addEventListener('input', () => {
      const v = parseFloat(qty.value); if (isNaN(v) || v <= 0) return;
      items[idx].qty = v; o.parts = serializeOrderParts(items); save();
    });
    wrap.append(el('span', { class:'pp-chip' + (unknown ? ' unknown' : ''), title: unknown ? 'This part is not on the Parts tab' : '' }, [
      el('span', { class:'pp-name' }, it.name + (unknown ? ' (not registered)' : '')),
      el('span', {}, '×'), qty,
      el('button', { class:'pp-del', title:'Remove', onclick:()=>{ items.splice(idx, 1); o.parts = serializeOrderParts(items); save(); renderOrders(); } }, '✕'),
    ]));
  });

  if (!knownNames.size){
    wrap.append(el('span', { class:'pp-empty' }, 'Add parts on the Parts tab first'));
  } else {
    const chosen = new Set(items.map(it => it.name));
    const OPT_CAP = 300; const avail = []; let more = 0;
    for (const p of state.parts){ if (!p.name || chosen.has(p.name)) continue; if (avail.length < OPT_CAP) avail.push(p.name); else { more++; } }
    const add = el('select', { class:'pp-add' });
    add.append(el('option', { value:'', selected:true }, '+ Add VC…'));
    avail.forEach(n => add.append(el('option', { value:n }, n)));
    if (more) add.append(el('option', { value:'', disabled:true }, `…${more} more (use the filter)`));
    if (!avail.length) add.disabled = true;
    add.addEventListener('change', () => {
      const name = add.value; if (!name) return;
      items.push({ name, qty:1 }); o.parts = serializeOrderParts(items); save(); renderOrders();
    });
    wrap.append(add);
  }
  return wrap;
}
byId('order-filter') && byId('order-filter').addEventListener('input', e => { orderFilter = e.target.value; renderOrders(); });
byId('order-add').onclick = () => { state.orders.push({ product:'', parts:'', time:'', so:'' }); save(); orderFilter=''; byId('order-filter').value=''; renderOrders(); };
byId('order-clear').onclick = () => { if (state.orders.length && confirm('Clear all orders?')){ state.orders=[]; save(); renderOrders(); } };
byId('order-clear2').onclick = () => {
  if (!state.orders.length){ toast('There are no orders', 'err'); return; }
  if (confirm(`Clear all ${state.orders.length} orders?`)){ state.orders = []; save(); renderOrders(); toast('Orders cleared', 'ok'); }
};
byId('order-rand').onclick = () => {
  const partNames = state.parts.map(p => p.name).filter(Boolean);
  if (!state.products.length || !partNames.length){ toast('Add data on the Products and Parts tabs first', 'err'); return; }
  const inp = prompt('How many random test orders? (uses existing products/parts with random parts, quantities, due dates and actual finish dates)', '10');
  if (inp == null) return;
  const n = Math.max(1, Math.min(2000, parseInt(inp, 10) || 0));
  const ri = (a, b) => a + Math.floor(Math.random() * (b - a + 1));
  const pick = arr => arr[ri(0, arr.length - 1)];
  const isoOff = off => { const d = new Date(); d.setDate(d.getDate() + off); return isoLocal(d); };
  const wnames = state.workers.map(w => w.name).filter(Boolean);
  for (let k = 0; k < n; k++){
    const product = pick(state.products);
    const kParts = ri(1, Math.min(8, partNames.length));
    const chosen = partNames.slice().sort(() => Math.random() - 0.5).slice(0, kParts);
    const parts = chosen.join('; ');                 // VCs never repeat in an order
    const dueOff = ri(1, 60);                         // due date is always after today
    state.orders.push({ product, parts,
      designer: wnames.length ? pick(wnames) : '', checker: wnames.length ? pick(wnames) : '',
      required: isoOff(dueOff),                        // col6 due date (deadline)
      time: isoOff(dueOff - ri(5, 40)),               // col7 created date (before due)
      actualFinish: isoOff(dueOff + ri(-5, 15)) });   // col8 actual finish
  }
  save(); renderOrders(); toast(`Generated ${n} random test orders`, 'ok');
};
byId('order-paste').onclick = () => openModal('Paste orders (copy three columns from Excel)',
  'Each line: product type + parts needed + submitted date (optional), separated by Tab. Inside the parts cell separate parts with ; or line breaks, and write quantities as ×N.',
  'Flagship\tCustom housing×1; Standard bracket×2; Standard screw pack×4\t2026-07-05', text => {
    let n=0;
    text.split('\n').forEach(line => {
      if (!line.trim()) return;
      const c = line.split('\t');
      state.orders.push({ product: clean(c[0]), parts: clean(c[1]), time: normDate(c[2]) });
      n++;
    });
    save(); renderOrders(); toast(`Imported ${n} orders`, 'ok');
  });
function normDate(v){
  v = clean(v); if (!v) return '';
  const m = v.match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
  return v;
}

/* =========================================================
   MODAL (shared bulk-text input)
   ========================================================= */
let modalCb = null;
function openModal(title, hint, placeholder, cb){
  byId('modal-title').textContent = title;
  byId('modal-hint').innerHTML = hint;
  const ta = byId('modal-text'); ta.value=''; ta.placeholder = placeholder || '';
  modalCb = cb; byId('modal').hidden = false; ta.focus();
}
byId('modal-cancel').onclick = () => { byId('modal').hidden = true; modalCb=null; };
byId('modal').addEventListener('click', e => { if (e.target.id==='modal') byId('modal-cancel').click(); });
byId('modal-ok').onclick = () => { const v = byId('modal-text').value; byId('modal').hidden = true; const cb=modalCb; modalCb=null; if (cb) cb(v); };

/* =========================================================
   IMPORT  (xlsx / json files)
   ========================================================= */
function importMode(){ return document.querySelector('input[name=impmode]:checked').value; }
const ioLog = (() => {
  const box = byId('io-log');
  return msg => { box.hidden=false; box.textContent += (box.textContent?'\n':'') + msg; box.scrollTop = box.scrollHeight; };
})();

function aoaFromSheet(ws){ return XLSX.utils.sheet_to_json(ws, { header:1, blankrows:false, defval:'' }); }

function classifySheet(headerRow){
  const fields = headerRow.map(normHeader);
  if (fields.includes('part') && fields.includes('category')) return 'parts';
  if (fields.includes('product') && fields.includes('parts')) return 'orders';
  if (fields.includes('worker')) return 'workers';
  return null;
}

function importPartsAoa(aoa, mode){
  const hdr = aoa[0].map(normHeader);
  const ci = { part: hdr.indexOf('part'), category: hdr.indexOf('category') };
  if (ci.part < 0) ci.part = 0;
  if (ci.category < 0) ci.category = 1;
  if (mode === 'replace') state.parts = [];
  let n = 0;
  for (let r = 1; r < aoa.length; r++){
    const name = clean(aoa[r][ci.part]); if (!name) continue;
    const cat = normCategory(aoa[r][ci.category]);
    const ex = state.parts.find(p => p.name === name);
    if (ex) ex.category = cat || ex.category; else state.parts.push({ name, category: cat });
    n++;
  }
  ioLog(`  parts: imported ${n} rows`);
}
function importOrdersAoa(aoa, mode){
  const hdr = aoa[0].map(normHeader);
  let cp = hdr.indexOf('product'), cl = hdr.indexOf('parts'), cd = hdr.indexOf('due');
  if (cp < 0) cp = 0; if (cl < 0) cl = 1; if (cd < 0) cd = 2;
  if (mode === 'replace') state.orders = [];
  let n = 0;
  for (let r = 1; r < aoa.length; r++){
    const row = aoa[r];
    const product = clean(row[cp]), parts = clean(row[cl]), time = normDate(row[cd]);
    if (!product && !parts) continue;
    state.orders.push({ product, parts, time }); n++;
  }
  ioLog(`  orders: imported ${n} orders`);
}
function importWorkersAoa(aoa, mode){
  const hdr = aoa[0];
  const nameIdx = Math.max(0, hdr.map(normHeader).indexOf('worker'));
  // each non-name column -> {product, role, i}; role '' = legacy (applies to both design/check)
  const cols = hdr.map((h,i)=>({ i, ...parseWorkerColRole(h) })).filter(c => c.i !== nameIdx && c.product);
  cols.forEach(c => { if (!state.products.includes(c.product)) addProduct(c.product); });
  if (mode === 'replace') state.workers = [];
  let n = 0;
  for (let r = 1; r < aoa.length; r++){
    const row = aoa[r]; const name = clean(row[nameIdx]); if (!name || isSkipWorker(name)) continue;
    let w = state.workers.find(x => x.name === name);
    if (!w){ w = { name, grades:{} }; state.products.forEach(p=>w.grades[p]=emptyGrade()); state.workers.push(w); }
    cols.forEach(c => {
      if (!w.grades[c.product] || typeof w.grades[c.product] !== 'object') w.grades[c.product] = emptyGrade();
      const g = normGrade(row[c.i]);
      if (c.role === 'd' || c.role === '') w.grades[c.product].d = g;
      if (c.role === 'v' || c.role === '') w.grades[c.product].v = g;
    });
    n++;
  }
  ioLog(`  workers: imported ${n} (${cols.length} columns: product × design/check)`);
}

async function handleFiles(files){
  const mode = importMode();
  for (const file of files){
    try {
      const lower = file.name.toLowerCase();
      if (lower.endsWith('.json')){
        const text = await file.text();
        const data = JSON.parse(text);
        if (Array.isArray(data.products) || Array.isArray(data)){
          const list = Array.isArray(data) ? data : data.products;
          if (mode === 'replace') { state.products = []; state.workers.forEach(w => w.grades = {}); }
          list.map(String).forEach(addProduct);
          ioLog(`✔ ${file.name}: ${list.length} products`);
        } else if (data.__partsAnalysisBackup || data.parts || data.orders){
          restoreBackup(data); ioLog(`✔ ${file.name}: full backup restored`);
        } else { ioLog(`⚠ ${file.name}: unrecognized JSON`); }
        continue;
      }
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type:'array' });
      ioLog(`✔ ${file.name}`);
      // hour-report log: intercept the whole workbook (its "By worker" sheets have a Worker header and must never fall into generic detection)
      if (wb.SheetNames.includes('Data (long)')){
        importHourReport(wb, file.name);
        continue;
      }
      // try each sheet; classify by name then by header
      let handled = 0;
      wb.SheetNames.forEach(sn => {
        const aoa = aoaFromSheet(wb.Sheets[sn]);
        if (!aoa.length) return;
        let kind = sn===SHEET.parts?'parts': sn===SHEET.workers?'workers': sn===SHEET.orders?'orders': LEGACY_SHEET[sn] || classifySheet(aoa[0]);
        // filename hint
        if (!kind){ if (lower.includes('part')) kind='parts'; else if (lower.includes('worker')) kind='workers'; else if (lower.includes('order')) kind='orders'; }
        if (kind==='parts'){ importPartsAoa(aoa, mode); handled++; }
        else if (kind==='workers'){ importWorkersAoa(aoa, mode); handled++; }
        else if (kind==='orders'){ importOrdersAoa(aoa, mode); handled++; }
      });
      if (!handled) ioLog(`  ⚠ Could not recognize the sheet type (headers need to include Part name / Worker name / Product type, etc.)`);
    } catch (err){
      console.error(err); ioLog(`✖ ${file.name}: ${err.message}`);
    }
  }
  save(); renderAll();
  toast('Import complete', 'ok');
}

const fileInput = byId('file-input');
byId('dropzone').addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', e => { if (e.target.files.length) handleFiles([...e.target.files]); fileInput.value=''; });
const dz = byId('dropzone');
['dragover','dragenter'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
['dragleave','drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
dz.addEventListener('drop', e => { if (e.dataTransfer.files.length) handleFiles([...e.dataTransfer.files]); });

// workers.xlsx import stays on the generic auto-detect picker
byId('worker-import') && (byId('worker-import').onclick = () => fileInput.click());

/* =========================================================
   HOURLOG  (hour log — HourReport_{year}_Analysis.xlsx exported by hour-report)
   Detected by the sheet "Data (long)". Fixed column order (columns are only ever appended, never reordered):
   0 Year · 2 Month# · 3 Group · 4 Team · 5 Worker# · 6 Worker · 7 Role · 9 Task · 12 Hours
   Leave sheet: 5 Worker# · 6 Worker · 7 Category · 8 Value
   Only aggregated results (worker × task/month) are stored, not raw rows, so localStorage doesn't blow up.
   ========================================================= */
function importHourReport(wb, filename){
  const aoa = aoaFromSheet(wb.Sheets['Data (long)']);
  const wmap = new Map();                       // key = Worker#¦name
  const years = new Map(); const months = new Set(); const taskTotals = new Map();
  let nRows = 0;
  for (let r = 1; r < aoa.length; r++){
    const row = aoa[r];
    const name = clean(row[6]); const hours = Number(row[12]);
    if (!name || !isFinite(hours) || hours <= 0) continue;
    const key = clean(row[5]) + '¦' + name;
    let w = wmap.get(key);
    if (!w){ w = { no: clean(row[5]), name, group: clean(row[3]), team: clean(row[4]), role: clean(row[7]),
      total: 0, byTask: {}, byMonth: {}, leave: {}, leaveTotal: 0 }; wmap.set(key, w); }
    const task = TASK_ALIASES[clean(row[9])] || clean(row[9]) || '(unnamed task)'; const m = Number(row[2]);
    w.total += hours;
    w.byTask[task] = (w.byTask[task] || 0) + hours;
    if (m >= 1 && m <= 12){ w.byMonth[m] = (w.byMonth[m] || 0) + hours; months.add(m); }
    taskTotals.set(task, (taskTotals.get(task) || 0) + hours);
    const y = Number(row[0]); if (isFinite(y) && y > 2000) years.set(y, (years.get(y) || 0) + 1);
    nRows++;
  }
  if (!nRows){ ioLog('  ⚠ No usable hour rows in Data (long)'); return; }
  if (wb.SheetNames.includes('Leave')){
    const la = aoaFromSheet(wb.Sheets['Leave']);
    for (let r = 1; r < la.length; r++){
      const row = la[r]; const name = clean(row[6]); const val = Number(row[8]);
      if (!name || !isFinite(val) || val === 0) continue;
      const key = clean(row[5]) + '¦' + name;
      let w = wmap.get(key);
      if (!w){ w = { no: clean(row[5]), name, group: clean(row[3]), team: clean(row[4]), role: '',
        total: 0, byTask: {}, byMonth: {}, leave: {}, leaveTotal: 0 }; wmap.set(key, w); }
      const cat = clean(row[7]) || '(uncategorized)';
      w.leave[cat] = (w.leave[cat] || 0) + val; w.leaveTotal += val;
    }
  }
  const year = [...years.entries()].sort((a, b) => b[1] - a[1]).map(e => e[0])[0] || null;
  state.hourlog = {
    year, source: filename || '', importedAt: todayISO(),
    months: [...months].sort((a, b) => a - b),
    tasks: [...taskTotals.entries()].map(([name, hours]) => ({ name, hours })).sort((a, b) => b.hours - a.hours),
    workers: [...wmap.values()].sort((a, b) => b.total - a.total),
  };
  ioLog(`  Hour log: ${state.hourlog.workers.length} workers · ${nRows} hour rows · ${state.hourlog.months.length} months (${year || '?'}) → see the Hour Log tab`);
}

byId('hourlog-import') && (byId('hourlog-import').onclick = () => fileInput.click());
byId('hourlog-clear') && (byId('hourlog-clear').onclick = () => {
  if (state.hourlog && confirm('Clear the imported hour log? (other data is not affected)')){ state.hourlog = null; save(); renderHourlog(); }
});
byId('hourlog-export') && (byId('hourlog-export').onclick = () =>
  exportReportFromDOM('hourlog-out', `Hour-log-${(state.hourlog && state.hourlog.year) || todayISO()}.xlsx`, 'Hour log (actual logged hours)'));

function renderHourlog(){
  const out = byId('hourlog-out'); if (!out) return; out.innerHTML = '';
  const HL = state.hourlog;
  if (!HL){
    out.append(el('div',{class:'empty-note'}, 'No hour log yet. In hour-report, click "Export Excel" on the Task Totals page to get HourReport_{year}_Analysis.xlsx, then click "Import" above or drop it on the Import/Export tab.'));
    return;
  }
  const td = (v,num)=>el('td', num?{class:'num'}:{}, v);
  const section = (t,sub)=>out.append(el('div',{class:'section-title'}, sub?el('span',{},[t,el('span',{style:'color:var(--muted);font-weight:500'},'  '+sub)]):t));
  const table = (headers, rows)=>{ const x=el('table',{class:'report'}); const headTr = el('tr',{}, headers.map(h=>el('th',{},h))); x.append(el('thead',{}, headTr)); const b=el('tbody',{}); rows.forEach(r=>b.append(r)); x.append(b); alignNumHeaders(headTr, rows[0]); out.append(el('div',{class:'report-wrap'}, x)); };
  const pct = x => x==null ? '—' : round(100*x,0)+'%';

  const design = HL.workers.reduce((a,w)=>a+(w.byTask['Order design']||0),0);
  const verify = HL.workers.reduce((a,w)=>a+(w.byTask['Order check']||0),0);
  const ot = HL.workers.reduce((a,w)=>a+(w.byTask['Overtime']||0),0);
  const total = HL.workers.reduce((a,w)=>a+w.total,0);
  const M = HL.months.length;
  // utilization is always based on the standard 8h shift (not HOURS_PER_DAY — once that's replaced by actual order hours it would be distorted here)
  const AL = state.settings.annualLeave >= 0 ? state.settings.annualLeave : 15;
  const capPerWorker = M > 0 ? (260 - AL) / 12 * M * DEFAULT_HOURS_PER_DAY : 0;

  const cards = el('div',{class:'cards'});
  const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  card('Year', HL.year || '—');
  card('Months covered', HL.months.join(', ') || '—', `(${M} months)`);
  card('Workers', HL.workers.length, '');
  card('Total logged', round(total), 'hours');
  card('Order design', round(design), 'hours');
  card('Order check', round(verify), 'hours');
  if (ot > 0) card('Overtime', round(ot), 'hours');
  out.append(cards);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    `Source ${HL.source || '—'} · imported ${HL.importedAt} · utilization uses a standard shift: ${M} months per person ≈ ${round(capPerWorker)}h (260 workdays per year − ${AL} days annual leave, × ${fmtHrs(DEFAULT_HOURS_PER_DAY)}h standard shift).`));

  // ---- actual-hours basis: hours per person per day actually spent on order design/check; can replace the global "hours per person per day" in one click ----
  const owOf = w => (w.byTask['Order design'] || 0) + (w.byTask['Order check'] || 0);
  const wdOf = w => Object.keys(w.byMonth || {}).length * 260 / 12;   // months with logged hours × ~21.7 workdays per month
  const orderWorkers = HL.workers.filter(w => owOf(w) > 0);
  const owSum = orderWorkers.reduce((a, w) => a + owOf(w), 0);
  const wdSum = orderWorkers.reduce((a, w) => a + wdOf(w), 0);
  const actualRate = wdSum > 0 ? round(owSum / wdSum, 2) : null;
  // actual check ratio = total order-check hours ÷ total order-design hours in the log
  const actualVR = design > 0 && verify > 0 ? round(verify / design, 2) : null;
  if (actualRate){
    const applied = Math.abs(HOURS_PER_DAY - actualRate) < 0.005 && (actualVR == null || Math.abs(VERIFY_RATIO - actualVR) < 0.005);
    const btn = el('button', { class:'btn primary', style:'margin-left:8px' }, applied ? '✔ Already applied' : 'Apply actual basis');
    if (applied) btn.disabled = true;
    btn.onclick = () => {
      applyHoursPerDay(actualRate);
      if (actualVR != null) applyVerifyRatio(actualVR);      // check ratio switches to the actual ratio (no longer fixed at 0.3)
      state.settings.annualLeave = 0;                        // actual logged hours already include leave; zero schedule annual leave to avoid double-counting
      save(); renderHourlog(); renderVerifySettings();
      ['analysis-out','sched-out','retro-out'].forEach(id => { const n = byId(id); if (n) n.innerHTML = ''; });
      toast(`Applied actual basis: hours per person per day ${actualRate}h${actualVR!=null?`, check ratio ${actualVR}`:''}, schedule annual leave = 0`, 'ok');
    };
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},[
      el('b',{},'📐 Actual-hours basis: '),
      `Across the ${orderWorkers.length} workers in the log with order hours, each person actually spends on average `,
      el('b',{}, `${actualRate} hours`),
      ` per day on order design/check (the rest goes to communication / meetings / support)`,
      actualVR != null ? el('span',{},[ `; actual check/design ratio = ${round(verify)}h ÷ ${round(design)}h = `, el('b',{}, String(actualVR)),
        ` (current check ratio ${VERIFY_RATIO})` ]) : '',
      `. Current "hours per person per day" = ${fmtHrs(HOURS_PER_DAY)}h. Click the button to apply the actual basis: ① hours per person per day → ${actualRate}h ("workdays" and capacity in Analysis / Retrospective / Schedule follow the real pace); ` +
      (actualVR != null ? `② check ratio → ${actualVR} (check hours = design × actual ratio, no longer fixed at 0.3); ` : '') +
      `③ schedule "annual leave" → 0 (actual data already includes leave, avoids double-counting). You can change these back any time in the Schedule/Retrospective settings and on the Parts tab.`,
      btn ]));
  }

  // name reconciliation (both directions, color-coded; names must match exactly to link up)
  const known = new Set(state.workers.map(w => w.name).filter(Boolean));
  const hlNames = new Set(HL.workers.map(w => w.name));
  const unknown = HL.workers.filter(w => !known.has(w.name)).map(w => w.name);        // 🟠 in log, not on roster
  const rosterOnly = state.workers.map(w => w.name).filter(n => n && !hlNames.has(n)); // 🔵 on roster, not in log
  if (unknown.length && state.workers.length)
    out.append(el('div',{class:'alert warn'}, `🟠 ${unknown.length} people logged hours but are not on the Workers list (orange rows below): ${unknown.slice(0,10).join(', ')}${unknown.length>10?` and ${unknown.length - 10} more`:''}. If the roster uses a different spelling (e.g. an English name), make both sides match and they will link up automatically.`));
  if (rosterOnly.length)
    out.append(el('div',{class:'alert info'}, `🔵 ${rosterOnly.length} people are on the Workers list but have no logged hours under that name: ${rosterOnly.slice(0,10).join(', ')}${rosterOnly.length>10?` and ${rosterOnly.length - 10} more`:''}. Either they really logged nothing, or the names are spelled differently — these people are also shown in blue under "Retrospective · Load per worker".`));

  section('Per worker · actual hours for the year', 'Sorted by total hours; utilization = total logged ÷ capacity for the period (basis above); 🟠 orange = not on the Workers list');
  const otherOf = w => Math.max(0, w.total - (w.byTask['Order design']||0) - (w.byTask['Order check']||0) - (w.byTask['Overtime']||0));
  table(['Worker','Group/Team','Role','Total hours','Order design','Order check','Overtime','Other tasks','Order hours/day','Leave','Utilization'],
    HL.workers.slice(0, ROW_CAP).map(w => {
      const util = capPerWorker > 0 ? w.total / capPerWorker : null;
      const wd = wdOf(w); const rate = (wd > 0 && owOf(w) > 0) ? round(owOf(w) / wd, 1) : null;
      return el('tr',{class: (state.workers.length && !known.has(w.name)) ? 'unmatched-hourlog' : ((util!=null && util>1.000001)?'infeasible':'')},[
        td(w.name), td((w.group?w.group+' / ':'')+(w.team||'—')), td(w.role||'—'),
        td(round(w.total)+'h',true), td(round(w.byTask['Order design']||0)+'h',true), td(round(w.byTask['Order check']||0)+'h',true),
        td(round(w.byTask['Overtime']||0)+'h',true), td(round(otherOf(w))+'h',true),
        td(rate==null?'—':rate+'h', true),
        td(w.leaveTotal ? round(w.leaveTotal,1) : '—', true),
        td(util==null?'—':pct(util), true) ]);
    }));
  if (HL.workers.length > ROW_CAP) out.append(el('p',{class:'hint'}, `Showing the first ${ROW_CAP} (of ${HL.workers.length}).`));
  out.append(el('p',{class:'hint'}, '"Leave" is the raw total per category from the hour-report Leave sheet (units as in that sheet); rows with utilization over 100% are red (usually overtime / cross-team support).'));

  section('Task totals', 'Everyone, whole year, sorted by hours');
  table(['Task','Total hours','Share'],
    HL.tasks.map(t => el('tr',{},[ td(t.name), td(round(t.hours)+'h',true), td(total>0?pct(t.hours/total):'—',true) ])));
}

function numOrBlank(v){
  if (v == null || String(v).trim() === '') return '';
  const n = Number(v); return isNaN(n) ? '' : String(n);
}
function pickFile(accept, cb){
  const f = el('input', { type:'file', accept, style:'display:none' });
  document.body.append(f);
  f.addEventListener('change', async e => {
    const file = e.target.files[0]; f.remove(); if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:true });
      cb(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, blankrows:false, defval:'' }));
    } catch (err){ console.error(err); toast('Import failed: ' + err.message, 'err'); }
  });
  f.click();
}

// parts.xlsx — positional: col1=VC code, col7=repeat design hours, col8=first-order design hours
byId('part-import') && (byId('part-import').onclick = () => pickFile('.xlsx,.xls', aoa => {
  if (aoa.length < 2){ toast('The file has no data rows', 'err'); return; }
  if (state.parts.length && confirm(`Replace the existing ${state.parts.length} VCs with the file's ${aoa.length-1} rows? (Cancel = merge)`)) state.parts = [];
  let n = 0;
  aoa.slice(1).forEach(r => {
    const code = clean(r[0]); if (!code) return;
    const t7 = numOrBlank(r[7]), t8 = numOrBlank(r[8]);     // repeat = column 8, first order = column 9
    const ex = state.parts.find(p => p.name === code);
    if (ex){ ex.t7 = t7; ex.t8 = t8; } else state.parts.push({ name:code, category:'', t7, t8 });
    n++;
  });
  save(); renderParts(); toast(`Imported ${n} VCs (set difficulty manually)`, 'ok');
}));

/* =========================================================
   Smart order import: find the header row automatically, match columns to fields by synonyms,
   show a mapping confirmation dialog (adjustable by hand), then bulk import. If nothing is recognized, prefill with the old template's column order.
   ========================================================= */
const ORDER_FIELDS = [
  { key:'salesOrder', label:'Sales Order (optional)', syn:['salesorder','销售订单','so号'], excl:['item','行项目'] },
  { key:'orderItem',  label:'Sales Order Item (optional)', syn:['salesorderitem','orderitem','行项目','item'] },
  { key:'vc',       label:'VC list (required)', syn:['vc','零件','图号','part'] },
  { key:'product',  label:'Product',         syn:['产品','型号','product'] },
  { key:'designer', label:'Designer',        syn:['设计','designer'], excl:['工时','时间','日期','小时','hours'] },
  { key:'checker',  label:'Checker',         syn:['校对','校验','审核','checker'], excl:['工时','时间','日期','小时','hours'] },
  { key:'required', label:'Due date',        syn:['交付','交期','截止','要求完成','deadline','due','required'] },
  { key:'time',     label:'Created date',    syn:['创建','建单','提交','下单','开单','created','create on','createon','create','submit'] },
  { key:'finish',   label:'Actual finish date', syn:['实际完成','完成日期','完成时间','实际结束','finish','结案'] },
];
const POSITIONAL_GUESS = { vc:0, designer:2, checker:3, product:4, required:5, time:6, finish:7 };   // old template column order
function normHeaderCell(s){ return String(s == null ? '' : s).toLowerCase().replace(/[\s　_\-：:（）()\[\]]/g, ''); }
function matchOrderField(f, h){
  if (!h) return false;
  if (f.excl && f.excl.some(x => h.includes(x))) return false;
  return f.syn.some(s => h.includes(normHeaderCell(s)));
}
// scan the whole sheet for the row that looks most like a header: greedily claim columns in field order; the row matching the most fields wins (needs ≥2).
// a row matching ≥5 fields is taken as the header right away (a data row almost never matches 5), avoiding a full scan of huge files.
function detectOrderHeader(aoa){
  let best = null;
  for (let r = 0; r < aoa.length; r++){
    const cells = (aoa[r] || []).map(normHeaderCell);
    if (!cells.some(Boolean)) continue;                 // empty row, skip
    const map = {}; const claimed = new Set();
    ORDER_FIELDS.forEach(f => {
      for (let c = 0; c < cells.length; c++){
        if (claimed.has(c)) continue;
        if (matchOrderField(f, cells[c])){ map[f.key] = c; claimed.add(c); break; }
      }
    });
    const score = Object.keys(map).length;
    if (score >= 2 && (!best || score > best.score)) best = { row: r, map, score };
    if (score >= 5) break;                              // confident enough, this is it
  }
  return best;
}
function commitOrdersImport(aoa, headerRow, map){
  const rows = aoa.slice(headerRow + 1);
  const get = (r, key) => map[key] != null && map[key] !== '' ? r[map[key]] : '';
  if (state.orders.length && confirm(`Replace the existing ${state.orders.length} orders with the file's ${rows.length} rows? (Cancel = merge/append)`)) state.orders = [];
  let n = 0;
  const haveP = new Set(state.products), newProds = [];
  const haveV = new Set(state.parts.map(p => p.name)), newVcs = [];
  const haveW = new Set(state.workers.map(w => w.name)), newWorkers = [];
  rows.forEach(r => {
    const vcs = String(get(r,'vc') || '').split(/[;；、,，]/).map(s => s.trim()).filter(Boolean);
    const product = clean(get(r,'product'));
    const designer = clean(get(r,'designer')), checker = clean(get(r,'checker'));
    if (!vcs.length && !product) return;
    // order number = Sales Order + '-' + Item (leading zeros stripped from Item; all zeros = 0)
    const soBase = String(get(r,'salesOrder') ?? '').trim();
    let soItem = String(get(r,'orderItem') ?? '').trim();
    if (soItem) soItem = soItem.replace(/^0+/, '') || '0';
    const so = soBase ? (soItem ? `${soBase}-${soItem}` : soBase) : '';
    state.orders.push({
      product, parts: vcs.join('; '), designer, checker, so,
      required: cellToISO(get(r,'required')),
      time: cellToISO(get(r,'time')),
      actualFinish: cellToISO(get(r,'finish')),
    });
    if (product && !haveP.has(product)){ haveP.add(product); newProds.push(product); }
    vcs.forEach(v => { if (!haveV.has(v)){ haveV.add(v); newVcs.push(v); } });
    [designer, checker].forEach(nm => { if (nm && !isSkipWorker(nm) && !haveW.has(nm)){ haveW.add(nm); newWorkers.push(nm); } });
    n++;
  });
  newProds.forEach(p => state.products.push(p));
  state.workers.forEach(w => newProds.forEach(p => { if (!(p in w.grades)) w.grades[p] = emptyGrade(); }));
  newVcs.forEach(v => state.parts.push({ name: v, category:'', t7:'', t8:'' }));
  newWorkers.forEach(nm => { const g = {}; state.products.forEach(p => g[p] = emptyGrade()); state.workers.push({ name: nm, grades: g }); });
  save(); renderOrders(); renderParts(); renderProducts(); renderWorkers();
  toast(`Imported ${n} orders` + (newWorkers.length?`, added ${newWorkers.length} workers`:'') + (newVcs.length?` and ${newVcs.length} VCs`:''), 'ok');
}
function smartImportOrders(aoa){
  if (!aoa.length){ toast('The file is empty', 'err'); return; }
  const det = detectOrderHeader(aoa);
  const headerRow = det ? det.row : 0;
  const map = det ? { ...det.map } : { ...POSITIONAL_GUESS };
  const header = aoa[headerRow] || [];
  const nCols = Math.max(header.length, ...(aoa.slice(headerRow+1, headerRow+4).map(r => (r||[]).length)), 1);
  const sampleOf = c => { for (let r = headerRow+1; r < Math.min(headerRow+6, aoa.length); r++){
    const v = (aoa[r]||[])[c]; if (v != null && String(v).trim() !== '') return String(v).slice(0, 14); } return ''; };
  // mapping confirmation dialog
  const backdrop = el('div', { class:'modal-backdrop' });
  const box = el('div', { class:'modal', style:'max-width:640px' });
  box.append(el('h3', {}, det ? `Import orders — header found (row ${headerRow+1}, ${det.score} fields matched)` : 'Import orders — no header found, prefilled with the old template column order'));
  box.append(el('p', { class:'hint' }, `${aoa.length - headerRow - 1} data rows. Confirm which column each field comes from (the first sample value is shown in the dropdown); choose "Don't import" for fields you don't need.`));
  const sels = {};
  const grid = el('div', { style:'display:grid;grid-template-columns:auto 1fr;gap:8px 12px;align-items:center;margin:10px 0' });
  ORDER_FIELDS.forEach(f => {
    const sel = el('select', {});
    sel.append(el('option', { value:'', selected: map[f.key] == null }, "(Don't import)"));
    for (let c = 0; c < nCols; c++){
      const label = `Col ${c+1}: ${String(header[c] ?? '').slice(0,16) || '(no header)'}${sampleOf(c) ? ` | e.g. ${sampleOf(c)}` : ''}`;
      sel.append(el('option', { value:String(c), selected: map[f.key] === c }, label));
    }
    sels[f.key] = sel;
    grid.append(el('label', {}, f.label), sel);
  });
  box.append(grid);
  const close = () => backdrop.remove();
  const cancel = el('button', { class:'btn', onclick: close }, 'Cancel');
  const ok = el('button', { class:'btn primary', onclick: () => {
    const m = {}; ORDER_FIELDS.forEach(f => { const v = sels[f.key].value; m[f.key] = v === '' ? null : parseInt(v, 10); });
    if (m.vc == null){ toast('You must pick a column for "VC list"', 'err'); return; }
    const dup = Object.values(m).filter(v => v != null);
    if (new Set(dup).size !== dup.length){ toast('Two fields use the same column, please check', 'err'); return; }
    close(); commitOrdersImport(aoa, headerRow, m);
  } }, '✔ Import');
  box.append(el('div', { style:'display:flex;gap:8px;justify-content:flex-end;margin-top:6px' }, [cancel, ok]));
  backdrop.append(box);
  backdrop.addEventListener('click', e => { if (e.target === backdrop) close(); });
  document.body.append(backdrop);
}
byId('order-import') && (byId('order-import').onclick = () => pickFile('.xlsx,.xls', smartImportOrders));

/* =========================================================
   EXPORT  (xlsx / json, analyze.py-compatible)
   ========================================================= */
function setCols(ws, widths){ ws['!cols'] = widths.map(w => ({ wch: w })); }
function buildPartsWb(){
  let aoa, widths;
  if (hoursDefault()){
    // default times -> keep the original 2-column template (byte-identical to before)
    aoa = [HDR.parts, ...state.parts.filter(p=>p.name).map(p => [p.name, p.category])];
    widths = [16, 20];
  } else {
    // custom times -> add a per-part "Hours per piece" column
    aoa = [[...HDR.parts, 'Hours per piece (h)'],
      ...state.parts.filter(p=>p.name).map(p => [p.name, p.category,
        CATEGORIES.includes(p.category) ? CATEGORY_HOURS[p.category] : ''])];
    widths = [16, 20, 14];
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa); setCols(ws, widths);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, SHEET.parts); return wb;
}
function buildWorkersWb(){
  // two columns per product: "<product>·Design" "<product>·Check"
  const header = [HDR.workersName];
  state.products.forEach(p => { header.push(`${p}·Design`, `${p}·Check`); });
  const rows = state.workers.filter(w=>w.name).map(w => {
    const row = [w.name];
    state.products.forEach(p => { const g = w.grades[p] || {}; row.push(g.d || '', g.v || ''); });
    return row;
  });
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]); setCols(ws, [14, ...state.products.flatMap(()=>[12,12])]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, SHEET.workers); return wb;
}
function buildOrdersWb(){
  const aoa = [HDR.orders, ...state.orders.filter(o=>o.product||o.parts).map(o => [o.product, o.parts, o.time])];
  const ws = XLSX.utils.aoa_to_sheet(aoa); setCols(ws, [14, 56, 26]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, SHEET.orders); return wb;
}
function exportOne(kind){
  if (kind==='parts'){ XLSX.writeFile(buildPartsWb(), 'parts.xlsx'); }
  else if (kind==='workers'){ XLSX.writeFile(buildWorkersWb(), 'workers.xlsx'); }
  else if (kind==='orders'){ XLSX.writeFile(buildOrdersWb(), 'orders.xlsx'); }
  else if (kind==='products'){ downloadText('products.json', JSON.stringify({ products: state.products }, null, 2)); }
}
document.querySelectorAll('[data-export]').forEach(b => b.onclick = () => exportOne(b.dataset.export));

/* Export the rendered Analysis/Schedule/Retrospective result to a local .xlsx (what you see = what you get).
   Walks the -out container in document order: cards -> Summary sheet, each table -> its own sheet,
   alert text -> Notes sheet. 100% local (SheetJS), no network. */
function sanitizeSheetName(name, used){
  let s = (name || 'Sheet').replace(/[:\\\/\?\*\[\]]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 28) || 'Sheet';
  let n = s, i = 2; while (used.has(n)){ n = `${s.slice(0,26)} ${i++}`; } used.add(n); return n;
}
function tableToAoa(tbl){
  const aoa = [];
  tbl.querySelectorAll('thead tr').forEach(tr => aoa.push([...tr.children].map(c => c.textContent.trim())));
  tbl.querySelectorAll('tbody tr').forEach(tr => aoa.push([...tr.children].map(c => c.textContent.trim())));
  return aoa;
}
function exportReportFromDOM(containerId, filename, title, extraSheets){
  const box = byId(containerId);
  if (!box || !box.children.length || !box.textContent.trim()){ toast('Run it once first, then export the result', 'err'); return; }
  const used = new Set();
  const wb = XLSX.utils.book_new();
  // Summary: cards (k/v) + alert lines
  const summary = [[title], [`Exported ${new Date().toLocaleString()}`], []];
  box.querySelectorAll('.cards .stat').forEach(s => {
    const k = (s.querySelector('.k')||{}).textContent || '';
    const v = (s.querySelector('.v')||{}).textContent || '';
    if (k.trim()) summary.push([k.trim(), v.trim()]);
  });
  const alerts = [...box.querySelectorAll('.alert')].map(a => a.textContent.trim().replace(/\s+/g,' ')).filter(Boolean);
  if (alerts.length){ summary.push([]); summary.push(['Conclusions / notes']); alerts.forEach(t => summary.push([t])); }
  const sumWs = XLSX.utils.aoa_to_sheet(summary); setCols(sumWs, [26, 60]);
  XLSX.utils.book_append_sheet(wb, sumWs, sanitizeSheetName('Summary', used));
  // each table -> its own sheet, named by the nearest preceding section title
  let lastTitle = '';
  box.querySelectorAll('.section-title, table.report').forEach(node => {
    if (node.classList.contains('section-title')){ lastTitle = node.textContent.trim().split(/\s{2,}/)[0]; return; }
    const aoa = tableToAoa(node); if (!aoa.length) return;
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    setCols(ws, aoa[0].map((_, c) => Math.min(40, Math.max(8, ...aoa.map(r => (r[c]||'').length + 2)))));
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(lastTitle || 'Details', used));
  });
  // extra sheets: generated straight from data (not affected by the page's "first N rows only" truncation)
  (extraSheets || []).forEach(s => {
    if (!s || !Array.isArray(s.aoa) || !s.aoa.length) return;
    const ws = XLSX.utils.aoa_to_sheet(s.aoa);
    setCols(ws, s.aoa[0].map((_, c) => Math.min(40, Math.max(8, ...s.aoa.map(r => String(r[c] == null ? '' : r[c]).length + 2)))));
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(s.name || 'Appendix', used));
  });
  XLSX.writeFile(wb, filename);
  toast('Report exported: ' + filename, 'ok');
}
byId('analysis-export') && (byId('analysis-export').onclick = () =>
  exportReportFromDOM('analysis-out', `Analysis-report-${todayISO()}.xlsx`, 'Analysis report (theoretical hours vs actual logged)'));
byId('sched-export') && (byId('sched-export').onclick = () => {
  // full timeline (every workday, untruncated) as an extra sheet
  const extra = [];
  if (lastSched && lastSched.days && lastSched.days.length){
    const S = lastSched;
    const dlByIdx = new Map(); S.deadlines.forEach(d => dlByIdx.set(d.T - 1, d));
    const n2 = x => Math.round(x * 100) / 100;   // write real numbers so Excel can sum/chart them directly
    const aoa = [['Date','Day','Capacity (h)','①/② sequential load (h)','①/② cumulative (h)','③ close to due (h)','③ cumulative (h)','Due that day (orders)','Due that day (h)']];
    for (let i = 0; i < S.N; i++){
      const dl = dlByIdx.get(i);
      aoa.push([ isoLocal(S.days[i]), WEEKDAYS[S.days[i].getDay()],
        n2(S.cap[i]), n2(S.eager[i]), n2(S.eagerCum[i]), n2(S.jit[i]), n2(S.jitCum[i]),
        dl ? dl.count : '', dl ? n2(dl.hours) : '' ]);
    }
    extra.push({ name:'Daily timeline (full)', aoa });
    // ①② full per-order tables (not cut off at the page's 200 rows)
    const seqAoa = rows => { const a = [['Seq','Order no.','Product','Designer','Checker','Hours (h)','Expected start','Expected finish','Due','Late (workdays, negative = early)']];
      rows.forEach((r, i) => a.push([ i+1, r.so || r.id, r.product || '', r.designer || '', r.checker || '', n2(r.hours),
        r.startIdx == null ? "Doesn't fit" : isoLocal(S.days[r.startIdx]),
        r.finishIdx == null ? "Doesn't fit" : isoLocal(S.days[r.finishIdx]),
        r.deadline ? isoLocal(r.deadline) : '', r.lateWD == null ? '' : r.lateWD ]));
      return a; };
    extra.push({ name:'① Sequential - per order', aoa: seqAoa(S.seqCont) });
    extra.push({ name:'② By due date - per order', aoa: seqAoa(S.seqSlot) });
    // ④ per-person daily load (full, not limited to the page's 120 rows; exports the mode currently selected on the page)
    const P = S.pw && S.pw[pwMode];
    if (P && P.day && P.day.size){
      const PW_LABEL = { slot:'② by due date', cont:'① sequential', jit:'③ close to due' };
      const a2 = [['Worker','Date','Day','Hours that day (h)','Capacity that day (h)','Orders handled']];
      [...P.day.keys()].sort().forEach(name => {
        const arr = P.day.get(name); const lab = P.orders.get(name) || {};
        for (let i = 0; i < S.N; i++){ if (arr[i] > 1e-9)
          a2.push([name, isoLocal(S.days[i]), WEEKDAYS[S.days[i].getDay()], n2(arr[i]), n2(S.dayCapW(name,i)), (lab[i]||[]).join(', ')]); }
      });
      extra.push({ name:`Per-person daily load (${PW_LABEL[pwMode]||''})`, aoa: a2 });
    }
    if (S.rebalance && S.rebalance.suggestions.length){
      const a3 = [['From','Current load (h)','To','Current load (h)','Suggested move (h)','Orders','Which orders']];
      S.rebalance.suggestions.forEach(s => a3.push([ s.from, n2(s.fromLoad), s.to, n2(s.toLoad), n2(s.amount), s.examples.length,
        s.examples.slice(0,300).map(e=>`${e.id}(${e.role==='d'?'design':'check'} ${Math.round(e.h*10)/10}h${e.sure?'':'·ability not filled'})`).join(', ')
        + (s.examples.length>300?` … ${s.examples.length} orders in total`:'') ]));
      extra.push({ name:'Reassignment suggestions', aoa: a3 });
    }
  }
  exportReportFromDOM('sched-out', `Schedule-report-${todayISO()}.xlsx`, 'Schedule report (by due date)', extra);
});
byId('retro-report') && (byId('retro-report').onclick = () =>
  exportReportFromDOM('retro-out', `Retrospective-report-${todayISO()}.xlsx`, 'Retrospective report'));
byId('export-all').onclick = async () => {
  const steps = ['parts','workers','orders','products'];
  for (const s of steps){ exportOne(s); await new Promise(r => setTimeout(r, 350)); }
  toast('Exported 4 files', 'ok');
};
function downloadText(name, text){
  const blob = new Blob([text], { type:'application/json' });
  const a = el('a', { href: URL.createObjectURL(blob), download: name }); document.body.append(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/* whole-app backup / restore */
byId('backup-save').onclick = () => downloadText(`parts-analysis-backup-${new Date().toISOString().slice(0,10)}.json`,
  JSON.stringify({ __partsAnalysisBackup:1, ...state }, null, 2));
function restoreBackup(data){
  state = { products:(data.products||[]).map(String),
    parts:(data.parts||[]).map(p=>({name:p.name||'',category:p.category||''})),
    workers:(data.workers||[]).map(w=>({name:w.name||'',grades:w.grades||{}})),
    orders:(data.orders||[]).map(o=>({product:o.product||'',parts:o.parts||'',time:o.time||'',actualFinish:o.actualFinish||'',designer:o.designer||'',checker:o.checker||'',required:o.required||'',so:o.so||''})),
    leaves:(data.leaves||[]).map(L=>({worker:L.worker||'',start:L.start||'',end:L.end||'',hours:(L.hours!=null&&L.hours!=='')?L.hours:(L.half?DEFAULT_HOURS_PER_DAY/2:'')})),
    hourlog:(data.hourlog && Array.isArray(data.hourlog.workers)) ? data.hourlog : null,
    settings:{ categoryHours:{ ...DEFAULT_CATEGORY_HOURS }, hoursPerDay:DEFAULT_HOURS_PER_DAY, startDate:'' } };
  applyCategoryHours(data.settings && data.settings.categoryHours);
  applyHoursPerDay(data.settings && data.settings.hoursPerDay);
  applyVerifyRatio(data.settings && data.settings.verifyRatio);
  if (data.settings && typeof data.settings.startDate === 'string') state.settings.startDate = data.settings.startDate;
  state.settings.schedBuffer = (data.settings && data.settings.schedBuffer >= 0) ? data.settings.schedBuffer : 2;
  state.settings.annualLeave = (data.settings && data.settings.annualLeave >= 0) ? data.settings.annualLeave : 15;
  save(); renderAll();
}
byId('backup-input').addEventListener('change', async e => {
  const f = e.target.files[0]; if (!f) return;
  try { restoreBackup(JSON.parse(await f.text())); toast('Backup restored', 'ok'); }
  catch(err){ toast('Restore failed: '+err.message, 'err'); }
  e.target.value='';
});
byId('wipe-all').onclick = () => { if (confirm('Clear ALL data (products/parts/workers/orders)? This cannot be undone.')){ state={products:[],parts:[],workers:[],orders:[],leaves:[],hourlog:null,settings:{categoryHours:{...DEFAULT_CATEGORY_HOURS},hoursPerDay:DEFAULT_HOURS_PER_DAY}}; applyCategoryHours(DEFAULT_CATEGORY_HOURS); applyHoursPerDay(DEFAULT_HOURS_PER_DAY); applyVerifyRatio(DEFAULT_VERIFY_RATIO); save(); renderAll(); toast('All data cleared', 'ok'); } };

/* =========================================================
   ANALYSIS  (full port of analyze.py — same numbers)
   ========================================================= */
byId('analysis-run').onclick = renderAnalysis;
function renderAnalysis(){
  const out = byId('analysis-out'); out.innerHTML='';
  let R;
  try { R = runFullAnalysis(state); }
  catch(e){ console.error(e); out.append(el('div',{class:'alert warn'}, 'Analysis error: '+e.message)); return; }

  // local builders
  const td = (v,num)=>el('td', num?{class:'num'}:{}, v);
  const section = (t,sub)=>out.append(el('div',{class:'section-title'}, sub?el('span',{},[t,el('span',{style:'color:var(--muted);font-weight:500'},'  '+sub)]):t));
  const table = (headers, rows)=>{ const t=el('table',{class:'report'});
    const headTr = el('tr',{}, headers.map(h=>el('th',{},h)));
    t.append(el('thead',{}, headTr));
    const b=el('tbody',{}); rows.forEach(r=>b.append(r)); t.append(b);
    alignNumHeaders(headTr, rows[0]); out.append(el('div',{class:'report-wrap'}, t)); };
  const days = h => h==null ? '—' : round(h/HOURS_PER_DAY, 2);

  // fatal-input errors (these make analyze.py refuse to run)
  if (R.errors && R.errors.length)
    out.append(el('div',{class:'alert warn'}, [el('b',{},'⚠ Needs fixing: '),
      el('ul',{}, R.errors.slice(0,20).map(e=>el('li',{},e)))]));

  if (!R.hasData){
    const need=[]; if(!state.orders.length)need.push('orders'); if(!state.parts.length)need.push('parts');
    if(!state.workers.length)need.push('workers'); if(!state.products.length)need.push('products');
    out.append(el('div',{class:'alert warn'}, `Analysis needs orders + parts + workers (+ products). Missing: ${need.join(', ')||'valid data'}.`));
    if (R.warnings && R.warnings.length)
      out.append(el('div',{class:'alert warn'}, [el('b',{},'Data notes: '), el('ul',{}, R.warnings.map(w=>el('li',{},w)))]));
    return;
  }

  const A = R.asAssigned;
  const pct = x => x==null ? '—' : round(100*x,0)+'%';
  const ROW = 200;
  const HL = state.hourlog;
  const nFin = R.orders.filter(o => o.actualFinish).length;

  /* ===================== ① Theoretical hours summary (all orders) ===================== */
  out.append(el('div',{class:'section-title'},'① Theoretical hours summary — all orders, from the VC table'));
  const cards = el('div',{class:'cards'});
  const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  card('Orders', R.nOrders, `(completed ${nFin} / open ${R.nOrders - nFin})`);
  card('Theoretical design hours', round(R.totalDesign), 'hours');
  card('Theoretical check hours', round(R.totalVerify), 'hours');
  card('Total', round(R.totalDemand), 'hours');
  const teamDaily = R.nWorkers * HOURS_PER_DAY;   // total order hours the whole team can put in per day
  card('Team daily capacity', round(teamDaily,1), `h/day (${R.nWorkers} people × ${fmtHrs(HOURS_PER_DAY)}h)`);
  if (teamDaily > 0) card('Whole team needs', round(R.totalDemand/teamDaily,1), 'workdays to finish');
  out.append(cards);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    `Design = Σ(VC hours) × coefficient (≤5 VCs ×1, 6–10 ×1.2, >10 ×1.5); check = design × ${VERIFY_RATIO} (check ratio is on the Parts tab and can be calibrated from the hour log). ${R.productNames.length} products · ${R.nParts} VCs · ${R.nWorkers} workers.`));

  /* ===================== ② Theory vs actual logged hours (everyone) ===================== */
  if (HL && Array.isArray(HL.workers) && HL.workers.length){
    const actD = HL.workers.reduce((a,w)=>a+(w.byTask['Order design']||0),0);
    const actV = HL.workers.reduce((a,w)=>a+(w.byTask['Order check']||0),0);
    section('② Theory vs actual logged hours · everyone', `Actual = order design/check hours from the hour log (${HL.year||'?'}, months ${HL.months.join(', ')})`);
    table(['Stage','Theory (h)','Actual logged (h)','Actual − theory (h)','Actual / theory'],
      [['Order design', R.totalDesign, actD], ['Order check', R.totalVerify, actV], ['Total', R.totalDemand, actD+actV]]
      .map(([k,t,a],i)=>el('tr',{},[
        i===2 ? el('td',{},el('b',{},k)) : td(k),
        td(round(t),true), td(round(a),true),
        td((a-t>=0?'+':'')+round(a-t),true), td(t>0?pct(a/t):'—',true) ])));
    out.append(el('p',{class:'hint',style:'margin:-2px 0 2px'},
      'For this comparison to make sense, the orders and the hour log must cover the same period (e.g. both 2025). Actual > theory = more time spent than the VC standard (or orders missing); actual < theory = faster than standard (or hours not fully logged). If the check ratio was just calibrated from the hour log, the "Order check" ratio will naturally sit close to the design row — that is expected.'));
  } else {
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
      'ℹ No hour log imported yet: import the hour report Excel on the Hour Log tab and this section will show the "② Theory vs actual" totals and the actual columns in ③.'));
  }

  /* ===================== ③ Per worker · theory vs actual ===================== */
  {
    const act = new Map(((HL && HL.workers) || []).map(w => [w.name, w]));
    const theo = new Map(A.workerLoads.map(w => [w.name, w]));
    const rows = [...new Set([...theo.keys(), ...act.keys()])].map(n => {
      const t = theo.get(n), a = act.get(n);
      const thD = t ? t.design : 0, thV = t ? t.verify : 0;
      const acD = a ? (a.byTask['Order design']||0) : 0, acV = a ? (a.byTask['Order check']||0) : 0;
      return { n, thD, thV, thT: thD+thV, acD, acV, acT: acD+acV };
    }).filter(r => r.thT > 1e-9 || r.acT > 1e-9).sort((x,y)=>(y.thT+y.acT)-(x.thT+x.acT));
    section('③ Per worker · theory vs actual', HL ? 'Theory = summed from order assignments (VC hours of orders this person designs/checks); actual = hours logged' : 'Theory = summed from order assignments; import an hour log to add actual columns');
    if (HL){
      table(['Worker','Theory design','Theory check','Theory total','Actual design','Actual check','Actual total','Actual/theory'],
        rows.slice(0,ROW_CAP).map(r => el('tr',{},[
          td(r.n), td(round(r.thD)+'h',true), td(round(r.thV)+'h',true), el('td',{class:'num'},el('b',{},round(r.thT)+'h')),
          td(round(r.acD)+'h',true), td(round(r.acV)+'h',true), el('td',{class:'num'},el('b',{},round(r.acT)+'h')),
          td((r.thT>0&&r.acT>0)?pct(r.acT/r.thT):'—',true) ])));
    } else {
      table(['Worker','Theory design','Theory check','Total','Workdays equiv.','Orders (design/check)'],
        A.workerLoads.slice(0,ROW_CAP).map(w => el('tr',{},[
          td(w.name), td(round(w.design)+'h',true), td(round(w.verify)+'h',true),
          el('td',{class:'num'}, el('b',{}, round(w.total)+'h')), td(round(w.days,2),true),
          td(`${w.asDesigner}/${w.asChecker}`,true) ])));
    }
    if (rows.length > ROW_CAP) out.append(el('p',{class:'hint'}, `Showing the first ${ROW_CAP} (of ${rows.length}).`));
  }

  /* ===================== ④ Order hours detail ===================== */
  section('④ Order hours detail',`Design = Σ(VC hours) × coefficient; check = design × ${VERIFY_RATIO}`);
  table(['Order','Product','Difficulty','VCs','Coeff.','Design h','Check h','Designer','Checker','Due date'],
    A.rows.slice(0,ROW).map(r=>el('tr',{},[
      td(r.so || r.id), td(r.product), td(r.difficulty), td(r.nVcs,true), td('×'+r.coeff,true),
      td(round(r.designHours,2),true), td(round(r.verifyHours,2),true),
      el('td',{}, r.designerOk? r.designer||'—' : el('span',{class:'badge bad'}, (r.designer||'empty')+' ✖')),
      el('td',{}, r.checkerOk? r.checker||'—' : el('span',{class:'badge bad'}, (r.checker||'empty')+' ✖')),
      td(r.deadline||'—') ])));
  if (A.rows.length > ROW) out.append(el('p',{class:'hint'}, `Showing the first ${ROW} (of ${A.rows.length}); the exported report includes every table on the page.`));

  if (R.warnings && R.warnings.length)
    out.append(el('div',{class:'alert warn'}, [el('b',{},'Data notes: '), el('ul',{}, R.warnings.slice(0,10).map(w=>el('li',{},w)))]));
}
function round(x, d=1){ const f = Math.pow(10,d); return String(Math.round(x*f)/f); }
/* right-align headers of numeric columns too (based on whether the first row's cell has .num), so they line up with the numbers */
function alignNumHeaders(headTr, firstRow){
  if (!firstRow || !firstRow.children || !headTr) return;
  [...firstRow.children].forEach((c, i) => {
    if (c.classList && c.classList.contains('num') && headTr.children[i]) headTr.children[i].classList.add('num');
  });
}

/* =========================================================
   RETROSPECTIVE
   ========================================================= */
function openRetro(){
  byId('retro-hpd').value = fmtHrs(HOURS_PER_DAY);
  renderRetroInputs();
}
byId('retro-hpd') && byId('retro-hpd').addEventListener('change', () => {
  const v = parseFloat(byId('retro-hpd').value);
  if (isNaN(v) || v <= 0){ byId('retro-hpd').value = fmtHrs(HOURS_PER_DAY); return; }
  applyHoursPerDay(v); save(); byId('retro-out').innerHTML = '';
});
byId('retro-run') && (byId('retro-run').onclick = renderRetro);

/* strict positional Excel for the retrospective: col1=product col2=parts col3=submit col4=actual finish */
const RETRO_HDR = ['Product type', 'Parts needed (name×N; …)', 'Submitted (YYYY-MM-DD)', 'Actual finish date (YYYY-MM-DD)'];
function cellToISO(v){
  if (v == null || v === '') return '';
  if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
  if (typeof v === 'number' && XLSX.SSF){ const d = XLSX.SSF.parse_date_code(v); if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`; }
  return normDate(String(v));
}
byId('retro-import') && (byId('retro-import').onclick = () => byId('retro-file').click());
byId('retro-file') && byId('retro-file').addEventListener('change', async e => {
  const file = e.target.files[0]; if (!file) return; e.target.value = '';
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:true });
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, blankrows:false, defval:'' });
    if (aoa.length < 2){ toast('The file has no data rows (row 1 is the header)', 'err'); return; }
    const orders = [];
    aoa.slice(1).forEach(r => {                          // skip header row, read strictly by column index
      const product = clean(r[0]), parts = clean(r[1]);
      if (!product && !parts) return;
      orders.push({ product, parts, time: cellToISO(r[2]), actualFinish: cellToISO(r[3]) });
    });
    if (!orders.length){ toast('No order rows to import', 'err'); return; }
    if (state.orders.length && !confirm(`Replace the current ${state.orders.length} orders with the ${orders.length} orders in the file (including actual finish dates). Continue?`)) return;
    state.orders = orders; save();
    renderOrders(); renderRetroInputs(); byId('retro-out').innerHTML = '';
    toast(`Imported ${orders.length} retrospective orders`, 'ok');
  } catch (err){ console.error(err); toast('Import failed: ' + err.message, 'err'); }
});
byId('retro-export') && (byId('retro-export').onclick = () => {
  const rows = state.orders.map(o => [o.product || '', o.parts || '', o.time || '', o.actualFinish || '']);
  if (!rows.length) rows.push(['Sample product', 'Custom housing×2; Standard bracket×4; Standard screw pack×8', '2025-01-06', '2025-01-20']);
  const ws = XLSX.utils.aoa_to_sheet([RETRO_HDR, ...rows]); setCols(ws, [16, 42, 22, 24]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Retrospective orders');
  XLSX.writeFile(wb, 'retrospective.xlsx');
});

function renderRetroInputs(){
  const box = byId('retro-inputs'); box.innerHTML = '';
  if (!state.orders.length){ box.append(el('div',{class:'empty-note'}, 'No orders yet. Add them on the Orders tab, or import historical orders on the Import/Export tab.')); return; }
  const t = el('table', { class:'grid retro-grid' });
  t.append(el('thead', {}, el('tr', {}, [
    el('th',{class:'rownum'},'#'), el('th',{style:'width:160px'},'Product'), el('th',{style:'width:300px'},'VC list'),
    el('th',{style:'width:130px'},'Created date'), el('th',{style:'width:200px'},'Actual finish date'), el('th',{class:'rt-spacer'},'') ])));
  const tb = el('tbody', {});
  state.orders.slice(0, ROW_CAP).forEach((o, i) => {
    const fin = el('input', { type:'date', value:o.actualFinish||'' });
    fin.addEventListener('input', () => { o.actualFinish = fin.value; save(); });
    tb.append(el('tr', {}, [
      el('td',{class:'rownum'}, String(i+1)),
      el('td',{}, o.product || el('span',{style:'color:var(--muted)'},'(empty)')),
      el('td',{class:'cell',style:'color:var(--muted);max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:9px 12px'}, o.parts || '—'),
      el('td',{class:'cell',style:'padding:9px 12px;color:var(--muted)'}, o.time || '—'),
      el('td',{}, fin),
      el('td',{class:'rt-spacer'},''),
    ]));
  });
  if (state.orders.length > ROW_CAP)
    tb.append(el('tr', {}, el('td', { colspan:'6', style:'padding:10px;color:var(--muted);text-align:center' }, `Showing the first ${ROW_CAP} (of ${state.orders.length}); the retrospective includes every order with an actual finish date.`)));
  t.append(tb); box.append(t);
}

function renderRetro(){
  const out = byId('retro-out'); out.innerHTML = '';
  let Ret; try { Ret = runRetrospective(state, {}); }
  catch(e){ console.error(e); out.append(el('div',{class:'alert warn'}, 'Retrospective error: '+e.message)); return; }

  const td = (v,num)=>el('td', num?{class:'num'}:{}, v);
  const section = (t,sub)=>out.append(el('div',{class:'section-title'}, sub?el('span',{},[t,el('span',{style:'color:var(--muted);font-weight:500'},'  '+sub)]):t));
  const table = (headers, rows)=>{ const x=el('table',{class:'report'});
    const headTr = el('tr',{}, headers.map(h=>el('th',{},h)));
    x.append(el('thead',{}, headTr)); const b=el('tbody',{}); rows.forEach(r=>b.append(r)); x.append(b);
    alignNumHeaders(headTr, rows[0]); out.append(el('div',{class:'report-wrap'}, x)); };
  const d2 = x => x==null ? '—' : round(x,2);
  const pct = x => x==null ? '—' : round(100*x,0)+'%';
  const fmtDate = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;

  if (!Ret.hasData){
    if (Ret.nAll > 0 && Ret.nFinished === 0){
      out.append(el('div',{class:'alert warn'}, `None of these ${Ret.nAll} orders has an "Actual finish date" yet — the retrospective only analyzes completed orders. Fill in finish dates in ① above, then click "Run retrospective"; for unfinished work see the Schedule tab.`));
      return;
    }
    const need=[]; if(!state.orders.length)need.push('orders'); if(!state.parts.length)need.push('parts');
    if(!state.workers.length)need.push('workers');
    out.append(el('div',{class:'alert warn'}, `The retrospective needs orders + parts + workers. Missing: ${need.join(', ')||'valid data'}.`));
    return;
  }
  if (!Ret.nWithFinish){
    out.append(el('div',{class:'alert warn'}, 'Fill in an "Actual finish date" for at least one order in ① above, then click "Run retrospective".'));
    return;
  }

  // ---- retrospective range: completed orders, earliest created date → today ----
  out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
    `Retrospective range: ${Ret.start ? fmtDate(Ret.start) : 'earliest created date'} → today (${todayISO()}), ${Ret.nWithFinish} completed orders included` +
    (Ret.nUnfinished > 0 ? `; ${Ret.nUnfinished} unfinished orders excluded (they are on the Schedule tab).` : '.')));

  // ---- ② real delivery performance (per order, on its own due/created dates, calendar days — most intuitive) ----
  out.append(el('div',{class:'section-title'}, el('span',{},['② Retrospective · actual delivery performance',
    el('span',{style:'color:var(--muted);font-weight:500'}, '  each order measured against its own due/created dates (calendar days)')])));
  const cardsD = el('div',{class:'cards'});
  const cardD=(k,v,u)=>cardsD.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  if (Ret.nWithDeadline){
    cardD('On-time rate', pct(Ret.onTimeRate));
    cardD('Late orders', Ret.nLate+' / '+Ret.nWithDeadline, 'orders');
    if (Ret.nLate){ cardD('Median delay', Ret.medianLate==null?'—':Ret.medianLate, 'calendar days'); cardD('Longest delay', Ret.maxLate, 'calendar days'); }
  } else {
    cardD('On-time rate', '—', 'no due dates');
  }
  cardD('Median turnaround', Ret.medianTurnaround==null?'—':Ret.medianTurnaround, 'calendar days');
  out.append(cardsD);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    'On-time rate / lateness use each order\'s own "Due date"; turnaround = actual finish − that order\'s created date. These only look at each order itself, unaffected by other orders or where the data starts, so they best reflect real delivery.'));

  // ---- completion speed (per order): plain-language percentiles + bucket bars. Elapsed time isn't summed (parallel work would double-count days) ----
  out.append(el('div',{class:'section-title',style:'margin-top:8px'}, el('span',{},['Completion speed · how long from created to finished',
    el('span',{style:'color:var(--muted);font-weight:500'}, '  workdays per order (Mon–Fri only)')])));
  if (Ret.nTheo){
    const cards = el('div',{class:'cards'});
    const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
    card('Half of orders', '≤ '+round(Ret.turnP50,1), 'workdays');
    card('90% of orders', '≤ '+round(Ret.turnP90,1), 'workdays');
    card('Slowest 5%', '≥ '+round(Ret.turnP95,1), `workdays (longest ${Ret.turnMax})`);
    card('Avg work per order (theory)', round(Ret.avgTheoHours,1), 'h/order');
    card('Whole team needs (theory)', Ret.nWorkers > 0 ? round(Ret.theoSumWD / Ret.nWorkers, 1) : '—',
      `workdays (total hours ÷ ${Ret.nWorkers} people × ${fmtHrs(Ret.H)}h/day)`);
    out.append(cards);
    out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
      `How to read: "Half of orders ≤ ${round(Ret.turnP50,1)} workdays" = 50% of orders took no more than ${round(Ret.turnP50,1)} workdays from created to actual finish (${Ret.nTheo} orders have both dates). ` +
      `Compare with theory: the design+check work per order averages only about ${round(Ret.avgTheoHours,1)} hours — slow completion is mostly queueing / working on other orders at the same time, not big jobs.`));
    // distribution table: where the bulk is and how long the tail is, at a glance
    section('Turnaround distribution', 'Which range each order\'s created-to-finished workdays falls into');
    const bar = share => el('div',{class:'sbar-cell'},[
      el('div',{class:'sbar'}, el('div',{class:'sbar-fill',style:`width:${Math.min(100, share*100)}%`})),
      el('span',{class:'sbar-num'}, pct(share)) ]);
    table(['Turnaround (workdays)','Orders','Share','Cumulative'],
      Ret.turnDist.filter(b => b.n > 0).map(b => el('tr',{},[
        td(b.label), td(b.n, true), el('td',{}, bar(b.share)), td('≤ '+pct(b.cumShare), true) ])));
  } else {
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
      'ℹ Completion speed unavailable: these completed orders have no "Created date". Add created dates to show it.'));
  }

  // ---- verdict ----  conclude from real delivery (on-time rate / lateness), not by subtracting "whole-batch span vs ideal hours"
  if (Ret.nWithDeadline){
    if (!Ret.nLate)
      out.append(el('div',{class:'alert ok'}, [el('b',{},'✔ Everything delivered on time. '),
        `All ${Ret.nWithDeadline} orders with a due date were finished on or before their due date.`]));
    else
      out.append(el('div',{class:'alert warn'}, [el('b',{},`✗ ${Ret.nLate} orders were late, on-time rate ${pct(Ret.onTimeRate)}.`),
        ` Longest delay ${Ret.maxLate} days, median delay ${Ret.medianLate} days (details in "Late orders" below).`,
        Ret.bestTraining ? el('div',{style:'margin-top:4px'},[el('b',{},'Without hiring: '),`"${Ret.bestTraining.label}" could cut the theoretical minimum by about ${d2(Ret.bestTraining.saved/Ret.H)} workdays.`]) : null ]));
  } else {
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'}, [el('b',{},'ℹ Can\'t judge on-time delivery. '),
      ` These orders have no due dates. ${Ret.nWithFinish} completed, median turnaround ${Ret.medianTurnaround==null?'—':Ret.medianTurnaround} calendar days. Add due dates to see the on-time rate.`]));
  }

  // ---- ③ improvement without new labor ----
  out.append(el('div',{class:'section-title'}, '③ Improvements without hiring'));
  const posLevers = Ret.trainingLevers.filter(l => l.saved/Ret.H >= 0.005);   // only list ones with a positive gain
  if (posLevers.length){
    section('② Train existing workers','Raise an existing worker\'s grade for a product by one level (no new hires) and its effect on the theoretical minimum (top 10)');
    table(['Action','Minimum after','Saves'],
      posLevers.slice(0,10).map((l,i)=> el('tr',{},[
        el('td',{}, i===0?el('b',{},l.label):l.label),
        td(d2(l.hours/Ret.H)+' days', true), td('↓ '+d2(l.saved/Ret.H)+' days', true) ])));
  }

  // ---- per-worker load ----
  section('Load per worker', Ret.loadsAreActual
    ? `Actual logged hours (order design + check, from the hour log) ÷ available hours per year ((260 − ${Ret.annualLeave} days annual leave) × ${fmtHrs(Ret.H)}h ≈ ${round(Ret.capYear)}h/person·year); over 100% is red; 🔵 blue = on the roster but not in the hour log (no hours logged, or names spelled differently)`
    : `No hour log imported — using as-assigned standard hours ÷ available hours per year (≈ ${round(Ret.capYear)}h/person·year) for now; over 100% is red`);
  table(['Worker', Ret.loadsAreActual ? 'Actual hours (logged)' : 'Assigned hours (standard)','Idle hours','Utilization','Suggestion'],
    Ret.workerLoads.map(w => {
      const noHL = Ret.loadsAreActual && !w.inHL;
      return el('tr',{class: w.util>1.000001 ? 'infeasible' : (noHL ? 'unmatched-roster' : '')},[
        td(w.name), td(round(w.hours)+'h',true), td(round(w.idle)+'h',true), td(pct(w.util),true),
        el('td',{}, noHL ? '🔵 Name not in the hour log (no hours logged, or names spelled differently)'
          : (w.trainTarget ? `Fairly idle, could be trained for ${w.trainTarget.product}/${w.trainTarget.cat}/${ROLE_LABEL[w.trainTarget.role]} (needs grade ${w.trainTarget.need})` : (w.util>=0.95?'Fully loaded':'—'))) ]);
    }));

  // ---- ④ standard hours vs actual logged (needs the hour log, linked by name) ----
  if (state.hourlog && Array.isArray(state.hourlog.workers) && state.hourlog.workers.length){
    const HL = state.hourlog;
    const std = new Map(((Ret.R.asAssigned && Ret.R.asAssigned.workerLoads) || []).map(w => [w.name, w]));
    const act = new Map(HL.workers.map(w => [w.name, w]));
    const rows = [];
    const rosterSet = new Set(state.workers.map(w => w.name).filter(Boolean));
    [...new Set([...std.keys(), ...act.keys()])].forEach(n => {
      const s = std.get(n), a = act.get(n);
      const sd = s ? s.design : 0, sv = s ? s.verify : 0;                       // standard: VC hours summed by assignment
      const ad = a ? (a.byTask['Order design'] || 0) : 0, av = a ? (a.byTask['Order check'] || 0) : 0;  // actual: hours logged
      if (sd + sv + ad + av <= 1e-9) return;
      rows.push({ n, sd, sv, ad, av,
        noHL: rosterSet.has(n) && !act.has(n),   // blue: on roster, not in log
        noRoster: !rosterSet.has(n),             // orange: in log, not on roster
        effD: (sd > 0 && ad > 0) ? sd / ad : null, effV: (sv > 0 && av > 0) ? sv / av : null });
    });
    if (rows.length){
      rows.sort((x, y) => (y.sd + y.ad + y.sv + y.av) - (x.sd + x.ad + x.sv + x.av));
      const effTxt = e => e == null ? '—' : round(100*e, 0) + '%' + (e >= 1.05 ? ' (faster than standard)' : e <= 0.95 ? ' (slower than standard)' : ' (≈ standard)');
      const sumSd = rows.reduce((t,r)=>t+r.sd,0), sumAd = rows.reduce((t,r)=>t+r.ad,0);
      const sumSv = rows.reduce((t,r)=>t+r.sv,0), sumAv = rows.reduce((t,r)=>t+r.av,0);
      section('④ Standard hours vs actual logged', `Orders' VC standard hours (summed by assignment) vs actual hours from the hour log (${HL.year || '?'}, months ${HL.months.join(', ')}); 🔵 blue = on roster but not in log, 🟠 orange = in log but not on roster`);
      if (sumAd > 0 || sumAv > 0)
        out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
          `Whole group: design standard ${round(sumSd)}h / actual ${round(sumAd)}h${sumAd>0&&sumSd>0?`, efficiency ${round(100*sumSd/sumAd,0)}%`:''}; check standard ${round(sumSv)}h / actual ${round(sumAv)}h${sumAv>0&&sumSv>0?`, efficiency ${round(100*sumSv/sumAv,0)}%`:''}. Efficiency = standard ÷ actual; >100% means faster than the VC standard.`));
      table(['Worker','Standard design','Actual · order design','Design efficiency','Standard check','Actual · order check','Check efficiency'],
        rows.slice(0, ROW_CAP).map(r => el('tr',{class:
          r.noHL ? 'unmatched-roster' : r.noRoster ? 'unmatched-hourlog'
          : ((r.effD!=null&&r.effD<=0.7)||(r.effV!=null&&r.effV<=0.7)?'infeasible':'')},[
          td(r.n), td(round(r.sd)+'h',true), td(round(r.ad)+'h',true), td(effTxt(r.effD),true),
          td(round(r.sv)+'h',true), td(round(r.av)+'h',true), td(effTxt(r.effV),true) ])));
      if (rows.length > ROW_CAP) out.append(el('p',{class:'hint'}, `Showing the first ${ROW_CAP} (of ${rows.length}).`));
      out.append(el('p',{class:'hint'},
        'Only comparable if both sides cover the same period (this batch of orders ≈ the hour-log year) and names match exactly; "Overtime" hours are not counted in actual design/check; rows with clearly low efficiency (≤70%) are red — this may be a different logging basis, missing orders, or genuinely slower than standard, so check the data before drawing conclusions.'));
    }
  }

  // ---- late orders (each order on its own calendar, no global anchor / ideal-hours index) ----
  const ROWL = 200;
  const late = Ret.lateVsDeadline || [];
  section('Late orders', 'Late = actual finish − that order\'s due date; turnaround = actual finish − that order\'s created date (both in calendar days, each order on its own, unaffected by other orders)');
  if (!Ret.nWithDeadline){
    out.append(el('div',{class:'alert warn'}, 'These orders have no "Due date", so lateness can\'t be judged. Add due dates on the Orders tab and run the retrospective again.'));
  } else if (!late.length){
    out.append(el('div',{class:'alert ok'}, '✔ No late orders: every order with an actual finish date was finished on or before its "Due date".'));
  } else {
    table(['Order','Product','Created date','Due date','Actual finish','Turnaround (calendar days)','Late (calendar days)'],
      late.slice(0,ROWL).map(o => el('tr',{class:'infeasible'},[
        td(o.so || o.id), td(o.product),
        td(o.submit?fmtDate(o.submit):'—'), td(o.deadline?fmtDate(o.deadline):'—'), td(fmtDate(o.finish)),
        td(o.turnaroundCal!=null?o.turnaroundCal:'—', true),
        td('+'+o.lateCal+' days', true) ])));
    if (late.length > ROWL) out.append(el('p',{class:'hint'}, `Showing the ${ROWL} latest orders (of ${late.length} late).`));
  }

  // ---- reconstructed optimal timeline (daily load) ----
  if (Ret.R.sim.daily.length){
    section('Reconstructed optimal timeline', 'Hours per difficulty and utilization by day (ideal schedule)');
    table(['Day','High (h)','Low (h)','None (h)','Day total','Utilization'],
      Ret.R.sim.daily.map((dd,i)=> el('tr',{},[ td('Day '+(i+1)),
        td(round(dd.High),true), td(round(dd.Low),true), td(round(dd.None),true), td(round(dd.total)+'h',true),
        td(Ret.R.sim.totalCapacity? round(100*dd.total/Ret.R.sim.totalCapacity)+'%':'—', true) ])));
  }
  foldify(out);
}

/* =========================================================
   SCHEDULE  (scheduling by due date)
   ========================================================= */
const WEEKDAYS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
function isoLocal(d){ return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function todayISO(){ return isoLocal(new Date()); }
// schedule start day: defaults to today (the schedule only covers unfinished work, forward from today).
// a future date set by the user is kept; past dates are pulled back to today automatically.
function defaultSchedStart(){
  const today = todayISO();
  const saved = state.settings.startDate;
  return (saved && saved >= today) ? saved : today;
}
function pctSafe(a,b){ return b>0 ? round(100*a/b,0)+'%' : '—'; }
/* Folding: wrap each "section title + the content after it" in a collapsible <details>; tables over 30 rows start collapsed.
   Elements with the nofold class are not folded (e.g. the "Notes" at the bottom). Exported reports aren't affected: folding is just display state, everything stays in the DOM. */
function foldify(out){
  const kids = [...out.children];
  const groups = []; let g = null;
  kids.forEach(k => {
    if (k.classList && k.classList.contains('nofold')){ g = null; return; }
    if (k.classList && k.classList.contains('section-title')){ g = { title:k, items:[] }; groups.push(g); }
    else if (g) g.items.push(k);
  });
  const real = groups.filter(x => x.items.length);
  if (real.length < 2) return;
  let nClosed = 0;
  real.forEach(x => {
    const det = el('details',{class:'fold'});
    const rows = x.items.reduce((n,it)=> n + it.querySelectorAll('tbody tr').length, 0);
    det.open = rows <= 30; if (!det.open) nClosed++;
    out.insertBefore(det, x.title);
    const sum = el('summary',{});
    sum.append(x.title);
    if (!det.open) sum.append(el('span',{class:'fold-hint'}, ` (${rows} rows, click to expand)`));
    det.append(sum);
    x.items.forEach(it => det.append(it));
  });
  const bar = el('div',{class:'toolbar fold-bar'},[
    el('span',{class:'hint'}, nClosed ? `Collapsed ${nClosed} long sections automatically; click a section title to expand/collapse:` : 'Click a section title to collapse/expand:'),
    (()=>{ const b=el('button',{class:'btn'},'Expand all'); b.onclick=()=>out.querySelectorAll('details.fold').forEach(d=>d.open=true); return b; })(),
    (()=>{ const b=el('button',{class:'btn'},'Collapse all'); b.onclick=()=>out.querySelectorAll('details.fold').forEach(d=>d.open=false); return b; })() ]);
  out.insertBefore(bar, out.firstChild);
}
function schedBar(load, C, over, lvl, idle){
  const w = C>0 ? Math.min(100, load/C*100) : 0;
  const bar = el('div',{class:'sbar'+(lvl?' lvl':'')+(over?' over':'')}, el('div',{class:'sbar-fill',style:`width:${w}%`}));
  const label = idle ? 'idle' : (load>1e-9 ? round(load)+'h ('+pctSafe(load,C)+')' : '—');
  return el('div',{class:'sbar-cell'}, [bar, el('span',{class:'sbar-num'}, label)]);
}

function openSchedule(){
  byId('sched-start').value = defaultSchedStart();
  byId('sched-hpd').value = fmtHrs(HOURS_PER_DAY);
  byId('sched-buffer') && (byId('sched-buffer').value = String(state.settings.schedBuffer));
  byId('sched-al') && (byId('sched-al').value = String(state.settings.annualLeave));
  renderSchedInputs();
}
byId('sched-start') && byId('sched-start').addEventListener('change', () => { state.settings.startDate = byId('sched-start').value; save(); byId('sched-out').innerHTML=''; });
byId('sched-hpd') && byId('sched-hpd').addEventListener('change', () => {
  const v = parseFloat(byId('sched-hpd').value);
  if (isNaN(v) || v <= 0){ byId('sched-hpd').value = fmtHrs(HOURS_PER_DAY); return; }
  applyHoursPerDay(v); save(); byId('sched-out').innerHTML='';
});
byId('sched-buffer') && byId('sched-buffer').addEventListener('change', () => {
  let v = parseInt(byId('sched-buffer').value, 10);
  if (isNaN(v) || v < 0 || v > 60){ byId('sched-buffer').value = String(state.settings.schedBuffer); return; }
  state.settings.schedBuffer = v; save(); byId('sched-out').innerHTML='';
});
byId('sched-al') && byId('sched-al').addEventListener('change', () => {
  let v = parseInt(byId('sched-al').value, 10);
  if (isNaN(v) || v < 0 || v > 100){ byId('sched-al').value = String(state.settings.annualLeave); return; }
  state.settings.annualLeave = v; save(); byId('sched-out').innerHTML='';
});
byId('sched-run') && (byId('sched-run').onclick = renderSchedule);

function renderSchedInputs(){
  const box = byId('sched-inputs'); box.innerHTML='';
  if (!state.orders.length){ box.append(el('div',{class:'empty-note'}, 'No orders yet. Add them on the Orders tab first.')); return; }
  const t = el('table', { class:'grid retro-grid' });
  t.append(el('thead', {}, el('tr', {}, [
    el('th',{class:'rownum'},'#'), el('th',{style:'width:160px'},'Product'), el('th',{style:'width:300px'},'VC list'),
    el('th',{style:'width:170px'},'Due date (deadline)'), el('th',{style:'width:140px'},'Created date'), el('th',{class:'rt-spacer'},'') ])));
  const tb = el('tbody', {});
  state.orders.slice(0, ROW_CAP).forEach((o, i) => {
    const dl = el('input', { type:'date', value:o.required||'' });      // due date = the schedule's deadline
    dl.addEventListener('input', () => { o.required = dl.value; save(); });
    tb.append(el('tr', {}, [
      el('td',{class:'rownum'}, String(i+1)),
      el('td',{}, o.product || el('span',{style:'color:var(--muted)'},'(empty)')),
      el('td',{class:'cell',style:'color:var(--muted);max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:9px 12px'}, o.parts || '—'),
      el('td',{}, dl),
      el('td',{class:'cell',style:'padding:9px 12px;color:var(--muted)'}, o.time || '—'),
      el('td',{class:'rt-spacer'},''),
    ]));
  });
  if (state.orders.length > ROW_CAP)
    tb.append(el('tr', {}, el('td', { colspan:'6', style:'padding:10px;color:var(--muted);text-align:center' }, `Showing the first ${ROW_CAP} (of ${state.orders.length}); scheduling includes all of them.`)));
  t.append(tb); box.append(t);
}

let lastSched = null;   // latest schedule result, used by "Export report" to build the full timeline (unaffected by display truncation)
let pwMode = 'slot';    // ④ per-worker schedule current mode: 'cont' ① sequential / 'slot' ② by due date / 'jit' ③ close to due
function renderSchedule(){
  const out = byId('sched-out'); out.innerHTML='';
  let S; try { S = runSchedule(state, { startDate: byId('sched-start').value, bufferDays: parseInt(byId('sched-buffer') && byId('sched-buffer').value, 10),
    annualLeaveDays: parseInt(byId('sched-al') && byId('sched-al').value, 10) }); }
  catch(e){ console.error(e); lastSched = null; out.append(el('div',{class:'alert warn'}, 'Scheduling error: '+e.message)); return; }
  lastSched = (S && S.days && S.days.length) ? S : null;

  const td = (v,num)=>el('td', num?{class:'num'}:{}, v);
  const section = (t,sub)=>out.append(el('div',{class:'section-title'}, sub?el('span',{},[t,el('span',{style:'color:var(--muted);font-weight:500'},'  '+sub)]):t));
  const tableIn = (host, headers, rows)=>{ const x=el('table',{class:'report'}); const headTr = el('tr',{}, headers.map(h=>el('th',{},h))); x.append(el('thead',{}, headTr)); const b=el('tbody',{}); rows.forEach(r=>b.append(r)); x.append(b); alignNumHeaders(headTr, rows[0]); host.append(el('div',{class:'report-wrap'}, x)); };
  const table = (headers, rows)=>tableIn(out, headers, rows);

  if (!S.hasData){ out.append(el('div',{class:'alert warn'}, 'Scheduling needs orders + parts + workers.')); return; }
  if (S.noSchedule){
    out.append(el('div',{class:'alert warn'},
      S.nFinished > 0
        ? `No unfinished orders: all ${S.nFinished} have an "Actual finish date". The schedule only covers unfinished work; for past performance see the Retrospective tab.`
        : 'No schedulable orders: add orders first (their parts must be registered on the Parts tab).'));
    return;
  }

  const C = S.capacity;
  // ---- cards ----
  out.append(el('div',{class:'section-title'}, '② Schedule results'));
  const cards = el('div',{class:'cards'});
  const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  card('Total hours', round(S.total), 'hours');
  card('Team daily capacity (after leave)', round(S.baseCap,1), 'h/day');
  card('All done by', 'Day '+S.eagerFinishDay, `workday (${isoLocal(S.days[Math.min(S.N,S.eagerFinishDay)-1])})`);
  card('② By due date: late', S.slotLate.length + ' / ' + S.scheduledOrders.length, 'orders');
  card('③ Close-to-due buffer', S.buffer, 'workdays');
  out.append(cards);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    `Start ${isoLocal(S.start)} · ${S.nWorkers} people × ${fmtHrs(S.H)}h = full ${round(C)}h/day, minus annual leave ${S.annualLeave} days/person·year (× ${round(S.leaveFactor*100)}%) ≈ ${round(S.baseCap,1)}h/day · workdays = Mon–Fri · last due date ${isoLocal(S.days[S.N-1])} (workday ${S.N})`));

  if (S.overdue && S.overdue.length)
    out.append(el('div',{class:'alert warn'}, [el('b',{},`⚠ ${S.overdue.length} orders are due before the start date: `),
      `${S.overdue.slice(0,8).join(', ')}${S.overdue.length>8?` and ${S.overdue.length - 8} more`:''}. They are marked red as "Over capacity" in the tables below; what it really means is their due date is before the start date so they can't fit — move the start date earlier or check the due dates.`]));

  if (S.farDeadlines && S.farDeadlines.length)
    out.append(el('div',{class:'alert warn'}, [el('b',{},`⚠ The timeline is stretched to ${isoLocal(S.lastDl)}: `),
      `${S.farDeadlines.length} orders have a due date more than 1 year after the latest created date (${S.maxCreated?isoLocal(S.maxCreated):'—'}), possibly a typo or import misalignment: ` +
      S.farDeadlines.slice(0,5).map(f=>`${f.id} (due ${isoLocal(f.deadline)})`).join(', ') +
      (S.farDeadlines.length>5?` and ${S.farDeadlines.length - 5} more`:'') +
      `. The timeline ends at the latest due date — search for the year (e.g. ${S.lastDl.getFullYear()}) on the Orders tab to check and fix it.`]));

  if (S.slotLate.length)
    out.append(el('div',{class:'alert warn'}, `✖ Over capacity: even with "② By due date" (earliest due first, no idle time) ${S.slotLate.length} orders will be late (red in table ②). Options: start earlier / overtime, or train more workers to share the bottleneck difficulty (see Analysis and Retrospective).`));
  else if (S.scheduledOrders.length)
    out.append(el('div',{class:'alert ok'}, `✔ With "② By due date" every order finishes on or before its due date; all work can be done by workday ${S.eagerFinishDay} (${isoLocal(S.days[Math.min(S.N,S.eagerFinishDay)-1])}) at the earliest. Expected start/finish dates per order are in tables ① and ② below.`));

  // ---- ①② per-order timetables ----
  const ROWL = 200;
  const lateTxt = r => r.lateWD == null ? '—' : r.lateWD > 0 ? `${r.lateWD} days late` : r.lateWD === 0 ? 'on due date' : `${-r.lateWD} days early`;
  const seqRows = rows => rows.slice(0, ROWL).map((r, i) => el('tr',{class:(r.lateWD != null && r.lateWD > 0) ? 'infeasible' : ''},[
    td(i+1, true), td(r.so || r.id), td(r.product || '—'),
    td(r.designer || '—'), td(r.checker || '—'),
    td(round(r.hours,1)+'h', true),
    td(r.startIdx == null ? "Doesn't fit" : isoLocal(S.days[r.startIdx])),
    td(r.finishIdx == null ? "Doesn't fit (>1500 days)" : isoLocal(S.days[r.finishIdx])),
    td(r.deadline ? isoLocal(r.deadline) : '—'),
    td(lateTxt(r), true) ]));
  const seqHead = ['Seq','Order','Product','Designer','Checker','Hours','Expected start','Expected finish','Due','Early/late (workdays)'];

  section('① Sequential · per-order timetable', `Ignores due dates: queue by created date, finish one then start the next, no idle time${S.contLate.length?` (${S.contLate.length} late, in red)`:''}`);
  table(seqHead, seqRows(S.seqCont));
  if (S.seqCont.length > ROWL) out.append(el('p',{class:'hint'}, `Page shows the first ${ROWL} (of ${S.seqCont.length}); the exported report has the full per-order table.`));

  section('② By due date (slot) · per-order timetable', `Earliest due first; when a batch finishes early, the freed time is filled with later tasks automatically${S.slotLate.length?` (${S.slotLate.length} late, in red)`:''}`);
  table(seqHead, seqRows(S.seqSlot));
  if (S.seqSlot.length > ROWL) out.append(el('p',{class:'hint'}, `Page shows the first ${ROWL} (of ${S.seqSlot.length}); the exported report has the full per-order table.`));

  // ---- ③ deadline check (close to due date, reference) ----
  if (S.deadlines.length){
  section('③ Close to due date (late start · reference) · per due date', `Start as late as possible and finish exactly ${S.buffer} workdays before the due date${S.bufferShort.length?` (${S.bufferShort.length} due dates are tight on capacity: finish lands close to the due date but not late)`:''}`);
  table(['Due date','Due (orders)','Due (hours)','Workday #','③ Close to due: done on','①/② earliest possible','Status'],
    S.deadlines.map(d => {
      const jitDate = S.days[Math.min(S.N-1, d.jitFinishIdx)];
      const eagerDate = S.days[Math.min(S.N-1, d.eagerFinishIdx)];
      const lead = d.jitLeadWD;
      return el('tr',{class:d.feasible?'':'infeasible'},[
        td(isoLocal(d.date)), td(d.count,true), td(round(d.hours)+'h',true), td('Day '+d.T,true),
        td(isoLocal(jitDate) + (d.feasible ? (lead>0?` (${lead} workdays early)`:' (on due date)') : ''), true),
        td(isoLocal(eagerDate), true),
        el('td',{}, d.feasible?el('span',{class:'badge ok'},'Achievable'):el('span',{class:'badge bad'},'Over capacity')) ]);
    }));

  }   // end if (S.deadlines.length) — ③ only makes sense for orders with a due date

  // ---- timeline ----
  section('Daily timeline', 'Blue bar = hours scheduled that day / capacity that day; ① and ② have the same total daily load (both run at full load continuously), they only differ in which order goes first');
  const dlByIdx = new Map(); S.deadlines.forEach(d => dlByIdx.set(d.T-1, d));
  const cap = 120, rows = [];
  for (let i=0;i<S.N && i<cap;i++){
    const dd = S.days[i], e = S.eager[i], j = S.jit[i], dl = dlByIdx.get(i);
    const idle = e <= 1e-9; const capI = S.cap[i]; const low = capI < S.baseCap - 1e-9;
    rows.push(el('tr',{class:(idle?'sched-idle ':'')+(dl?'sched-dl':'')},[
      td(isoLocal(dd)),
      el('td',{class:[0,6].includes(dd.getDay())?'weekend':''}, WEEKDAYS[dd.getDay()]),
      el('td',{class:'num'+(low?' cap-low':'')}, round(capI,1)+'h'+(low?' ⤵':'')),
      el('td',{}, schedBar(e, capI, false, false, idle)),
      el('td',{}, schedBar(j, capI, j>capI+1e-6, true, j<=1e-9)),
      el('td',{}, dl ? el('span',{class:'dl-flag'+(dl.feasible?'':' bad')}, `▼ due ${dl.count} orders`) : '') ]));
  }
  table(['Date','Day','Capacity','①/② sequential load','③ close to due (ref.)','Due'], rows);
  if (S.N > cap) out.append(el('p',{class:'hint'}, `Page shows the first ${cap} workdays (of ${S.N}); click "📤 Export report .xlsx" above to get the full timeline for all ${S.N} workdays plus the full per-order tables.`));

  // ---- ④ per worker · personal schedule (three switchable modes) ----
  if (S.pw && (S.pw.slot.workers.length || S.pw.cont.workers.length || S.pw.jit.workers.length)){
    section('④ Per worker · personal schedule', `Design hours go on the designer's own timeline; check goes on the checker's timeline after that order's design is done; leave deducted per person, annual leave spread evenly; in forward modes the same person finishes their design queue before checking (approximation)${S.pwUnassigned?`; ${S.pwUnassigned} orders missing an assigned worker are not included`:''}`);
    const PW_MODES = [
      ['slot','② By due date (earliest due first)'],
      ['cont','① Sequential (queue by created date)'],
      ['jit', `③ Close to due date (late start · backward from due · ${S.buffer}-workday buffer)`] ];
    if (!PW_MODES.some(m => m[0] === pwMode)) pwMode = 'slot';
    const pickMode = el('select',{});
    PW_MODES.forEach(([v,t]) => pickMode.append(el('option',{value:v}, t)));
    pickMode.value = pwMode;
    const modeBox = el('div',{});
    const renderPW = () => {
      modeBox.innerHTML = '';
      const P = S.pw[pwMode];
      if (!P.workers.length){ modeBox.append(el('p',{class:'hint'}, 'No personal tasks to schedule in this mode.')); return; }
      if (pwMode === 'jit' && P.skipped > 0)
        modeBox.append(el('p',{class:'hint'}, `③ only schedules orders with a due date: ${P.skipped} orders without one are not in this mode (they still take part in ① and ②).`));
      const lastAll = Math.max(-1, ...P.workers.map(w => w.lastIdx));
      tableIn(modeBox, ['Worker','Design (orders/hours)','Check (orders/hours)','Total hours','Busy days','Daily avg','Earliest start','Busy until','Status'],
        P.workers.slice(0, ROW_CAP).map(w => el('tr',{class:(w.unplaced>1e-9||(pwMode!=='jit'&&w.lastIdx===lastAll))?'infeasible':''},[
          td(w.name), td(`${w.nD} orders / ${round(w.dH)}h`,true), td(`${w.nV} orders / ${round(w.vH)}h`,true),
          el('td',{class:'num'}, el('b',{}, round(w.total)+'h')),
          td(w.busyDays,true), td(round(w.avg,1)+'h',true),
          td(w.firstIdx>=0?isoLocal(S.days[w.firstIdx]):'—'),
          td(w.lastIdx>=0?isoLocal(S.days[w.lastIdx]):'—'),
          el('td',{}, w.unplaced>1e-9 ? el('span',{class:'badge bad'}, pwMode==='jit' ? `Overloaded ${round(w.unplaced)}h, can't make due dates` : `Doesn't fit (${round(w.unplaced)}h)`)
            : (pwMode!=='jit' && w.lastIdx===lastAll) ? el('span',{class:'badge bad'},'Sets total duration')
            : el('span',{class:'badge ok'}, pwMode==='jit' ? 'On time' : 'Can finish')) ])));
      if (P.workers.length > ROW_CAP) modeBox.append(el('p',{class:'hint'}, `Showing the first ${ROW_CAP} (of ${P.workers.length}).`));
      // day-by-day progress viewer
      const pick = el('select',{});
      P.workers.forEach(w => pick.append(el('option',{value:w.name}, `${w.name} (${round(w.total)}h)`)));
      const btnDay = el('button',{class:'btn'},"View this worker's daily progress");
      const holder = el('div',{});
      btnDay.onclick = () => {
        const name = pick.value; holder.innerHTML='';
        const arr = P.day.get(name); const lab = P.orders.get(name) || {};
        const rows2 = []; let shown2 = 0;
        for (let i=0; i<S.N && shown2<120; i++){
          if (!(arr && arr[i] > 1e-9)) continue; shown2++;
          const capI = S.dayCapW(name, i);
          rows2.push(el('tr',{},[
            td(isoLocal(S.days[i])),
            el('td',{class:[0,6].includes(S.days[i].getDay())?'weekend':''}, WEEKDAYS[S.days[i].getDay()]),
            el('td',{}, schedBar(arr[i], capI, arr[i] > capI + 1e-6, false, false)),
            td((lab[i]||[]).slice(0,6).join(', ') + ((lab[i]||[]).length>6?` … ${lab[i].length} orders`:'')) ]));
        }
        if (!rows2.length){ holder.append(el('p',{class:'hint'}, 'This worker has no order work scheduled in this mode.')); return; }
        tableIn(holder, ['Date','Day',"Progress that day (÷ this person's capacity)",'Orders handled'], rows2);
        holder.append(el('p',{class:'hint'}, 'Only workdays with work are listed (up to 120 rows); full data is in the "Per-person daily load" sheet of the exported report (exports the currently selected mode).'));
      };
      modeBox.append(el('div',{class:'toolbar',style:'margin:4px 0'},[pick, btnDay]));
      modeBox.append(holder);
    };
    pickMode.onchange = () => { pwMode = pickMode.value; renderPW(); };
    out.append(el('div',{class:'toolbar',style:'margin:4px 0'},[el('span',{class:'hint'},'Schedule mode: '), pickMode]));
    out.append(modeBox);
    renderPW();
  }

  // ---- ⑤ reassignment suggestions ----
  if (S.rebalance){
    const RB = S.rebalance;
    section('⑤ Reassignment suggestions', `Balance by each person's total assigned hours (group average ${round(RB.avgLoad)}h/person); graded people by grade (never someone whose grade can't do it); when no graded candidate fits, "whoever is free" (flagged "ability not filled"); design only moves to designers, check only to checkers`);
    if (!RB.suggestions.length){
      out.append(el('div',{class:'alert ok'}, `✔ Load is fairly balanced: busiest ${round(RB.maxBefore)}h vs average ${round(RB.avgLoad)}h, no worthwhile suggestions (or nobody suitable can take the overloaded work).`));
    } else {
      out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
        `With the suggestions below, the busiest worker's load drops from about ${round(RB.maxBefore)}h to about ${round(RB.maxAfter)}h. These are suggestions only and don't change any data — to actually reassign, change the assigned workers on the Orders tab.`));
      table(['From (current load)','To (current load)','Suggested move','Which orders (stage · hours)'],
        RB.suggestions.slice(0,12).map(s => el('tr',{},[
          td(`${s.from} (${round(s.fromLoad)}h)`), td(`${s.to} (${round(s.toLoad)}h)`),
          td('~'+round(s.amount)+'h', true),
          td(s.examples.slice(0,6).map(e=>`${e.id} (${e.role==='d'?'design':'check'} ${round(e.h,1)}h${e.sure?'':'·ability not filled'})`).join(', ')
            + (s.examples.length>6?` … ${s.examples.length} orders in total (full list in the exported report)`:'')) ])));
      if (RB.suggestions.length > 12)
        out.append(el('p',{class:'hint'}, `${RB.suggestions.length} suggestions in total; the page shows the 12 largest moves; the full list is in the "Reassignment suggestions" sheet of the exported report.`));
      out.append(el('p',{class:'hint'}, '"Ability not filled" = the receiver\'s A/B/C grade for that product and stage is blank — they were suggested as "whoever is free", so confirm their ability manually first; people whose grade explicitly can\'t do it never appear. Design and check are moved separately and never mixed.'));
    }
  }

  // ---- notes ----
  const notes = [];
  notes.push(`The schedule only covers unfinished orders (${S.nOpen} in total); the ${S.nFinished} with an "Actual finish date" are excluded and use no capacity — see the Retrospective tab for them. The start date defaults to today and can be changed.`);
  if (S.lastDlOrders && S.lastDlOrders.length)
    notes.push(`The last due date ${isoLocal(S.lastDl)} is set by ${S.lastDlOrders.slice(0,3).join(', ')}${S.lastDlOrders.length>3?` and ${S.lastDlOrders.length - 3} more`:''} — the timeline runs up to that day.`);
  if (S.noDeadline.length) notes.push(`${S.noDeadline.length} orders have no due date: they still queue in ① and ② (last in ②), but are not part of ③ or lateness checks.`);
  if (S.leaveDaysAffected) notes.push(`Leave included: capacity reduced on ${S.leaveDaysAffected} workdays (about ${round(S.leaveHoursLost)} hours less ≈ ${round(S.leaveHoursLost/S.H,1)} person-days); days marked ⤵ in the capacity column are leave days. Leave records are kept on the Workers tab and only affect this page.`);
  notes.push(`Three methods: "① Sequential" ignores due dates and works through orders by created date one after another; "② By due date" does the earliest due first and fills freed time with later tasks — ① and ② have the same daily total load, they only differ in which order goes first, so per-order finish dates differ; "③ Close to due date" starts as late as possible and finishes exactly ${S.buffer} workdays before the due date (buffer adjustable, reference only).`);
  notes.push(`Annual leave is spread at ${S.annualLeave} days per person per year: 260 workdays a year (52 weeks × 5 days) minus ${S.annualLeave} days → daily capacity × ${round(S.leaveFactor*100)}%, already included above. If the leave records already contain annual leave, set annual leave above to 0 to avoid double-counting.`);
  notes.push('This page schedules against total team capacity; for ability bottlenecks ("who can do which difficulty") see Analysis.');
  out.append(el('div',{class:'alert warn nofold'}, [el('b',{},'Notes: '), el('ul',{}, notes.map(n=>el('li',{},n)))]));
  foldify(out);
}

/* =========================================================
   INIT
   ========================================================= */
function renderAll(){ renderProducts(); renderParts(); renderWorkers(); renderOrders(); renderHourlog(); updateCounts(); }
load();
renderAll();
