/* Parts-Analysis local input tool.
   100% local: data lives in localStorage, Excel handled in-browser via vendored SheetJS.
   Export formats match analyze.py exactly (sheet names, headers, separators). */
'use strict';

/* ---------- domain constants (mirror analyze.py) ---------- */
const DEFAULT_CATEGORY_HOURS = { '高': 4.0, '低': 2.0, '无': 0.25 };
// Live, user-customizable hours-per-difficulty. Mutated in place (never reassigned)
// so analysis.js, which reads it at call-time, always sees the current values.
const CATEGORY_HOURS = { '高': 4.0, '低': 2.0, '无': 0.25 };
const CATEGORIES = ['高', '低', '无'];                 // hardest -> easiest
const ABILITY_SKILLS = { A: ['高', '低', '无'], B: ['低', '无'], C: ['无'] };
const GRADES = ['', 'A', 'B', 'C'];
// Worker ability is split into two roles: 设计 (design, 'd') and 校验 (verify, 'v').
// Each part needs BOTH a design pass and a verify pass; w.grades[product] = { d, v }.
const ROLES = ['d', 'v'];
const ROLE_LABEL = { d: '设计', v: '校验' };
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
const DEFAULT_HOURS_PER_DAY = 8;   // 工作日固定周一至周五（见 analysis.js WORKDAYS_PER_WEEK）
// live, user-customizable shift length (mutated, never re-declared) so analysis.js sees it
let HOURS_PER_DAY = 8;
/* 校对系数：校对工时 = 设计工时 × VERIFY_RATIO。默认 0.3；
   可在「工时台账」用实际报工一键校准（= 台账 订单校对总工时 ÷ 订单设计总工时）。 */
const DEFAULT_VERIFY_RATIO = 0.3;
let VERIFY_RATIO = DEFAULT_VERIFY_RATIO;
function applyVerifyRatio(v){
  v = (v == null || v === '' || isNaN(v) || v <= 0 || v > 3) ? DEFAULT_VERIFY_RATIO : Number(v);
  VERIFY_RATIO = v; state.settings = { ...(state.settings || {}), verifyRatio: v };
}

// header text -> internal field (subset of analyze.py COLUMN_ALIASES)
const COLUMN_ALIASES = {
  '零件名称':'part','part':'part','partname':'part','零件':'part',
  '耗时等级':'category','硬度':'category','难度':'category','category':'category','level':'category','hardness':'category','等级':'category',
  '单件工时':'hours','hours':'hours','工时':'hours',
  '产品类型':'product','类型':'product','产品名称':'product','产品':'product','product':'product','producttype':'product',
  '所需零件':'parts','零件清单':'parts','所需零件清单':'parts','parts':'parts','partsneeded':'parts','partlist':'parts',
  '提交时间':'due','提交日期':'due','交期':'due','due':'due','duedate':'due','submission':'due','submissiontime':'due',
  '工人姓名':'worker','worker':'worker','name':'worker','姓名':'worker',
};
const CATEGORY_ALIASES = {
  '高':'高','high':'高','h':'高','难':'高',
  '低':'低','low':'低','l':'低','中':'低',
  '无':'无','none':'无','no':'无','n':'无','易':'无',
};
const PART_SPLIT_RE = /[;；、\n]+/;
const QTY_TAIL_RE = /[×✕✖]\s*([0-9]+(?:\.[0-9]+)?)\s*$/;
const CODE_PREFIX_RE = /^\s*\d+\s*[:：]\s*/;

// exact headers used by the existing data/*.xlsx (kept byte-for-byte for compatibility)
const SHEET = { parts:'零件清单', workers:'工人能力', orders:'订单明细' };
const HDR = {
  parts: ['零件名称', '耗时等级(高/低/无)'],
  workersName: '工人姓名',
  orders: ['产品类型', '所需零件(用、或,分隔; 数量×N)', '提交时间(可选,YYYY-MM-DD)'],
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
  state.parts = (state.parts || []).map(p => ({ name: p.name || '', category: p.category || '', t7: p.t7 ?? '', t8: p.t8 ?? '' }));
  state.workers = (state.workers || []).map(w => ({ name: w.name || '', grades: migrateGrades(w.grades) }))
    .filter(w => !isSkipWorker(w.name));   // 剔除 null / special order / order canceled 占位名

  state.orders = (state.orders || []).map(o => ({ product: o.product || '', parts: o.parts || '', time: o.time || '', actualFinish: o.actualFinish || '', designer: o.designer || '', checker: o.checker || '', required: o.required || '', so: o.so || '' }));
  state.leaves = (state.leaves || []).map(L => ({ worker: L.worker || '', start: L.start || '', end: L.end || '',
    hours: (L.hours != null && L.hours !== '') ? L.hours : (L.half ? HOURS_PER_DAY / 2 : '') }));   // half → 半个班（旧数据兼容）
  if (!state.hourlog || typeof state.hourlog !== 'object' || !Array.isArray(state.hourlog.workers)) state.hourlog = null;   // 工时台账（hour-report 导入，聚合后存储）
  applyCategoryHours(state.settings && state.settings.categoryHours);
  applyHoursPerDay(state.settings && state.settings.hoursPerDay);
  applyVerifyRatio(state.settings && state.settings.verifyRatio);
  if (typeof state.settings.startDate !== 'string') state.settings.startDate = '';
  if (!(state.settings.schedBuffer >= 0)) state.settings.schedBuffer = 2;   // 排期提前完成缓冲（工作日）
  if (!(state.settings.annualLeave >= 0)) state.settings.annualLeave = 15;  // 年假（天/人·年），均摊扣排期日产能
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
      if (!quotaWarned){ quotaWarned = true; toast('数据过大，本地保存失败（浏览器存储已满）。导出备份后可继续使用。', 'err'); }
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
    const el = byId('dl-' + k); if (el) el.textContent = state[k].length + ' 行';
  });
  const dp = byId('dl-products'); if (dp) dp.textContent = state.products.length + ' 个';
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
  if (state.products.some(p => p === name)){ toast('产品已存在：'+name, 'err'); return false; }
  state.products.push(name);
  state.workers.forEach(w => { if (!(name in w.grades)) w.grades[name] = emptyGrade(); });
  save(); renderProducts();
  return true;
}
function renameProduct(idx, next){
  next = clean(next);
  const prev = state.products[idx];
  if (!next || next === prev) { renderProducts(); return; }
  if (state.products.some((p,i) => p === next && i !== idx)){ toast('已有同名产品', 'err'); renderProducts(); return; }
  state.products[idx] = next;
  state.workers.forEach(w => { w.grades[next] = w.grades[prev] || emptyGrade(); delete w.grades[prev]; });
  state.orders.forEach(o => { if (o.product === prev) o.product = next; });
  save(); renderProducts();
}
function deleteProduct(idx){
  const name = state.products[idx];
  const used = state.orders.filter(o => o.product === name).length;
  let msg = `删除产品「${name}」？将移除工人能力中的该列`;
  if (used) msg += `；有 ${used} 张订单引用了它（订单不会被删除，但产品列会变空）`;
  if (!confirm(msg + '。')) return;
  state.products.splice(idx, 1);
  state.workers.forEach(w => delete w.grades[name]);
  save(); renderProducts();
}
function renderProducts(){
  const ul = byId('product-list'); ul.innerHTML = '';
  if (!state.products.length) ul.append(el('li', { class:'empty-note', style:'list-style:none' }, '还没有产品。上面输入名称添加，或在「导入/导出」里导入。'));
  state.products.forEach((p, i) => {
    const input = el('input', { class:'pname', value:p, spellcheck:'false' });
    input.addEventListener('change', () => renameProduct(i, input.value));
    input.addEventListener('keydown', e => { if (e.key==='Enter') input.blur(); });
    const li = el('li', { class:'chip' }, [
      input,
      el('button', { class:'mini mv', title:'上移', onclick:()=>moveProduct(i,-1) }, '↑'),
      el('button', { class:'mini mv', title:'下移', onclick:()=>moveProduct(i, 1) }, '↓'),
      el('button', { class:'mini', title:'删除', onclick:()=>deleteProduct(i) }, '✕'),
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
byId('product-bulk').onclick = () => openModal('批量添加产品', '每行一个产品名。已存在的会自动跳过。', '', text => {
  let n = 0; text.split('\n').forEach(line => { if (addProduct(line)) n++; });
  toast(`已添加 ${n} 个产品`, 'ok');
});

/* =========================================================
   PARTS  (+ custom hours-per-difficulty settings)
   ========================================================= */
function renderHoursSettings(){
  const bar = byId('hours-settings'); if (!bar) return;
  bar.innerHTML = '';
  bar.append(el('span', { class:'settings-label' }, '单件工时（小时）：'));
  CATEGORIES.forEach(c => {
    const inp = el('input', { type:'number', min:'0', step:'0.25', value:fmtHrs(CATEGORY_HOURS[c]) });
    inp.addEventListener('change', () => {
      let v = parseFloat(inp.value);
      if (isNaN(v) || v < 0){ v = CATEGORY_HOURS[c]; inp.value = fmtHrs(v); toast('请输入 ≥ 0 的数字', 'err'); return; }
      CATEGORY_HOURS[c] = v; state.settings.categoryHours[c] = v; save();
      renderParts();              // refresh dropdown labels + hours column
      byId('analysis-out').innerHTML = '';  // analysis is now stale
      renderHoursSettings();      // refresh the "已改" marker
      toast(`已设 ${c} = ${fmtHrs(v)}h`, 'ok');
    });
    bar.append(el('label', { class:'h-'+c }, [c, inp, el('span', { class:'unit' }, 'h')]));
  });
  bar.append(el('button', { class:'btn ghost', onclick:resetHours }, '恢复默认'));
  bar.append(el('span', { class:'settings-hint' },
    hoursDefault()
      ? '改这里会影响下拉标签、工时统计与分析。'
      : el('span', {}, [el('span', { class:'dirty' }, '已自定义'), '：导出 parts.xlsx 时会附带「单件工时」列，analyze.py 会照此计算。'])));
}
function resetHours(){
  CATEGORIES.forEach(c => { CATEGORY_HOURS[c] = DEFAULT_CATEGORY_HOURS[c]; state.settings.categoryHours[c] = DEFAULT_CATEGORY_HOURS[c]; });
  save(); renderHoursSettings(); renderParts(); byId('analysis-out').innerHTML = '';
  toast('已恢复默认 (高4 / 低2 / 无0.25)', 'ok');
}

/* 校对系数设置栏（零件页）：校对工时 = 设计 × 系数 */
function renderVerifySettings(){
  const bar = byId('verify-settings'); if (!bar) return;
  bar.innerHTML = '';
  bar.append(el('span', { class:'settings-label' }, '校对系数：'));
  const inp = el('input', { type:'number', min:'0.01', max:'3', step:'0.01', value:String(VERIFY_RATIO), style:'width:90px' });
  inp.addEventListener('change', () => {
    const v = parseFloat(inp.value);
    if (isNaN(v) || v <= 0 || v > 3){ inp.value = String(VERIFY_RATIO); toast('请输入 0–3 之间的数字', 'err'); return; }
    applyVerifyRatio(v); save();
    ['analysis-out','sched-out','retro-out'].forEach(id => { const n = byId(id); if (n) n.innerHTML = ''; });   // 口径变了，旧结果作废
    renderVerifySettings();
    toast(`已设校对系数 = ${v}（校对工时 = 设计 × ${v}）`, 'ok');
  });
  bar.append(el('label', {}, [inp]));
  bar.append(el('span', { class:'settings-hint' },
    el('span', {}, [
      `校对工时 = 设计工时 × 系数（当前 ${VERIFY_RATIO}${VERIFY_RATIO === DEFAULT_VERIFY_RATIO ? '，默认' : ''}）。`,
      '在「工时台账」导入 hour report 后可按实际报工一键校准（= 订单校对总工时 ÷ 订单设计总工时）。' ])));
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
    if (shown >= ROW_CAP) return;     // cap rendered rows; use 筛选 to narrow
    shown++;
    const nameI = el('input', { value:p.name, placeholder:'VC 代码' });
    nameI.addEventListener('input', () => { p.name = nameI.value; save(); });
    const sel = el('select', { class:'cell cat-'+p.category });
    [['','—'],['高','高'],['低','低'],['无','无']].forEach(([v,t]) =>
      sel.append(el('option', { value:v, selected:v===p.category }, t)));
    sel.addEventListener('change', () => { p.category = sel.value; sel.className = 'cell cat-'+p.category; save(); });
    const t7I = el('input', { type:'number', min:'0', step:'0.01', value:p.t7, placeholder:'col8', title:'重复订单的设计工时 (来自第8列)' });
    t7I.addEventListener('input', () => { p.t7 = t7I.value; save(); });
    const t8I = el('input', { type:'number', min:'0', step:'0.01', value:p.t8, placeholder:'col9', title:'首单的设计工时 (来自第9列)' });
    t8I.addEventListener('input', () => { p.t8 = t8I.value; save(); });
    tb.append(el('tr', {}, [
      el('td', { class:'rownum' }, String(i+1)),
      el('td', {}, nameI),
      el('td', {}, sel),
      el('td', {}, t7I),
      el('td', {}, t8I),
      el('td', { class:'act' }, el('button', { class:'row-del', title:'删除', onclick:()=>{ state.parts.splice(i,1); save(); renderParts(); } }, '✕')),
    ]));
  });
  if (matched > shown)
    byId('part-table').querySelector('tbody').append(el('tr', {}, el('td', { colspan:'6', style:'padding:12px;color:var(--muted);text-align:center' },
      `仅显示前 ${shown} 个（共 ${matched}${f?' 个匹配':' 个'}）。用上方「筛选」缩小范围。`)));
  updateCounts();
}
byId('part-add').onclick = () => { state.parts.push({ name:'', category:'', t7:'', t8:'' }); save(); partFilter=''; byId('part-filter').value=''; renderParts(); };
byId('part-filter').addEventListener('input', e => { partFilter = e.target.value; renderParts(); });
byId('part-clear').onclick = () => { if (state.parts.length && confirm('清空所有零件？')){ state.parts=[]; save(); renderParts(); } };
byId('part-paste').onclick = () => openModal('粘贴零件（从 Excel 复制两列）',
  '每行：零件名称 + 等级，用 Tab 或逗号分隔。等级可写 高/低/无（也认 high/low/none）。无等级也可，留空即可。',
  '定制外壳\t高\n标准支架\t低\n标准螺丝包\t无', text => {
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
    save(); renderParts(); toast(`已导入 ${n} 行零件`, 'ok');
  });

/* =========================================================
   WORKERS (matrix)
   ========================================================= */
function renderWorkers(){
  renderLeaves();   // 请假登记（不受「先加产品」限制，独立渲染）
  const noprod = byId('worker-noproducts');
  const table = byId('worker-table');
  if (!state.products.length){ noprod.hidden = false; table.style.display='none'; updateCounts(); return; }
  noprod.hidden = true; table.style.display='';
  const thead = table.querySelector('thead'); const tb = table.querySelector('tbody');
  thead.innerHTML=''; tb.innerHTML='';
  const prods = state.products;                          // show ALL products
  // two-row header: product name spans 设计|校验; second row = the sub-columns
  const r1 = el('tr', {}, [ el('th', { class:'rownum', rowspan:'2' }, '#'), el('th', { class:'wname', rowspan:'2' }, '工人姓名') ]);
  prods.forEach(p => r1.append(el('th', { class:'role-head grp-start', colspan:'2' }, p)));
  r1.append(el('th', { class:'act', rowspan:'2' }, ''));
  const r2 = el('tr', {});
  prods.forEach(() => { r2.append(el('th', { class:'sub-d grp-start' }, '设计')); r2.append(el('th', { class:'sub-v' }, '校验')); });
  thead.append(r1, r2);

  const mkSel = (g, role) => {
    const sel = el('select', { class:'mini-grade g-'+(g[role]||''), title: ROLE_LABEL[role] });
    GRADES.forEach(v => sel.append(el('option', { value:v, selected:v===(g[role]||'') }, v || '—')));
    sel.addEventListener('change', () => { g[role] = sel.value; sel.className='mini-grade g-'+sel.value; save(); });
    return sel;
  };
  state.workers.slice(0, ROW_CAP).forEach((w, i) => {
    const nameI = el('input', { value:w.name, placeholder:'姓名' });
    nameI.addEventListener('input', () => { w.name = nameI.value; save(); });
    const tr = el('tr', {}, [ el('td', { class:'rownum' }, String(i+1)), el('td', { class:'matrix-name' }, nameI) ]);
    prods.forEach(p => {
      if (!w.grades[p] || typeof w.grades[p] !== 'object') w.grades[p] = emptyGrade();
      const g = w.grades[p];
      tr.append(el('td', { class:'gradecell grp-start' }, mkSel(g,'d')));
      tr.append(el('td', { class:'gradecell' }, mkSel(g,'v')));
    });
    tr.append(el('td', { class:'act' }, el('button', { class:'row-del', title:'删除', onclick:()=>{ state.workers.splice(i,1); save(); renderWorkers(); } }, '✕')));
    tb.append(tr);
  });
  if (state.workers.length > ROW_CAP)
    tb.append(el('tr', {}, el('td', { colspan: String(2 + prods.length*2 + 1), style:'padding:12px;color:var(--muted);text-align:center' },
      `仅显示前 ${ROW_CAP} 名（共 ${state.workers.length} 名）。`)));
  updateCounts();
}
function addWorker(name){
  const g = {}; state.products.forEach(p => g[p]=emptyGrade());
  state.workers.push({ name: clean(name), grades:g });
}
byId('worker-add').onclick = () => { if (!state.products.length){ toast('请先添加产品', 'err'); return; } addWorker(''); save(); renderWorkers(); };
byId('worker-clear').onclick = () => { if (state.workers.length && confirm('清空所有工人？')){ state.workers=[]; save(); renderWorkers(); } };

/* ---------- 请假登记（仅影响「排期」的每日团队产能） ---------- */
function leaveDailyHours(L){                       // 每天请假的小时数：留空/非法 = 整班 H
  const v = +L.hours;
  return (L.hours !== '' && L.hours != null && isFinite(v) && v > 0) ? Math.min(HOURS_PER_DAY, v) : HOURS_PER_DAY;
}
function leaveSpanDays(L){                         // 起止都含在内的日历天
  const s = parseISO(L.start); if (!s) return 0;
  let e = parseISO(L.end) || s; if (e.getTime() < s.getTime()) e = s;
  return Math.round((e.getTime() - s.getTime()) / 86400000) + 1;
}
function leaveTotalText(L){
  const d = leaveSpanDays(L); if (!d) return '—';
  const h = leaveDailyHours(L);
  return `${round(d * h, 1)}h（${d}天 × ${fmtHrs(h)}h）`;
}
function workerNameOptions(selected){
  const names = [...new Set(state.workers.map(w => w.name).filter(Boolean))];
  if (selected && !names.includes(selected)) names.unshift(selected);   // 保留已删工人的名字，别丢记录
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
    el('th',{class:'rownum'},'#'), el('th',{style:'width:200px'},'工人'),
    el('th',{style:'width:150px'},'起（含）'), el('th',{style:'width:150px'},'止（含）'),
    el('th',{style:'width:150px'},`每日请假(小时)`), el('th',{class:'num',style:'width:150px'},'合计工时'),
    el('th',{class:'act'},'') ])));
  const tb = el('tbody', {});
  leaves.forEach((L, i) => {
    const sel = el('select', {});
    sel.append(el('option', { value:'', selected:!L.worker }, '（选择工人）'));
    workerNameOptions(L.worker).forEach(n => sel.append(el('option', { value:n, selected:n===L.worker }, n)));
    sel.addEventListener('change', () => { L.worker = sel.value; save(); });
    const totCell = el('td', { class:'num' }, leaveTotalText(L));
    const si = el('input', { type:'date', value:L.start || '' });
    si.addEventListener('change', () => { L.start = si.value; if (!L.end) L.end = si.value; save(); renderLeaves(); });
    const ei = el('input', { type:'date', value:L.end || '' });
    ei.addEventListener('change', () => { L.end = ei.value; save(); renderLeaves(); });
    const hi = el('input', { type:'number', step:'0.5', min:'0', max:String(HOURS_PER_DAY),
      placeholder:`整天(${fmtHrs(HOURS_PER_DAY)})`, value:(L.hours===''||L.hours==null) ? '' : L.hours, style:'width:110px' });
    hi.addEventListener('input', () => {                        // 边打边更新合计，不整表重绘（不抢焦点）
      L.hours = hi.value === '' ? '' : (isNaN(parseFloat(hi.value)) ? '' : parseFloat(hi.value));
      totCell.textContent = leaveTotalText(L); save();
    });
    hi.addEventListener('change', () => {                       // 失焦时把非法/超班值规整
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
      el('td',{class:'act'}, el('button',{class:'row-del',title:'删除',onclick:()=>{ state.leaves.splice(i,1); save(); renderLeaves(); }}, '✕')) ]));
  });
  t.append(tb); wrap.append(t);
}
byId('leave-add') && (byId('leave-add').onclick = () => {
  (state.leaves || (state.leaves = [])).push({ worker:'', start:'', end:'', half:false });
  save(); renderLeaves();
});
byId('worker-rand').onclick = () => {
  if (!state.products.length){ toast('请先添加产品', 'err'); return; }
  if (!state.workers.length){ toast('请先添加/导入工人', 'err'); return; }
  const ri = (a,b) => a + Math.floor(Math.random()*(b-a+1));
  const grade = () => ['A','B','C'][ri(0,2)];
  state.workers.forEach(w => {
    const role = ri(0,4);                         // 0-1 只设计, 2-3 只校验, 4 两者都会
    const doD = role <= 1 || role === 4, doV = role >= 2;
    state.products.forEach(p => {
      const g = { d:'', v:'' };
      if (doD && Math.random() < 0.85) g.d = grade();   // 偶尔留空 = 该产品不参与
      if (doV && Math.random() < 0.85) g.v = grade();
      w.grades[p] = g;
    });
  });
  save(); renderWorkers();
  if (byId('analysis-out')) byId('analysis-out').innerHTML = '';
  toast('已随机生成工人能力（测试）', 'ok');
};
byId('worker-paste').onclick = () => {
  const head = '工人姓名\t' + (state.products.length ? state.products.map(p=>`${p}·设计\t${p}·校验`).join('\t') : '产品A·设计\t产品A·校验');
  openModal('粘贴工人能力矩阵（从 Excel 复制）',
    '第一行为表头：<b>工人姓名</b> + 每个产品两列 <b>产品·设计</b> / <b>产品·校验</b>（也认 design/verify）。格子填 A/B/C，留空=该角色不参与。无角色后缀的列视为对该产品设计+校验同等级。',
    head + '\n张三\tA\tB', text => {
      const rows = text.split('\n').map(r => r.split('\t')).filter(r => r.some(c => c.trim()));
      if (rows.length < 2){ toast('需要表头行 + 至少一行数据', 'err'); return; }
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
      save(); renderWorkers(); renderProducts(); toast(`已导入 ${n} 位工人`, 'ok');
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
  if (current && !hasCur) s.append(el('option', { value:current, selected:true }, current + '（未登记）'));
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
    if (shown >= ROW_CAP) return;     // cap rendered rows; use 搜索 to narrow
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
    const soI = el('input', { value:o.so||'', placeholder:'如 123456789-1200', style:'width:150px' });
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
      el('td', { class:'act' }, el('button', { class:'row-del', title:'删除', onclick:()=>{ state.orders.splice(i,1); save(); renderOrders(); } }, '✕')),
    ]));
  });
  if (matched > shown)
    tb.append(el('tr', {}, el('td', { colspan:'9', style:'padding:12px;color:var(--muted);text-align:center' },
      `仅显示前 ${shown} 张（共 ${matched}${f?' 张匹配':' 张'}${f?'':'，全部已导入'}）。${f?'':'用上方「搜索」定位某张订单。'}分析时会计入全部 ${state.orders.length} 张。`)));
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
  return items.map(it => it.qty === 1 ? it.name : `${it.name}×${fmtQty(it.qty)}`).join('、');
}
/* pick parts that already exist in the 零件 tab; quantities via steppers.
   knownNames is built once per render (passed in) to avoid O(orders×parts) work. */
function buildPartsPicker(o, knownNames){
  if (!knownNames) knownNames = new Set(state.parts.map(p => p.name).filter(Boolean));
  const items = orderPartsItems(o);
  const wrap = el('div', { class:'parts-picker' });

  items.forEach((it, idx) => {
    const unknown = !knownNames.has(it.name);
    const qty = el('input', { class:'pp-qty', type:'number', min:'0', step:'1', value:fmtQty(it.qty), title:'数量' });
    qty.addEventListener('input', () => {
      const v = parseFloat(qty.value); if (isNaN(v) || v <= 0) return;
      items[idx].qty = v; o.parts = serializeOrderParts(items); save();
    });
    wrap.append(el('span', { class:'pp-chip' + (unknown ? ' unknown' : ''), title: unknown ? '该零件不在「零件」标签里' : '' }, [
      el('span', { class:'pp-name' }, it.name + (unknown ? ' (未登记)' : '')),
      el('span', {}, '×'), qty,
      el('button', { class:'pp-del', title:'移除', onclick:()=>{ items.splice(idx, 1); o.parts = serializeOrderParts(items); save(); renderOrders(); } }, '✕'),
    ]));
  });

  if (!knownNames.size){
    wrap.append(el('span', { class:'pp-empty' }, '请先在「零件」标签添加零件'));
  } else {
    const chosen = new Set(items.map(it => it.name));
    const OPT_CAP = 300; const avail = []; let more = 0;
    for (const p of state.parts){ if (!p.name || chosen.has(p.name)) continue; if (avail.length < OPT_CAP) avail.push(p.name); else { more++; } }
    const add = el('select', { class:'pp-add' });
    add.append(el('option', { value:'', selected:true }, '+ 添加 VC…'));
    avail.forEach(n => add.append(el('option', { value:n }, n)));
    if (more) add.append(el('option', { value:'', disabled:true }, `…还有 ${more} 个（用筛选）`));
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
byId('order-clear').onclick = () => { if (state.orders.length && confirm('清空所有订单？')){ state.orders=[]; save(); renderOrders(); } };
byId('order-clear2').onclick = () => {
  if (!state.orders.length){ toast('当前没有订单', 'err'); return; }
  if (confirm(`清空全部 ${state.orders.length} 张订单？`)){ state.orders = []; save(); renderOrders(); toast('已清空订单', 'ok'); }
};
byId('order-rand').onclick = () => {
  const partNames = state.parts.map(p => p.name).filter(Boolean);
  if (!state.products.length || !partNames.length){ toast('请先在「产品」「零件」标签添加数据', 'err'); return; }
  const inp = prompt('随机生成多少张测试订单？（用现有产品/零件，随机零件·数量·交付日期·实际完成日期）', '10');
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
    const parts = chosen.join('、');                 // VCs never repeat in an order
    const dueOff = ri(1, 60);                         // 交付日期一定在今天之后
    state.orders.push({ product, parts,
      designer: wnames.length ? pick(wnames) : '', checker: wnames.length ? pick(wnames) : '',
      required: isoOff(dueOff),                        // col6 交付日期(截止)
      time: isoOff(dueOff - ri(5, 40)),               // col7 创建日期(早于交付)
      actualFinish: isoOff(dueOff + ri(-5, 15)) });   // col8 实际完成
  }
  save(); renderOrders(); toast(`已随机生成 ${n} 张测试订单`, 'ok');
};
byId('order-paste').onclick = () => openModal('粘贴订单（从 Excel 复制三列）',
  '每行：产品类型 + 所需零件 + 提交时间(可选)，用 Tab 分隔。零件一格里用 、; 或换行分隔，数量写 ×N。',
  '旗舰款\t定制外壳×1、标准支架×2、标准螺丝包×4\t2026-07-05', text => {
    let n=0;
    text.split('\n').forEach(line => {
      if (!line.trim()) return;
      const c = line.split('\t');
      state.orders.push({ product: clean(c[0]), parts: clean(c[1]), time: normDate(c[2]) });
      n++;
    });
    save(); renderOrders(); toast(`已导入 ${n} 张订单`, 'ok');
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
  ioLog(`  parts: 导入 ${n} 行`);
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
  ioLog(`  orders: 导入 ${n} 张`);
}
function importWorkersAoa(aoa, mode){
  const hdr = aoa[0];
  const nameIdx = Math.max(0, hdr.map(normHeader).indexOf('worker'));
  // each non-name column -> {product, role, i}; role '' = legacy (applies to both 设计/校验)
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
  ioLog(`  workers: 导入 ${n} 位（${cols.length} 列：产品×设计/校验）`);
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
          ioLog(`✔ ${file.name}: 产品 ${list.length} 个`);
        } else if (data.__partsAnalysisBackup || data.parts || data.orders){
          restoreBackup(data); ioLog(`✔ ${file.name}: 整体备份已恢复`);
        } else { ioLog(`⚠ ${file.name}: 无法识别的 JSON`); }
        continue;
      }
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type:'array' });
      ioLog(`✔ ${file.name}`);
      // hour-report 台账：整本拦截（其中 By worker 等表的表头含 Worker，绝不能落进通用识别）
      if (wb.SheetNames.includes('Data (long)')){
        importHourReport(wb, file.name);
        continue;
      }
      // try each sheet; classify by name then by header
      let handled = 0;
      wb.SheetNames.forEach(sn => {
        const aoa = aoaFromSheet(wb.Sheets[sn]);
        if (!aoa.length) return;
        let kind = sn===SHEET.parts?'parts': sn===SHEET.workers?'workers': sn===SHEET.orders?'orders': classifySheet(aoa[0]);
        // filename hint
        if (!kind){ if (lower.includes('part')) kind='parts'; else if (lower.includes('worker')) kind='workers'; else if (lower.includes('order')) kind='orders'; }
        if (kind==='parts'){ importPartsAoa(aoa, mode); handled++; }
        else if (kind==='workers'){ importWorkersAoa(aoa, mode); handled++; }
        else if (kind==='orders'){ importOrdersAoa(aoa, mode); handled++; }
      });
      if (!handled) ioLog(`  ⚠ 无法识别工作表类型（表头需含 零件名称/工人姓名/产品类型 等）`);
    } catch (err){
      console.error(err); ioLog(`✖ ${file.name}: ${err.message}`);
    }
  }
  save(); renderAll();
  toast('导入完成', 'ok');
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
   HOURLOG  (工时台账 — hour-report 导出的 HourReport_{年}_Analysis.xlsx)
   识别依据：工作表 "Data (long)"。列序固定（只会追加不会重排）：
   0 Year · 2 Month# · 3 Group · 4 Team · 5 Worker# · 6 Worker · 7 Role · 9 Task · 12 Hours
   Leave 表：5 Worker# · 6 Worker · 7 Category · 8 Value
   只存聚合结果（工人×任务/月份），不存原始行，避免撑爆 localStorage。
   ========================================================= */
function importHourReport(wb, filename){
  const aoa = aoaFromSheet(wb.Sheets['Data (long)']);
  const wmap = new Map();                       // key = Worker#¦名字
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
    const task = clean(row[9]) || '(未命名任务)'; const m = Number(row[2]);
    w.total += hours;
    w.byTask[task] = (w.byTask[task] || 0) + hours;
    if (m >= 1 && m <= 12){ w.byMonth[m] = (w.byMonth[m] || 0) + hours; months.add(m); }
    taskTotals.set(task, (taskTotals.get(task) || 0) + hours);
    const y = Number(row[0]); if (isFinite(y) && y > 2000) years.set(y, (years.get(y) || 0) + 1);
    nRows++;
  }
  if (!nRows){ ioLog('  ⚠ Data (long) 里没有可用的工时行'); return; }
  if (wb.SheetNames.includes('Leave')){
    const la = aoaFromSheet(wb.Sheets['Leave']);
    for (let r = 1; r < la.length; r++){
      const row = la[r]; const name = clean(row[6]); const val = Number(row[8]);
      if (!name || !isFinite(val) || val === 0) continue;
      const key = clean(row[5]) + '¦' + name;
      let w = wmap.get(key);
      if (!w){ w = { no: clean(row[5]), name, group: clean(row[3]), team: clean(row[4]), role: '',
        total: 0, byTask: {}, byMonth: {}, leave: {}, leaveTotal: 0 }; wmap.set(key, w); }
      const cat = clean(row[7]) || '(未分类)';
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
  ioLog(`  工时台账: ${state.hourlog.workers.length} 位工人 · ${nRows} 行报工 · ${state.hourlog.months.length} 个月（${year || '?'} 年）→ 见「工时台账」页`);
}

byId('hourlog-import') && (byId('hourlog-import').onclick = () => fileInput.click());
byId('hourlog-clear') && (byId('hourlog-clear').onclick = () => {
  if (state.hourlog && confirm('清除已导入的工时台账？（不影响其他数据）')){ state.hourlog = null; save(); renderHourlog(); }
});
byId('hourlog-export') && (byId('hourlog-export').onclick = () =>
  exportReportFromDOM('hourlog-out', `工时台账-${(state.hourlog && state.hourlog.year) || todayISO()}.xlsx`, '工时台账（实际报工）'));

function renderHourlog(){
  const out = byId('hourlog-out'); if (!out) return; out.innerHTML = '';
  const HL = state.hourlog;
  if (!HL){
    out.append(el('div',{class:'empty-note'}, '还没有台账数据。从 hour-report 的 Task Totals 页点「Export Excel」得到 HourReport_{年}_Analysis.xlsx，再点上方「导入」或拖进「导入/导出」页即可。'));
    return;
  }
  const td = (v,num)=>el('td', num?{class:'num'}:{}, v);
  const section = (t,sub)=>out.append(el('div',{class:'section-title'}, sub?el('span',{},[t,el('span',{style:'color:var(--muted);font-weight:500'},'  '+sub)]):t));
  const table = (headers, rows)=>{ const x=el('table',{class:'report'}); const headTr = el('tr',{}, headers.map(h=>el('th',{},h))); x.append(el('thead',{}, headTr)); const b=el('tbody',{}); rows.forEach(r=>b.append(r)); x.append(b); alignNumHeaders(headTr, rows[0]); out.append(el('div',{class:'report-wrap'}, x)); };
  const pct = x => x==null ? '—' : round(100*x,0)+'%';

  const design = HL.workers.reduce((a,w)=>a+(w.byTask['订单设计']||0),0);
  const verify = HL.workers.reduce((a,w)=>a+(w.byTask['订单校对']||0),0);
  const ot = HL.workers.reduce((a,w)=>a+(w.byTask['加班']||0),0);
  const total = HL.workers.reduce((a,w)=>a+w.total,0);
  const M = HL.months.length;
  // 利用率固定按 8h 标准班算（不用 HOURS_PER_DAY——把它换成实际订单工时后，这里会失真）
  const AL = state.settings.annualLeave >= 0 ? state.settings.annualLeave : 15;
  const capPerWorker = M > 0 ? (260 - AL) / 12 * M * DEFAULT_HOURS_PER_DAY : 0;

  const cards = el('div',{class:'cards'});
  const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  card('年份', HL.year || '—');
  card('覆盖月份', HL.months.join('、') || '—', `月（${M} 个）`);
  card('工人', HL.workers.length, '位');
  card('总报工', round(total), '小时');
  card('订单设计', round(design), '小时');
  card('订单校对', round(verify), '小时');
  if (ot > 0) card('加班', round(ot), '小时');
  out.append(cards);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    `来源 ${HL.source || '—'} · 导入于 ${HL.importedAt} · 利用率按标准班计：每人 ${M} 个月 ≈ ${round(capPerWorker)}h（一年 260 个工作日 − 年假 ${AL} 天，× ${fmtHrs(DEFAULT_HOURS_PER_DAY)}h 标准班）。`));

  // ---- 实际工时口径：每人每天真正花在订单设计/校对上的小时数，可一键替换全局「每人每日工时」----
  const owOf = w => (w.byTask['订单设计'] || 0) + (w.byTask['订单校对'] || 0);
  const wdOf = w => Object.keys(w.byMonth || {}).length * 260 / 12;   // 该人有报工的月份 × 每月约21.7个工作日
  const orderWorkers = HL.workers.filter(w => owOf(w) > 0);
  const owSum = orderWorkers.reduce((a, w) => a + owOf(w), 0);
  const wdSum = orderWorkers.reduce((a, w) => a + wdOf(w), 0);
  const actualRate = wdSum > 0 ? round(owSum / wdSum, 2) : null;
  // 校对系数实际值 = 台账 订单校对总工时 ÷ 订单设计总工时
  const actualVR = design > 0 && verify > 0 ? round(verify / design, 2) : null;
  if (actualRate){
    const applied = Math.abs(HOURS_PER_DAY - actualRate) < 0.005 && (actualVR == null || Math.abs(VERIFY_RATIO - actualVR) < 0.005);
    const btn = el('button', { class:'btn primary', style:'margin-left:8px' }, applied ? '✔ 已是当前值' : '应用实际口径');
    if (applied) btn.disabled = true;
    btn.onclick = () => {
      applyHoursPerDay(actualRate);
      if (actualVR != null) applyVerifyRatio(actualVR);      // 校对系数改用实际比例（不再固定 0.3）
      state.settings.annualLeave = 0;                        // 实际报工已包含请假影响，排期年假归零避免重复扣
      save(); renderHourlog(); renderVerifySettings();
      ['analysis-out','sched-out','retro-out'].forEach(id => { const n = byId(id); if (n) n.innerHTML = ''; });
      toast(`已应用实际口径：每人每日工时 ${actualRate}h${actualVR!=null?`，校对系数 ${actualVR}`:''}，排期年假 = 0`, 'ok');
    };
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},[
      el('b',{},'📐 实际工时口径：'),
      `台账里 ${orderWorkers.length} 位有订单报工的工人，平均每人每天实际花在订单设计/校对上 `,
      el('b',{}, `${actualRate} 小时`),
      `（其余时间被沟通/会议/支持占用）`,
      actualVR != null ? el('span',{},[ `；实际校对/设计比 = ${round(verify)}h ÷ ${round(design)}h = `, el('b',{}, String(actualVR)),
        `（当前校对系数 ${VERIFY_RATIO}）` ]) : '',
      `。当前「每人每日工时」= ${fmtHrs(HOURS_PER_DAY)}h。点按钮一键应用实际口径：① 每人每日工时 → ${actualRate}h（分析/复盘/排期的“工作日”与产能按真实节奏算）；` +
      (actualVR != null ? `② 校对系数 → ${actualVR}（校对工时 = 设计 × 实际比例，不再固定 0.3）；` : '') +
      `③ 排期「年假」→ 0（实际数据已含请假影响，避免重复扣）。随时可在排期/复盘设置和「零件」页手动改回。`,
      btn ]));
  }

  // 名字对账（两个方向，分色标注；名字需完全一致才能联动）
  const known = new Set(state.workers.map(w => w.name).filter(Boolean));
  const hlNames = new Set(HL.workers.map(w => w.name));
  const unknown = HL.workers.filter(w => !known.has(w.name)).map(w => w.name);        // 🟠 台账有、名册无
  const rosterOnly = state.workers.map(w => w.name).filter(n => n && !hlNames.has(n)); // 🔵 名册有、台账无
  if (unknown.length && state.workers.length)
    out.append(el('div',{class:'alert warn'}, `🟠 ${unknown.length} 位在台账里报了工，但不在「工人能力」名单中（下表中橙底标出）：${unknown.slice(0,10).join('、')}${unknown.length>10?` 等 ${unknown.length} 位`:''}。若名册里用的是另一种写法（如英文名），把两边改一致即可自动对上。`));
  if (rosterOnly.length)
    out.append(el('div',{class:'alert info'}, `🔵 ${rosterOnly.length} 位在「工人能力」名单中，但台账里没有这个名字的报工：${rosterOnly.slice(0,10).join('、')}${rosterOnly.length>10?` 等 ${rosterOnly.length} 位`:''}。要么确实没报工，要么两边名字写法不同——这些人在「复盘·每位工人负载」里也会以蓝底标出。`));

  section('每位工人 · 全年实际报工', '按总工时降序；利用率 = 总报工 ÷ 该时段产能（口径见上）；🟠 橙底 = 不在「工人能力」名单中');
  const otherOf = w => Math.max(0, w.total - (w.byTask['订单设计']||0) - (w.byTask['订单校对']||0) - (w.byTask['加班']||0));
  table(['工人','组/队','角色','总工时','订单设计','订单校对','加班','其他任务','订单工时/日','请假','利用率'],
    HL.workers.slice(0, ROW_CAP).map(w => {
      const util = capPerWorker > 0 ? w.total / capPerWorker : null;
      const wd = wdOf(w); const rate = (wd > 0 && owOf(w) > 0) ? round(owOf(w) / wd, 1) : null;
      return el('tr',{class: (state.workers.length && !known.has(w.name)) ? 'unmatched-hourlog' : ((util!=null && util>1.000001)?'infeasible':'')},[
        td(w.name), td((w.group?w.group+' / ':'')+(w.team||'—')), td(w.role||'—'),
        td(round(w.total)+'h',true), td(round(w.byTask['订单设计']||0)+'h',true), td(round(w.byTask['订单校对']||0)+'h',true),
        td(round(w.byTask['加班']||0)+'h',true), td(round(otherOf(w))+'h',true),
        td(rate==null?'—':rate+'h', true),
        td(w.leaveTotal ? round(w.leaveTotal,1) : '—', true),
        td(util==null?'—':pct(util), true) ]);
    }));
  if (HL.workers.length > ROW_CAP) out.append(el('p',{class:'hint'}, `仅显示前 ${ROW_CAP} 位（共 ${HL.workers.length}）。`));
  out.append(el('p',{class:'hint'}, '「请假」为 hour-report Leave 表按分类合计的原始数值（单位以该表为准）；利用率超 100% 的行标红（多为加班/跨队支援）。'));

  section('任务合计', '全员全年，按小时降序');
  table(['任务','总工时','占比'],
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
    } catch (err){ console.error(err); toast('导入失败: ' + err.message, 'err'); }
  });
  f.click();
}

// parts.xlsx — positional: col1=VC代码, col7=重复设计工时, col8=首单设计工时
byId('part-import') && (byId('part-import').onclick = () => pickFile('.xlsx,.xls', aoa => {
  if (aoa.length < 2){ toast('文件没有数据行', 'err'); return; }
  if (state.parts.length && confirm(`用文件的 ${aoa.length-1} 行替换现有 ${state.parts.length} 个 VC？（取消=合并）`)) state.parts = [];
  let n = 0;
  aoa.slice(1).forEach(r => {
    const code = clean(r[0]); if (!code) return;
    const t7 = numOrBlank(r[7]), t8 = numOrBlank(r[8]);     // 重复=第8列, 首单=第9列
    const ex = state.parts.find(p => p.name === code);
    if (ex){ ex.t7 = t7; ex.t8 = t8; } else state.parts.push({ name:code, category:'', t7, t8 });
    n++;
  });
  save(); renderParts(); toast(`已导入 ${n} 个 VC（难度请手动设置）`, 'ok');
}));

/* =========================================================
   订单智能导入：自动在前 10 行里找表头行，按同义词把列匹配到字段，
   弹映射确认框（可手动调整），确认后批量导入。识别不了就按旧模板列序预填。
   ========================================================= */
const ORDER_FIELDS = [
  { key:'salesOrder', label:'订单号 Sales Order（可选）', syn:['salesorder','销售订单','so号'], excl:['item','行项目'] },
  { key:'orderItem',  label:'行项目 Item（可选）',        syn:['salesorderitem','orderitem','行项目','item'] },
  { key:'vc',       label:'VC 列表（必选）', syn:['vc','零件','图号','part'] },
  { key:'product',  label:'产品',            syn:['产品','型号','product'] },
  { key:'designer', label:'设计工人',        syn:['设计','designer'], excl:['工时','时间','日期','小时','hours'] },
  { key:'checker',  label:'校对工人',        syn:['校对','校验','审核','checker'], excl:['工时','时间','日期','小时','hours'] },
  { key:'required', label:'交付日期',        syn:['交付','交期','截止','要求完成','deadline','due','required'] },
  { key:'time',     label:'创建日期',        syn:['创建','建单','提交','下单','开单','created','create on','createon','create','submit'] },
  { key:'finish',   label:'实际完成日期',    syn:['实际完成','完成日期','完成时间','实际结束','finish','结案'] },
];
const POSITIONAL_GUESS = { vc:0, designer:2, checker:3, product:4, required:5, time:6, finish:7 };   // 旧模板列序
function normHeaderCell(s){ return String(s == null ? '' : s).toLowerCase().replace(/[\s　_\-：:（）()\[\]]/g, ''); }
function matchOrderField(f, h){
  if (!h) return false;
  if (f.excl && f.excl.some(x => h.includes(x))) return false;
  return f.syn.some(s => h.includes(normHeaderCell(s)));
}
// 遍历整张表找最像表头的一行：按字段顺序贪心认领列，命中字段数最多者胜（≥2 才算找到）。
// 某行命中 ≥5 个字段时直接认定为表头（数据行几乎不可能同时命中 5 个），避免超大文件全表扫描。
function detectOrderHeader(aoa){
  let best = null;
  for (let r = 0; r < aoa.length; r++){
    const cells = (aoa[r] || []).map(normHeaderCell);
    if (!cells.some(Boolean)) continue;                 // 整行为空，跳过
    const map = {}; const claimed = new Set();
    ORDER_FIELDS.forEach(f => {
      for (let c = 0; c < cells.length; c++){
        if (claimed.has(c)) continue;
        if (matchOrderField(f, cells[c])){ map[f.key] = c; claimed.add(c); break; }
      }
    });
    const score = Object.keys(map).length;
    if (score >= 2 && (!best || score > best.score)) best = { row: r, map, score };
    if (score >= 5) break;                              // 足够确定，就是它
  }
  return best;
}
function commitOrdersImport(aoa, headerRow, map){
  const rows = aoa.slice(headerRow + 1);
  const get = (r, key) => map[key] != null && map[key] !== '' ? r[map[key]] : '';
  if (state.orders.length && confirm(`用文件的 ${rows.length} 行替换现有 ${state.orders.length} 张订单？（取消=合并追加）`)) state.orders = [];
  let n = 0;
  const haveP = new Set(state.products), newProds = [];
  const haveV = new Set(state.parts.map(p => p.name)), newVcs = [];
  const haveW = new Set(state.workers.map(w => w.name)), newWorkers = [];
  rows.forEach(r => {
    const vcs = String(get(r,'vc') || '').split(/[;；、,，]/).map(s => s.trim()).filter(Boolean);
    const product = clean(get(r,'product'));
    const designer = clean(get(r,'designer')), checker = clean(get(r,'checker'));
    if (!vcs.length && !product) return;
    // 订单号 = Sales Order + '-' + Item（Item 去掉前导 0；全 0 记作 0）
    const soBase = String(get(r,'salesOrder') ?? '').trim();
    let soItem = String(get(r,'orderItem') ?? '').trim();
    if (soItem) soItem = soItem.replace(/^0+/, '') || '0';
    const so = soBase ? (soItem ? `${soBase}-${soItem}` : soBase) : '';
    state.orders.push({
      product, parts: vcs.join('、'), designer, checker, so,
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
  toast(`已导入 ${n} 张订单` + (newWorkers.length?`，新增 ${newWorkers.length} 名工人`:'') + (newVcs.length?`、${newVcs.length} 个 VC`:''), 'ok');
}
function smartImportOrders(aoa){
  if (!aoa.length){ toast('文件是空的', 'err'); return; }
  const det = detectOrderHeader(aoa);
  const headerRow = det ? det.row : 0;
  const map = det ? { ...det.map } : { ...POSITIONAL_GUESS };
  const header = aoa[headerRow] || [];
  const nCols = Math.max(header.length, ...(aoa.slice(headerRow+1, headerRow+4).map(r => (r||[]).length)), 1);
  const sampleOf = c => { for (let r = headerRow+1; r < Math.min(headerRow+6, aoa.length); r++){
    const v = (aoa[r]||[])[c]; if (v != null && String(v).trim() !== '') return String(v).slice(0, 14); } return ''; };
  // 映射确认框
  const backdrop = el('div', { class:'modal-backdrop' });
  const box = el('div', { class:'modal', style:'max-width:640px' });
  box.append(el('h3', {}, det ? `导入订单 — 已识别表头（第 ${headerRow+1} 行，命中 ${det.score} 个字段）` : '导入订单 — 未识别到表头，已按旧模板列序预填'));
  box.append(el('p', { class:'hint' }, `共 ${aoa.length - headerRow - 1} 行数据。请确认每个字段对应哪一列（下拉里附首个样例值）；不需要的字段选「不导入」。`));
  const sels = {};
  const grid = el('div', { style:'display:grid;grid-template-columns:auto 1fr;gap:8px 12px;align-items:center;margin:10px 0' });
  ORDER_FIELDS.forEach(f => {
    const sel = el('select', {});
    sel.append(el('option', { value:'', selected: map[f.key] == null }, '（不导入）'));
    for (let c = 0; c < nCols; c++){
      const label = `列${c+1}：${String(header[c] ?? '').slice(0,16) || '(无表头)'}${sampleOf(c) ? `｜例 ${sampleOf(c)}` : ''}`;
      sel.append(el('option', { value:String(c), selected: map[f.key] === c }, label));
    }
    sels[f.key] = sel;
    grid.append(el('label', {}, f.label), sel);
  });
  box.append(grid);
  const close = () => backdrop.remove();
  const cancel = el('button', { class:'btn', onclick: close }, '取消');
  const ok = el('button', { class:'btn primary', onclick: () => {
    const m = {}; ORDER_FIELDS.forEach(f => { const v = sels[f.key].value; m[f.key] = v === '' ? null : parseInt(v, 10); });
    if (m.vc == null){ toast('至少要选「VC 列表」对应的列', 'err'); return; }
    const dup = Object.values(m).filter(v => v != null);
    if (new Set(dup).size !== dup.length){ toast('有两个字段选了同一列，请检查', 'err'); return; }
    close(); commitOrdersImport(aoa, headerRow, m);
  } }, '✔ 导入');
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
    // custom times -> add a per-part 单件工时 column so analyze.py uses these hours
    aoa = [[...HDR.parts, '单件工时(小时)'],
      ...state.parts.filter(p=>p.name).map(p => [p.name, p.category,
        CATEGORIES.includes(p.category) ? CATEGORY_HOURS[p.category] : ''])];
    widths = [16, 20, 14];
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa); setCols(ws, widths);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, SHEET.parts); return wb;
}
function buildWorkersWb(){
  // two columns per product: 「<产品>·设计」「<产品>·校验」
  const header = [HDR.workersName];
  state.products.forEach(p => { header.push(`${p}·设计`, `${p}·校验`); });
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

/* Export the rendered 分析/排期/复盘 result to a local .xlsx (what you see = what you get).
   Walks the -out container in document order: cards -> 摘要 sheet, each table -> its own sheet,
   alert text -> 说明 sheet. 100% local (SheetJS), no network. */
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
  if (!box || !box.children.length || !box.textContent.trim()){ toast('请先运行一次，得到结果后再导出', 'err'); return; }
  const used = new Set();
  const wb = XLSX.utils.book_new();
  // 摘要: cards (k/v) + alert lines
  const summary = [[title], [`导出时间 ${new Date().toLocaleString()}`], []];
  box.querySelectorAll('.cards .stat').forEach(s => {
    const k = (s.querySelector('.k')||{}).textContent || '';
    const v = (s.querySelector('.v')||{}).textContent || '';
    if (k.trim()) summary.push([k.trim(), v.trim()]);
  });
  const alerts = [...box.querySelectorAll('.alert')].map(a => a.textContent.trim().replace(/\s+/g,' ')).filter(Boolean);
  if (alerts.length){ summary.push([]); summary.push(['结论 / 说明']); alerts.forEach(t => summary.push([t])); }
  const sumWs = XLSX.utils.aoa_to_sheet(summary); setCols(sumWs, [26, 60]);
  XLSX.utils.book_append_sheet(wb, sumWs, sanitizeSheetName('摘要', used));
  // each table -> its own sheet, named by the nearest preceding section title
  let lastTitle = '';
  box.querySelectorAll('.section-title, table.report').forEach(node => {
    if (node.classList.contains('section-title')){ lastTitle = node.textContent.trim().split(/\s{2,}/)[0]; return; }
    const aoa = tableToAoa(node); if (!aoa.length) return;
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    setCols(ws, aoa[0].map((_, c) => Math.min(40, Math.max(8, ...aoa.map(r => (r[c]||'').length + 2)))));
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(lastTitle || '明细', used));
  });
  // 附加工作表：直接由数据生成（不受页面「仅显示前 N 行」截断影响）
  (extraSheets || []).forEach(s => {
    if (!s || !Array.isArray(s.aoa) || !s.aoa.length) return;
    const ws = XLSX.utils.aoa_to_sheet(s.aoa);
    setCols(ws, s.aoa[0].map((_, c) => Math.min(40, Math.max(8, ...s.aoa.map(r => String(r[c] == null ? '' : r[c]).length + 2)))));
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(s.name || '附表', used));
  });
  XLSX.writeFile(wb, filename);
  toast('已导出报告 ' + filename, 'ok');
}
byId('analysis-export') && (byId('analysis-export').onclick = () =>
  exportReportFromDOM('analysis-out', `分析报告-${todayISO()}.xlsx`, '分析报告（理论工时 vs 实际报工）'));
byId('sched-export') && (byId('sched-export').onclick = () => {
  // 完整时间线（全部工作日，不截断）作为附加工作表
  const extra = [];
  if (lastSched && lastSched.days && lastSched.days.length){
    const S = lastSched;
    const dlByIdx = new Map(); S.deadlines.forEach(d => dlByIdx.set(d.T - 1, d));
    const n2 = x => Math.round(x * 100) / 100;   // 写真正的数值，Excel 里可直接求和/画图
    const aoa = [['日期','周','当日产能(h)','①/② 连续负荷(h)','①/② 累计(h)','③ 贴近交期(h)','③ 累计(h)','当日交期(单)','当日应交(h)']];
    for (let i = 0; i < S.N; i++){
      const dl = dlByIdx.get(i);
      aoa.push([ isoLocal(S.days[i]), WEEKDAYS[S.days[i].getDay()],
        n2(S.cap[i]), n2(S.eager[i]), n2(S.eagerCum[i]), n2(S.jit[i]), n2(S.jitCum[i]),
        dl ? dl.count : '', dl ? n2(dl.hours) : '' ]);
    }
    extra.push({ name:'每日时间线-完整', aoa });
    // ①② 完整逐单表（不受页面 200 行截断）
    const seqAoa = rows => { const a = [['顺序','订单号','产品','设计工人','校对工人','工时(h)','预计开始','预计完成','交期','迟交(工作日,负=提前)']];
      rows.forEach((r, i) => a.push([ i+1, r.so || r.id, r.product || '', r.designer || '', r.checker || '', n2(r.hours),
        r.startIdx == null ? '排不下' : isoLocal(S.days[r.startIdx]),
        r.finishIdx == null ? '排不下' : isoLocal(S.days[r.finishIdx]),
        r.deadline ? isoLocal(r.deadline) : '', r.lateWD == null ? '' : r.lateWD ]));
      return a; };
    extra.push({ name:'① 连续排-逐单', aoa: seqAoa(S.seqCont) });
    extra.push({ name:'② 按交期-逐单', aoa: seqAoa(S.seqSlot) });
    // ④ 每人每日负荷（完整，不受页面 120 行限制；导出页面当前选中的口径）
    const P = S.pw && S.pw[pwMode];
    if (P && P.day && P.day.size){
      const PW_LABEL = { slot:'②按交期', cont:'①连续排', jit:'③贴近交期' };
      const a2 = [['工人','日期','周','当日工时(h)','当日产能(h)','经手订单']];
      [...P.day.keys()].sort().forEach(name => {
        const arr = P.day.get(name); const lab = P.orders.get(name) || {};
        for (let i = 0; i < S.N; i++){ if (arr[i] > 1e-9)
          a2.push([name, isoLocal(S.days[i]), WEEKDAYS[S.days[i].getDay()], n2(arr[i]), n2(S.dayCapW(name,i)), (lab[i]||[]).join('、')]); }
      });
      extra.push({ name:`每人每日负荷(${PW_LABEL[pwMode]||''})`, aoa: a2 });
    }
    if (S.rebalance && S.rebalance.suggestions.length){
      const a3 = [['从','当前负载(h)','转给','当前负载(h)','建议转移(h)','单数','转哪些单']];
      S.rebalance.suggestions.forEach(s => a3.push([ s.from, n2(s.fromLoad), s.to, n2(s.toLoad), n2(s.amount), s.examples.length,
        s.examples.slice(0,300).map(e=>`${e.id}(${e.role==='d'?'设计':'校对'} ${Math.round(e.h*10)/10}h${e.sure?'':'·能力未填'})`).join('、')
        + (s.examples.length>300?` 等 ${s.examples.length} 张`:'') ]));
      extra.push({ name:'重新分配建议', aoa: a3 });
    }
  }
  exportReportFromDOM('sched-out', `排期报告-${todayISO()}.xlsx`, '排期报告（按交期排程）', extra);
});
byId('retro-report') && (byId('retro-report').onclick = () =>
  exportReportFromDOM('retro-out', `复盘报告-${todayISO()}.xlsx`, '复盘报告'));
byId('export-all').onclick = async () => {
  const steps = ['parts','workers','orders','products'];
  for (const s of steps){ exportOne(s); await new Promise(r => setTimeout(r, 350)); }
  toast('已导出 4 个文件，放进 parts-analysis/data/ 即可', 'ok');
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
  try { restoreBackup(JSON.parse(await f.text())); toast('备份已恢复', 'ok'); }
  catch(err){ toast('恢复失败: '+err.message, 'err'); }
  e.target.value='';
});
byId('wipe-all').onclick = () => { if (confirm('确定清空全部数据（产品/零件/工人/订单）？此操作不可撤销。')){ state={products:[],parts:[],workers:[],orders:[],leaves:[],hourlog:null,settings:{categoryHours:{...DEFAULT_CATEGORY_HOURS},hoursPerDay:DEFAULT_HOURS_PER_DAY}}; applyCategoryHours(DEFAULT_CATEGORY_HOURS); applyHoursPerDay(DEFAULT_HOURS_PER_DAY); applyVerifyRatio(DEFAULT_VERIFY_RATIO); save(); renderAll(); toast('已清空', 'ok'); } };

/* =========================================================
   ANALYSIS  (full port of analyze.py — same numbers)
   ========================================================= */
byId('analysis-run').onclick = renderAnalysis;
function renderAnalysis(){
  const out = byId('analysis-out'); out.innerHTML='';
  let R;
  try { R = runFullAnalysis(state); }
  catch(e){ console.error(e); out.append(el('div',{class:'alert warn'}, '分析出错: '+e.message)); return; }

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
    out.append(el('div',{class:'alert warn'}, [el('b',{},'⚠ 需要修正（analyze.py 会因此拒绝运行）：'),
      el('ul',{}, R.errors.slice(0,20).map(e=>el('li',{},e)))]));

  if (!R.hasData){
    const need=[]; if(!state.orders.length)need.push('订单'); if(!state.parts.length)need.push('零件');
    if(!state.workers.length)need.push('工人'); if(!state.products.length)need.push('产品');
    out.append(el('div',{class:'alert warn'}, `需要先有 订单 + 零件 + 工人（+产品）才能分析。缺少：${need.join('、')||'有效数据'}。`));
    if (R.warnings && R.warnings.length)
      out.append(el('div',{class:'alert warn'}, [el('b',{},'数据提示：'), el('ul',{}, R.warnings.map(w=>el('li',{},w)))]));
    return;
  }

  const A = R.asAssigned;
  const pct = x => x==null ? '—' : round(100*x,0)+'%';
  const ROW = 200;
  const HL = state.hourlog;
  const nFin = R.orders.filter(o => o.actualFinish).length;

  /* ===================== ① 理论工时汇总（全部订单） ===================== */
  out.append(el('div',{class:'section-title'},'① 理论工时汇总 — 全部订单按 VC 表计算'));
  const cards = el('div',{class:'cards'});
  const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  card('订单', R.nOrders, `张（已完成 ${nFin} / 未完成 ${R.nOrders - nFin}）`);
  card('设计理论工时', round(R.totalDesign), '小时');
  card('校对理论工时', round(R.totalVerify), '小时');
  card('合计', round(R.totalDemand), '小时');
  const teamDaily = R.nWorkers * HOURS_PER_DAY;   // 全队一天总共能投入的订单工时
  card('团队日产能', round(teamDaily,1), `小时/天（${R.nWorkers} 人 × ${fmtHrs(HOURS_PER_DAY)}h）`);
  if (teamDaily > 0) card('全队需要', round(R.totalDemand/teamDaily,1), '个工作日做完');
  out.append(cards);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    `设计 = Σ(VC工时) × 系数（≤5 个 VC ×1，6–10 ×1.2，>10 ×1.5）；校对 = 设计 × ${VERIFY_RATIO}（校对系数见「零件」页，可用工时台账实际值校准）。产品 ${R.productNames.length} 种 · VC ${R.nParts} 种 · 工人 ${R.nWorkers} 人。`));

  /* ===================== ② 理论 vs 实际报工（全员合计） ===================== */
  if (HL && Array.isArray(HL.workers) && HL.workers.length){
    const actD = HL.workers.reduce((a,w)=>a+(w.byTask['订单设计']||0),0);
    const actV = HL.workers.reduce((a,w)=>a+(w.byTask['订单校对']||0),0);
    section('② 理论 vs 实际报工 · 全员合计', `实际 = 工时台账（${HL.year||'?'} 年 ${HL.months.join('、')} 月）的订单设计/校对报工`);
    table(['环节','理论(h)','实际报工(h)','实际 − 理论(h)','实际 / 理论'],
      [['订单设计', R.totalDesign, actD], ['订单校对', R.totalVerify, actV], ['合计', R.totalDemand, actD+actV]]
      .map(([k,t,a],i)=>el('tr',{},[
        i===2 ? el('td',{},el('b',{},k)) : td(k),
        td(round(t),true), td(round(a),true),
        td((a-t>=0?'+':'')+round(a-t),true), td(t>0?pct(a/t):'—',true) ])));
    out.append(el('p',{class:'hint',style:'margin:-2px 0 2px'},
      '比对前提：订单批次与台账时段要覆盖同一段时间（比如都是 2025 年），否则差值没有意义。实际 > 理论 = 实际花的比 VC 标准多（或订单没录全）；实际 < 理论 = 比标准快（或报工没报全）。校对系数若刚用台账校准过，「订单校对」行的比值会自然贴近设计行——属预期。'));
  } else {
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
      'ℹ 尚未导入工时台账：到「工时台账」页导入 hour report 的 Excel 后，这里会出现「② 理论 vs 实际报工」总量对比和 ③ 的实际列。'));
  }

  /* ===================== ③ 每位工人 · 理论 vs 实际 ===================== */
  {
    const act = new Map(((HL && HL.workers) || []).map(w => [w.name, w]));
    const theo = new Map(A.workerLoads.map(w => [w.name, w]));
    const rows = [...new Set([...theo.keys(), ...act.keys()])].map(n => {
      const t = theo.get(n), a = act.get(n);
      const thD = t ? t.design : 0, thV = t ? t.verify : 0;
      const acD = a ? (a.byTask['订单设计']||0) : 0, acV = a ? (a.byTask['订单校对']||0) : 0;
      return { n, thD, thV, thT: thD+thV, acD, acV, acT: acD+acV };
    }).filter(r => r.thT > 1e-9 || r.acT > 1e-9).sort((x,y)=>(y.thT+y.acT)-(x.thT+x.acT));
    section('③ 每位工人 · 理论 vs 实际', HL ? '理论 = 按订单指派汇总（该人做设计/校对的订单 VC 工时）；实际 = 台账报工' : '理论 = 按订单指派汇总；导入工时台账后这里会多出实际列');
    if (HL){
      table(['工人','理论设计','理论校对','理论合计','实际设计','实际校对','实际合计','实际/理论'],
        rows.slice(0,ROW_CAP).map(r => el('tr',{},[
          td(r.n), td(round(r.thD)+'h',true), td(round(r.thV)+'h',true), el('td',{class:'num'},el('b',{},round(r.thT)+'h')),
          td(round(r.acD)+'h',true), td(round(r.acV)+'h',true), el('td',{class:'num'},el('b',{},round(r.acT)+'h')),
          td((r.thT>0&&r.acT>0)?pct(r.acT/r.thT):'—',true) ])));
    } else {
      table(['工人','理论设计','理论校对','合计','折合工作日','订单(设计/校对)'],
        A.workerLoads.slice(0,ROW_CAP).map(w => el('tr',{},[
          td(w.name), td(round(w.design)+'h',true), td(round(w.verify)+'h',true),
          el('td',{class:'num'}, el('b',{}, round(w.total)+'h')), td(round(w.days,2),true),
          td(`${w.asDesigner}/${w.asChecker}`,true) ])));
    }
    if (rows.length > ROW_CAP) out.append(el('p',{class:'hint'}, `仅显示前 ${ROW_CAP} 位（共 ${rows.length}）。`));
  }

  /* ===================== ④ 订单工时明细 ===================== */
  section('④ 订单工时明细',`设计 = Σ(VC工时) × 系数；校对 = 设计 × ${VERIFY_RATIO}`);
  table(['订单','产品','难度','VC数','系数','设计h','校对h','设计工人','校对工人','交付日期'],
    A.rows.slice(0,ROW).map(r=>el('tr',{},[
      td(r.so || r.id), td(r.product), td(r.difficulty), td(r.nVcs,true), td('×'+r.coeff,true),
      td(round(r.designHours,2),true), td(round(r.verifyHours,2),true),
      el('td',{}, r.designerOk? r.designer||'—' : el('span',{class:'badge bad'}, (r.designer||'空')+' ✖')),
      el('td',{}, r.checkerOk? r.checker||'—' : el('span',{class:'badge bad'}, (r.checker||'空')+' ✖')),
      td(r.deadline||'—') ])));
  if (A.rows.length > ROW) out.append(el('p',{class:'hint'}, `仅显示前 ${ROW} 张（共 ${A.rows.length}）；导出报告含页面各表。`));

  if (R.warnings && R.warnings.length)
    out.append(el('div',{class:'alert warn'}, [el('b',{},'数据提示：'), el('ul',{}, R.warnings.slice(0,10).map(w=>el('li',{},w)))]));
}
function round(x, d=1){ const f = Math.pow(10,d); return String(Math.round(x*f)/f); }
/* 让数字列的表头也右对齐（依据首行单元格是否带 .num），与单元格数字对齐更好扫读 */
function alignNumHeaders(headTr, firstRow){
  if (!firstRow || !firstRow.children || !headTr) return;
  [...firstRow.children].forEach((c, i) => {
    if (c.classList && c.classList.contains('num') && headTr.children[i]) headTr.children[i].classList.add('num');
  });
}

/* =========================================================
   RETROSPECTIVE  (复盘)
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
const RETRO_HDR = ['产品类型', '所需零件(名×N、…)', '提交时间(YYYY-MM-DD)', '实际完成日期(YYYY-MM-DD)'];
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
    if (aoa.length < 2){ toast('文件没有数据行（第 1 行为表头）', 'err'); return; }
    const orders = [];
    aoa.slice(1).forEach(r => {                          // skip header row, read strictly by column index
      const product = clean(r[0]), parts = clean(r[1]);
      if (!product && !parts) return;
      orders.push({ product, parts, time: cellToISO(r[2]), actualFinish: cellToISO(r[3]) });
    });
    if (!orders.length){ toast('没有可导入的订单行', 'err'); return; }
    if (state.orders.length && !confirm(`将用文件中的 ${orders.length} 张订单替换当前 ${state.orders.length} 张订单（含实际完成日期）。继续？`)) return;
    state.orders = orders; save();
    renderOrders(); renderRetroInputs(); byId('retro-out').innerHTML = '';
    toast(`已导入 ${orders.length} 张复盘订单`, 'ok');
  } catch (err){ console.error(err); toast('导入失败: ' + err.message, 'err'); }
});
byId('retro-export') && (byId('retro-export').onclick = () => {
  const rows = state.orders.map(o => [o.product || '', o.parts || '', o.time || '', o.actualFinish || '']);
  if (!rows.length) rows.push(['示例产品', '定制外壳×2、标准支架×4、标准螺丝包×8', '2025-01-06', '2025-01-20']);
  const ws = XLSX.utils.aoa_to_sheet([RETRO_HDR, ...rows]); setCols(ws, [16, 42, 22, 24]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '复盘订单');
  XLSX.writeFile(wb, 'retrospective.xlsx');
});

function renderRetroInputs(){
  const box = byId('retro-inputs'); box.innerHTML = '';
  if (!state.orders.length){ box.append(el('div',{class:'empty-note'}, '还没有订单。请先在「订单」标签添加，或在「导入/导出」里导入历史订单。')); return; }
  const t = el('table', { class:'grid retro-grid' });
  t.append(el('thead', {}, el('tr', {}, [
    el('th',{class:'rownum'},'#'), el('th',{style:'width:160px'},'产品'), el('th',{style:'width:300px'},'VC 列表'),
    el('th',{style:'width:130px'},'创建日期'), el('th',{style:'width:200px'},'实际完成日期'), el('th',{class:'rt-spacer'},'') ])));
  const tb = el('tbody', {});
  state.orders.slice(0, ROW_CAP).forEach((o, i) => {
    const fin = el('input', { type:'date', value:o.actualFinish||'' });
    fin.addEventListener('input', () => { o.actualFinish = fin.value; save(); });
    tb.append(el('tr', {}, [
      el('td',{class:'rownum'}, String(i+1)),
      el('td',{}, o.product || el('span',{style:'color:var(--muted)'},'（空）')),
      el('td',{class:'cell',style:'color:var(--muted);max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:9px 12px'}, o.parts || '—'),
      el('td',{class:'cell',style:'padding:9px 12px;color:var(--muted)'}, o.time || '—'),
      el('td',{}, fin),
      el('td',{class:'rt-spacer'},''),
    ]));
  });
  if (state.orders.length > ROW_CAP)
    tb.append(el('tr', {}, el('td', { colspan:'6', style:'padding:10px;color:var(--muted);text-align:center' }, `仅显示前 ${ROW_CAP} 张（共 ${state.orders.length}）；复盘会计入全部有实际完成日期的订单。`)));
  t.append(tb); box.append(t);
}

function renderRetro(){
  const out = byId('retro-out'); out.innerHTML = '';
  let Ret; try { Ret = runRetrospective(state, {}); }
  catch(e){ console.error(e); out.append(el('div',{class:'alert warn'}, '复盘出错: '+e.message)); return; }

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
      out.append(el('div',{class:'alert warn'}, `这 ${Ret.nAll} 张订单都还没填「实际完成日期」——复盘只分析已完成的订单。请在上方 ① 填写完成日期，再点「复盘分析」；未完成的活请看「排期」页。`));
      return;
    }
    const need=[]; if(!state.orders.length)need.push('订单'); if(!state.parts.length)need.push('零件');
    if(!state.workers.length)need.push('工人');
    out.append(el('div',{class:'alert warn'}, `需要 订单 + 零件 + 工人 才能复盘。缺少：${need.join('、')||'有效数据'}。`));
    return;
  }
  if (!Ret.nWithFinish){
    out.append(el('div',{class:'alert warn'}, '请在上方 ① 给至少一张订单填写「实际完成日期」，再点「复盘分析」。'));
    return;
  }

  // ---- 复盘范围：已完成的订单，最早创建日期 → 今天 ----
  out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
    `复盘范围：${Ret.start ? fmtDate(Ret.start) : '最早创建日'} → 今天（${todayISO()}），共 ${Ret.nWithFinish} 张已完成订单参与` +
    (Ret.nUnfinished > 0 ? `；${Ret.nUnfinished} 张未完成不计入（它们在「排期」页）。` : '。')));

  // ---- ② 真实交付表现（按每张单自己的交付/创建日期，日历天，最直观）----
  out.append(el('div',{class:'section-title'}, el('span',{},['② 复盘结果 · 实际交付表现',
    el('span',{style:'color:var(--muted);font-weight:500'}, '  按每张单自己的交付/创建日期计（日历天）')])));
  const cardsD = el('div',{class:'cards'});
  const cardD=(k,v,u)=>cardsD.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  if (Ret.nWithDeadline){
    cardD('准时率', pct(Ret.onTimeRate));
    cardD('迟交单数', Ret.nLate+' / '+Ret.nWithDeadline, '张');
    if (Ret.nLate){ cardD('中位迟交', Ret.medianLate==null?'—':Ret.medianLate, '日历天'); cardD('最长迟交', Ret.maxLate, '日历天'); }
  } else {
    cardD('准时率', '—', '缺交付日期');
  }
  cardD('中位实际用时', Ret.medianTurnaround==null?'—':Ret.medianTurnaround, '日历天');
  out.append(cardsD);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    '准时率/迟交按每张单自己的「交付日期」；实际用时 = 实际完成 − 该单创建日期。这些只看每张单自身，不受其他订单或数据起点影响，最能反映真实交付。'));

  // ---- 完成速度（逐单）：分位数说人话 + 区间分布条。经过时间不加总（并行会重复计天）----
  out.append(el('div',{class:'section-title',style:'margin-top:8px'}, el('span',{},['完成速度 · 从创建到完成要多久',
    el('span',{style:'color:var(--muted);font-weight:500'}, '  按每张单自己算的工作日（只数周一至周五）')])));
  if (Ret.nTheo){
    const cards = el('div',{class:'cards'});
    const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
    card('一半订单', '≤ '+round(Ret.turnP50,1), '工作日完成');
    card('九成订单', '≤ '+round(Ret.turnP90,1), '工作日完成');
    card('最慢 5%', '≥ '+round(Ret.turnP95,1), `工作日（最长 ${Ret.turnMax}）`);
    card('单均活量(理论)', round(Ret.avgTheoHours,1), '小时/单');
    card('全队需要(理论)', Ret.nWorkers > 0 ? round(Ret.theoSumWD / Ret.nWorkers, 1) : '—',
      `个工作日（总工时 ÷ ${Ret.nWorkers} 人 × ${fmtHrs(Ret.H)}h/天）`);
    out.append(cards);
    out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
      `读法：「一半订单 ≤ ${round(Ret.turnP50,1)} 工作日完成」= 50% 的单从创建日到实际完成日不超过 ${round(Ret.turnP50,1)} 个工作日（共 ${Ret.nTheo} 张有创建+完成日期）。` +
      `对比理论：平均每张单的设计+校对活儿只有约 ${round(Ret.avgTheoHours,1)} 小时——完成慢主要是排队/同时在做别的单，不是活儿大。`));
    // 分布表：大头在哪、拖尾多长，一眼看清
    section('实际用时分布', '每张单从创建到完成的工作日数落在哪个区间');
    const bar = share => el('div',{class:'sbar-cell'},[
      el('div',{class:'sbar'}, el('div',{class:'sbar-fill',style:`width:${Math.min(100, share*100)}%`})),
      el('span',{class:'sbar-num'}, pct(share)) ]);
    table(['实际用时(工作日)','单数','占比','累计'],
      Ret.turnDist.filter(b => b.n > 0).map(b => el('tr',{},[
        td(b.label), td(b.n, true), el('td',{}, bar(b.share)), td('≤ '+pct(b.cumShare), true) ])));
  } else {
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
      'ℹ 无法统计完成速度：这些已完成的订单缺「创建日期」。补上创建日期后即可显示。'));
  }

  // ---- verdict ----  以真实交付（准时率/迟交）下结论，不做「全批跨度 vs 理想工时」的相减
  if (Ret.nWithDeadline){
    if (!Ret.nLate)
      out.append(el('div',{class:'alert ok'}, [el('b',{},'✔ 全部按时交付。'),
        `${Ret.nWithDeadline} 张有交付日期的订单，都在各自交期当天或之前完成。`]));
    else
      out.append(el('div',{class:'alert warn'}, [el('b',{},`✗ 有 ${Ret.nLate} 张迟交，准时率 ${pct(Ret.onTimeRate)}。`),
        ` 最长迟交 ${Ret.maxLate} 天、中位迟交 ${Ret.medianLate} 天（明细见下方「迟交订单」）。`,
        Ret.bestTraining ? el('div',{style:'margin-top:4px'},[el('b',{},'不加人的方向：'),`「${Ret.bestTraining.label}」可把理论最短再缩短约 ${d2(Ret.bestTraining.saved/Ret.H)} 个工作日。`]) : null ]));
  } else {
    out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'}, [el('b',{},'ℹ 无法判断准时。'),
      ` 这些订单没有填交付日期。已完成 ${Ret.nWithFinish} 张，中位实际用时 ${Ret.medianTurnaround==null?'—':Ret.medianTurnaround} 日历天。补上交付日期后即可看准时率。`]));
  }

  // ---- ③ improvement without new labor ----
  out.append(el('div',{class:'section-title'}, '③ 不加人的改进办法'));
  const posLevers = Ret.trainingLevers.filter(l => l.saved/Ret.H >= 0.005);   // 只列有正收益的
  if (posLevers.length){
    section('② 培训现有工人','把现有工人某产品的等级提一级（不招新人），对理论最短的影响（前 10）');
    table(['措施','措施后最短','可省'],
      posLevers.slice(0,10).map((l,i)=> el('tr',{},[
        el('td',{}, i===0?el('b',{},l.label):l.label),
        td(d2(l.hours/Ret.H)+' 天', true), td('↓ '+d2(l.saved/Ret.H)+' 天', true) ])));
  }

  // ---- per-worker load ----
  section('每位工人负载', Ret.loadsAreActual
    ? `实际报工（订单设计+校对，来自工时台账）÷ 一年可用工时（(260 − 年假 ${Ret.annualLeave} 天) × ${fmtHrs(Ret.H)}h ≈ ${round(Ret.capYear)}h/人·年）；超 100% 标红；🔵 蓝底 = 名册里有、但台账里没有这个名字（未报工或两边名字写法不同）`
    : `未导入工时台账——暂以按指派的标准工时 ÷ 一年可用工时（≈ ${round(Ret.capYear)}h/人·年）；超 100% 标红`);
  table(['工人', Ret.loadsAreActual ? '实际工时(报工)' : '指派工时(标准)','闲置工时','利用率','建议'],
    Ret.workerLoads.map(w => {
      const noHL = Ret.loadsAreActual && !w.inHL;
      return el('tr',{class: w.util>1.000001 ? 'infeasible' : (noHL ? 'unmatched-roster' : '')},[
        td(w.name), td(round(w.hours)+'h',true), td(round(w.idle)+'h',true), td(pct(w.util),true),
        el('td',{}, noHL ? '🔵 台账中无此名字（未报工，或两边名字写法不同）'
          : (w.trainTarget ? `闲置较多，可培训其胜任 ${w.trainTarget.product}/${w.trainTarget.cat}/${ROLE_LABEL[w.trainTarget.role]}（需 ${w.trainTarget.need} 级）` : (w.util>=0.95?'已满负荷':'—'))) ]);
    }));

  // ---- ④ 标准工时 vs 实际报工（需要「工时台账」，按名字联动）----
  if (state.hourlog && Array.isArray(state.hourlog.workers) && state.hourlog.workers.length){
    const HL = state.hourlog;
    const std = new Map(((Ret.R.asAssigned && Ret.R.asAssigned.workerLoads) || []).map(w => [w.name, w]));
    const act = new Map(HL.workers.map(w => [w.name, w]));
    const rows = [];
    const rosterSet = new Set(state.workers.map(w => w.name).filter(Boolean));
    [...new Set([...std.keys(), ...act.keys()])].forEach(n => {
      const s = std.get(n), a = act.get(n);
      const sd = s ? s.design : 0, sv = s ? s.verify : 0;                       // 标准：VC 工时按指派汇总
      const ad = a ? (a.byTask['订单设计'] || 0) : 0, av = a ? (a.byTask['订单校对'] || 0) : 0;  // 实际：台账报工
      if (sd + sv + ad + av <= 1e-9) return;
      rows.push({ n, sd, sv, ad, av,
        noHL: rosterSet.has(n) && !act.has(n),   // 蓝：名册有、台账无
        noRoster: !rosterSet.has(n),             // 橙：台账有、名册无
        effD: (sd > 0 && ad > 0) ? sd / ad : null, effV: (sv > 0 && av > 0) ? sv / av : null });
    });
    if (rows.length){
      rows.sort((x, y) => (y.sd + y.ad + y.sv + y.av) - (x.sd + x.ad + x.sv + x.av));
      const effTxt = e => e == null ? '—' : round(100*e, 0) + '%' + (e >= 1.05 ? '（比标准快）' : e <= 0.95 ? '（比标准慢）' : '（≈标准）');
      const sumSd = rows.reduce((t,r)=>t+r.sd,0), sumAd = rows.reduce((t,r)=>t+r.ad,0);
      const sumSv = rows.reduce((t,r)=>t+r.sv,0), sumAv = rows.reduce((t,r)=>t+r.av,0);
      section('④ 标准工时 vs 实际报工', `订单的 VC 标准工时（按指派汇总）对比工时台账实际报工（${HL.year || '?'} 年 ${HL.months.join('、')} 月）；🔵 蓝底 = 名册有但台账无此名，🟠 橙底 = 台账有但名册无此名`);
      if (sumAd > 0 || sumAv > 0)
        out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
          `全组：设计 标准 ${round(sumSd)}h / 实际报工 ${round(sumAd)}h${sumAd>0&&sumSd>0?`，效率 ${round(100*sumSd/sumAd,0)}%`:''}；校对 标准 ${round(sumSv)}h / 实际报工 ${round(sumAv)}h${sumAv>0&&sumSv>0?`，效率 ${round(100*sumSv/sumAv,0)}%`:''}。效率 = 标准 ÷ 实际，>100% 表示比 VC 标准干得快。`));
      table(['工人','标准设计','实际·订单设计','设计效率','标准校对','实际·订单校对','校对效率'],
        rows.slice(0, ROW_CAP).map(r => el('tr',{class:
          r.noHL ? 'unmatched-roster' : r.noRoster ? 'unmatched-hourlog'
          : ((r.effD!=null&&r.effD<=0.7)||(r.effV!=null&&r.effV<=0.7)?'infeasible':'')},[
          td(r.n), td(round(r.sd)+'h',true), td(round(r.ad)+'h',true), td(effTxt(r.effD),true),
          td(round(r.sv)+'h',true), td(round(r.av)+'h',true), td(effTxt(r.effV),true) ])));
      if (rows.length > ROW_CAP) out.append(el('p',{class:'hint'}, `仅显示前 ${ROW_CAP} 位（共 ${rows.length}）。`));
      out.append(el('p',{class:'hint'},
        '前提：两边须覆盖同一时间段（这批订单 ≈ 台账年份）且名字完全一致才有可比性；「加班」任务的报工没算进实际设计/校对；效率明显偏低（≤70%）的行标红——可能是报工口径不同、订单没录全，或确实比标准慢，建议先核对数据再下结论。'));
    }
  }

  // ---- 迟交订单（按每张单自身的日历口径，不用全局锚点/理想工时序号）----
  const ROWL = 200;
  const late = Ret.lateVsDeadline || [];
  section('迟交订单', '迟交 = 实际完成日 − 该单的交付日期；实际用时 = 实际完成日 − 该单的创建日期（都是日历天，只看每张单自己，不受其他订单影响）');
  if (!Ret.nWithDeadline){
    out.append(el('div',{class:'alert warn'}, '这些订单没有填「交付日期」，无法判断是否迟交。请在「订单」页补上交付日期后再复盘。'));
  } else if (!late.length){
    out.append(el('div',{class:'alert ok'}, '✔ 没有迟交订单：凡是填了实际完成日期的订单，都在各自「交付日期」当天或之前完成。'));
  } else {
    table(['订单','产品','创建日期','交付日期','实际完成日','实际用时(日历天)','迟交(日历天)'],
      late.slice(0,ROWL).map(o => el('tr',{class:'infeasible'},[
        td(o.so || o.id), td(o.product),
        td(o.submit?fmtDate(o.submit):'—'), td(o.deadline?fmtDate(o.deadline):'—'), td(fmtDate(o.finish)),
        td(o.turnaroundCal!=null?o.turnaroundCal:'—', true),
        td('+'+o.lateCal+' 天', true) ])));
    if (late.length > ROWL) out.append(el('p',{class:'hint'}, `仅显示迟交最多的前 ${ROWL} 张（共 ${late.length} 张迟交）。`));
  }

  // ---- reconstructed optimal timeline (daily load) ----
  if (Ret.R.sim.daily.length){
    section('重建的最优时间线', '逐日各难度工时与利用率（理想排程）');
    table(['第几天','高 (h)','低 (h)','无 (h)','当日合计','利用率'],
      Ret.R.sim.daily.map((dd,i)=> el('tr',{},[ td('第'+(i+1)+'天'),
        td(round(dd.高),true), td(round(dd.低),true), td(round(dd.无),true), td(round(dd.total)+'h',true),
        td(Ret.R.sim.totalCapacity? round(100*dd.total/Ret.R.sim.totalCapacity)+'%':'—', true) ])));
  }
  foldify(out);
}

/* =========================================================
   SCHEDULE  (排期 — 按交期排程)
   ========================================================= */
const WEEKDAYS = ['日','一','二','三','四','五','六'];
function isoLocal(d){ return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function todayISO(){ return isoLocal(new Date()); }
// 排程开工日：默认今天（排期只排未完成的活，从今天向后排）。
// 用户手动设过未来日期则保留；过去的日期自动拉回今天。
function defaultSchedStart(){
  const today = todayISO();
  const saved = state.settings.startDate;
  return (saved && saved >= today) ? saved : today;
}
function pctSafe(a,b){ return b>0 ? round(100*a/b,0)+'%' : '—'; }
/* 折叠：把每个「小节标题 + 其后的内容」包成可开合的 <details>；超过 30 行的表默认折叠。
   带 nofold 类的元素不入折叠组（如页尾「说明」）。导出报告不受影响：折叠只是显示状态，内容都在 DOM 里。 */
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
    if (!det.open) sum.append(el('span',{class:'fold-hint'}, `（含 ${rows} 行，点击展开）`));
    det.append(sum);
    x.items.forEach(it => det.append(it));
  });
  const bar = el('div',{class:'toolbar fold-bar'},[
    el('span',{class:'hint'}, nClosed ? `已自动折叠 ${nClosed} 个长小节，点小节标题可展开/收起：` : '点小节标题可折叠/展开：'),
    (()=>{ const b=el('button',{class:'btn'},'全部展开'); b.onclick=()=>out.querySelectorAll('details.fold').forEach(d=>d.open=true); return b; })(),
    (()=>{ const b=el('button',{class:'btn'},'全部折叠'); b.onclick=()=>out.querySelectorAll('details.fold').forEach(d=>d.open=false); return b; })() ]);
  out.insertBefore(bar, out.firstChild);
}
function schedBar(load, C, over, lvl, idle){
  const w = C>0 ? Math.min(100, load/C*100) : 0;
  const bar = el('div',{class:'sbar'+(lvl?' lvl':'')+(over?' over':'')}, el('div',{class:'sbar-fill',style:`width:${w}%`}));
  const label = idle ? '空窗' : (load>1e-9 ? round(load)+'h ('+pctSafe(load,C)+')' : '—');
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
  if (!state.orders.length){ box.append(el('div',{class:'empty-note'}, '还没有订单。请先在「订单」标签添加。')); return; }
  const t = el('table', { class:'grid retro-grid' });
  t.append(el('thead', {}, el('tr', {}, [
    el('th',{class:'rownum'},'#'), el('th',{style:'width:160px'},'产品'), el('th',{style:'width:300px'},'VC 列表'),
    el('th',{style:'width:170px'},'交付日期（截止）'), el('th',{style:'width:140px'},'创建日期'), el('th',{class:'rt-spacer'},'') ])));
  const tb = el('tbody', {});
  state.orders.slice(0, ROW_CAP).forEach((o, i) => {
    const dl = el('input', { type:'date', value:o.required||'' });      // 交付日期 = 排期用的截止
    dl.addEventListener('input', () => { o.required = dl.value; save(); });
    tb.append(el('tr', {}, [
      el('td',{class:'rownum'}, String(i+1)),
      el('td',{}, o.product || el('span',{style:'color:var(--muted)'},'（空）')),
      el('td',{class:'cell',style:'color:var(--muted);max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:9px 12px'}, o.parts || '—'),
      el('td',{}, dl),
      el('td',{class:'cell',style:'padding:9px 12px;color:var(--muted)'}, o.time || '—'),
      el('td',{class:'rt-spacer'},''),
    ]));
  });
  if (state.orders.length > ROW_CAP)
    tb.append(el('tr', {}, el('td', { colspan:'6', style:'padding:10px;color:var(--muted);text-align:center' }, `仅显示前 ${ROW_CAP} 张（共 ${state.orders.length}）；排程会计入全部。`)));
  t.append(tb); box.append(t);
}

let lastSched = null;   // 最近一次排程结果，供「导出报告」生成完整时间线（不受显示截断影响）
let pwMode = 'slot';    // ④ 每位工人排期当前口径：'cont' ① 连续排 / 'slot' ② 按交期 / 'jit' ③ 贴近交期
function renderSchedule(){
  const out = byId('sched-out'); out.innerHTML='';
  let S; try { S = runSchedule(state, { startDate: byId('sched-start').value, bufferDays: parseInt(byId('sched-buffer') && byId('sched-buffer').value, 10),
    annualLeaveDays: parseInt(byId('sched-al') && byId('sched-al').value, 10) }); }
  catch(e){ console.error(e); lastSched = null; out.append(el('div',{class:'alert warn'}, '排程出错: '+e.message)); return; }
  lastSched = (S && S.days && S.days.length) ? S : null;

  const td = (v,num)=>el('td', num?{class:'num'}:{}, v);
  const section = (t,sub)=>out.append(el('div',{class:'section-title'}, sub?el('span',{},[t,el('span',{style:'color:var(--muted);font-weight:500'},'  '+sub)]):t));
  const tableIn = (host, headers, rows)=>{ const x=el('table',{class:'report'}); const headTr = el('tr',{}, headers.map(h=>el('th',{},h))); x.append(el('thead',{}, headTr)); const b=el('tbody',{}); rows.forEach(r=>b.append(r)); x.append(b); alignNumHeaders(headTr, rows[0]); host.append(el('div',{class:'report-wrap'}, x)); };
  const table = (headers, rows)=>tableIn(out, headers, rows);

  if (!S.hasData){ out.append(el('div',{class:'alert warn'}, '需要 订单 + 零件 + 工人 才能排程。')); return; }
  if (S.noSchedule){
    out.append(el('div',{class:'alert warn'},
      S.nFinished > 0
        ? `没有未完成的订单：全部 ${S.nFinished} 张都已填「实际完成日期」。排期只排未完成的活；历史表现请看「复盘」页。`
        : '没有可排程的订单：请先添加订单（其零件需已在「零件」里登记）。'));
    return;
  }

  const C = S.capacity;
  // ---- cards ----
  out.append(el('div',{class:'section-title'}, '② 排程结果'));
  const cards = el('div',{class:'cards'});
  const card=(k,v,u)=>cards.append(el('div',{class:'stat'},[el('div',{class:'k'},k),el('div',{class:'v',html:v+(u?` <span class="u">${u}</span>`:'')})]));
  card('总工时', round(S.total), '小时');
  card('团队日产能(扣年假)', round(S.baseCap,1), '小时/天');
  card('全部做完', '第 '+S.eagerFinishDay, `工作日（${isoLocal(S.days[Math.min(S.N,S.eagerFinishDay)-1])}）`);
  card('② 按交期迟交', S.slotLate.length + ' / ' + S.scheduledOrders.length, '张');
  card('③ 贴交期缓冲', S.buffer, '工作日');
  out.append(cards);
  out.append(el('p',{class:'hint',style:'margin:-6px 0 2px'},
    `开工 ${isoLocal(S.start)} · ${S.nWorkers} 人 × ${fmtHrs(S.H)}h = 满编 ${round(C)}h/天，扣年假 ${S.annualLeave} 天/人·年（× ${round(S.leaveFactor*100)}%）≈ ${round(S.baseCap,1)}h/天 · 工作日=周一至周五 · 末交期 ${isoLocal(S.days[S.N-1])}（第 ${S.N} 个工作日）`));

  if (S.overdue && S.overdue.length)
    out.append(el('div',{class:'alert warn'}, [el('b',{},`⚠ ${S.overdue.length} 张订单的交期早于开工日：`),
      `${S.overdue.slice(0,8).join('、')}${S.overdue.length>8?` 等 ${S.overdue.length} 张`:''}。下表里它们会标红为「产能不足」，实际含义是交期在开工日之前排不进去——请前移开工日或核对交期。`]));

  if (S.farDeadlines && S.farDeadlines.length)
    out.append(el('div',{class:'alert warn'}, [el('b',{},`⚠ 时间线被拉到 ${isoLocal(S.lastDl)}：`),
      `有 ${S.farDeadlines.length} 张订单的交付日期比最晚创建日期（${S.maxCreated?isoLocal(S.maxCreated):'—'}）还晚超过 1 年，可能是填错或导入错位：` +
      S.farDeadlines.slice(0,5).map(f=>`${f.id}（交付 ${isoLocal(f.deadline)}）`).join('、') +
      (S.farDeadlines.length>5?` 等 ${S.farDeadlines.length} 张`:'') +
      `。排期时间线的终点 = 最晚交付日期，请到「订单」页用搜索框搜年份（如 ${S.lastDl.getFullYear()}）核对修正。`]));

  if (S.slotLate.length)
    out.append(el('div',{class:'alert warn'}, `✖ 产能不足：即便按「② 按交期」排法（交期最早的先做、连续不空闲）也有 ${S.slotLate.length} 张会迟交（②表标红）。办法：提前开工 / 加班，或培训更多工人分担瓶颈难度（见「分析」「复盘」）。`));
  else if (S.scheduledOrders.length)
    out.append(el('div',{class:'alert ok'}, `✔ 按「② 按交期」排法所有订单都能在各自交期当天或之前完成；全部活最快第 ${S.eagerFinishDay} 个工作日（${isoLocal(S.days[Math.min(S.N,S.eagerFinishDay)-1])}）做完。逐单的预计开始/完成日期见下面 ①② 两张表。`));

  // ---- ①② 逐单时间表 ----
  const ROWL = 200;
  const lateTxt = r => r.lateWD == null ? '—' : r.lateWD > 0 ? `迟 ${r.lateWD} 天` : r.lateWD === 0 ? '交期当天' : `早 ${-r.lateWD} 天`;
  const seqRows = rows => rows.slice(0, ROWL).map((r, i) => el('tr',{class:(r.lateWD != null && r.lateWD > 0) ? 'infeasible' : ''},[
    td(i+1, true), td(r.so || r.id), td(r.product || '—'),
    td(r.designer || '—'), td(r.checker || '—'),
    td(round(r.hours,1)+'h', true),
    td(r.startIdx == null ? '排不下' : isoLocal(S.days[r.startIdx])),
    td(r.finishIdx == null ? '排不下(超1500天)' : isoLocal(S.days[r.finishIdx])),
    td(r.deadline ? isoLocal(r.deadline) : '—'),
    td(lateTxt(r), true) ]));
  const seqHead = ['顺序','订单','产品','设计工人','校对工人','工时','预计开始','预计完成','交期','提前/迟交(工作日)'];

  section('① 连续排 · 逐单时间表', `不管交期：按创建日期排队，干完一个做下一个、连续不空闲${S.contLate.length?`（迟交 ${S.contLate.length} 张，标红)`:''}`);
  table(seqHead, seqRows(S.seqCont));
  if (S.seqCont.length > ROWL) out.append(el('p',{class:'hint'}, `页面仅显示前 ${ROWL} 张（共 ${S.seqCont.length}）；导出报告含完整逐单表。`));

  section('② 按交期(slot) · 逐单时间表', `交期最早的先做；某批活提前做完，空出的时间自动排后面的任务${S.slotLate.length?`（迟交 ${S.slotLate.length} 张，标红)`:''}`);
  table(seqHead, seqRows(S.seqSlot));
  if (S.seqSlot.length > ROWL) out.append(el('p',{class:'hint'}, `页面仅显示前 ${ROWL} 张（共 ${S.seqSlot.length}）；导出报告含完整逐单表。`));

  // ---- ③ deadline check (贴近交期，参考) ----
  if (S.deadlines.length){
  section('③ 贴近交期(晚开工·参考) · 各交期达成', `尽量晚开工、在交期前 ${S.buffer} 个工作日刚好完成${S.bufferShort.length?`（${S.bufferShort.length} 个交期产能较紧，完成日贴近交期但不迟交）`:''}`);
  table(['交期','应交(单)','应交(工时)','工作日序','③ 贴近交期：做完日','①/② 连续最早可交','状态'],
    S.deadlines.map(d => {
      const jitDate = S.days[Math.min(S.N-1, d.jitFinishIdx)];
      const eagerDate = S.days[Math.min(S.N-1, d.eagerFinishIdx)];
      const lead = d.jitLeadWD;
      return el('tr',{class:d.feasible?'':'infeasible'},[
        td(isoLocal(d.date)), td(d.count,true), td(round(d.hours)+'h',true), td('第'+d.T+'天',true),
        td(isoLocal(jitDate) + (d.feasible ? (lead>0?`（提前 ${lead} 个工作日）`:'（交期当天）') : ''), true),
        td(isoLocal(eagerDate), true),
        el('td',{}, d.feasible?el('span',{class:'badge ok'},'可达'):el('span',{class:'badge bad'},'产能不足')) ]);
    }));

  }   // end if (S.deadlines.length) —— ③ 只对有交期的单有意义

  // ---- timeline ----
  section('每日时间线', '蓝条 = 当日安排工时 / 当日产能；①②的每日总负荷相同（都是连续满负荷），区别只在先做哪张单');
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
      el('td',{}, dl ? el('span',{class:'dl-flag'+(dl.feasible?'':' bad')}, `▼ 交期 ${dl.count}单`) : '') ]));
  }
  table(['日期','周','产能','①/② 连续负荷','③ 贴近交期(参考)','交期'], rows);
  if (S.N > cap) out.append(el('p',{class:'hint'}, `页面仅显示前 ${cap} 个工作日（共 ${S.N}）；点上方「📤 导出报告 .xlsx」可得到全部 ${S.N} 个工作日的完整时间线与完整逐单表。`));

  // ---- ④ 每位工人 · 个人排期（三种口径可切换）----
  if (S.pw && (S.pw.slot.workers.length || S.pw.cont.workers.length || S.pw.jit.workers.length)){
    section('④ 每位工人 · 个人排期', `设计工时排进设计工人自己的时间线，校对在该单设计完成后排进校对工人时间线；请假按人扣、年假均摊；正排口径下同一人先做完设计队列再做校对（近似）${S.pwUnassigned?`；${S.pwUnassigned} 张缺指派工人的订单未包含`:''}`);
    const PW_MODES = [
      ['slot','② 按交期（交期最早的先做）'],
      ['cont','① 连续排（按创建日期排队）'],
      ['jit', `③ 贴近交期（晚开工·从交期倒排·缓冲 ${S.buffer} 个工作日）`] ];
    if (!PW_MODES.some(m => m[0] === pwMode)) pwMode = 'slot';
    const pickMode = el('select',{});
    PW_MODES.forEach(([v,t]) => pickMode.append(el('option',{value:v}, t)));
    pickMode.value = pwMode;
    const modeBox = el('div',{});
    const renderPW = () => {
      modeBox.innerHTML = '';
      const P = S.pw[pwMode];
      if (!P.workers.length){ modeBox.append(el('p',{class:'hint'}, '这个口径下没有可排的个人任务。')); return; }
      if (pwMode === 'jit' && P.skipped > 0)
        modeBox.append(el('p',{class:'hint'}, `③ 只排有交期的订单：${P.skipped} 张没填交期的订单不在此口径内（它们在 ①② 里照常参与）。`));
      const lastAll = Math.max(-1, ...P.workers.map(w => w.lastIdx));
      tableIn(modeBox, ['工人','设计(单/工时)','校对(单/工时)','总工时','忙碌天数','日均','最早开工','预计忙到','状态'],
        P.workers.slice(0, ROW_CAP).map(w => el('tr',{class:(w.unplaced>1e-9||(pwMode!=='jit'&&w.lastIdx===lastAll))?'infeasible':''},[
          td(w.name), td(`${w.nD} 单 / ${round(w.dH)}h`,true), td(`${w.nV} 单 / ${round(w.vH)}h`,true),
          el('td',{class:'num'}, el('b',{}, round(w.total)+'h')),
          td(w.busyDays,true), td(round(w.avg,1)+'h',true),
          td(w.firstIdx>=0?isoLocal(S.days[w.firstIdx]):'—'),
          td(w.lastIdx>=0?isoLocal(S.days[w.lastIdx]):'—'),
          el('td',{}, w.unplaced>1e-9 ? el('span',{class:'badge bad'}, pwMode==='jit' ? `超载 ${round(w.unplaced)}h 贴不进交期` : `排不下(${round(w.unplaced)}h)`)
            : (pwMode!=='jit' && w.lastIdx===lastAll) ? el('span',{class:'badge bad'},'决定总工期')
            : el('span',{class:'badge ok'}, pwMode==='jit' ? '可按期' : '可完成')) ])));
      if (P.workers.length > ROW_CAP) modeBox.append(el('p',{class:'hint'}, `仅显示前 ${ROW_CAP} 位（共 ${P.workers.length}）。`));
      // 逐日进度查看器
      const pick = el('select',{});
      P.workers.forEach(w => pick.append(el('option',{value:w.name}, `${w.name}（${round(w.total)}h）`)));
      const btnDay = el('button',{class:'btn'},'查看该工人的每日进度');
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
            td((lab[i]||[]).slice(0,6).join('、') + ((lab[i]||[]).length>6?` 等 ${lab[i].length} 张`:'')) ]));
        }
        if (!rows2.length){ holder.append(el('p',{class:'hint'}, '这位工人在此口径下没有被排到任何订单工作。')); return; }
        tableIn(holder, ['日期','周','当日进度（÷ 该人当日产能）','经手订单'], rows2);
        holder.append(el('p',{class:'hint'}, '只列有活的工作日（最多 120 行）；完整数据见「导出报告」的「每人每日负荷」表（导出当前选中的口径）。'));
      };
      modeBox.append(el('div',{class:'toolbar',style:'margin:4px 0'},[pick, btnDay]));
      modeBox.append(holder);
    };
    pickMode.onchange = () => { pwMode = pickMode.value; renderPW(); };
    out.append(el('div',{class:'toolbar',style:'margin:4px 0'},[el('span',{class:'hint'},'排程口径：'), pickMode]));
    out.append(modeBox);
    renderPW();
  }

  // ---- ⑤ 重新分配建议 ----
  if (S.rebalance){
    const RB = S.rebalance;
    section('⑤ 重新分配建议', `按每人被指派的总工时均衡（全组平均 ${round(RB.avgLoad)}h/人）；有评级按评级（评级不胜任的绝不推荐），没有能胜任的评级人选时按「谁空谁来」（标「能力未填」）；设计环节只转给设计、校对环节只转给校对`);
    if (!RB.suggestions.length){
      out.append(el('div',{class:'alert ok'}, `✔ 负载相对均衡：最忙 ${round(RB.maxBefore)}h vs 平均 ${round(RB.avgLoad)}h，暂无值得调整的建议（或超载者的活找不到合适的人接）。`));
    } else {
      out.append(el('div',{class:'alert',style:'border:1px solid var(--line)'},
        `按下列建议调整后，最忙工人的负载可从约 ${round(RB.maxBefore)}h 降到约 ${round(RB.maxAfter)}h。这只是建议，不会改动数据——实际调整请到「订单」页改指派工人。`));
      table(['从（当前负载）','转给（当前负载）','建议转移','转哪些单（环节·工时）'],
        RB.suggestions.slice(0,12).map(s => el('tr',{},[
          td(`${s.from}（${round(s.fromLoad)}h）`), td(`${s.to}（${round(s.toLoad)}h）`),
          td('约 '+round(s.amount)+'h', true),
          td(s.examples.slice(0,6).map(e=>`${e.id}（${e.role==='d'?'设计':'校对'} ${round(e.h,1)}h${e.sure?'':'·能力未填'}）`).join('、')
            + (s.examples.length>6?` 等 ${s.examples.length} 张（完整清单见导出报告）`:'')) ])));
      if (RB.suggestions.length > 12)
        out.append(el('p',{class:'hint'}, `共 ${RB.suggestions.length} 条建议，页面仅显示转移量最大的前 12 条；完整清单见「导出报告」的「重新分配建议」表。`));
      out.append(el('p',{class:'hint'}, '「能力未填」= 接收人该产品该环节的 A/B/C 等级空白——是按「谁空谁来」推荐的，请先人工确认其能力；评级明确不胜任的人不会出现在建议里。设计与校对环节分开转移，不会混。'));
    }
  }

  // ---- notes ----
  const notes = [];
  notes.push(`排期只排未完成的订单（共 ${S.nOpen} 张）；已填「实际完成日期」的 ${S.nFinished} 张不参与、不占产能——它们的表现看「复盘」页。开工日期默认今天，可手动修改。`);
  if (S.lastDlOrders && S.lastDlOrders.length)
    notes.push(`末交期 ${isoLocal(S.lastDl)} 由 ${S.lastDlOrders.slice(0,3).join('、')}${S.lastDlOrders.length>3?` 等 ${S.lastDlOrders.length} 张`:''} 的交付日期决定——时间线铺到这一天为止。`);
  if (S.noDeadline.length) notes.push(`${S.noDeadline.length} 张订单没填交期：照常参与 ①② 排队（②里排在最后），但不参与 ③ 和迟交判定。`);
  if (S.leaveDaysAffected) notes.push(`已计入请假：${S.leaveDaysAffected} 个工作日产能被调低（合计少约 ${round(S.leaveHoursLost)} 小时 ≈ ${round(S.leaveHoursLost/S.H,1)} 人·天），产能列里标 ⤵ 的即请假日。请假记录在「工人」页维护，只影响本页。`);
  notes.push(`三种排法：「① 连续排」不管交期，按创建日期排队干完一个做下一个；「② 按交期」交期最早的先做、提前做完自动把后面的任务填进空余——①② 每日总负荷相同，区别在先做哪张单，所以逐单完成日不同；「③ 贴近交期」尽量晚开工、在交期前 ${S.buffer} 个工作日刚好完成（缓冲可调，仅供参考）。`);
  notes.push(`年假按每人每年 ${S.annualLeave} 天均摊：一年 260 个工作日（52 周 × 5 天）扣 ${S.annualLeave} 天 → 日产能 × ${round(S.leaveFactor*100)}%，已计入上表。若「请假登记」里登记的就是年假，请把上方年假改为 0，避免重复扣。`);
  notes.push('本页按团队总产能排时间；“谁能做哪种难度”的能力瓶颈请看「分析」。');
  out.append(el('div',{class:'alert warn nofold'}, [el('b',{},'说明：'), el('ul',{}, notes.map(n=>el('li',{},n)))]));
  foldify(out);
}

/* =========================================================
   INIT
   ========================================================= */
function renderAll(){ renderProducts(); renderParts(); renderWorkers(); renderOrders(); renderHourlog(); updateCounts(); }
load();
renderAll();
