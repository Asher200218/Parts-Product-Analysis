/* analysis.js — in-browser engine (VC-time model, role-aware).
    • Each part = a VC (3-digit code) with a difficulty (High/Low/None, for worker eligibility),
      repeat design hours t7, first-order design hours t8.
    • Per (product, VC): the chronologically earliest order (by created date) uses t8
      (first order); later ones use t7 (repeat). t8 blank → use t7. Any value 0 → 0.03.
    • Order design = (Σ its VC times) × coeff(VC count: ≤5→1, 6–10→1.2, >10→1.5).
      Order check = design × VERIFY_RATIO (default 0.3, can be calibrated from the hour log).   (precedence: design → check)
    • Two views: ① as assigned (fixed design/check worker per order)  ② what-if (optimal).
   Relies on app.js globals: CATEGORIES, ABILITY_SKILLS, GRADE_RANK, HOURS_PER_DAY,
   ROLES, ROLE_LABEL, clean, normGrade, normCategory, normDate, parseOrderParts. */
'use strict';

const MAX_EXACT_GROUPS = 22;
const MAX_SIM_DAYS = 3650;
const GRADE_RANK = { A: 3, B: 2, C: 1 };
const DIFF_RANK = { 'High': 3, 'Low': 2, 'None': 1 };
const SPECIAL_VC = { '999': 0.08 };   // special VC code -> fixed design hours
const WORKDAYS_PER_WEEK = 5;          // workdays are fixed = Monday–Friday (weekends carry no hours/scheduling)
// design/check worker set to one of these placeholders -> the whole order is void (excluded from all analysis)
const SKIP_WORKERS = new Set(['special order', 'order canceled', 'order cancelled', 'null']);
function isSkipWorker(name){ return SKIP_WORKERS.has(String(name || '').trim().toLowerCase()); }

/* ---------- helpers ---------- */
function roleName(r){ return ROLE_LABEL[r] || r; }
function roleGrade(w, product, role){ const g = w.grades[product]; return (g && g[role]) || ''; }
function workerCan(w, product, category, role){
  const g = roleGrade(w, product, role);
  return !!g && ABILITY_SKILLS[g] && ABILITY_SKILLS[g].includes(category);
}
function workerBestRank(w){
  let r = 0;
  for (const p in w.grades){ const g = w.grades[p]; if (!g) continue;
    for (const role of ROLES){ const rk = GRADE_RANK[g[role]]; if (rk > r) r = rk; } }
  return r;
}
function upgradeGrade(g){ return ({ C:'B', B:'A' })[g] || null; }
function neededGrade(cat){ return cat === 'High' ? 'A' : (cat === 'Low' ? 'B' : 'C'); }
function catIndex(c){ return CATEGORIES.indexOf(c); }
function gkey(g){ return g.prod + '¦' + g.cat + '¦' + g.role; }
function numOrNull(v){ if (v == null || String(v).trim() === '') return null; const n = Number(v); return isNaN(n) ? null : n; }
function parseISO(s){ if (!s) return null; const m = String(s).match(/(\d{4})-(\d{1,2})-(\d{1,2})/); return m ? new Date(+m[1], +m[2]-1, +m[3]) : null; }
const DAY_MS = 86400000;
function diffDays(a, b){ return Math.round((a - b) / DAY_MS); }
// number of workdays elapsed between a→b (Mon–Fri only; same day = 0)
function businessDaysBetween(a, b){
  const days = diffDays(b, a); if (days <= 0) return 0;
  let n = Math.floor(days / 7) * 5;
  const rem = days % 7, dow = a.getDay();
  for (let i = 1; i <= rem; i++){ const d = (dow + i) % 7; if (d !== 0 && d !== 6) n++; }
  return n;
}

/* ---------- input resolution ---------- */
function buildAnalysisInputs(state){
  const errors = [], warnings = [];
  const productNames = state.products.slice();

  // VC parts: code -> { category, t7, t8 }
  const partMap = new Map();
  state.parts.forEach(p => {
    const code = clean(p.name); if (!code) return;
    partMap.set(code, { category: normCategory(p.category), t7: numOrNull(p.t7), t8: numOrNull(p.t8) });
  });

  // workers (design/check grades). Keep workers even with NO grades filled in — the ① as-assigned
  // view needs them (times come from the VC table, not from ability); grades only power the
  // capability check + the ② what-if.  `graded` = has at least one A/B/C.
  const workers = [];
  state.workers.forEach(w => {
    const name = clean(w.name); if (!name) return;
    const grades = {}; let any = false;
    for (const p in (w.grades || {})){
      const src = w.grades[p] || {};
      const d = normGrade(src.d), v = normGrade(src.v);
      if (d || v){ grades[p] = { d, v }; any = true; }
    }
    workers.push({ name, grades, graded: any });
  });
  const workerByName = new Map(workers.map(w => [w.name, w]));

  // orders
  const orders = []; let seq = 0, skipped = 0;
  state.orders.forEach((o, idx) => {
    seq++;
    const product = clean(o.product);
    const designer = clean(o.designer), checker = clean(o.checker);
    // design/check worker is special order / order canceled / null -> void the whole order
    if (isSkipWorker(designer) || isSkipWorker(checker)){ skipped++; return; }
    const vcs = parseOrderParts(o.parts).map(x => x.name).filter(Boolean);
    if (!product && !vcs.length) return;
    // any VC not yet known -> add as a blank VC (will get the 0.03 fallback)
    vcs.forEach(c => { if (!partMap.has(c)) partMap.set(c, { category:'', t7:null, t8:null }); });
    orders.push({ id:`Order${seq}`, srcIndex: idx, product, vcs, so: clean(o.so),
      created: clean(o.time), deadline: clean(o.required),
      designer, checker, actualFinish: clean(o.actualFinish) });
  });
  if (skipped) warnings.push(`Voided ${skipped} orders (design/check worker is special order / order canceled / null)`);

  computeOrderTimes(orders, partMap);

  return { productNames, partMap, nPartsTotal: partMap.size, workers, workerByName, orders, errors, warnings };
}

/* ---------- VC time computation (first-time vs repeat, coefficient, check = design × VERIFY_RATIO) ---------- */
function computeOrderTimes(orders, partMap){
  // chronological order by created date (blank dates last, stable)
  const order = orders.map((_, i) => i);
  const tval = o => { const d = parseISO(o.created); return d ? d.getTime() : Infinity; };
  order.sort((a, b) => (tval(orders[a]) - tval(orders[b])) || (a - b));
  const seen = new Set();                         // product¦vc  -> already used in an earlier order
  const firstFlag = orders.map(() => ({}));       // order idx -> { vc: isFirstOrderForThisProduct }
  order.forEach(i => {
    const o = orders[i];
    o.vcs.forEach(vc => { const k = o.product + '¦' + vc; if (!seen.has(k)){ seen.add(k); firstFlag[i][vc] = true; } else firstFlag[i][vc] = false; });
  });

  orders.forEach((o, i) => {
    let base = 0; const detail = [];
    o.vcs.forEach(vc => {
      const p = partMap.get(vc) || { t7:null, t8:null, category:'' };
      const first = !!firstFlag[i][vc];
      let t;
      if (SPECIAL_VC[vc] != null) t = SPECIAL_VC[vc];        // special codes like 999: fixed hours
      else {
        t = first ? (p.t8 != null ? p.t8 : p.t7) : p.t7;    // first order uses t8 (t7 if blank); repeats use t7
        if (t == null) t = 0;
        if (t === 0) t = 0.03;                               // 0 → base time 0.03
      }
      base += t; detail.push({ vc, time: t, first });
    });
    const nv = o.vcs.length;
    const coeff = nv <= 5 ? 1 : (nv <= 10 ? 1.2 : 1.5);
    o.nVcs = nv; o.coeff = coeff; o.vcDetail = detail;
    o.designHours = base * coeff;
    o.verifyHours = o.designHours * (typeof VERIFY_RATIO !== 'undefined' ? VERIFY_RATIO : 0.3);   // check ratio can be calibrated from the hour log
    o.total = o.designHours + o.verifyHours;
    // order difficulty = hardest VC difficulty (blank VC → None)
    let diff = 'None', rank = 0;
    o.vcs.forEach(vc => { const c = (partMap.get(vc) || {}).category; const cc = CATEGORIES.includes(c) ? c : 'None'; if (DIFF_RANK[cc] > rank){ rank = DIFF_RANK[cc]; diff = cc; } });
    o.difficulty = diff;
  });
}

/* ---------- what-if groups: order design/check hours bucketed by (product, difficulty, role) ---------- */
function buildGroups(orders){
  const map = new Map();
  orders.forEach(o => {
    [['d', o.designHours], ['v', o.verifyHours]].forEach(([role, hrs]) => {
      if (hrs <= 1e-12) return;
      const g = { prod: o.product, cat: o.difficulty, role };
      const k = gkey(g); const cur = map.get(k) || { ...g, hours: 0 }; cur.hours += hrs; map.set(k, cur);
    });
  });
  return [...map.values()];
}

/* ---------- exact minimum makespan (Hall/subset bound) ---------- */
function minMakespan(items, workers, longestUnit = 0){
  const filt = items.filter(it => it.hours > 1e-12);
  const n = filt.length;
  if (n === 0) return { hours: 0, binding: [], infeasible: [] };
  const masks = workers.map(w => { let m = 0; for (let i = 0; i < n; i++) if (workerCan(w, filt[i].prod, filt[i].cat, filt[i].role)) m |= (1 << i); return m; });
  const bd = i => ({ prod: filt[i].prod, cat: filt[i].cat, role: filt[i].role });
  const infeasible = [];
  for (let i = 0; i < n; i++) if (!masks.some(mk => mk & (1 << i))) infeasible.push(bd(i));
  if (infeasible.length) return { hours: null, binding: [], infeasible };
  let bestVal = -1, binding = [];
  if (n <= MAX_EXACT_GROUPS){
    const total = 1 << n; let bestS = 0;
    for (let S = 1; S < total; S++){
      let num = 0; for (let i = 0; i < n; i++) if (S & (1 << i)) num += filt[i].hours;
      let den = 0; for (const mk of masks) if (mk & S) den++;
      const val = num / den; if (val > bestVal){ bestVal = val; bestS = S; }
    }
    for (let i = 0; i < n; i++) if (bestS & (1 << i)) binding.push(bd(i));
  } else {
    for (let i = 0; i < n; i++){ let den = 0; for (const mk of masks) if (mk & (1 << i)) den++; const v = filt[i].hours / den; if (v > bestVal){ bestVal = v; binding = [bd(i)]; } }
    const full = (1 << n) - 1; let den = 0; for (const mk of masks) if (mk & full) den++;
    const tot = filt.reduce((a, x) => a + x.hours, 0);
    if (tot / den > bestVal){ bestVal = tot / den; binding = filt.map((_, i) => bd(i)); }
  }
  return { hours: Math.max(bestVal, longestUnit), binding, infeasible: [] };
}

function formatDuration(hours){
  if (hours === null) return 'infeasible';
  const days = hours / HOURS_PER_DAY, whole = Math.floor(days), rem = hours - whole * HOURS_PER_DAY;
  return `${hours.toFixed(1)} h/person ≈ ${days.toFixed(2)} workdays (${whole}d ${rem.toFixed(1)}h)`;
}
function describeBinding(binding, workers, infeasible){
  if (infeasible && infeasible.length)
    return `Infeasible: no worker can do ${infeasible.map(g=>`${g.prod} ${g.cat} difficulty (${roleName(g.role)})`).join(', ')}`;
  if (!binding.length) return 'No clear bottleneck';
  const prods = [...new Set(binding.map(b=>b.prod))].sort();
  const cats = [...new Set(binding.map(b=>b.cat))].sort((a,b)=>catIndex(a)-catIndex(b));
  const roles = [...new Set(binding.map(b=>b.role))].map(roleName);
  const elig = workers.filter(w => binding.some(g => workerCan(w,g.prod,g.cat,g.role))).length;
  return `${prods.join('/')} ${cats.join('/')} difficulty (${roles.join('/')}) (only ${elig} can do it)`;
}
function bindingEnd(binding){
  if (!binding.length) return 'None';
  const cats = new Set(binding.map(b=>b.cat));
  if (cats.has('High')) return 'High-end (grade A)';
  if (cats.size === 1 && cats.has('None')) return 'Low-end (grade C is fine)';
  return 'Mid-range (grade B)';
}
function unblockSuggestions(infeasible, workers){
  return infeasible.map(g => {
    const need = neededGrade(g.cat);
    const trainable = workers.filter(w => { const cur = roleGrade(w, g.prod, g.role); return cur && GRADE_RANK[cur] < GRADE_RANK[need]; }).map(w => w.name).sort();
    return { product: g.prod, category: g.cat, role: g.role, need, trainable };
  });
}
function feasibleMinMakespan(items, workers, longestUnit, infeasible){
  const bad = new Set(infeasible.map(gkey));
  const feasible = items.filter(it => !bad.has(gkey(it)));
  if (!feasible.length) return 0;
  return minMakespan(feasible, workers, longestUnit).hours;
}

/* ---------- improvement: train / hire / reassign (role-aware) ---------- */
function cloneWorkers(workers){ return workers.map(w => ({ name: w.name, grades: Object.fromEntries(Object.entries(w.grades).map(([p,g]) => [p, { ...g }])) })); }
function hiredWorker(product, role, grade){ return { name:`(new hire ${product}-${roleName(role)}${grade})`, grades:{ [product]: role==='d'?{d:grade,v:''}:{d:'',v:grade} } }; }
function improvementAnalysis(items, workers, longestUnit){
  const base = minMakespan(items, workers, longestUnit);
  const levers = [];
  if (base.hours === null) return { base, levers };
  const pairMap = new Map();
  base.binding.forEach(g => { const k = g.prod + '¦' + g.role; const e = pairMap.get(k) || { prod: g.prod, role: g.role, need: 'C' }; if (GRADE_RANK[neededGrade(g.cat)] > GRADE_RANK[e.need]) e.need = neededGrade(g.cat); pairMap.set(k, e); });
  const pairs = [...pairMap.values()];
  pairs.forEach(pr => {
    const nw = workers.concat([hiredWorker(pr.prod, pr.role, pr.need)]);
    const h = minMakespan(items, nw, longestUnit).hours;
    levers.push({ label:`Hire 1 grade-${pr.need} ${roleName(pr.role)} worker for ${pr.prod}`, kind:'hire', cost: 3.0 + (3 - GRADE_RANK[pr.need]) * 0.1, hours: h, saved: base.hours - h });
  });
  // training: any worker with the same (product, role, currentGrade) gives the identical result,
  // so test only ONE representative per key (keeps this O(pairs×grades), scalable to tens of thousands of workers)
  const seenTrain = new Set();
  for (const w of workers) for (const pr of pairs){
    const cur = roleGrade(w, pr.prod, pr.role); const up = upgradeGrade(cur); if (!up) continue;
    const key = pr.prod + '¦' + pr.role + '¦' + cur; if (seenTrain.has(key)) continue; seenTrain.add(key);
    const nw = workers.map(x => { if (x.name !== w.name) return x;   // reuse refs; deep-copy only this worker
      const g2 = { ...x.grades, [pr.prod]: { ...(x.grades[pr.prod] || { d:'', v:'' }) } }; g2[pr.prod][pr.role] = up; return { name:x.name, grades:g2 }; });
    const h = minMakespan(items, nw, longestUnit).hours;
    if (base.hours - h > 1e-9) levers.push({ label:`Train ${w.name}: ${pr.prod} ${roleName(pr.role)} ${cur}→${up}`, kind:'train', cost:2.0, hours:h, saved: base.hours - h });
  }
  return { base, levers };
}
function bestLever(levers){ const u = levers.filter(l => l.saved != null && l.saved > 1e-6); return u.length ? u.sort((a,b)=>(b.saved-a.saved)||(a.cost-b.cost))[0] : null; }
function hireSensitivity(items, workers, longestUnit, product, role, grade, maxAdd = 6){
  const out = []; const extra = [];
  for (let k = 0; k <= maxAdd; k++){ const h = minMakespan(items, workers.concat(extra), longestUnit).hours; out.push({ added: k, hours: h, days: h == null ? null : h / HOURS_PER_DAY }); extra.push(hiredWorker(product, role, grade)); }
  return out;
}

/* ---------- what-if greedy schedule (optimal assignment, order-level design→check) ---------- */
function schedKey(o){ return o.deadline ? [0, o.deadline, o.id] : [1, o.created || '', o.id]; }
function cmpKey(a, b){ for (let i=0;i<a.length;i++){ if (a[i]<b[i]) return -1; if (a[i]>b[i]) return 1; } return 0; }
function simulate(orders, workers){
  const ordered = orders.slice().sort((a,b)=>cmpKey(schedKey(a), schedKey(b)));
  const orderIds = ordered.map(o=>o.id);
  const tasks = []; const orderRemaining = {};
  ordered.forEach(o => {
    const d = { order:o.id, product:o.product, cat:o.difficulty, role:'d', remaining:o.designHours, design:null };
    const v = { order:o.id, product:o.product, cat:o.difficulty, role:'v', remaining:o.verifyHours, design:d };
    if (o.designHours > 1e-12) tasks.push(d);
    if (o.verifyHours > 1e-12) tasks.push(v);
    orderRemaining[o.id] = o.designHours + o.verifyHours;
  });
  const completion = {}; orderIds.forEach(id => { if (!(orderRemaining[id] > 0)) completion[id] = 0; });
  const avail = t => t.remaining > 1e-9 && (t.role === 'd' || t.design.remaining <= 1e-9);
  const simWorkers = workers.slice().sort((a,b) => (workerBestRank(b) - workerBestRank(a)) || (a.name<b.name?-1:a.name>b.name?1:0));
  const daily = []; const allocations = []; let stuck = []; let day = 0;
  const totalCapacity = workers.length * HOURS_PER_DAY;
  while (tasks.some(t => t.remaining > 1e-9) && day < MAX_SIM_DAYS){
    day++; const used = { High:0, Low:0, None:0 }; let progress = 0;
    for (const worker of simWorkers){
      let capacity = HOURS_PER_DAY;
      for (const cat of CATEGORIES) for (const role of ROLES){
        for (const task of tasks){
          if (capacity <= 1e-9) break;
          if (task.cat !== cat || task.role !== role || !avail(task)) continue;
          if (!workerCan(worker, task.product, cat, role)) continue;
          const take = Math.min(capacity, task.remaining);
          task.remaining -= take; capacity -= take; progress += take; used[cat] += take;
          orderRemaining[task.order] -= take;
          allocations.push({ day, worker:worker.name, grade:roleGrade(worker, task.product, role), order:task.order, product:task.product, cat, role, hours:take });
          if (orderRemaining[task.order] <= 1e-9 && !(task.order in completion)) completion[task.order] = day;
        }
        if (capacity <= 1e-9) break;
      }
    }
    used.total = used.High + used.Low + used.None; daily.push(used);
    if (progress <= 1e-9){
      const s = new Map(); tasks.forEach(t => { if (t.remaining > 1e-9){ const g = { prod:t.product, cat:t.cat, role:t.role }; s.set(gkey(g), g); } });
      stuck = [...s.values()]; day--; daily.pop(); break;
    }
  }
  return { orderSequence: orderIds, completion, days: day, daily, totalCapacity, allocations, stuck, unfinishedHours: tasks.reduce((a,t)=>a+t.remaining, 0) };
}

/* ---------- as-assigned view: fixed design/check per order ---------- */
function runAsAssigned(orders, workerByName, H, wpw){
  // capability: if the worker has NO grade for this (product,role) we treat it as UNKNOWN and
  // don't flag (they were assigned, and we may just lack ability data). Only flag a grade that's
  // present but insufficient for the order's difficulty.
  const can = (name, product, diff, role) => {
    if (!name) return true;
    const w = workerByName.get(name); if (!w) return true;
    const g = roleGrade(w, product, role); if (!g) return true;   // unknown ability -> don't flag
    return ABILITY_SKILLS[g] && ABILITY_SKILLS[g].includes(diff);
  };
  // anchor for deadline mapping = earliest created (fallback earliest deadline / today)
  const dates = orders.map(o => parseISO(o.created) || parseISO(o.deadline)).filter(Boolean);
  const anchor = dates.length ? new Date(Math.min(...dates)) : new Date();
  const calToWork = cal => Math.max(0, Math.round(cal * (wpw || 7) / 7));

  const ordered = orders.slice().sort((a,b)=>cmpKey(schedKey(a), schedKey(b)));
  const free = {};                                  // worker -> cumulative busy hours
  const rows = [];
  ordered.forEach(o => {
    const dStart = free[o.designer] || 0, dEnd = dStart + o.designHours; free[o.designer] = dEnd;
    const vStart = Math.max(free[o.checker] || 0, dEnd), vEnd = vStart + o.verifyHours; free[o.checker] = vEnd;
    const dl = parseISO(o.deadline);
    const deadlineWD = dl ? calToWork(diffDays(dl, anchor) + 1) : null;
    const finishWD = vEnd / H;
    rows.push({ id:o.id, so:o.so, product:o.product, designer:o.designer, checker:o.checker,
      designHours:o.designHours, verifyHours:o.verifyHours, total:o.total, nVcs:o.nVcs, coeff:o.coeff, difficulty:o.difficulty,
      deadline:o.deadline, finishWD, deadlineWD,
      late: (deadlineWD != null) ? (finishWD > deadlineWD + 1e-9) : null,
      designerOk: !o.designer || can(o.designer, o.product, o.difficulty, 'd'),
      checkerOk: !o.checker || can(o.checker, o.product, o.difficulty, 'v') });
  });

  // per-worker load (design as designer + verify as checker)
  const load = new Map();
  const bump = (name, k, v) => { if (!name) return; const e = load.get(name) || { name, design:0, verify:0, asDesigner:0, asChecker:0 }; e[k] += v; load.set(name, e); };
  orders.forEach(o => { bump(o.designer, 'design', o.designHours); bump(o.designer, 'asDesigner', 1); bump(o.checker, 'verify', o.verifyHours); bump(o.checker, 'asChecker', 1); });
  const workerLoads = [...load.values()].map(e => ({ ...e, total: e.design + e.verify, days: (e.design + e.verify) / H })).sort((a,b)=>b.total-a.total);

  const makespanDays = Math.max(0, ...Object.values(free).map(h => h / H));
  const lateOrders = rows.filter(r => r.late).sort((a,b)=>(b.finishWD-(b.deadlineWD||0))-(a.finishWD-(a.deadlineWD||0)));
  const capMismatch = rows.filter(r => !r.designerOk || !r.checkerOk);
  const noAssign = rows.filter(r => !r.designer || !r.checker);
  return { rows, workerLoads, makespanDays, lateOrders, capMismatch, noAssign, anchor, wpw: wpw || 7,
    totalDesign: orders.reduce((a,o)=>a+o.designHours,0), totalVerify: orders.reduce((a,o)=>a+o.verifyHours,0) };
}

/* ---------- order summaries (for schedule/retro reuse) ---------- */
function summarizeOrders(orders){
  const out = {};
  orders.forEach(o => { out[o.id] = { product:o.product, due:o.created, deadline:o.deadline, total:o.total,
    design:o.designHours, verify:o.verifyHours, hours:{High:0,Low:0,None:0}, partsTxt:o.vcs.join(', ') };
    out[o.id].hours[o.difficulty] = o.total; });
  return out;
}
function capacityTable(productNames, groups, workers){
  const byKey = new Map(groups.map(g => [gkey(g), g.hours]));
  const rows = [];
  productNames.forEach(prod => CATEGORIES.forEach(cat => ROLES.forEach(role => {
    const hours = byKey.get(gkey({ prod, cat, role })) || 0; if (hours <= 0) return;
    const capable = workers.filter(w => workerCan(w, prod, cat, role));
    rows.push({ prod, cat, role, hours, cap: capable.length, names: capable.map(w=>w.name), infeasible: capable.length===0 });
  })));
  return rows;
}

/* ====================================================================
   Top level
   ==================================================================== */
function runFullAnalysis(state){
  const b = buildAnalysisInputs(state);
  const { productNames, partMap, nPartsTotal, workers, workerByName, orders, errors, warnings } = b;
  const usedVcs = new Set(); orders.forEach(o => o.vcs.forEach(v => usedVcs.add(v)));
  const nGraded = workers.filter(w => w.graded).length;
  const result = { productNames, nPartsTotal, errors, warnings, nGraded,
    nOrders: orders.length, nWorkers: workers.length, nParts: usedVcs.size,
    hasData: orders.length > 0 && workers.length > 0 };
  if (!result.hasData) return result;

  const groups = buildGroups(orders);
  const demandByCat = { High:0, Low:0, None:0 };
  groups.forEach(g => demandByCat[g.cat] += g.hours);
  const totalDesign = orders.reduce((a,o)=>a+o.designHours,0);
  const totalVerify = orders.reduce((a,o)=>a+o.verifyHours,0);
  const totalDemand = totalDesign + totalVerify;
  const longestUnit = Math.max(0, ...orders.map(o => o.total));   // a single order's design+check chain

  // what-if (optimal)
  const { base, levers } = improvementAnalysis(groups, workers, longestUnit);
  const minHours = base.hours, binding = base.binding, infeasible = base.infeasible;
  const bottleneckLabel = describeBinding(binding, workers, infeasible);
  const feasibleHours = minHours === null ? feasibleMinMakespan(groups, workers, longestUnit, infeasible) : null;
  // greedy schedule is O(days×workers×tasks) — skip it on large data (only the retrospective uses it)
  const sim = (orders.length <= 800 && workers.length <= 200)
    ? simulate(orders, workers)
    : { orderSequence: orders.map(o=>o.id), completion:{}, days: minHours==null?0:Math.ceil(minHours/HOURS_PER_DAY),
        daily:[], totalCapacity: workers.length*HOURS_PER_DAY, allocations:[], stuck:[], unfinishedHours:0, skipped:true };

  let sensitivity = null;
  if (minHours !== null && binding.length){
    const bindCats = new Set(binding.map(x=>x.cat));
    const grade = bindCats.has('High') ? 'A' : (bindCats.has('Low') ? 'B' : 'C');
    const first = binding.slice().sort((a,b)=>(a.prod<b.prod?-1:a.prod>b.prod?1:0)||(a.role<b.role?-1:a.role>b.role?1:0))[0];
    sensitivity = { product: first.prod, role: first.role, grade, rows: hireSensitivity(groups, workers, longestUnit, first.prod, first.role, grade, 6) };
  }

  // as-assigned (fixed workers); workdays fixed = Monday–Friday
  const wpw = WORKDAYS_PER_WEEK;
  const asAssigned = runAsAssigned(orders, workerByName, HOURS_PER_DAY, wpw);

  Object.assign(result, {
    groups, demandByCat, totalDemand, totalDesign, totalVerify, minHours, binding, infeasible, bottleneckLabel,
    feasibleHours, levers, bestLever: bestLever(levers), longestUnit,
    sim, orderInfo: summarizeOrders(orders), orders, workers, partMap,
    capacity: capacityTable(productNames, groups, workers),
    unblock: infeasible.length ? unblockSuggestions(infeasible, workers) : [],
    sensitivity, asAssigned,
    formatDuration, bindingEnd: bindingEnd(binding),
  });
  return result;
}

/* ====================================================================
   RETROSPECTIVE (uses created date as submit, actual finish as finish)
   ==================================================================== */
function runRetrospective(state, opts){
  const wpw = WORKDAYS_PER_WEEK;
  const H = HOURS_PER_DAY;
  // The retrospective only looks at COMPLETED orders (actual finish date filled in): range = earliest created date → today.
  // Unfinished orders belong to the Schedule and are excluded from every retrospective stat (baseline hours, worker load, ④ comparison).
  const finishedOrders = (state.orders || []).filter(o => String(o.actualFinish || '').trim());
  const R = runFullAnalysis({ ...state, orders: finishedOrders });
  const out = { R, hasData: R.hasData, H, wpw,
    nAll: (state.orders || []).length, nFinished: finishedOrders.length,
    nUnfinished: (state.orders || []).length - finishedOrders.length };
  if (!R.hasData) return out;

  // benchmark = the AS-ASSIGNED schedule (given designer/checker), which scales & needs no greedy sim
  const A = R.asAssigned;
  const expById = new Map(A.rows.map(r => [r.id, Math.max(1, Math.ceil(r.finishWD))]));   // expected completion (work-day)
  const orders = R.orders.map(o => ({ id:o.id, so:o.so, product:o.product, partsTxt:o.vcs.join(', '),
    submit: parseISO(o.created), finish: parseISO(o.actualFinish), deadline: parseISO(o.deadline),
    simDay: expById.has(o.id) ? expById.get(o.id) : null, hours: o.total }));
  out.orders = orders; out.nOrders = orders.length;
  const withFinish = orders.filter(o => o.finish); out.nWithFinish = withFinish.length;

  out.infeasible = R.minHours === null;
  out.optimumDays = out.infeasible ? null : R.minHours / H;   // what-if theoretical lower bound (reference)
  out.expectedDays = A.makespanDays;                          // expected duration as assigned
  out.simDays = A.makespanDays; out.totalDemand = R.totalDemand; out.nWorkers = R.nWorkers;

  // available hours per person = 260 workdays per year − annual leave, × hours per person per day (denominator is "one year", not the busiest worker's span)
  const AL = (state.settings && state.settings.annualLeave >= 0) ? +state.settings.annualLeave : 15;
  const cap = (260 - AL) * H;
  out.capYear = cap; out.annualLeave = AL;
  const loadByName = new Map(A.workerLoads.map(w => [w.name, w]));
  // prefer actual logged hours from the hour log (order design + order check); fall back to as-assigned standard hours if no log imported
  const HLw = new Map(((state.hourlog && state.hourlog.workers) || []).map(w =>
    [w.name, (w.byTask['Order design'] || 0) + (w.byTask['Order check'] || 0)]));
  out.loadsAreActual = HLw.size > 0;
  out.workerLoads = (R.workers || []).map(w => {
    const L = loadByName.get(w.name);
    const inHL = HLw.has(w.name);                       // is this name in the hour log? (for reconciliation colors)
    const hrs = out.loadsAreActual ? (HLw.get(w.name) || 0) : (L ? L.total : 0);
    const gap = (R.binding || []).find(g => !workerCan(w, g.prod, g.cat, g.role));
    return { name: w.name, hours: hrs, idle: Math.max(0, cap - hrs), util: cap > 0 ? hrs / cap : 0, inHL,
      trainTarget: (cap > 0 && hrs / cap < 0.6 && gap) ? { product: gap.prod, cat: gap.cat, role: gap.role, need: neededGrade(gap.cat) } : null };
  }).sort((a, b) => a.util - b.util);

  out.trainingLevers = (R.levers || []).filter(l => l.kind === 'train').sort((a, b) => (b.saved - a.saved) || (a.cost - b.cost));
  out.bestTraining = out.trainingLevers[0] || null;
  out.bottleneckLabel = R.bottleneckLabel; out.unblock = R.unblock;
  out.achievableUtil = cap > 0 ? R.totalDemand / (R.nWorkers * cap) : 0;
  if (!withFinish.length) return out;

  const submits = orders.map(o => o.submit).filter(Boolean);
  const finishes = withFinish.map(o => o.finish);
  const start = new Date(Math.min(...(submits.length ? submits : finishes)));
  out.start = start;
  const calToWork = cal => Math.max(1, Math.round(cal * wpw / 7));
  withFinish.forEach(o => { o.leadCalDays = diffDays(o.finish, o.submit || start); o.actualWorkDay = calToWork(Math.max(0, diffDays(o.finish, start)) + 1); o.slack = (o.simDay != null) ? (o.actualWorkDay - o.simDay) : null;
    // per order, on its own calendar (not affected by other orders / a global anchor)
    o.turnaroundCal = o.submit ? Math.max(0, diffDays(o.finish, o.submit)) : null;   // actual duration = actual finish − created
    o.lateCal = o.deadline ? diffDays(o.finish, o.deadline) : null;                  // lateness = actual finish − due (late only if > 0)
    // theory vs actual (per order, in workdays): theory = this order's design+check done back to back; actual = Mon–Fri days between created → finished
    o.theoWD = o.hours / H;
    o.turnWD = o.submit ? businessDaysBetween(o.submit, o.finish) : null;
  });
  out.lastFinish = new Date(Math.max(...finishes)); out.spanCalDays = diffDays(out.lastFinish, start) + 1;
  out.actualMakespanDays = Math.max(...withFinish.map(o => o.actualWorkDay));
  out.reschedSaveDays = Math.max(0, out.actualMakespanDays - A.makespanDays);
  out.availableActual = R.nWorkers * H * out.actualMakespanDays;
  out.utilActual = out.availableActual > 0 ? R.totalDemand / out.availableActual : 0;
  out.idleActual = Math.max(0, out.availableActual - R.totalDemand);
  out.lateOrders = withFinish.filter(o => o.slack != null && o.slack > 0).sort((a, b) => b.slack - a.slack);
  // truly late orders (actual finish after due date), sorted by days late desc — used by the retrospective table
  out.lateVsDeadline = withFinish.filter(o => o.lateCal != null && o.lateCal > 0)
    .sort((a, b) => (b.lateCal - a.lateCal) || (a.id < b.id ? -1 : 1));
  out.nWithDeadline = withFinish.filter(o => o.deadline).length;
  // real delivery metrics (each order on its own calendar days, no global anchor): on-time rate / late count / median turnaround / longest delay
  const median = arr => { if (!arr.length) return null; const s = arr.slice().sort((a,b)=>a-b); const m = s.length >> 1;
    return s.length % 2 ? s[m] : Math.round((s[m-1] + s[m]) / 2); };
  out.nLate = out.lateVsDeadline.length;
  out.nOnTime = Math.max(0, out.nWithDeadline - out.nLate);
  out.onTimeRate = out.nWithDeadline > 0 ? out.nOnTime / out.nWithDeadline : null;
  out.medianTurnaround = median(withFinish.map(o => o.turnaroundCal).filter(v => v != null));
  out.maxLate = out.lateVsDeadline.length ? out.lateVsDeadline[0].lateCal : 0;
  out.medianLate = median(out.lateVsDeadline.map(o => o.lateCal));
  // completion speed (per order, workdays). Elapsed times can't be summed (orders run in parallel, the same day counts for every open order),
  // so express it intuitively: "X% of orders finish in ≤ N workdays" + a bucket distribution; only pure hours (total workload) can be summed.
  const withBoth = withFinish.filter(o => o.turnWD != null);
  out.nTheo = withBoth.length;
  out.theoSumWD = withBoth.reduce((a, o) => a + o.theoWD, 0);                        // total workload (person·workdays)
  out.avgTheoHours = out.nTheo ? withBoth.reduce((a, o) => a + o.hours, 0) / out.nTheo : null;  // average work per order (hours)
  const turns = withBoth.map(o => o.turnWD).sort((a, b) => a - b);
  const pctile = q => turns.length ? turns[Math.min(turns.length - 1, Math.max(0, Math.ceil(q * turns.length) - 1))] : null;
  out.turnP50 = pctile(0.5); out.turnP90 = pctile(0.9); out.turnP95 = pctile(0.95);
  out.turnMax = turns.length ? turns[turns.length - 1] : null;
  const BUCKETS = [[0,0,'Same day'],[1,5,'1–5 days (~1 week)'],[6,10,'6–10 days (~2 weeks)'],
    [11,21,'11–21 days (~1 month)'],[22,65,'22–65 days (~1 quarter)'],[66,Infinity,'Over 65 days (> 1 quarter)']];
  let cum = 0;
  out.turnDist = BUCKETS.map(([lo, hi, label]) => { const n = turns.filter(t => t >= lo && t <= hi).length; cum += n;
    return { label, lo, hi, n, share: out.nTheo ? n / out.nTheo : 0, cumShare: out.nTheo ? cum / out.nTheo : 0 }; });
  return out;
}

/* ====================================================================
   SCHEDULE (deadline = due date / required)
   ==================================================================== */
function isWorkDay(d){ const dow = d.getDay(); return dow !== 0 && dow !== 6; }   // Monday–Friday
function runSchedule(state, opts){
  const wpw = WORKDAYS_PER_WEEK;
  const H = HOURS_PER_DAY;
  const R = runFullAnalysis(state);
  const out = { R, hasData: R.hasData, H, wpw };
  if (!R.hasData) return out;
  const nWorkers = R.nWorkers, C = nWorkers * H; out.nWorkers = nWorkers; out.capacity = C;
  const start = parseISO(opts && opts.startDate) || new Date(); start.setHours(0,0,0,0); out.start = start;

  // The schedule only covers UNFINISHED orders (no actual finish date), scheduled forward from the start day (default today);
  // completed work belongs to the Retrospective and takes no part in scheduling or capacity.
  const openOrders = R.orders.filter(o => !o.actualFinish);
  out.nFinished = R.orders.length - openOrders.length;
  const orders = openOrders.map((o, i) => ({ id:o.id, so:o.so, seq:i, product:o.product, hours:o.total,
    designH:o.designHours, verifyH:o.verifyHours, difficulty:o.difficulty, partsTxt:o.vcs.join(', '),
    designer:o.designer || '', checker:o.checker || '', deadline: parseISO(o.deadline), created: parseISO(o.created) }));
  out.noDeadline = orders.filter(o => !o.deadline).map(o => o.id);
  const schedulable = orders.filter(o => o.hours > 1e-9);          // ①② sequential/slot schedule every unfinished order with hours (including ones without a due date)
  const withDl = schedulable.filter(o => o.deadline); out.scheduledOrders = withDl;
  out.nOpen = schedulable.length;
  const ec = orders.filter(o => o.created).sort((a, b) => a.created - b.created)[0];
  out.earliestCreated = ec ? { id: ec.id, date: ec.created } : null;
  if (!schedulable.length){ out.noSchedule = true; return out; }

  const lastDl = withDl.length ? new Date(Math.max(...withDl.map(o => o.deadline.getTime()))) : new Date(start);
  // who sets the last due date + typo detection: a due date more than 1 year after the latest created date of all orders is most likely a typo / import misalignment
  out.lastDl = lastDl;
  out.lastDlOrders = withDl.filter(o => o.deadline.getTime() === lastDl.getTime()).map(o => o.id);
  const createds = R.orders.map(o => parseISO(o.created)).filter(Boolean);
  const maxCreated = createds.length ? new Date(Math.max(...createds.map(d => d.getTime()))) : null;
  out.maxCreated = maxCreated;
  out.farDeadlines = maxCreated ? withDl.filter(o => diffDays(o.deadline, maxCreated) > 365)
    .sort((a, b) => b.deadline - a.deadline).map(o => ({ id: o.id, deadline: o.deadline })) : [];
  const total = schedulable.reduce((a, o) => a + o.hours, 0); out.total = total;   // all hours (including no due date)
  // annual leave: AL days per person per year with no specific dates, spread over 52×5=260 workdays as a daily capacity discount (schedule only)
  const AL = Math.max(0, Math.min(100, Math.round(
    (opts && opts.annualLeaveDays != null && isFinite(+opts.annualLeaveDays)) ? +opts.annualLeaveDays
    : (state.settings && state.settings.annualLeave != null && isFinite(+state.settings.annualLeave)) ? +state.settings.annualLeave : 15)));
  const leaveFactor = Math.max(0, (260 - AL) / 260);
  const baseCap = C * leaveFactor;   // everyday team capacity after annual leave
  out.annualLeave = AL; out.leaveFactor = leaveFactor; out.baseCap = baseCap;
  // leave records: subtract the leave hours that fall on each workday from that day's capacity (exact, dated deduction)
  // each record hpd = leave hours per day (blank/invalid = full shift H, at most one full shift)
  const leaves = (state.leaves || []).map(L => {
    const s = parseISO(L && L.start); let e = parseISO(L && L.end) || s;
    const v = L && L.hours; const hpd = (v !== '' && v != null && isFinite(+v) && +v > 0) ? Math.min(H, +v) : H;
    return { w: String((L && L.worker) || '').trim(), s, e, hpd }; }).filter(L => L.s);
  leaves.forEach(L => { if (L.e.getTime() < L.s.getTime()){ const t = L.s; L.s = L.e; L.e = t; } });
  const leavesByWorker = new Map();
  leaves.forEach(L => { if (!L.w) return; if (!leavesByWorker.has(L.w)) leavesByWorker.set(L.w, []); leavesByWorker.get(L.w).push(L); });
  const dayCap = d => { let offH = 0; const t = d.getTime();
    for (const L of leaves) if (t >= L.s.getTime() && t <= L.e.getTime()) offH += L.hpd;
    return Math.max(0, baseCap - Math.min(baseCap, offH)); };
  // timeline: extend to the last due date, and at least until all hours can be done (sequential may run past the last due date), capped at 1500 workdays.
  // in the per-person simulation the busiest worker doing their own queue serially may be slower than the whole team — extend the timeline far enough for them too.
  const loadPreview = new Map();
  schedulable.forEach(o => {
    if (o.designer && o.designH > 1e-9) loadPreview.set(o.designer, (loadPreview.get(o.designer) || 0) + o.designH);
    if (o.checker && o.verifyH > 1e-9) loadPreview.set(o.checker, (loadPreview.get(o.checker) || 0) + o.verifyH);
  });
  const maxPersonal = Math.max(0, ...loadPreview.values());
  const perWorkerDay = H * leaveFactor;
  const needPersonalDays = perWorkerDay > 1e-9 ? Math.min(1500, Math.ceil(maxPersonal / perWorkerDay) + 40) : 0;
  const days = []; const cap = []; const cur = new Date(start); let capSum = 0;
  while (days.length < 1500 && (cur.getTime() <= lastDl.getTime() || capSum < total - 1e-9 || days.length < needPersonalDays)){
    if (isWorkDay(cur)){ const c = dayCap(cur); days.push(new Date(cur)); cap.push(c); capSum += c; }
    cur.setDate(cur.getDate() + 1);
    if (baseCap <= 1e-9 && cur.getTime() > lastDl.getTime()) break;   // avoid spinning when capacity is 0
  }
  if (!days.length){ days.push(new Date(start)); cap.push(dayCap(start)); }
  const N = days.length; out.days = days; out.N = N;
  const capCum = []; { let a = 0; for (let i = 0; i < N; i++){ a += cap[i]; capCum.push(a); } }
  out.cap = cap; out.capCum = capCum;
  out.leaveHoursLost = cap.reduce((s, c) => s + (baseCap - c), 0);
  out.leaveDaysAffected = cap.filter(c => c < baseCap - 1e-9).length;
  const dlIndex = d => { let c = 0; for (let i = 0; i < N; i++){ if (days[i].getTime() <= d.getTime()) c++; else break; } return c; };
  withDl.forEach(o => o.Tg = dlIndex(o.deadline));
  out.overdue = withDl.filter(o => o.deadline.getTime() < start.getTime()).map(o => o.id);   // due before start day = already overdue

  // ①② per-order queue: consume capacity in sequence to get each order's expected start/finish day
  const seqSchedule = seq => {
    const rows = []; let cum = 0, p = 0;
    for (const o of seq){
      const sH = cum, fH = cum + o.hours; cum = fH;
      while (p < N && capCum[p] <= sH + 1e-9) p++;
      const startIdx = p < N ? p : null;
      while (p < N && capCum[p] < fH - 1e-9) p++;
      const finishIdx = p < N ? p : null;
      const dlT = o.deadline ? dlIndex(o.deadline) : null;                       // which workday the due date is (1-based)
      const lateWD = (finishIdx != null && dlT != null) ? (finishIdx + 1) - dlT : null;   // >0 late, <0 early, 0 on the day
      rows.push({ id:o.id, so:o.so, product:o.product, hours:o.hours, designH:o.designH, verifyH:o.verifyH,
        created:o.created, deadline:o.deadline,
        designer:o.designer, checker:o.checker, startIdx, finishIdx, dlT, lateWD });
    }
    return rows;
  };
  const byCreated = (a, b) => ((a.created ? a.created.getTime() : Infinity) - (b.created ? b.created.getTime() : Infinity)) || (a.seq - b.seq);
  out.seqCont = seqSchedule(schedulable.slice().sort(byCreated));                                  // ① sequential: by created date
  out.seqSlot = seqSchedule(schedulable.slice().sort((a, b) =>                                     // ② by due date: earliest due first, no due date last
    ((a.deadline ? a.deadline.getTime() : Infinity) - (b.deadline ? b.deadline.getTime() : Infinity)) || byCreated(a, b)));
  out.slotLate = out.seqSlot.filter(r => r.lateWD != null && r.lateWD > 0);
  out.contLate = out.seqCont.filter(r => r.lateWD != null && r.lateWD > 0);
  const byT = new Map();
  withDl.forEach(o => { const e = byT.get(o.Tg) || { T:o.Tg, hours:0, count:0, date:o.deadline }; e.hours += o.hours; e.count++; if (o.deadline < e.date) e.date = o.deadline; byT.set(o.Tg, e); });
  const dls = [...byT.values()].sort((a, b) => a.T - b.T);
  let cum = 0; dls.forEach(d => { cum += d.hours; d.cum = cum; d.feasible = d.T > 0 && d.cum <= capCum[Math.min(N, d.T) - 1] + 1e-6; });
  out.deadlines = dls;
  const eager = new Array(N).fill(0); let rem = total;
  for (let i = 0; i < N && rem > 1e-9; i++){ eager[i] = Math.min(cap[i], rem); rem -= eager[i]; }
  let eagerFinish = 0, acc = 0; for (let i = 0; i < N; i++){ acc += eager[i]; if (acc >= total - 1e-9){ eagerFinish = i + 1; break; } }
  out.eager = eager; out.eagerFinishDay = eagerFinish || N; out.eagerIdleDays = Math.max(0, N - (eagerFinish || N));
  // close to due date (backward / JIT): work back from each due date, first trying to finish `buffer` workdays before it;
  // if capacity is short, borrow earlier days; only then use the buffer (finish right at the due date but not late); still short = truly overloaded (red).
  // due dates are processed latest first: later work claims the later days, so naturally "orders due later aren't rushed early".
  const buffer = Math.max(0, Math.min(60, Math.round(
    (opts && opts.bufferDays != null && isFinite(+opts.bufferDays)) ? +opts.bufferDays
    : (state.settings && state.settings.schedBuffer != null && isFinite(+state.settings.schedBuffer)) ? +state.settings.schedBuffer : 2)));
  out.buffer = buffer;
  const jit = new Array(N).fill(0); const remCap = cap.slice();
  const overloadDeadlines = []; const bufferShort = [];
  dls.slice().sort((a, b) => b.T - a.T).forEach(d => {
    let need = d.hours; let lastUsed = -1;
    const end = Math.min(N - 1, d.T - 1);            // last usable workday without being late
    const idealEnd = end - buffer;                   // target finish day (buffer workdays before due)
    for (let i = idealEnd; i >= 0 && need > 1e-9; i--){ const t = Math.min(remCap[i], need); if (t > 0){ jit[i] += t; remCap[i] -= t; need -= t; if (i > lastUsed) lastUsed = i; } }
    if (need > 1e-9){                                // can't finish early → use the buffer (still not late)
      bufferShort.push(d);
      for (let i = Math.max(0, idealEnd + 1); i <= end && need > 1e-9; i++){ const t = Math.min(remCap[i], need); if (t > 0){ jit[i] += t; remCap[i] -= t; need -= t; if (i > lastUsed) lastUsed = i; } }
    }
    if (need > 1e-9){ overloadDeadlines.push(d); jit[Math.max(0, end)] += need; lastUsed = Math.max(0, end); }   // truly overloaded
    d.jitFinishIdx = lastUsed < 0 ? Math.max(0, Math.min(idealEnd, end)) : lastUsed;   // last day this batch is scheduled on
    d.jitLeadWD = Math.max(0, (d.T - 1) - d.jitFinishIdx);                             // how many workdays ahead of the due date
  });
  out.jit = jit; out.overloadDeadlines = overloadDeadlines; out.feasibleAll = dls.every(d => d.feasible);
  out.bufferShort = bufferShort.filter(d => !overloadDeadlines.includes(d));
  out.jitPeak = Math.max(0, ...jit); out.eagerPeak = Math.max(0, ...eager);
  out.eagerCum = []; out.jitCum = []; let a2 = 0, b2 = 0;
  for (let i = 0; i < N; i++){ a2 += eager[i]; b2 += jit[i]; out.eagerCum.push(a2); out.jitCum.push(b2); }
  // earliest day each due date can be delivered under the "finish ASAP" reference schedule (first day eagerCum covers the cumulative amount due)
  { let k = 0;
    dls.forEach(d => { while (k < N - 1 && out.eagerCum[k] < d.cum - 1e-6) k++; d.eagerFinishIdx = k; }); }

  // ---- ④ per-person simulation: design hours go on the designer's own timeline, check goes on the checker's timeline after that order's design is done;
  //      leave is deducted exactly per person, annual leave spread evenly. Three modes:
  //      cont=① sequential (forward by created order), slot=② by due date (earliest due first, forward),
  //      jit =③ close to due date (backward from due: check finishes `buffer` days before due, design finishes before check starts; only orders with a due date).
  //      In forward modes, if the same person both designs and checks, they finish their own design queue before checking (an approximation, see page note). ----
  const baseW = H * leaveFactor;                                   // available hours per person per day (with annual leave spread)
  const dayCapW = (name, i) => { let c = baseW; const ls = leavesByWorker.get(name);
    if (ls){ const t = days[i].getTime(); for (const L of ls) if (t >= L.s.getTime() && t <= L.e.getTime()) c -= L.hpd; }
    return Math.max(0, c); };
  out.dayCapW = dayCapW;
  out.pwUnassigned = schedulable.filter(o => (!o.designer && o.designH > 1e-9) || (!o.checker && o.verifyH > 1e-9)).length;
  const dLoad = new Map(), vLoad = new Map(), dCnt = new Map(), vCnt = new Map();
  schedulable.forEach(o => {
    if (o.designer && o.designH > 1e-9){ dLoad.set(o.designer, (dLoad.get(o.designer) || 0) + o.designH); dCnt.set(o.designer, (dCnt.get(o.designer) || 0) + 1); }
    if (o.checker && o.verifyH > 1e-9){ vLoad.set(o.checker, (vLoad.get(o.checker) || 0) + o.verifyH); vCnt.set(o.checker, (vCnt.get(o.checker) || 0) + 1); }
  });
  const pwSummary = (pwDay, unplacedBy) => [...pwDay.keys()].map(name => {
    const arr = pwDay.get(name); let totalH = 0, busy = 0, last = -1, first = -1;
    for (let i = 0; i < N; i++){ if (arr[i] > 1e-9){ totalH += arr[i]; busy++; last = i; if (first < 0) first = i; } }
    return { name, total: totalH, busyDays: busy, lastIdx: last, firstIdx: first, avg: busy ? totalH / busy : 0,
      dH: dLoad.get(name) || 0, vH: vLoad.get(name) || 0, nD: dCnt.get(name) || 0, nV: vCnt.get(name) || 0,
      unplaced: unplacedBy.get(name) || 0 };
  }).sort((a, b) => b.total - a.total);
  // forward (cont / slot): put design on each timeline in the given order, then check in order of "this order's design finish day"
  const simForward = seq => {
    const pwDay = new Map(), pwOrders = new Map(), pwCursor = new Map(), unplacedBy = new Map();
    const ensure = name => { if (!pwDay.has(name)){ pwDay.set(name, new Float64Array(N)); pwOrders.set(name, {}); pwCursor.set(name, { d: 0 }); } };
    const alloc = (name, hours, minDay, label) => {
      ensure(name);
      const arr = pwDay.get(name), curP = pwCursor.get(name), lab = pwOrders.get(name);
      let d = Math.max(curP.d, minDay), rem = hours, first = null;
      while (rem > 1e-9 && d < N){
        const free = dayCapW(name, d) - arr[d];
        if (free > 1e-9){ const take = Math.min(free, rem); arr[d] += take; rem -= take;
          if (first == null) first = d;
          (lab[d] || (lab[d] = [])).push(label);
          if (rem <= 1e-9) break; }
        d++;
      }
      curP.d = Math.max(curP.d, d);
      if (rem > 1e-9) unplacedBy.set(name, (unplacedBy.get(name) || 0) + rem);
      return { start: first, finish: rem > 1e-9 ? null : d, unplaced: rem };
    };
    const designFin = new Map();
    seq.forEach(r => {
      if (r.designer && r.designH > 1e-9){ const a = alloc(r.designer, r.designH, 0, r.so || r.id); designFin.set(r.id, a.finish == null ? Infinity : a.finish); }
      else designFin.set(r.id, 0);
    });
    seq.filter(r => r.checker && r.verifyH > 1e-9)
      .slice().sort((a, b) => (designFin.get(a.id) - designFin.get(b.id)) || ((a.dlT != null ? a.dlT : 1e9) - (b.dlT != null ? b.dlT : 1e9)))
      .forEach(r => { const dv = designFin.get(r.id); alloc(r.checker, r.verifyH, isFinite(dv) ? dv : N, (r.so || r.id) + '·chk'); });
    return { day: pwDay, orders: pwOrders, workers: pwSummary(pwDay, unplacedBy), skipped: 0 };
  };
  // backward (jit): later due dates claim later days first; check searches backward from "buffer days before due", design finishes before the day check starts.
  // hours that don't fit before the due date pile onto the last usable day (that day exceeds capacity, shown red = this person can't make the due date).
  const simJit = () => {
    const pwDay = new Map(), pwOrders = new Map(), unplacedBy = new Map();
    const ensure = name => { if (!pwDay.has(name)){ pwDay.set(name, new Float64Array(N)); pwOrders.set(name, {}); } };
    const allocB = (name, hours, idealEnd, hardEnd, label) => {
      ensure(name);
      const arr = pwDay.get(name), lab = pwOrders.get(name);
      let rem = hours, first = null;
      for (let d = Math.min(idealEnd, N - 1); d >= 0 && rem > 1e-9; d--){
        const free = dayCapW(name, d) - arr[d];
        if (free > 1e-9){ const take = Math.min(free, rem); arr[d] += take; rem -= take;
          first = d; (lab[d] || (lab[d] = [])).push(label); }
      }
      for (let d = Math.max(0, idealEnd + 1); d <= Math.min(hardEnd, N - 1) && rem > 1e-9; d++){   // use the buffer (still not late)
        const free = dayCapW(name, d) - arr[d];
        if (free > 1e-9){ const take = Math.min(free, rem); arr[d] += take; rem -= take;
          if (first == null) first = d;
          (lab[d] || (lab[d] = [])).push(label); }
      }
      if (rem > 1e-9){                                             // truly overloaded: can't fit before the due date
        const d = Math.max(0, Math.min(hardEnd, N - 1));
        arr[d] += rem; (lab[d] || (lab[d] = [])).push(label);
        if (first == null) first = d;
        unplacedBy.set(name, (unplacedBy.get(name) || 0) + rem);
      }
      return { start: first };
    };
    const withDlRows = out.seqSlot.filter(r => r.dlT != null).slice().sort((a, b) => b.dlT - a.dlT);
    withDlRows.forEach(r => {
      const end = Math.max(0, Math.min(N - 1, r.dlT - 1));         // due day = last day without being late
      const ideal = Math.max(0, end - buffer);
      let vStart = ideal;
      if (r.checker && r.verifyH > 1e-9){ const a = allocB(r.checker, r.verifyH, ideal, end, (r.so || r.id) + '·chk'); vStart = a.start == null ? ideal : a.start; }
      if (r.designer && r.designH > 1e-9) allocB(r.designer, r.designH, vStart, vStart, r.so || r.id);
    });
    return { day: pwDay, orders: pwOrders, workers: pwSummary(pwDay, unplacedBy), skipped: out.seqSlot.length - withDlRows.length };
  };
  out.pw = { cont: simForward(out.seqCont), slot: simForward(out.seqSlot), jit: simJit() };

  // ---- ⑤ reassignment suggestions: move overloaded workers' work to less busy people (suggestions only, no data is changed).
  //      With ability grades: only move to someone whose grade covers "this product · this role · this difficulty";
  //      if no graded candidate can do it, fall back to "whoever is free" — the least busy person with blank grades (flagged "ability not filled", confirm manually first);
  //      people whose grade explicitly says they can't do it are never suggested. Design only goes to designers, check only to checkers — never mixed. ----
  const wByName = new Map((R.workers || []).map(w => [w.name, w]));
  const loads = new Map(loadPreview);
  (R.workers || []).forEach(w => { if (!loads.has(w.name)) loads.set(w.name, 0); });
  const load0 = new Map(loads);                            // original load before adjustment (for display)
  const pool = [...loads.keys()];
  const totalLoad = [...loads.values()].reduce((a, b) => a + b, 0);
  const avgLoad = pool.length ? totalLoad / pool.length : 0;
  const canDo = (name, o, role) => {                       // 'ok' | 'unknown' (ability not filled) | null (grade can't do it / not on roster)
    const w = wByName.get(name); if (!w) return null;
    const g = roleGrade(w, o.product, role);
    if (!g) return 'unknown';
    return (ABILITY_SKILLS[g] && ABILITY_SKILLS[g].includes(o.difficulty)) ? 'ok' : null;
  };
  const tasksOf = name => { const list = [];
    schedulable.forEach(o => {
      if (o.designer === name && o.designH > 1e-9) list.push({ o, role:'d', h:o.designH });
      if (o.checker === name && o.verifyH > 1e-9) list.push({ o, role:'v', h:o.verifyH });
    }); return list.sort((a, b) => b.h - a.h); };
  const moves = []; const usedTask = new Set();            // each role of an order is moved at most once
  if (avgLoad > 0){
    const overs = pool.filter(n => loads.get(n) > avgLoad * 1.15).sort((a, b) => loads.get(b) - loads.get(a));
    for (const from of overs){
      let lastTo = null;                                   // bundling: keep giving small orders to the same person to form a solid block (instead of an hour to everyone)
      for (const t of tasksOf(from)){                      // start moving from the biggest task
        if (loads.get(from) <= avgLoad * 1.05 + 1e-9) break;      // close to average, enough
        const key = t.o.id + '¦' + t.role;
        if (usedTask.has(key) || t.h < 0.02) continue;            // only skip near-zero crumbs (many orders are a fraction of an hour each; don't skip them for being small)
        if (loads.get(from) - t.h < avgLoad * 0.8) continue;      // moving the whole order would overcorrect → try a smaller one
        // receiver cap: generally not above average; when the order itself is bigger than average, relax to "average + half an order", otherwise big orders never move
        const recvCap = Math.max(avgLoad * 1.05, avgLoad + t.h / 2);
        const fits = n => n !== from && loads.get(n) + t.h <= recvCap + 1e-9;
        const st = n => canDo(n, t.o, t.role);
        let to = null, sure = false;
        if (lastTo && fits(lastTo) && st(lastTo) === 'ok'){ to = lastTo; sure = true; }   // previous receiver whose grade qualifies → keep giving to them
        else {
          const rated = pool.filter(n => fits(n) && st(n) === 'ok');                      // graded people by grade
          if (rated.length){ to = rated[0]; for (const n of rated) if (loads.get(n) < loads.get(to)) to = n; sure = true; }
          else if (lastTo && fits(lastTo) && st(lastTo) === 'unknown'){ to = lastTo; }    // otherwise whoever is free (grade not filled)
          else {
            const blank = pool.filter(n => fits(n) && st(n) === 'unknown');
            if (blank.length){ to = blank[0]; for (const n of blank) if (loads.get(n) < loads.get(to)) to = n; }
          }
        }
        if (!to) continue;                                        // everyone else's grade explicitly can't → don't move this order
        lastTo = to;
        usedTask.add(key);
        moves.push({ from, to, id: t.o.so || t.o.id, role: t.role, h: t.h, sure });
        loads.set(from, loads.get(from) - t.h);
        loads.set(to, loads.get(to) + t.h);
      }
    }
  }
  const byPair = new Map();                                // aggregate into "from A → to B" suggestion rows (one row may hold many small orders)
  moves.forEach(m => { const k = m.from + '¦' + m.to;
    if (!byPair.has(k)) byPair.set(k, { from: m.from, to: m.to, fromLoad: load0.get(m.from) || 0, toLoad: load0.get(m.to) || 0, amount: 0, examples: [] });
    const p = byPair.get(k); p.amount += m.h; p.examples.push({ id: m.id, role: m.role, h: m.h, sure: m.sure }); });
  const sugg = [...byPair.values()].sort((a, b) => b.amount - a.amount)
    .filter(p => p.amount >= Math.max(0.5, avgLoad * 0.01));      // don't show crumb rows totalling under ~half an hour
  out.rebalance = { avgLoad, suggestions: sugg,
    maxBefore: Math.max(0, ...load0.values()),
    maxAfter: Math.max(0, ...loads.values()) };
  return out;
}

window.runFullAnalysis = runFullAnalysis;
window.runRetrospective = runRetrospective;
window.runSchedule = runSchedule;
window.fmtDuration = formatDuration;
