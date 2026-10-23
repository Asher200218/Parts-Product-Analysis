/* analysis.js — in-browser engine (VC-time model, role-aware).
   • Each part = a VC (3-digit code) with a 难度 (高/低/无, for worker eligibility),
     重复设计工时 t7, 首单设计工时 t8.
   • Per (product, VC): the chronologically earliest order (by 创建日期) uses t8
     (首单); later ones use t7 (重复). t8 blank → use t7. Any value 0 → 0.03.
   • Order 设计 = (Σ its VC times) × coeff(VC数: ≤5→1, 6–10→1.2, >10→1.5).
     Order 校对 = 设计 × VERIFY_RATIO (默认 0.3，可由工时台账实际值校准).   (precedence: 设计 → 校对)
   • Two views: ① 按指派 (fixed 设计/校对 worker per order)  ② 什么-if (optimal).
   Relies on app.js globals: CATEGORIES, ABILITY_SKILLS, GRADE_RANK, HOURS_PER_DAY,
   ROLES, ROLE_LABEL, clean, normGrade, normCategory, normDate, parseOrderParts. */
'use strict';

const MAX_EXACT_GROUPS = 22;
const MAX_SIM_DAYS = 3650;
const GRADE_RANK = { A: 3, B: 2, C: 1 };
const DIFF_RANK = { '高': 3, '低': 2, '无': 1 };
const SPECIAL_VC = { '999': 0.08 };   // 特殊 VC 码 -> 固定设计工时 (小时)
const WORKDAYS_PER_WEEK = 5;          // 工作日固定 = 周一至周五（周末不计工时/排程）
// 设计/校对工人为这些占位值 -> 整张订单作废 (不纳入任何分析)
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
function neededGrade(cat){ return cat === '高' ? 'A' : (cat === '低' ? 'B' : 'C'); }
function catIndex(c){ return CATEGORIES.indexOf(c); }
function gkey(g){ return g.prod + '¦' + g.cat + '¦' + g.role; }
function numOrNull(v){ if (v == null || String(v).trim() === '') return null; const n = Number(v); return isNaN(n) ? null : n; }
function parseISO(s){ if (!s) return null; const m = String(s).match(/(\d{4})-(\d{1,2})-(\d{1,2})/); return m ? new Date(+m[1], +m[2]-1, +m[3]) : null; }
const DAY_MS = 86400000;
function diffDays(a, b){ return Math.round((a - b) / DAY_MS); }
// a→b 之间经过的工作日数（只数周一至周五；同一天=0）
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

  // workers (设计/校验 grades). Keep workers even with NO grades filled in — the ① 按指派
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
    // 设计/校对工人为 special order / order canceled / null -> 作废整张订单
    if (isSkipWorker(designer) || isSkipWorker(checker)){ skipped++; return; }
    const vcs = parseOrderParts(o.parts).map(x => x.name).filter(Boolean);
    if (!product && !vcs.length) return;
    // any VC not yet known -> add as a blank VC (will get the 0.03 fallback)
    vcs.forEach(c => { if (!partMap.has(c)) partMap.set(c, { category:'', t7:null, t8:null }); });
    orders.push({ id:`订单${seq}`, srcIndex: idx, product, vcs, so: clean(o.so),
      created: clean(o.time), deadline: clean(o.required),
      designer, checker, actualFinish: clean(o.actualFinish) });
  });
  if (skipped) warnings.push(`已作废 ${skipped} 张订单 (设计/校对工人为 special order / order canceled / null)`);

  computeOrderTimes(orders, partMap);

  return { productNames, partMap, nPartsTotal: partMap.size, workers, workerByName, orders, errors, warnings };
}

/* ---------- VC time computation (first-time vs repeat, coefficient, 校对=设计×VERIFY_RATIO) ---------- */
function computeOrderTimes(orders, partMap){
  // chronological order by 创建日期 (blank dates last, stable)
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
      if (SPECIAL_VC[vc] != null) t = SPECIAL_VC[vc];        // 999 等特殊码: 固定工时
      else {
        t = first ? (p.t8 != null ? p.t8 : p.t7) : p.t7;    // 首单用 t8(无则 t7); 重复用 t7
        if (t == null) t = 0;
        if (t === 0) t = 0.03;                               // 0 → 基本时间 0.03
      }
      base += t; detail.push({ vc, time: t, first });
    });
    const nv = o.vcs.length;
    const coeff = nv <= 5 ? 1 : (nv <= 10 ? 1.2 : 1.5);
    o.nVcs = nv; o.coeff = coeff; o.vcDetail = detail;
    o.designHours = base * coeff;
    o.verifyHours = o.designHours * (typeof VERIFY_RATIO !== 'undefined' ? VERIFY_RATIO : 0.3);   // 校对系数可由工时台账校准
    o.total = o.designHours + o.verifyHours;
    // order difficulty = hardest VC difficulty (blank VC → 无)
    let diff = '无', rank = 0;
    o.vcs.forEach(vc => { const c = (partMap.get(vc) || {}).category; const cc = CATEGORIES.includes(c) ? c : '无'; if (DIFF_RANK[cc] > rank){ rank = DIFF_RANK[cc]; diff = cc; } });
    o.difficulty = diff;
  });
}

/* ---------- what-if groups: order 设计/校对 hours bucketed by (product, 难度, role) ---------- */
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
  if (hours === null) return '不可完成 infeasible';
  const days = hours / HOURS_PER_DAY, whole = Math.floor(days), rem = hours - whole * HOURS_PER_DAY;
  return `${hours.toFixed(1)} 工时/人 ≈ ${days.toFixed(2)} 个工作日 (${whole}天${rem.toFixed(1)}小时)`;
}
function describeBinding(binding, workers, infeasible){
  if (infeasible && infeasible.length)
    return `无法完成: ${infeasible.map(g=>`${g.prod}的${g.cat}难度(${roleName(g.role)})`).join('、')} 没有任何工人能做`;
  if (!binding.length) return '无明显瓶颈';
  const prods = [...new Set(binding.map(b=>b.prod))].sort();
  const cats = [...new Set(binding.map(b=>b.cat))].sort((a,b)=>catIndex(a)-catIndex(b));
  const roles = [...new Set(binding.map(b=>b.role))].map(roleName);
  const elig = workers.filter(w => binding.some(g => workerCan(w,g.prod,g.cat,g.role))).length;
  return `${prods.join('/')} 的 ${cats.join('/')} 难度（${roles.join('/')}）(仅 ${elig} 人可做)`;
}
function bindingEnd(binding){
  if (!binding.length) return '无';
  const cats = new Set(binding.map(b=>b.cat));
  if (cats.has('高')) return '高端(A级)';
  if (cats.size === 1 && cats.has('无')) return '低端(C级也可)';
  return '中端(B级)';
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
function hiredWorker(product, role, grade){ return { name:`(新招 ${product}-${roleName(role)}${grade})`, grades:{ [product]: role==='d'?{d:grade,v:''}:{d:'',v:grade} } }; }
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
    levers.push({ label:`招聘1名 ${pr.prod} 的 ${pr.need}级 ${roleName(pr.role)}工`, kind:'hire', cost: 3.0 + (3 - GRADE_RANK[pr.need]) * 0.1, hours: h, saved: base.hours - h });
  });
  // training: any worker with the same (product, role, currentGrade) gives the identical result,
  // so test only ONE representative per key (keeps this O(pairs×grades), scalable to 万级 workers)
  const seenTrain = new Set();
  for (const w of workers) for (const pr of pairs){
    const cur = roleGrade(w, pr.prod, pr.role); const up = upgradeGrade(cur); if (!up) continue;
    const key = pr.prod + '¦' + pr.role + '¦' + cur; if (seenTrain.has(key)) continue; seenTrain.add(key);
    const nw = workers.map(x => { if (x.name !== w.name) return x;   // reuse refs; deep-copy only this worker
      const g2 = { ...x.grades, [pr.prod]: { ...(x.grades[pr.prod] || { d:'', v:'' }) } }; g2[pr.prod][pr.role] = up; return { name:x.name, grades:g2 }; });
    const h = minMakespan(items, nw, longestUnit).hours;
    if (base.hours - h > 1e-9) levers.push({ label:`培训 ${w.name}: ${pr.prod} ${roleName(pr.role)} ${cur}→${up}`, kind:'train', cost:2.0, hours:h, saved: base.hours - h });
  }
  return { base, levers };
}
function bestLever(levers){ const u = levers.filter(l => l.saved != null && l.saved > 1e-6); return u.length ? u.sort((a,b)=>(b.saved-a.saved)||(a.cost-b.cost))[0] : null; }
function hireSensitivity(items, workers, longestUnit, product, role, grade, maxAdd = 6){
  const out = []; const extra = [];
  for (let k = 0; k <= maxAdd; k++){ const h = minMakespan(items, workers.concat(extra), longestUnit).hours; out.push({ added: k, hours: h, days: h == null ? null : h / HOURS_PER_DAY }); extra.push(hiredWorker(product, role, grade)); }
  return out;
}

/* ---------- what-if greedy schedule (optimal assignment, order-level 设计→校对) ---------- */
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
    day++; const used = { 高:0, 低:0, 无:0 }; let progress = 0;
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
    used.total = used.高 + used.低 + used.无; daily.push(used);
    if (progress <= 1e-9){
      const s = new Map(); tasks.forEach(t => { if (t.remaining > 1e-9){ const g = { prod:t.product, cat:t.cat, role:t.role }; s.set(gkey(g), g); } });
      stuck = [...s.values()]; day--; daily.pop(); break;
    }
  }
  return { orderSequence: orderIds, completion, days: day, daily, totalCapacity, allocations, stuck, unfinishedHours: tasks.reduce((a,t)=>a+t.remaining, 0) };
}

/* ---------- as-assigned view: fixed 设计/校对 per order ---------- */
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
    design:o.designHours, verify:o.verifyHours, hours:{高:0,低:0,无:0}, partsTxt:o.vcs.join('、') };
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
  const demandByCat = { 高:0, 低:0, 无:0 };
  groups.forEach(g => demandByCat[g.cat] += g.hours);
  const totalDesign = orders.reduce((a,o)=>a+o.designHours,0);
  const totalVerify = orders.reduce((a,o)=>a+o.verifyHours,0);
  const totalDemand = totalDesign + totalVerify;
  const longestUnit = Math.max(0, ...orders.map(o => o.total));   // a single order's 设计+校对 chain

  // what-if (optimal)
  const { base, levers } = improvementAnalysis(groups, workers, longestUnit);
  const minHours = base.hours, binding = base.binding, infeasible = base.infeasible;
  const bottleneckLabel = describeBinding(binding, workers, infeasible);
  const feasibleHours = minHours === null ? feasibleMinMakespan(groups, workers, longestUnit, infeasible) : null;
  // greedy schedule is O(days×workers×tasks) — skip it on large data (only 复盘 uses it)
  const sim = (orders.length <= 800 && workers.length <= 200)
    ? simulate(orders, workers)
    : { orderSequence: orders.map(o=>o.id), completion:{}, days: minHours==null?0:Math.ceil(minHours/HOURS_PER_DAY),
        daily:[], totalCapacity: workers.length*HOURS_PER_DAY, allocations:[], stuck:[], unfinishedHours:0, skipped:true };

  let sensitivity = null;
  if (minHours !== null && binding.length){
    const bindCats = new Set(binding.map(x=>x.cat));
    const grade = bindCats.has('高') ? 'A' : (bindCats.has('低') ? 'B' : 'C');
    const first = binding.slice().sort((a,b)=>(a.prod<b.prod?-1:a.prod>b.prod?1:0)||(a.role<b.role?-1:a.role>b.role?1:0))[0];
    sensitivity = { product: first.prod, role: first.role, grade, rows: hireSensitivity(groups, workers, longestUnit, first.prod, first.role, grade, 6) };
  }

  // as-assigned (fixed workers); 工作日固定 = 周一至周五
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
   RETROSPECTIVE (uses 创建日期 as submit, 实际完成 as finish)
   ==================================================================== */
function runRetrospective(state, opts){
  const wpw = WORKDAYS_PER_WEEK;
  const H = HOURS_PER_DAY;
  // 复盘只看「已完成」的订单（填了实际完成日期）：范围 = 最早创建日期 → 今天。
  // 未完成的订单归「排期」管，不参与复盘的任何统计（含基准工时、工人负载、④对比）。
  const finishedOrders = (state.orders || []).filter(o => String(o.actualFinish || '').trim());
  const R = runFullAnalysis({ ...state, orders: finishedOrders });
  const out = { R, hasData: R.hasData, H, wpw,
    nAll: (state.orders || []).length, nFinished: finishedOrders.length,
    nUnfinished: (state.orders || []).length - finishedOrders.length };
  if (!R.hasData) return out;

  // benchmark = the AS-ASSIGNED schedule (given designer/checker), which scales & needs no greedy sim
  const A = R.asAssigned;
  const expById = new Map(A.rows.map(r => [r.id, Math.max(1, Math.ceil(r.finishWD))]));   // expected completion (work-day)
  const orders = R.orders.map(o => ({ id:o.id, so:o.so, product:o.product, partsTxt:o.vcs.join('、'),
    submit: parseISO(o.created), finish: parseISO(o.actualFinish), deadline: parseISO(o.deadline),
    simDay: expById.has(o.id) ? expById.get(o.id) : null, hours: o.total }));
  out.orders = orders; out.nOrders = orders.length;
  const withFinish = orders.filter(o => o.finish); out.nWithFinish = withFinish.length;

  out.infeasible = R.minHours === null;
  out.optimumDays = out.infeasible ? null : R.minHours / H;   // 什么-if 理论下界（参考）
  out.expectedDays = A.makespanDays;                          // 按指派应完成工期
  out.simDays = A.makespanDays; out.totalDemand = R.totalDemand; out.nWorkers = R.nWorkers;

  // 每人可用工时 = 一年 260 个工作日 − 年假，× 每人每日工时（分母是"一年"，不是最忙者的工期）
  const AL = (state.settings && state.settings.annualLeave >= 0) ? +state.settings.annualLeave : 15;
  const cap = (260 - AL) * H;
  out.capYear = cap; out.annualLeave = AL;
  const loadByName = new Map(A.workerLoads.map(w => [w.name, w]));
  // 优先用工时台账的实际报工（订单设计+订单校对）；没导台账时退回按指派的标准工时
  const HLw = new Map(((state.hourlog && state.hourlog.workers) || []).map(w =>
    [w.name, (w.byTask['订单设计'] || 0) + (w.byTask['订单校对'] || 0)]));
  out.loadsAreActual = HLw.size > 0;
  out.workerLoads = (R.workers || []).map(w => {
    const L = loadByName.get(w.name);
    const inHL = HLw.has(w.name);                       // 台账里有没有这个名字（对账标色用）
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
    // 按单自身的日历口径（不受其他订单/全局锚点影响）
    o.turnaroundCal = o.submit ? Math.max(0, diffDays(o.finish, o.submit)) : null;   // 实际用时 = 实际完成 − 创建
    o.lateCal = o.deadline ? diffDays(o.finish, o.deadline) : null;                  // 迟交 = 实际完成 − 交付 (>0 才算迟)
    // 理论 vs 实际（逐单，工作日口径）：理论 = 该单设计+校对连续做完；实际 = 创建→完成之间的周一至周五天数
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
  // 真实迟交（实际完成晚于交付日期），按迟交天数降序 —— 复盘表用这个
  out.lateVsDeadline = withFinish.filter(o => o.lateCal != null && o.lateCal > 0)
    .sort((a, b) => (b.lateCal - a.lateCal) || (a.id < b.id ? -1 : 1));
  out.nWithDeadline = withFinish.filter(o => o.deadline).length;
  // 真实交付口径（按每张单自身的日历天，不用全局锚点）：准时率 / 迟交单数 / 中位实际用时 / 最长迟交
  const median = arr => { if (!arr.length) return null; const s = arr.slice().sort((a,b)=>a-b); const m = s.length >> 1;
    return s.length % 2 ? s[m] : Math.round((s[m-1] + s[m]) / 2); };
  out.nLate = out.lateVsDeadline.length;
  out.nOnTime = Math.max(0, out.nWithDeadline - out.nLate);
  out.onTimeRate = out.nWithDeadline > 0 ? out.nOnTime / out.nWithDeadline : null;
  out.medianTurnaround = median(withFinish.map(o => o.turnaroundCal).filter(v => v != null));
  out.maxLate = out.lateVsDeadline.length ? out.lateVsDeadline[0].lateCal : 0;
  out.medianLate = median(out.lateVsDeadline.map(o => o.lateCal));
  // 完成速度（逐单，工作日）。经过时间不能加总（订单并行，同一天会被每张在途单重复计），
  // 所以用直观表达：「X% 的单 ≤ N 个工作日完成」+ 区间分布；只有纯工时（总工作量）可以加总。
  const withBoth = withFinish.filter(o => o.turnWD != null);
  out.nTheo = withBoth.length;
  out.theoSumWD = withBoth.reduce((a, o) => a + o.theoWD, 0);                        // 总工作量（人·工作日）
  out.avgTheoHours = out.nTheo ? withBoth.reduce((a, o) => a + o.hours, 0) / out.nTheo : null;  // 单均活量（小时）
  const turns = withBoth.map(o => o.turnWD).sort((a, b) => a - b);
  const pctile = q => turns.length ? turns[Math.min(turns.length - 1, Math.max(0, Math.ceil(q * turns.length) - 1))] : null;
  out.turnP50 = pctile(0.5); out.turnP90 = pctile(0.9); out.turnP95 = pctile(0.95);
  out.turnMax = turns.length ? turns[turns.length - 1] : null;
  const BUCKETS = [[0,0,'当天'],[1,5,'1–5 天（约 1 周）'],[6,10,'6–10 天（约 2 周）'],
    [11,21,'11–21 天（约 1 个月）'],[22,65,'22–65 天（约 1 季度）'],[66,Infinity,'超过 65 天（> 1 季度）']];
  let cum = 0;
  out.turnDist = BUCKETS.map(([lo, hi, label]) => { const n = turns.filter(t => t >= lo && t <= hi).length; cum += n;
    return { label, lo, hi, n, share: out.nTheo ? n / out.nTheo : 0, cumShare: out.nTheo ? cum / out.nTheo : 0 }; });
  return out;
}

/* ====================================================================
   SCHEDULE (deadline = 交付日期/required)
   ==================================================================== */
function isWorkDay(d){ const dow = d.getDay(); return dow !== 0 && dow !== 6; }   // 周一至周五
function runSchedule(state, opts){
  const wpw = WORKDAYS_PER_WEEK;
  const H = HOURS_PER_DAY;
  const R = runFullAnalysis(state);
  const out = { R, hasData: R.hasData, H, wpw };
  if (!R.hasData) return out;
  const nWorkers = R.nWorkers, C = nWorkers * H; out.nWorkers = nWorkers; out.capacity = C;
  const start = parseISO(opts && opts.startDate) || new Date(); start.setHours(0,0,0,0); out.start = start;

  // 排期只排「未完成」的订单（没填实际完成日期），从开工日（默认今天）向后排；
  // 已完成的活归「复盘」管，不参与排期、不占产能。
  const openOrders = R.orders.filter(o => !o.actualFinish);
  out.nFinished = R.orders.length - openOrders.length;
  const orders = openOrders.map((o, i) => ({ id:o.id, so:o.so, seq:i, product:o.product, hours:o.total,
    designH:o.designHours, verifyH:o.verifyHours, difficulty:o.difficulty, partsTxt:o.vcs.join('、'),
    designer:o.designer || '', checker:o.checker || '', deadline: parseISO(o.deadline), created: parseISO(o.created) }));
  out.noDeadline = orders.filter(o => !o.deadline).map(o => o.id);
  const schedulable = orders.filter(o => o.hours > 1e-9);          // ①②连续/slot 排全部有工时的未完成单（无交期也排）
  const withDl = schedulable.filter(o => o.deadline); out.scheduledOrders = withDl;
  out.nOpen = schedulable.length;
  const ec = orders.filter(o => o.created).sort((a, b) => a.created - b.created)[0];
  out.earliestCreated = ec ? { id: ec.id, date: ec.created } : null;
  if (!schedulable.length){ out.noSchedule = true; return out; }

  const lastDl = withDl.length ? new Date(Math.max(...withDl.map(o => o.deadline.getTime()))) : new Date(start);
  // 末交期是谁定的 + 疑似填错检测：交付日期比全部订单里最晚的创建日期还晚 1 年以上，多半是笔误/导入错位
  out.lastDl = lastDl;
  out.lastDlOrders = withDl.filter(o => o.deadline.getTime() === lastDl.getTime()).map(o => o.id);
  const createds = R.orders.map(o => parseISO(o.created)).filter(Boolean);
  const maxCreated = createds.length ? new Date(Math.max(...createds.map(d => d.getTime()))) : null;
  out.maxCreated = maxCreated;
  out.farDeadlines = maxCreated ? withDl.filter(o => diffDays(o.deadline, maxCreated) > 365)
    .sort((a, b) => b.deadline - a.deadline).map(o => ({ id: o.id, deadline: o.deadline })) : [];
  const total = schedulable.reduce((a, o) => a + o.hours, 0); out.total = total;   // 全部工时（含无交期）
  // 年假：每人每年 AL 天，没有具体日期，按一年 52×5=260 个工作日均摊成日产能折扣（仅影响排期）
  const AL = Math.max(0, Math.min(100, Math.round(
    (opts && opts.annualLeaveDays != null && isFinite(+opts.annualLeaveDays)) ? +opts.annualLeaveDays
    : (state.settings && state.settings.annualLeave != null && isFinite(+state.settings.annualLeave)) ? +state.settings.annualLeave : 15)));
  const leaveFactor = Math.max(0, (260 - AL) / 260);
  const baseCap = C * leaveFactor;   // 扣完年假后的日常团队产能
  out.annualLeave = AL; out.leaveFactor = leaveFactor; out.baseCap = baseCap;
  // 请假登记：把落在每个工作日的请假小时数从当天产能里再扣掉（有具体日期的精确扣）
  // 每条记录 hpd = 每天请假的小时数（留空/非法 = 整班 H，最多扣满一个班）
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
  // 时间线：铺到末交期，且至少铺到全部工时能做完为止（连续排可能超出末交期），上限 1500 个工作日。
  // 个人级模拟里最忙的人串行干完自己的队列可能比团队合力还慢——把时间线也铺到够他用。
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
    if (baseCap <= 1e-9 && cur.getTime() > lastDl.getTime()) break;   // 产能为 0 时避免空转
  }
  if (!days.length){ days.push(new Date(start)); cap.push(dayCap(start)); }
  const N = days.length; out.days = days; out.N = N;
  const capCum = []; { let a = 0; for (let i = 0; i < N; i++){ a += cap[i]; capCum.push(a); } }
  out.cap = cap; out.capCum = capCum;
  out.leaveHoursLost = cap.reduce((s, c) => s + (baseCap - c), 0);
  out.leaveDaysAffected = cap.filter(c => c < baseCap - 1e-9).length;
  const dlIndex = d => { let c = 0; for (let i = 0; i < N; i++){ if (days[i].getTime() <= d.getTime()) c++; else break; } return c; };
  withDl.forEach(o => o.Tg = dlIndex(o.deadline));
  out.overdue = withDl.filter(o => o.deadline.getTime() < start.getTime()).map(o => o.id);   // 交期早于开工日 = 已逾期

  // ①② 逐单排队：按顺序连续消耗产能，得到每张单的预计开始/完成日
  const seqSchedule = seq => {
    const rows = []; let cum = 0, p = 0;
    for (const o of seq){
      const sH = cum, fH = cum + o.hours; cum = fH;
      while (p < N && capCum[p] <= sH + 1e-9) p++;
      const startIdx = p < N ? p : null;
      while (p < N && capCum[p] < fH - 1e-9) p++;
      const finishIdx = p < N ? p : null;
      const dlT = o.deadline ? dlIndex(o.deadline) : null;                       // 交期是第几个工作日（1-based）
      const lateWD = (finishIdx != null && dlT != null) ? (finishIdx + 1) - dlT : null;   // >0 迟交, <0 提前, 0 当天
      rows.push({ id:o.id, so:o.so, product:o.product, hours:o.hours, designH:o.designH, verifyH:o.verifyH,
        created:o.created, deadline:o.deadline,
        designer:o.designer, checker:o.checker, startIdx, finishIdx, dlT, lateWD });
    }
    return rows;
  };
  const byCreated = (a, b) => ((a.created ? a.created.getTime() : Infinity) - (b.created ? b.created.getTime() : Infinity)) || (a.seq - b.seq);
  out.seqCont = seqSchedule(schedulable.slice().sort(byCreated));                                  // ① 连续排：创建日期先后
  out.seqSlot = seqSchedule(schedulable.slice().sort((a, b) =>                                     // ② 按交期：交期最早优先，无交期垫底
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
  // 贴近交期（倒排/JIT）：从每个交期往前倒推，先争取在交期前 buffer 个工作日完成；
  // 产能不够就再往前借天；还不够才动用缓冲期（完成贴到交期但不迟交）；仍不够 = 真超载（标红）。
  // 交期从晚到早处理：晚交的活先占靠后的日子，天然实现「晚交的单不提前赶」。
  const buffer = Math.max(0, Math.min(60, Math.round(
    (opts && opts.bufferDays != null && isFinite(+opts.bufferDays)) ? +opts.bufferDays
    : (state.settings && state.settings.schedBuffer != null && isFinite(+state.settings.schedBuffer)) ? +state.settings.schedBuffer : 2)));
  out.buffer = buffer;
  const jit = new Array(N).fill(0); const remCap = cap.slice();
  const overloadDeadlines = []; const bufferShort = [];
  dls.slice().sort((a, b) => b.T - a.T).forEach(d => {
    let need = d.hours; let lastUsed = -1;
    const end = Math.min(N - 1, d.T - 1);            // 不迟交的最后一个可用工作日
    const idealEnd = end - buffer;                   // 目标完成日（交期前 buffer 个工作日）
    for (let i = idealEnd; i >= 0 && need > 1e-9; i--){ const t = Math.min(remCap[i], need); if (t > 0){ jit[i] += t; remCap[i] -= t; need -= t; if (i > lastUsed) lastUsed = i; } }
    if (need > 1e-9){                                // 提前完不成 → 动用缓冲期（仍不迟交）
      bufferShort.push(d);
      for (let i = Math.max(0, idealEnd + 1); i <= end && need > 1e-9; i++){ const t = Math.min(remCap[i], need); if (t > 0){ jit[i] += t; remCap[i] -= t; need -= t; if (i > lastUsed) lastUsed = i; } }
    }
    if (need > 1e-9){ overloadDeadlines.push(d); jit[Math.max(0, end)] += need; lastUsed = Math.max(0, end); }   // 真超载
    d.jitFinishIdx = lastUsed < 0 ? Math.max(0, Math.min(idealEnd, end)) : lastUsed;   // 这批活最后一天排在哪
    d.jitLeadWD = Math.max(0, (d.T - 1) - d.jitFinishIdx);                             // 比交期提前几个工作日
  });
  out.jit = jit; out.overloadDeadlines = overloadDeadlines; out.feasibleAll = dls.every(d => d.feasible);
  out.bufferShort = bufferShort.filter(d => !overloadDeadlines.includes(d));
  out.jitPeak = Math.max(0, ...jit); out.eagerPeak = Math.max(0, ...eager);
  out.eagerCum = []; out.jitCum = []; let a2 = 0, b2 = 0;
  for (let i = 0; i < N; i++){ a2 += eager[i]; b2 += jit[i]; out.eagerCum.push(a2); out.jitCum.push(b2); }
  // 每个交期在「尽早做完」参考排法下最早哪天可交（eagerCum 首次覆盖应交累计的那天）
  { let k = 0;
    dls.forEach(d => { while (k < N - 1 && out.eagerCum[k] < d.cum - 1e-6) k++; d.eagerFinishIdx = k; }); }

  // ---- ④ 个人级模拟：设计工时进设计工人自己的时间线，校对在该单设计完成后进校对工人的时间线；
  //      请假按人精确扣、年假均摊。三种口径：
  //      cont=① 连续排（按创建顺序正排）、slot=② 按交期（交期最早先做，正排）、
  //      jit =③ 贴近交期（从交期往前倒排：校对贴在交期前 buffer 天完成，设计在校对开始前完成；只含有交期的单）。
  //      正排口径下，同一人既设计又校对时先做完自己的设计队列再做校对（近似，见页面说明）。----
  const baseW = H * leaveFactor;                                   // 每人每日可用（含年假均摊）
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
  // 正排（cont / slot）：按给定顺序把设计排进各自时间线，校对按「该单设计完成日」先后排
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
      .forEach(r => { const dv = designFin.get(r.id); alloc(r.checker, r.verifyH, isFinite(dv) ? dv : N, (r.so || r.id) + '·校'); });
    return { day: pwDay, orders: pwOrders, workers: pwSummary(pwDay, unplacedBy), skipped: 0 };
  };
  // 倒排（jit）：交期晚的先占靠后的日子；校对从「交期前 buffer 天」往前找空，设计在校对开始那天之前完成。
  // 交期前塞不下的小时数堆在最后可用日上（当天会超产能、页面标红 = 该人贴不进交期）。
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
      for (let d = Math.max(0, idealEnd + 1); d <= Math.min(hardEnd, N - 1) && rem > 1e-9; d++){   // 动用缓冲期（仍不迟交）
        const free = dayCapW(name, d) - arr[d];
        if (free > 1e-9){ const take = Math.min(free, rem); arr[d] += take; rem -= take;
          if (first == null) first = d;
          (lab[d] || (lab[d] = [])).push(label); }
      }
      if (rem > 1e-9){                                             // 真超载：贴不进交期
        const d = Math.max(0, Math.min(hardEnd, N - 1));
        arr[d] += rem; (lab[d] || (lab[d] = [])).push(label);
        if (first == null) first = d;
        unplacedBy.set(name, (unplacedBy.get(name) || 0) + rem);
      }
      return { start: first };
    };
    const withDlRows = out.seqSlot.filter(r => r.dlT != null).slice().sort((a, b) => b.dlT - a.dlT);
    withDlRows.forEach(r => {
      const end = Math.max(0, Math.min(N - 1, r.dlT - 1));         // 交期当天 = 不迟交的最后一天
      const ideal = Math.max(0, end - buffer);
      let vStart = ideal;
      if (r.checker && r.verifyH > 1e-9){ const a = allocB(r.checker, r.verifyH, ideal, end, (r.so || r.id) + '·校'); vStart = a.start == null ? ideal : a.start; }
      if (r.designer && r.designH > 1e-9) allocB(r.designer, r.designH, vStart, vStart, r.so || r.id);
    });
    return { day: pwDay, orders: pwOrders, workers: pwSummary(pwDay, unplacedBy), skipped: out.seqSlot.length - withDlRows.length };
  };
  out.pw = { cont: simForward(out.seqCont), slot: simForward(out.seqSlot), jit: simJit() };

  // ---- ⑤ 重新分配建议：把超载者的活转给较闲的人（只是建议，不改任何数据）。
  //      有能力评级的按评级：只转给评级显示能胜任「该产品·该环节·该难度」的人；
  //      没有任何评级能胜任的人选时，退回「谁空谁来」——转给最闲且评级空白的人（标「能力未填」，请先人工确认）；
  //      评级明确写着不胜任的人永远不推荐。设计环节只转给别人当设计、校对环节只转给别人当校对，绝不混。----
  const wByName = new Map((R.workers || []).map(w => [w.name, w]));
  const loads = new Map(loadPreview);
  (R.workers || []).forEach(w => { if (!loads.has(w.name)) loads.set(w.name, 0); });
  const load0 = new Map(loads);                            // 调整前的原始负载（页面展示用）
  const pool = [...loads.keys()];
  const totalLoad = [...loads.values()].reduce((a, b) => a + b, 0);
  const avgLoad = pool.length ? totalLoad / pool.length : 0;
  const canDo = (name, o, role) => {                       // 'ok' | 'unknown'(能力未填) | null(评级不胜任/不在名册)
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
  const moves = []; const usedTask = new Set();            // 同一张单的同一环节只转一次
  if (avgLoad > 0){
    const overs = pool.filter(n => loads.get(n) > avgLoad * 1.15).sort((a, b) => loads.get(b) - loads.get(a));
    for (const from of overs){
      let lastTo = null;                                   // 打包：小单尽量持续转给同一个人，凑成整块（而不是散给所有人每人一小时）
      for (const t of tasksOf(from)){                      // 从最大的活开始转
        if (loads.get(from) <= avgLoad * 1.05 + 1e-9) break;      // 已接近平均，够了
        const key = t.o.id + '¦' + t.role;
        if (usedTask.has(key) || t.h < 0.02) continue;            // 只跳过几乎为 0 的碎屑（很多单每张只有零点几小时，不能因小而不转）
        if (loads.get(from) - t.h < avgLoad * 0.8) continue;      // 整单转走会矫枉过正 → 试更小的单
        // 接收上限：一般不超过平均值；单子本身比平均值大时放宽到「平均 + 半个单」，否则大单永远转不动
        const recvCap = Math.max(avgLoad * 1.05, avgLoad + t.h / 2);
        const fits = n => n !== from && loads.get(n) + t.h <= recvCap + 1e-9;
        const st = n => canDo(n, t.o, t.role);
        let to = null, sure = false;
        if (lastTo && fits(lastTo) && st(lastTo) === 'ok'){ to = lastTo; sure = true; }   // 评级胜任的上一接收者 → 继续给他
        else {
          const rated = pool.filter(n => fits(n) && st(n) === 'ok');                      // 有评级按评级
          if (rated.length){ to = rated[0]; for (const n of rated) if (loads.get(n) < loads.get(to)) to = n; sure = true; }
          else if (lastTo && fits(lastTo) && st(lastTo) === 'unknown'){ to = lastTo; }    // 否则谁空谁来（未填评级）
          else {
            const blank = pool.filter(n => fits(n) && st(n) === 'unknown');
            if (blank.length){ to = blank[0]; for (const n of blank) if (loads.get(n) < loads.get(to)) to = n; }
          }
        }
        if (!to) continue;                                        // 其余人评级都写明不胜任 → 这单不转
        lastTo = to;
        usedTask.add(key);
        moves.push({ from, to, id: t.o.so || t.o.id, role: t.role, h: t.h, sure });
        loads.set(from, loads.get(from) - t.h);
        loads.set(to, loads.get(to) + t.h);
      }
    }
  }
  const byPair = new Map();                                // 汇总成「从 A → 转给 B」的建议行（一行可含很多张小单）
  moves.forEach(m => { const k = m.from + '¦' + m.to;
    if (!byPair.has(k)) byPair.set(k, { from: m.from, to: m.to, fromLoad: load0.get(m.from) || 0, toLoad: load0.get(m.to) || 0, amount: 0, examples: [] });
    const p = byPair.get(k); p.amount += m.h; p.examples.push({ id: m.id, role: m.role, h: m.h, sure: m.sure }); });
  const sugg = [...byPair.values()].sort((a, b) => b.amount - a.amount)
    .filter(p => p.amount >= Math.max(0.5, avgLoad * 0.01));      // 合计不到半小时级别的碎条不展示
  out.rebalance = { avgLoad, suggestions: sugg,
    maxBefore: Math.max(0, ...load0.values()),
    maxAfter: Math.max(0, ...loads.values()) };
  return out;
}

window.runFullAnalysis = runFullAnalysis;
window.runRetrospective = runRetrospective;
window.runSchedule = runSchedule;
window.fmtDuration = formatDuration;
