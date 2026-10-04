import { firebaseConfig, OWNER_EMAIL, EMAIL_DOMAIN } from './config.js?v=202610041055';
import { kickoff, visitRecap, marketUpdate, dailyStore, dailyMarket } from './msgs.js?v=202610041055';
import {
  STORES, DISTRICTS, canonicalStore, isKnownStore, parseRsa, rangeFromFileName, parseTeamRoster, resolveReportNames,
  paceFactor, DEFAULT_GOALS, cidOf, status, fmt, goalsFor, TEAM_FOCUS, pickStoreFocus
} from './base.js?v=202610041055';
import {
  iso, fromIso, addDays, daysApart, weekStartOf, DAY_NAMES, DAY_LONG, dow, DEFAULT_OFF, validOff, safeOff, VISIT_DAYS, STORE_GOALS,
  parseDaily, needScore, band, pct, environment, buildPlan, pivotSuggestion, ELEMENTS, SEGMENTS, AORS, PRACTICE, VISIT_TYPES, kindToType, visitScore, visitSummary, consultantCoaching, drillFor, draggers, helpers, STORE_TO_RSA, hasCommitment, commitmentText, blackoutFor, offChoicesFor, storeFocus, rsaPicks, consultantWeeks, teamSignals,
  STORE_METRICS, slug, COACHING, METRICS, PLAIN, isOutlet, driveMin, driveText, MAX_SPLIT_MIN, LEVERS, leverStatus, suggestLever,
  consultantTrends, TREND_ROWS, trendFmt, trendRead, TREND_LABEL,
  OFFER_DEFAULT, PLAY, PLAY_CHECKS, PLAY_CHECKS_REMOTE, offerActive, offerMath, FLIQ_CHECKS, FLIQ_CHECKS_REMOTE, FLIQ_DAILY
} from './ml.js?v=202610041055';

// Legacy Sunday-start weeks, read as the Monday week that replaced them.
function fromSundayPlan(p, week) {
  if (!p) return null;
  const end = addDays(week, 6);
  const days = (p.days || []).filter(d => d.date >= week && d.date <= end);
  if (!safeOff(p.off).includes(0) && !days.some(d => d.date === end) && days.length < VISIT_DAYS) days.push({ date: end, store: null, kind: 'open', status: 'planned', score: null });
  return { ...p, weekStart: week, days, fromSunday: p.weekStart };
}
const sundayShift = (x, week) => x ? { ...x, weekStart: week, fromSunday: x.weekStart } : null;
const sundayId = id => { const m = String(id).match(/^(.*)_(\d{4}-\d{2}-\d{2})$/); return m && dow(m[2]) === 1 ? `${m[1]}_${addDays(m[2], -1)}` : '__none'; };
const sundayDoc = (d, id) => d ? { ...d, id, weekStart: id.slice(-10), fromSunday: d.weekStart } : null;
// A plan has days set by hand when the leader changed a day in that week (swaps, stops, edits).
const handSet = p => (p?.pivots || []).some(x => x.from !== 'Anchor' && x.date >= p.weekStart && x.date <= addDays(p.weekStart, 6));
const DEMO = !firebaseConfig.apiKey || firebaseConfig.apiKey.startsWith('PASTE');
const FB = 'https://www.gstatic.com/firebasejs/10.12.2/';
const ROLES = [['admin', 'Admin'], ['exec', 'Executive (view all)'], ['director', 'Director'], ['leader', 'Market Leader']];
const FIELD = ['leader', 'director'];
const roleLabel = r => (ROLES.find(x => x[0] === r) || [r, r])[1];
const S = { tab: null };

// ---------------------------------------------------------------- helpers
const $ = sel => document.querySelector(sel);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const today = () => iso(new Date());
const shortDate = s => { const d = fromIso(s); return `${d.getMonth() + 1}/${d.getDate()}`; };
const dayLabel = s => `${DAY_NAMES[dow(s)]} ${shortDate(s)}`;
const longDate = s => fromIso(s).toLocaleDateString('en-US', { weekday: 'long', month: 'short', day: 'numeric' });
const titleName = n => /^[A-Z\s.'-]+$/.test(n) ? n.toLowerCase().replace(/(^|[\s'-])([a-z])/g, (a, b, c) => b + c.toUpperCase()) : n;
function toast(msg, bad) {
  const t = $('#toast'); t.textContent = msg; t.className = 'toast show' + (bad ? ' bad' : '');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.className = 'toast', 3800);
}
function parseCsvText(text) {
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  const clean = rows.filter(r => r.some(c => c.trim() !== ''));
  if (!clean.length) return [];
  const head = clean[0].map(h => h.trim());
  return clean.slice(1).map(r => Object.fromEntries(head.map((h, i) => [h, r[i] ?? ''])));
}
// ---------------------------------------------------------------- daily budget (monthly workbook)
// Each store sheet: MONTHLY TARGETS (label in A, value in D), then Date | Day | Traffic | Sales rows.
// Stored as config/budget_YYYY-MM: { month, stores: { store: { m: {...}, d: { 'YYYY-MM-DD': [traffic, sales] } } } }.
function parseBudgetBook(wb, XLSX) {
  const stores = {}, skipped = []; let month = null;
  const isoOf = v => v instanceof Date ? iso(new Date(v.getFullYear(), v.getMonth(), v.getDate())) : /^\d{4}-\d{2}-\d{2}/.test(String(v)) ? String(v).slice(0, 10) : null;
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null });
    const bi = rows.findIndex(r => /daily budget breakdown/i.test(String(r?.[0] || '')));
    if (bi < 0) continue;
    const store = canonicalStore(String(rows[0]?.[0] || name).trim()) || canonicalStore(name);
    if (!isKnownStore(store)) { skipped.push(name); continue; }
    const m = {}, lab = { 'total traffic': 'traffic', 'total written sales': 'sales', 'close rate': 'closeRate', 'average ticket': 'avgTicket', 'sales per guest (spg)': 'spg', 'bedding spg (bspg)': 'bspg', 'effective gm %': 'egm' };
    rows.slice(0, bi).forEach(r => { const k = lab[String(r?.[0] || '').trim().toLowerCase()]; if (k && typeof r[3] === 'number') m[k] = r[3]; });
    const d = {};
    rows.slice(bi + 2).forEach(r => { const dt = isoOf(r?.[0]); if (dt && typeof r[2] === 'number' && typeof r[3] === 'number') { d[dt] = [Math.round(r[2] * 10) / 10, Math.round(r[3])]; month ||= dt.slice(0, 7); } });
    if (Object.keys(d).length) stores[store] = { m, d };
  }
  return Object.keys(stores).length ? { month, stores, skipped } : null;
}
// Budget for one store on one day: revenue, traffic and SPG (sales divided by traffic).
function budgetFor(store, date) {
  const b = S.budgets?.[date?.slice(0, 7)]?.stores?.[store]; const x = b?.d?.[date];
  return x ? { traffic: x[0], sales: x[1], spg: x[0] ? x[1] / x[0] : null } : null;
}
function budgetToDate(store, date) {
  const b = S.budgets?.[date?.slice(0, 7)]?.stores?.[store]; if (!b) return null;
  let traffic = 0, sales = 0; Object.entries(b.d).forEach(([k, v]) => { if (k <= date) { traffic += v[0]; sales += v[1]; } });
  return { traffic, sales, spg: traffic ? sales / traffic : null, month: b.m };
}
// ---------------------------------------------------------------- open carts (Storis open cart detail)
// Rolled up by consultant and store. Guest phone and email are dropped on the way in and never stored;
// the top carts keep only the guest's first name and last initial so the leader knows who to call.
const CART_DUE = [1, 3, 7]; // follow-up days: 1, 3 and 7 days after the cart was started
function parseCarts(rows, file) {
  const num = v => { const n = parseFloat(String(v ?? '').replace(/[$,]/g, '')); return isFinite(n) ? n : 0; };
  const get = (r, re) => { const k = Object.keys(r).find(k => re.test(k)); return k ? r[k] : ''; };
  const people = {}, stores = {}, bad = new Set(); let total = 0, n = 0;
  const short = g => { const t = String(g || '').trim().split(/\s+/).filter(Boolean); return !t.length ? 'Guest' : t.length === 1 ? t[0] : `${t[0]} ${t[t.length - 1][0]}.`; };
  for (const r of rows) {
    const st0 = String(get(r, /^store$/i)).trim(), who = String(get(r, /^associate$/i)).trim();
    if (!st0 || !who) continue;
    const store = canonicalStore(st0); if (!isKnownStore(store)) { bad.add(st0); continue; }
    const value = num(get(r, /cart value/i)), age = Math.round(num(get(r, /age/i))), lines = num(get(r, /^lines$/i));
    const cid = cidOf(who);
    const p = people[cid] ||= { name: who, store, n: 0, value: 0, wk: 0, wkValue: 0, due: 0, old: 0, top: [] };
    p.n++; p.value += value; if (age <= 7) { p.wk++; p.wkValue += value; } if (CART_DUE.includes(age)) p.due++; if (age > 14) p.old++;
    p.top.push({ g: short(get(r, /^guest$/i)), v: Math.round(value), a: age, l: lines });
    const sx = stores[store] ||= { n: 0, value: 0, wk: 0, wkValue: 0, due: 0, old: 0 };
    sx.n++; sx.value += value; if (age <= 7) { sx.wk++; sx.wkValue += value; } if (CART_DUE.includes(age)) sx.due++; if (age > 14) sx.old++;
    total += value; n++;
  }
  Object.values(people).forEach(p => { p.top = p.top.sort((a, b) => (b.a <= 14) - (a.a <= 14) || b.v - a.v).slice(0, 5); p.value = Math.round(p.value); p.wkValue = Math.round(p.wkValue); });
  Object.values(stores).forEach(x => { x.value = Math.round(x.value); x.wkValue = Math.round(x.wkValue); });
  const m = String(file).match(/(\d{4}-\d{2}-\d{2})/);
  return { date: m ? m[1] : today(), people, stores, total: Math.round(total), n, bad: [...bad] };
}
function prepCarts(file, rows) {
  const c = parseCarts(rows, file);
  if (!c.n) return { file, error: 'No open carts found. Check the file has Store, Associate and the cart value columns.' };
  return { file, kind: 'carts', label: 'open carts',
    summary: `Open carts as of <b>${esc(shortDate(c.date))}</b> · <b>${c.n.toLocaleString('en-US')}</b> carts · about <b>${$k(c.total)}</b> estimated value · ${Object.keys(c.people).length} consultants in ${Object.keys(c.stores).length} stores. Guest phone numbers and emails are left out.`,
    extra: c.bad.length ? `<p class="small warn">Rows skipped (store not recognized): ${c.bad.map(esc).join(', ')}</p>` : '',
    publish: () => S.be.saveCarts(c) };
}
const cartsFor = cid => S.carts?.people?.[cid] || null;
const cartLine = c => `${c.n} open cart${c.n === 1 ? '' : 's'} · about ${$k(c.value)}${c.wk ? ` · ${c.wk} started this week (${$k(c.wkValue)})` : ''}${c.due ? ` · ${c.due} due a follow-up today` : ''}`;
function prepBudget(file, b) {
  const n = Object.keys(b.stores).length, tot = Object.values(b.stores).reduce((a, x) => a + Object.values(x.d).reduce((y, v) => y + v[1], 0), 0);
  const missing = STORES.filter(st => !b.stores[st.name]).map(st => st.name);
  return { file, kind: 'budget', label: 'daily budget',
    summary: `Daily budget for <b>${esc(new Date(b.month + '-15T12:00').toLocaleString('en-US', { month: 'long', year: 'numeric' }))}</b> · ${n} stores · $${Math.round(tot).toLocaleString('en-US')} in sales for the month. Market Leaders see each store's revenue and SPG budget for the day.`,
    extra: (missing.length ? `<p class="small warn">No budget in this file for: ${missing.map(esc).join(', ')}</p>` : '') + (b.skipped.length ? `<p class="small muted">Sheets skipped (market totals or not a store): ${b.skipped.map(esc).join(', ')}</p>` : ''),
    publish: () => S.be.saveBudget(b) };
}
async function readSpreadsheet(file) {
  if (/\.csv$/i.test(file.name)) return parseCsvText(await file.text());
  const XLSX = await import('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm');
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  // The monthly daily budget workbook: one sheet per store with a DAILY BUDGET BREAKDOWN table.
  const budget = parseBudgetBook(wb, XLSX);
  if (budget) return Object.assign([], { budget });
  // Roster exports keep the team on a "Sales Team" sheet; everything else is the first sheet.
  const name = wb.SheetNames.find(n => /sales team|store leaders/i.test(n)) || wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
}
const storeOptions = (selected, list = STORES.map(s => s.name), extra = '') => {
  const groups = {};
  list.forEach(n => { const st = STORES.find(x => x.name === n); if (st) (groups[st.district] ||= []).push(st); });
  return extra + Object.entries(groups).map(([k, arr]) => `<optgroup label="${esc(DISTRICTS[k])}">${arr.map(st =>
    `<option value="${esc(st.name)}" ${st.name === selected ? 'selected' : ''}>${esc(st.name)}</option>`).join('')}</optgroup>`).join('');
};

// ---------------------------------------------------------------- Firebase backend
async function firebaseBackend() {
  const [{ initializeApp }, A, F] = await Promise.all([
    import(FB + 'firebase-app.js'), import(FB + 'firebase-auth.js'), import(FB + 'firebase-firestore.js')
  ]);
  const app = initializeApp(firebaseConfig);
  const auth = A.getAuth(app);
  const db = F.getFirestore(app);
  const email = () => (auth.currentUser?.email || '').toLowerCase();
  const get = async (c, id) => { const s = await F.getDoc(F.doc(db, c, id)); return s.exists() ? { id: s.id, ...s.data() } : null; };
  const all = async (c, ...wheres) => (await F.getDocs(F.query(F.collection(db, c), ...wheres.map(([a, b]) => F.where(a, '==', b))))).docs.map(d => ({ id: d.id, ...d.data() }));
  const be = {
    demo: false,
    onAuth: cb => A.onAuthStateChanged(auth, u => cb(u ? { email: u.email.toLowerCase(), verified: u.emailVerified } : null)),
    signIn: (e, p) => A.signInWithEmailAndPassword(auth, e, p),
    async register(e, p) { const c = await A.createUserWithEmailAndPassword(auth, e, p); await A.sendEmailVerification(c.user); },
    resendVerify: () => A.sendEmailVerification(auth.currentUser),
    async refresh() { await A.reload(auth.currentUser); await auth.currentUser.getIdToken(true); return auth.currentUser.emailVerified; },
    reset: e => A.sendPasswordResetEmail(auth, e),
    signOut: () => A.signOut(auth),
    async profile() {
      const e = email();
      const p = await get('users', e);
      if (p) return p;
      if (e === OWNER_EMAIL.toLowerCase()) {
        const me = { email: e, name: 'Frank Pina', role: 'admin', stores: [], off: DEFAULT_OFF };
        await F.setDoc(F.doc(db, 'users', e), me);
        return me;
      }
      return null;
    },
    meta: async () => (await get('config', 'meta')) || {},
    daily: date => get('daily', date),
    async publishDaily(d) {
      const now = new Date().toISOString();
      await F.setDoc(F.doc(db, 'daily', d.date), { date: d.date, periods: d.periods, stores: d.stores, by: email(), at: now, file: d.file });
      const meta = await be.meta();
      const dates = [...new Set([...(meta.dailyDates || []), d.date])].sort().reverse().slice(0, 60);
      await F.setDoc(F.doc(db, 'config', 'meta'), { ...meta, dailyDates: dates, latestDaily: dates[0], lastDaily: { by: email(), at: now, file: d.file, date: d.date } });
    },
    rsa: () => get('rsa', 'latest'),
    rsaAt: date => get('rsa', date),
    rsaYtd: () => get('rsa', 'ytd'),
    async publishRsaYtd(r) {
      const now = new Date().toISOString();
      await F.setDoc(F.doc(db, 'rsa', 'ytd'), { ...r, by: email(), at: now });
      const meta = await be.meta();
      await F.setDoc(F.doc(db, 'config', 'meta'), { ...meta, lastRsaYtd: { by: email(), at: now, file: r.file, from: r.from, to: r.to, people: r.people.length } });
    },
    // Every upload is kept by date so the app can take this week out of the month-to-date numbers.
    async publishRsa(r) {
      const now = new Date().toISOString();
      const meta = await be.meta();
      await F.setDoc(F.doc(db, 'rsa', r.to), { ...r, by: email(), at: now });
      const newest = !meta.lastRsa?.to || r.to >= meta.lastRsa.to;
      if (newest) await F.setDoc(F.doc(db, 'rsa', 'latest'), { ...r, by: email(), at: now });
      const dates = [...new Set([...(meta.rsaDates || []), r.to])].sort().reverse().slice(0, 60);
      await F.setDoc(F.doc(db, 'config', 'meta'), { ...meta, rsaDates: dates, ...(newest ? { lastRsa: { by: email(), at: now, file: r.file, to: r.to, people: r.people.length } } : {}) });
    },
    roster: async () => (await get('config', 'roster'))?.people || [],
    markets: async () => (await get('config', 'markets'))?.markets || [],
    storeLeaders: async () => (await get('config', 'storeLeaders'))?.leaders || [],
    saveStoreLeaders: leaders => F.setDoc(F.doc(db, 'config', 'storeLeaders'), { leaders, at: new Date().toISOString() }),
    saveMarkets: markets => F.setDoc(F.doc(db, 'config', 'markets'), { markets, at: new Date().toISOString() }),
    saveRoster: people => F.setDoc(F.doc(db, 'config', 'roster'), { people, at: new Date().toISOString() }),
    budget: month => get('config', 'budget_' + month),
    carts: () => get('config', 'carts'),
    saveCarts: c => F.setDoc(F.doc(db, 'config', 'carts'), { ...c, by: email(), at: new Date().toISOString() }),
    saveBudget: b => F.setDoc(F.doc(db, 'config', 'budget_' + b.month), { month: b.month, stores: b.stores, by: email(), at: new Date().toISOString() }),
    users: () => all('users'),
    saveUser: u => F.setDoc(F.doc(db, 'users', u.email), u),
    deleteUser: e => F.deleteDoc(F.doc(db, 'users', e)),
    // Weeks used to start on Sunday. Anything saved under a Sunday start is read as that Monday's week.
    plan: async (e, week) => (await get('plans', `${e}_${week}`)) || (dow(week) === 1 ? fromSundayPlan(await get('plans', `${e}_${addDays(week, -1)}`), week) : null),
    plansForWeek: async week => { const a = await all('plans', ['weekStart', week]); if (dow(week) !== 1) return a; const old = await all('plans', ['weekStart', addDays(week, -1)]).catch(() => []); return [...a, ...old.filter(o => !a.some(x => x.email === o.email)).map(o => fromSundayPlan(o, week))]; },
    savePlan: p => F.setDoc(F.doc(db, 'plans', `${p.email}_${p.weekStart}`), p),
    visits: () => all('visits'),
    saveVisit: v => F.setDoc(F.doc(db, 'visits', v.id), v),
    photos: visitId => all('photos', ['visitId', visitId]),
    savePhoto: ph => F.setDoc(F.doc(db, 'photos', ph.id), ph),
    deletePhoto: id => F.deleteDoc(F.doc(db, 'photos', id)),
    timeOff: async (e, week) => (await get('timeoff', `${e}_${week}`)) || (dow(week) === 1 ? sundayShift(await get('timeoff', `${e}_${addDays(week, -1)}`), week) : null),
    timeOffForWeek: async week => { const a = await all('timeoff', ['weekStart', week]); if (dow(week) !== 1) return a; const old = await all('timeoff', ['weekStart', addDays(week, -1)]).catch(() => []); return [...a, ...old.filter(o => !a.some(x => x.email === o.email)).map(o => sundayShift(o, week))]; },
    saveTimeOff: t => F.setDoc(F.doc(db, 'timeoff', `${t.email}_${t.weekStart}`), t),
    oneOnOne: async id => (await get('oneonones', id)) || sundayDoc(await get('oneonones', sundayId(id)), id),
    saveOneOnOne: d => F.setDoc(F.doc(db, 'oneonones', d.id), d),
    onePrivate: async id => (await get('oneprivate', id)) || sundayDoc(await get('oneprivate', sundayId(id)), id),
    saveOnePrivate: d => F.setDoc(F.doc(db, 'oneprivate', d.id), d),
    alerts: () => all('alerts'),
    offer: async () => (await get('config', 'offer'))?.offer || null,
    saveOffer: o => F.setDoc(F.doc(db, 'config', 'offer'), { offer: o, at: new Date().toISOString() }),
    saveAlert: a => F.setDoc(F.doc(db, 'alerts', a.id), a)
  };
  return be;
}

// ---------------------------------------------------------------- demo backend (in memory, made-up numbers)
function demoBackend() {
  let seed = 7; const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
  const t = today(), week = weekStartOf(t);
  const leaders = [
    { email: 'east@demo', name: 'Demo Market Leader (Jacksonville)', role: 'leader', stores: ['Town Center', 'Orange Park', 'Yulee', 'St. Augustine'], off: DEFAULT_OFF },
    { email: 'nc@demo', name: 'Demo Market Leader (Carolinas)', role: 'leader', pilots: ['frontlineiq'], stores: ['Greensboro', 'Winston Salem', 'Burlington', 'Danville', 'Outlet Greensboro'], off: DEFAULT_OFF },
    { email: 'director@demo', name: 'Demo Director (East)', role: 'director', stores: ['Town Center', 'North', 'Orange Park', 'Brunswick', 'Yulee', 'St. Augustine', 'Outlet Regency'], off: [0, 6] },
    { email: 'gulf@demo', name: 'Demo Market Leader (Gulf Coast)', role: 'leader', stores: ['Mobile', "D'Iberville", 'Spanish Fort', 'Pensacola', 'Crestview', 'Ft. Walton'], off: [2, 3] }
  ];
  const users = {
    [OWNER_EMAIL]: { email: OWNER_EMAIL, name: 'Frank Pina', role: 'admin', stores: [], off: DEFAULT_OFF },
    'exec@demo': { email: 'exec@demo', name: 'Demo Executive', role: 'exec', stores: [] }
  };
  leaders.forEach(l => users[l.email] = l);
  // Per-store profile: sales vs budget (MTD, WTD), SPG vs LY, close rate bps, traffic vs LY.
  const fixed = {
    'Town Center': [-9, -7, -5, -120, -6], 'Orange Park': [2, 4, 5, 60, 3], 'Yulee': [-6, -4, -3, -80, -18], 'St. Augustine': [-3, -2, 1, 20, 2],
    'Winston Salem': [-11, -13, -6, -150, -4], 'Burlington': [-4, -6, -2, -60, 1], 'Danville': [-19, -24, -14, -380, -21], 'Greensboro': [3, 6, 4, 90, 5], 'Outlet Greensboro': [-8, -3, -5, -120, -9]
  };
  const dailyRows = (date, override = {}) => STORES.flatMap(st => {
    const f = override[st.name] || fixed[st.name] || [-15 + rnd() * 22, -18 + rnd() * 26, -12 + rnd() * 18, Math.round(-400 + rnd() * 550), -15 + rnd() * 20];
    const [sb, wb, spg, cr, tr] = f;
    const ns = 180000 + rnd() * 250000, wk = ns / 4;
    const bad = sb < -10 ? 1 : 0;
    const m = (metric, mtd, bud = '', ly = '', wtd = mtd, wbud = bud, wly = ly) => ({ 'Report Date': date, Segment: st.name, Metric: metric,
      'Day TY': String(/^(Net Sales|Traffic|Cancellations|Gross)/.test(metric) ? Math.round(Number(wtd) / 3) : wtd), 'Day LY': String(wly), 'Day Budget': wbud === '' ? '' : String((Number(wbud) + (rnd() - 0.5) * 16).toFixed(1)),
      'WTD TY': String(wtd), 'WTD LY': String(wly), 'WTD Budget': String(wbud), 'MTD TY': String(mtd), 'MTD LY': String(ly), 'MTD Budget': String(bud) });
    return [
      m('Net Sales (Stores)', ns.toFixed(0), sb.toFixed(1), '', wk.toFixed(0), wb.toFixed(1)),
      m('Sales per Guest w. Cancellations', (430 + rnd() * 160).toFixed(2), '', spg.toFixed(1)),
      m('Close Rate', (24 + rnd() * 8 - bad * 4).toFixed(1), String(cr)),
      m('Traffic', Math.round(900 + rnd() * 700), '', tr.toFixed(1), Math.round(220 + rnd() * 150), '', tr.toFixed(1)),
      m('Sales per Hour', (320 + rnd() * 140 - bad * 60).toFixed(2)), m('Avg Ticket w. Del.', (1900 + rnd() * 600).toFixed(2)),
      m('Eff. Margin', (54 + rnd() * 4).toFixed(2)), m('Finance % of Sales', (48 + rnd() * 20 - bad * 6).toFixed(2)),
      m('Finance Apps to Traffic', (7 + rnd() * 6).toFixed(2)), m('Bedding % of Sales', (12 + rnd() * 10).toFixed(2)),
      m('Bedding SPH', (40 + rnd() * 30).toFixed(2)), m('Protection % of Sales', (6 + rnd() * 4).toFixed(2)),
      m('Protection SPH', (22 + rnd() * 18).toFixed(2)), m('Protection Attachment', (46 + rnd() * 22 - bad * 8).toFixed(2)),
      m('Delivery % of sales', (6 + rnd() * 3).toFixed(2)),
      m('Cancellations', (-ns * (0.03 + rnd() * 0.05 + bad * 0.03)).toFixed(0)), m('Gross Sales', (ns * 1.08).toFixed(0))
    ];
  });
  const sat = addDays(week, -1), yest = addDays(t, -1) > sat ? addDays(t, -1) : sat;
  const daily = {};
  const d1 = parseDaily(dailyRows(sat)); daily[sat] = { date: sat, periods: d1.periods, stores: d1.stores };
  let meta = { rsaDates: [], dailyDates: Object.keys(daily).sort().reverse(), latestDaily: yest, lastDaily: { by: OWNER_EMAIL, at: new Date().toISOString(), file: `daily-report-${yest}.csv`, date: yest } };

  // Consultants: six per demo store, made-up names.
  const first = ['Maria', 'Devon', 'Alyssa', 'Marcus', 'Priya', 'Tyler', 'Jasmine', 'Chris', 'Nina', 'Omar', 'Keisha', 'Luis', 'Grace', 'Andre', 'Tessa'];
  const last = ['Alvarez', 'Brooks', 'Chen', 'Dawson', 'Ellis', 'Foster', 'Grant', 'Hayes', 'Ibarra', 'Jordan', 'Kim', 'Lopez', 'Moss', 'Nolan'];
  const roster = [], rsaRows = [];
  leaders.flatMap(l => l.stores).forEach((store, si) => {
    for (let i = 0; i < 6; i++) {
      const n = si * 6 + i, name = `${first[n % 15]} ${last[(n * 5 + Math.floor(n / 15)) % 14]}`;
      const skill = i === 5 ? 0.52 : i === 4 ? 0.72 : 0.82 + rnd() * 0.7;
      const hours = 110 + rnd() * 40, sph = 400 * skill * (store.startsWith('Outlet') ? 0.55 : 1);
      const pc = (a, b) => (a + rnd() * (b - a)).toFixed(2) + '%';
      roster.push({ cid: cidOf(name), name, store, title: i === 0 ? 'ASM' : 'RSA', aliases: [] });
      rsaRows.push({ 'Sales Associate': name, 'Net Sales': (sph * hours).toFixed(2), 'Cancellation %': pc(1, 9), 'Discount %': pc(4, 17),
        'Credit Apps #': Math.round(18 * skill * (0.5 + rnd() * 0.6)), 'Eff. Margin': pc(53, 59), 'SPH': sph.toFixed(2),
        'Avg Ticket w. Del.': (1700 + rnd() * 1000).toFixed(2), 'Fin. % of Sales': pc(45, 78), 'Bed. % of Sales': pc(9, 28),
        'Prot. % of Sales': pc(4, 11), 'Del. % of Sales': pc(5, 10) });
    }
  });
  const pr = parseRsa(rsaRows);
  const monthStart = yest.slice(0, 8) + '01';
  let rsa = { from: monthStart, to: yest, file: `rsa_report_${monthStart}_to_${yest}.csv`, people: pr.people.map(p => ({ ...p, store: roster.find(r => r.cid === p.cid)?.store || null })) };
  // Saturday's copy of the RSA report. The third consultant in each store has a rough week.
  const rsaHist = { [yest]: rsa };
  if (yest !== sat && sat.slice(0, 7) === yest.slice(0, 7)) {
    rsaHist[sat] = { from: monthStart, to: sat, people: rsa.people.map((p, i) => {
      const wkHours = 18 + rnd() * 10, factor = i % 6 === 2 ? 0.45 : i % 6 === 3 ? 1.45 : 0.85 + rnd() * 0.3;
      const hours = Math.max(1, p.hours - wkHours), sales = p.k.netSales - wkHours * p.k.sph * factor;
      const k = { ...p.k, netSales: sales, sph: sales / hours };
      if (i % 6 === 2) { k.financePct = Math.min(95, (p.k.financePct || 50) * 1.25); k.protectionPct = (p.k.protectionPct || 6) * 1.3; }
      if (i % 6 === 3) { k.beddingPct = (p.k.beddingPct || 15) * 0.8; }
      return { cid: p.cid, name: p.name, store: p.store, hours, k };
    }) };
  }

  // This week's plans, built from Saturday's numbers, with the past days already visited.
  const plans = {}, visits = {}, photos = {};
  const timeoff = { [`nc@demo_${addDays(week, 7)}`]: { email: 'nc@demo', weekStart: addDays(week, 7), off: [2, 4], at: t, by: 'nc@demo' } };
  const lastVisits = { 'Yulee': addDays(week, -16), 'Town Center': addDays(week, -5), 'Orange Park': addDays(week, -9) };
  Object.entries(lastVisits).forEach(([store, date]) => {
    const id = `east@demo_${date}_${slug(store)}`;
    visits[id] = { id, email: 'east@demo', name: users['east@demo'].name, role: 'leader', store, date, kind: 'first', vtype: 'Priority', status: 'done',
      leaderWin: { name: 'Store leader', text: 'Ran a tight huddle and knew every number.' }, working: 'Huddle ran on time. Team is presenting finance early.',
      actions: [{ behavior: 'Leader walks every guest over 20 minutes before they leave', owner: 'Store leader' }, { behavior: 'Protection check at close every night', owner: 'Closing leader' }, {}],
      checks: { facilities: { 0: 'yes', 1: 'partial' } }, aor: { 'Bedroom': 'needs' }, notes: '', at: date };
  });
  { // A couple of sample photos on the Town Center visit so the visit log shows them.
    const vid = `east@demo_${lastVisits['Town Center']}_${slug('Town Center')}`;
    const svg = (c, t) => 'data:image/svg+xml;base64,' + btoa(`<svg xmlns="http://www.w3.org/2000/svg" width="400" height="300"><rect width="400" height="300" fill="${c}"/><text x="200" y="160" font-size="28" text-anchor="middle" fill="#fff" font-family="sans-serif">${t}</text></svg>`);
    [['#3F738D', 'Front entrance', 'Front Entrance and Windows'], ['#003B4A', 'Dining', 'Dining']].forEach(([c, t, item], i) => { const id = `${vid}_${i}`; photos[id] = { id, visitId: vid, email: 'east@demo', store: 'Town Center', date: lastVisits['Town Center'], el: 'visual', item, caption: i ? 'Two tables missing price tags.' : 'Windows clean, sign is current.', data: svg(c, t) }; });
  }
  { const d = addDays(week, 1), id = `east@demo_${d}_${slug('Yulee')}_remote`;
    visits[id] = { id, remote: true, email: 'east@demo', name: users['east@demo'].name, role: 'leader', store: 'Yulee', date: d, kind: 'remote', vtype: 'Video call', status: 'done',
      leaderWin: { name: 'Store leader', text: 'Had the numbers ready before the call.' }, actions: [{ what: 'Finance % of sales', from: '48%', to: '55%', how: 'Get every guest their buying power', owner: 'Store leader' }, {}, {}], at: d }; }
  for (const l of leaders) {
    const lv = latestVisitMap(Object.values(visits));
    const scores = Object.fromEntries(l.stores.map(s => [s, needScore(daily[sat].stores[s], { lastVisit: lv[s], today: week })]));
    const p = { email: l.email, name: l.name, ...buildPlan({ weekStart: week, stores: l.stores, scores, off: l.off, role: l.role }), builtAt: week + 'T08:00:00', basisDate: sat, pivots: [], dismissed: [] };
    p.days.forEach(d => {
      if (d.date < t && d.store) {
        d.status = 'done';
        const id = `${l.email}_${d.date}_${slug(d.store)}`;
        visits[id] = { id, email: l.email, name: l.name, role: l.role, store: d.store, date: d.date, kind: d.kind, vtype: kindToType(d.kind), status: 'done',
          leaderWin: { name: 'Store leader', text: 'Every guest greeted at the door within 10 seconds.' },
          working: 'Leaders are greeting at the door and taking every up in turn.',
          actions: [{ behavior: 'Every guest gets their buying power', owner: 'All consultants' }, { behavior: 'Leader turnover before any guest walks', owner: 'Leader on duty' }, {}],
          checks: { culture: { 0: 'yes', 1: 'partial', 2: 'yes', 3: 'no' }, facilities: { 0: 'yes', 1: 'yes', 2: 'partial', 3: 'yes', 4: 'yes' } },
          segs: { 0: { 0: 'yes', 1: 'partial', 2: 'no', 3: 'yes' } }, aor: { 'Front Entrance and Windows': 'pass', 'Dining': 'needs' },
          notes: 'Two RSAs coached on the monthly payment talk track.', at: d.date };
      }
    });
    plans[`${l.email}_${week}`] = p;
  }
  // Mid-week swing: a store the Jacksonville leader already visited fell apart since Sunday,
  // so the app suggests a pivot for the rest of the week.
  if (yest !== sat) {
    const ep = plans[`east@demo_${week}`];
    const hit = ep.days.filter(d => d.status === 'done').sort((a, b) => (ep.basis[a.store] ?? 0) - (ep.basis[b.store] ?? 0))[0];
    const d2 = parseDaily(dailyRows(yest, hit ? { [hit.store]: [-21, -32, -15, -420, -5] } : {}));
    daily[yest] = { date: yest, periods: d2.periods, stores: d2.stores };
  }
  const dayBefore = addDays(yest, -1);
  if (dayBefore > sat && dayBefore.slice(0, 7) === yest.slice(0, 7)) {
    rsaHist[dayBefore] = { from: monthStart, to: dayBefore, people: rsa.people.map((p, i) => {
      const dh = 6 + rnd() * 3, f = i % 6 === 2 ? 0.3 : i % 6 === 3 ? 1.8 : 0.7 + rnd() * 0.6;
      const hours = Math.max(1, p.hours - dh), sales = p.k.netSales - dh * p.k.sph * f;
      return { cid: p.cid, name: p.name, store: p.store, hours, k: { netSales: sales, sph: sales / hours } };
    }) };
  }
  meta.rsaDates = Object.keys(rsaHist).sort().reverse();
  let storeLeaders = [];
  // Demo open carts: a few per consultant.
  let carts = { date: yest, people: {}, stores: {}, total: 0, n: 0 };
  rsa.people.forEach((p, i) => { const k = 2 + (i % 6) * 2, top = Array.from({ length: Math.min(k, 5) }, (_, j) => ({ g: ['Maria L.', 'James P.', 'Tasha W.', 'Kevin R.', 'Ana S.'][j], v: 1200 + ((i + j) % 5) * 900, a: [1, 3, 6, 9, 20][j], l: 2 + j })); const value = top.reduce((a, x) => a + x.v, 0) * k / top.length;
    carts.people[p.cid] = { name: p.name, store: p.store, n: k, value: Math.round(value), wk: Math.ceil(k / 2), wkValue: Math.round(value / 2), due: 2, old: i % 3, top };
    const sx = carts.stores[p.store] ||= { n: 0, value: 0, wk: 0, wkValue: 0, due: 0, old: 0 }; sx.n += k; sx.value += Math.round(value); sx.wk += Math.ceil(k / 2); sx.wkValue += Math.round(value / 2); sx.due += 2; sx.old += i % 3; carts.n += k; carts.total += Math.round(value); });
  // Demo daily budget for this month and last: a flat monthly number spread by weekday.
  const budgets = {};
  [yest.slice(0, 7), t.slice(0, 7)].forEach(mo => {
    if (budgets[mo]) return;
    const [yy, mm] = mo.split('-').map(Number), nd = new Date(yy, mm, 0).getDate(), w = [1.1, 0.75, 0.75, 0.8, 0.8, 1, 1.85];
    const stores = {};
    STORES.forEach((st, i) => { const monthly = 300000 + (i % 7) * 60000, traffic = Math.round(monthly / 560); const tw = Array.from({ length: nd }, (_, k) => w[new Date(yy, mm - 1, k + 1).getDay()]); const sw = tw.reduce((a, b) => a + b, 0);
      stores[st.name] = { m: { traffic, sales: monthly, spg: monthly / traffic }, d: Object.fromEntries(tw.map((x, k) => [`${mo}-${String(k + 1).padStart(2, '0')}`, [Math.round(traffic * x / sw * 10) / 10, Math.round(monthly * x / sw)]])) }; });
    budgets[mo] = { month: mo, stores };
  });
  let markets = [
    { id: 'jax', name: 'Jacksonville', leader: 'east@demo', director: 'director@demo', stores: ['Town Center', 'Orange Park', 'Yulee', 'St. Augustine'] },
    { id: 'nc', name: 'Carolinas', leader: 'nc@demo', director: '', stores: ['Greensboro', 'Winston Salem', 'Burlington', 'Danville', 'Outlet Greensboro'] },
    { id: 'gulf', name: 'Gulf Coast', leader: 'gulf@demo', director: '', stores: ['Mobile', "D'Iberville", 'Spanish Fort', 'Pensacola', 'Crestview', 'Ft. Walton'] }
  ];
  // Last week, for the 1 on 1: a plan and visits for each leader, the RSA copy from the Saturday
  // before it, and the Jacksonville leader's 1 on 1 from two weeks back with commitments to review.
  const lw = addDays(week, -7), sat2 = addDays(lw, -1);
  if (rsaHist[sat] && sat2.slice(0, 7) === sat.slice(0, 7)) {
    rsaHist[sat2] = { from: monthStart, to: sat2, people: rsaHist[sat].people.map((p, i) => {
      const h = 24 + rnd() * 10, f = i % 6 === 1 ? 1.55 : i % 6 === 5 ? 0.35 : i % 6 === 4 ? 0.6 : 0.8 + rnd() * 0.4;
      const hours = Math.max(1, p.hours - h), sales = Math.max(0, p.k.netSales - h * p.k.sph * f);
      return { cid: p.cid, name: p.name, store: p.store, hours, k: { netSales: sales, sph: sales / hours } };
    }) };
  }
  for (const l of leaders) {
    const scores = Object.fromEntries(l.stores.map(s => [s, needScore(daily[sat].stores[s], { today: lw })]));
    const p = { email: l.email, name: l.name, ...buildPlan({ weekStart: lw, stores: l.stores, scores, off: l.off, role: l.role }), builtAt: lw + 'T08:00:00', basisDate: addDays(lw, -1), pivots: [], dismissed: [] };
    let skipped = false;
    p.days.forEach(d => {
      if (!d.store) return;
      if (l.email === 'east@demo' && !skipped && d.kind !== 'first') { skipped = true; return; }
      d.status = 'done';
      const id = `${l.email}_${d.date}_${slug(d.store)}`;
      if (visits[id]) return;
      const team = rsa.people.filter(x => x.store === d.store).slice(0, 2);
      visits[id] = { id, email: l.email, name: l.name, role: l.role, store: d.store, date: d.date, kind: d.kind, vtype: kindToType(d.kind), status: 'done',
        leaderWin: { name: 'Store leader', text: 'Huddle ran on time with every number posted.' },
        consultants: team.map((x, j) => ({ cid: x.cid, name: x.name, notes: 'Worked the monthly payment talk track.', drill: j === 0 ? { key: 'financePct', scored: { 0: 'yes', 1: 'partial', 2: 'yes', 3: 'no' } } : {} })),
        actions: [{ what: 'Finance % of sales', key: 'financePct', from: '48%', to: '55%', how: 'Get every guest their buying power', owner: 'Store leader', due: addDays(lw, 6) }, {}, {}], at: d.date };
    });
    plans[`${l.email}_${lw}`] = p;
  }
  // Year to date through the end of last month, and last month's final copy. Every person's year
  // runs a little different from this month so the trend columns have something to say.
  const lmEnd = addDays(monthStart, -1), lmStart = lmEnd.slice(0, 8) + '01';
  const tilt = (p, i, f) => { const k = { ...p.k }; ['financePct', 'protectionPct', 'beddingPct', 'deliveryPct'].forEach((m, j) => { if (k[m] != null) k[m] = Math.round(k[m] * (((i + j) % 3 === 0) ? 1 + f : ((i + j) % 3 === 1) ? 1 - f : 1) * 100) / 100; }); return k; };
  rsaHist[lmEnd] = { from: lmStart, to: lmEnd, file: `rsa_report_${lmStart}_to_${lmEnd}.csv`, people: rsa.people.map((p, i) => { const hours = 150 + (i % 5) * 8, k = tilt(p, i, 0.15); k.sph = p.k.sph * (i % 4 === 0 ? 0.8 : 1.05); k.netSales = k.sph * hours; return { cid: p.cid, name: p.name, store: p.store, hours, k }; }) };
  let rsaYtd = lmEnd.slice(5, 7) === '12' ? null : { from: lmEnd.slice(0, 4) + '-01-01', to: lmEnd, file: `rsa_report_${lmEnd.slice(0, 4)}-01-01_to_${lmEnd}.csv`, people: rsa.people.map((p, i) => { const hours = 1300 + (i % 7) * 40, k = tilt(p, i, 0.3); k.sph = p.k.sph * (i % 4 === 1 ? 1.2 : 0.9); k.netSales = k.sph * hours; k.creditApps = Math.round((p.k.creditApps || 10) * 8); return { cid: p.cid, name: p.name, store: p.store, hours, k }; }) };
  meta.rsaDates = Object.keys(rsaHist).sort().reverse();
  const ones = {};
  let offerDoc = null;
  const alerts = { demo1: { id: 'demo1', type: 'schedule', email: 'nc@demo', name: users['nc@demo'].name, weekStart: week, at: addDays(t, -1) + 'T07:40:00', seen: false,
    reason: 'Called into a store (issue or emergency)', note: 'Two closers out at Danville. Covering the floor Tuesday, moved Winston Salem to Friday.',
    changes: [{ date: addDays(week, 2), from: 'Winston Salem', to: 'Danville' }, { date: addDays(week, 5), from: 'Burlington', to: 'Winston Salem' }] } };
  ones[`east@demo_${addDays(lw, -7)}`] = { id: `east@demo_${addDays(lw, -7)}`, email: 'east@demo', name: users['east@demo'].name, weekStart: addDays(lw, -7), status: 'done', heldAt: addDays(lw, 1) + 'T09:00:00', heldBy: OWNER_EMAIL, heldByName: 'Frank Pina',
    coaching: '', actions: [
      { key: 'closeRate', what: 'Close Rate across the market', from: '24%', to: '28%', how: 'No guest leaves without a TO. Leaders track TOs at every huddle.', owner: users['east@demo'].name, due: addDays(lw, 6) },
      { store: 'Yulee', key: 'financePct', what: 'Yulee: Finance %', from: '44%', to: '55%', how: 'Full-day visit Tuesday. Every guest gets their buying power.', owner: users['east@demo'].name, due: addDays(lw, 6) },
      { what: 'Full-day visits', from: '3 of 5', to: '5 of 5', how: 'Days off set by Sunday.', owner: users['east@demo'].name, due: addDays(lw, 6) }], support: 'Help backfill a closing leader at Yulee', supportBy: addDays(lw, 3) };
  let current = 'east@demo';
  const clone = x => structuredClone(x);
  return {
    demo: true,
    demoUsers: () => Object.values(users).filter(u => u.email in users),
    setDemoUser(e) { current = e; this._cb?.({ email: e, verified: true }); },
    onAuth(cb) { this._cb = cb; setTimeout(() => cb({ email: current, verified: true }), 0); },
    signIn: async () => {}, register: async () => {}, resendVerify: async () => {}, refresh: async () => true, reset: async () => {}, signOut: async () => {},
    profile: async () => clone(users[current]),
    meta: async () => clone(meta),
    daily: async d => clone(daily[d] || null),
    async publishDaily(d) { daily[d.date] = { date: d.date, periods: d.periods, stores: d.stores }; const dates = [...new Set([...meta.dailyDates, d.date])].sort().reverse(); meta = { ...meta, dailyDates: dates, latestDaily: dates[0], lastDaily: { by: current, at: new Date().toISOString(), file: d.file, date: d.date } }; },
    rsa: async () => clone(rsa),
    rsaAt: async d => clone(rsaHist[d] || null),
    rsaYtd: async () => clone(rsaYtd),
    async publishRsaYtd(r) { rsaYtd = clone(r); meta.lastRsaYtd = { by: current, at: new Date().toISOString(), file: r.file, from: r.from, to: r.to, people: r.people.length }; },
    async publishRsa(r) { rsaHist[r.to] = clone(r); meta.rsaDates = Object.keys(rsaHist).sort().reverse(); if (!meta.lastRsa?.to || r.to >= meta.lastRsa.to) { rsa = clone(r); meta.lastRsa = { by: current, at: new Date().toISOString(), file: r.file, to: r.to, people: r.people.length }; } },
    roster: async () => clone(roster),
    markets: async () => clone(markets),
    storeLeaders: async () => clone(storeLeaders),
    async saveStoreLeaders(l) { storeLeaders = clone(l); },
    async saveMarkets(m) { markets = clone(m); },
    async saveRoster(p) { roster.splice(0, roster.length, ...clone(p)); },
    budget: async m => clone(budgets[m] || null),
    carts: async () => clone(carts),
    async saveCarts(c) { carts = clone(c); },
    async saveBudget(b) { budgets[b.month] = clone(b); },
    users: async () => clone(Object.values(users)),
    async saveUser(u) { users[u.email] = clone(u); },
    async deleteUser(e) { delete users[e]; },
    plan: async (e, w) => clone(plans[`${e}_${w}`] || null),
    plansForWeek: async w => clone(Object.values(plans).filter(p => p.weekStart === w)),
    async savePlan(p) { plans[`${p.email}_${p.weekStart}`] = clone(p); },
    visits: async () => clone(Object.values(visits)),
    async saveVisit(v) { visits[v.id] = clone(v); },
    photos: async id => clone(Object.values(photos).filter(p => p.visitId === id)),
    async savePhoto(ph) { photos[ph.id] = clone(ph); },
    async deletePhoto(id) { delete photos[id]; },
    timeOff: async (e, w) => clone(timeoff[`${e}_${w}`] || null),
    timeOffForWeek: async w => clone(Object.values(timeoff).filter(x => x.weekStart === w)),
    async saveTimeOff(x) { timeoff[`${x.email}_${x.weekStart}`] = clone(x); },
    oneOnOne: async id => clone(ones[id] || null),
    async saveOneOnOne(d) { ones[d.id] = clone(d); },
    onePrivate: async id => clone(ones['p_' + id] || null),
    async saveOnePrivate(d) { ones['p_' + d.id] = clone(d); },
    alerts: async () => clone(Object.values(alerts)),
    offer: async () => clone(offerDoc),
    async saveOffer(o) { offerDoc = clone(o); },
    async saveAlert(a) { alerts[a.id] = clone(a); }
  };
}
function latestVisitMap(visits) {
  const out = {};
  visits.forEach(v => { if (v.remote) return; if (!out[v.store] || v.date > out[v.store]) out[v.store] = v.date; });
  return out;
}

// ---------------------------------------------------------------- boot
async function boot() {
  try { S.be = DEMO ? demoBackend() : await firebaseBackend(); }
  catch (e) { $('#app').innerHTML = `<div class="panel narrow"><h2>Could not load</h2><p>${esc(e.message)}</p></div>`; return; }
  if (DEMO) {
    $('#demoBar').hidden = false;
    const sel = $('#demoRole');
    sel.innerHTML = S.be.demoUsers().map(u => `<option value="${esc(u.email)}">${esc(roleLabel(u.role))}: ${esc(u.name)}</option>`).join('');
    sel.value = 'east@demo';
    sel.onchange = () => { S.tab = null; S.viewEmail = null; S.visit = null; S.be.setDemoUser(sel.value); };
  }
  S.be.onAuth(async u => {
    if (!u) return renderSignIn();
    if (!u.verified) return renderVerify(u.email);
    try { S.user = await S.be.profile(); } catch (e) { S.user = null; }
    if (!S.user) return renderNotRostered(u.email);
    await loadShared();
    renderShell();
  });
}
// Field team Frank named. They're added as logins the first time an admin opens the app, so they show
// in the Market Leader and Director dropdowns right away. Stores come from the markets they're put on.
const FIELD_TEAM = [
  { email: 'ocruz@1915south.com', name: 'Orlando Cruz', role: 'director' },
  { email: 'jmccord@1915south.com', name: 'Jourdain McCord', role: 'leader', pilots: ['frontlineiq'] },
  { email: 'msevert@1915south.com', name: 'Meagan Severt', role: 'leader' },
  { email: 'ccarritz@1915south.com', name: 'Cole Carritz', role: 'leader' },
  { email: 'jkeene@1915south.com', name: 'Jonathan Keene', role: 'leader' }
];
async function seedFieldTeam() {
  if (S.user?.role !== 'admin') return;
  for (const t of FIELD_TEAM) {
    const have = S.users.find(u => u.email === t.email);
    // Pilots are set once; after that, Setup controls them.
    if (have && t.pilots && have.pilots === undefined) { have.pilots = t.pilots; try { await S.be.saveUser(have); } catch (e) {} }
    if (have) continue;
    const u = { ...t, stores: [], off: DEFAULT_OFF };
    try { await S.be.saveUser(u); S.users.push(u); } catch (e) { console.warn('Could not add', t.email, e); }
  }
}
// Weeks for consultant numbers run Monday to Sunday.
const mondayOf = d => addDays(d, -((dow(d) + 6) % 7));
const isMtdCopy = r => !!r && (!r.from || !r.to || r.from === r.to.slice(0, 8) + '01');
// Year to date, last month, this month, this week and yesterday for every consultant, from the kept RSA copies.
// Periods follow today's date: on October 1 there is no October yet, and September is last month.
async function loadTrends(rsa) {
  if (!rsa?.to) return {};
  const t = today(), cur = t.slice(0, 7);
  const dates = S.meta.rsaDates || [];
  const cache = {}, at = d => d ? (cache[d] ||= (d === rsa.to ? Promise.resolve(rsa) : S.be.rsaAt(d).catch(() => null)).then(r => isMtdCopy(r) ? r : null)) : Promise.resolve(null);
  const finalOf = async month => { for (const d of dates.filter(x => x.slice(0, 7) === month).sort().reverse()) { const r = await at(d); if (r) return r; } return null; };
  const onOrBefore = async (day) => { for (const d of dates.filter(x => x <= day && x.slice(0, 7) === day.slice(0, 7)).sort().reverse()) { const r = await at(d); if (r) return r; } return null; };
  const prevMonth = m => { const [y, mo] = m.split('-').map(Number); return mo === 1 ? `${y - 1}-12` : `${y}-${String(mo - 1).padStart(2, '0')}`; };
  const monthEnd = m => { const [y, mo] = m.split('-').map(Number); return iso(new Date(y, mo, 0)); };
  const mtd = rsa.to.slice(0, 7) === cur ? rsa : null;
  const lastMonth = await finalOf(prevMonth(cur));
  // This week: everything after last Sunday. A week can cross into a new month.
  const sun = addDays(mondayOf(t), -1);
  const wk = [];
  if (rsa.to > sun) {
    if (sun.slice(0, 7) === rsa.to.slice(0, 7)) { const b = await onOrBefore(sun); if (b) wk.push({ plus: rsa, minus: b }); }
    else {
      // Week started last month: (last month's final copy minus last Sunday's) plus this month so far.
      let ok = true;
      if (sun !== monthEnd(sun.slice(0, 7))) { const e = await finalOf(sun.slice(0, 7)), b = await onOrBefore(sun); if (e && b) { if (e.to > b.to) wk.push({ plus: e, minus: b }); } else ok = false; }
      if (ok) wk.push({ plus: rsa, minus: null }); else wk.length = 0;
    }
  }
  // Yesterday (the last day in the newest copy).
  const day = [];
  if (rsa.to.endsWith('-01')) day.push({ plus: rsa, minus: null });
  else { const b = await at(dates.includes(addDays(rsa.to, -1)) ? addDays(rsa.to, -1) : null); if (b) day.push({ plus: rsa, minus: b }); }
  let ytd = await (S.be.rsaYtd ? S.be.rsaYtd().catch(() => null) : null);
  if (ytd && ytd.to?.slice(0, 4) !== cur.slice(0, 4) && ytd.to?.slice(0, 4) !== rsa.to.slice(0, 4)) ytd = null;
  const ytdSnap = ytd && mtd && ytd.to < mtd.to && ytd.to.slice(0, 7) === cur ? await at(dates.includes(ytd.to) ? ytd.to : null) : null;
  const [y, mo] = cur.split('-').map(Number);
  const monthEnds = ytd ? [] : (await Promise.all(Array.from({ length: mo - 1 }, (_, i) => finalOf(`${y}-${String(i + 1).padStart(2, '0')}`)))).filter(Boolean);
  S.trendInfo = { cur, lastMonth: lastMonth?.to || prevMonth(cur) + '-01', week: { from: mondayOf(t), to: rsa.to }, day: rsa.to, wkOk: wk.length > 0,
    ytd: ytd ? { from: ytd.from, to: ytd.to, uploaded: true } : monthEnds.length ? { from: monthEnds.map(r => r.from || r.to).sort()[0], to: rsa.to, uploaded: false } : null };
  return consultantTrends({ mtd, lastMonth, ytd, ytdSnap, monthEnds, wk, day });
}
async function loadShared() {
  $('#app').innerHTML = '<p class="loading">Loading the latest numbers…</p>';
  S.meta = await S.be.meta();
  const [daily, rsa, users, visits, roster, markets, storeLeaders] = await Promise.all([
    S.meta.latestDaily ? S.be.daily(S.meta.latestDaily) : null, S.be.rsa(), S.be.users().catch(() => [S.user]), S.be.visits(), S.be.roster(), S.be.markets().catch(() => []), S.be.storeLeaders().catch(() => [])
  ]);
  Object.assign(S, { daily, rsa, users, visits, roster, markets, storeLeaders });
  // Each store's numbers carry the store name, so store goals (like the outlet ticket goal) apply.
  Object.entries(daily?.stores || {}).forEach(([name, snap]) => Object.values(snap || {}).forEach(per => { if (per && typeof per === 'object' && per.k) per.store = name; }));
  S.alerts = seesAll() ? await S.be.alerts().catch(() => []) : [];
  S.offer = (await S.be.offer().catch(() => null)) || OFFER_DEFAULT;
  await seedFieldTeam();
  // Consultant week: compare today's RSA upload with the one through last Saturday.
  const sat = addDays(mondayOf(today()), -1);
  const baseDate = (S.meta.rsaDates || []).filter(d => d <= sat && (!rsa?.to || d < rsa.to)).sort().reverse()[0];
  S.rsaBase = baseDate ? await S.be.rsaAt(baseDate).then(r => isMtdCopy(r) ? r : null).catch(() => null) : null;
  S.weeks = consultantWeeks(rsa, S.rsaBase);
  // Yesterday by consultant: today's RSA upload minus the one before it.
  const prevDate = (S.meta.rsaDates || []).filter(d => rsa?.to && d < rsa.to).sort().reverse()[0];
  S.rsaPrev = prevDate ? await S.be.rsaAt(prevDate).catch(() => null) : null;
  S.days = S.rsaPrev ? consultantWeeks(rsa, S.rsaPrev) : {};
  { const t0 = today(), months = [...new Set([t0.slice(0, 7), addDays(t0, -1).slice(0, 7), (S.meta.latestDaily || t0).slice(0, 7)])];
    const got = await Promise.all(months.map(m => S.be.budget ? S.be.budget(m).catch(() => null) : null));
    S.budgets = Object.fromEntries(months.map((m, i) => [m, got[i]]).filter(x => x[1])); }
  S.carts = await (S.be.carts ? S.be.carts().catch(() => null) : null);
  S.trends = await loadTrends(rsa).catch(e => { console.warn('trends', e); return {}; });
  S.teams = {};
  for (const st of STORES) if ((rsa?.people || []).some(p => p.store === st.name)) S.teams[st.name] = teamSignals(rsa.people, S.weeks, st.name, DEFAULT_GOALS);
  S.lastVisit = latestVisitMap(visits);
  S.scores = scoresFor(daily);
}
// Priority = need, weighted by revenue. A big store a few points behind is more dollars than a small
// store far behind, so volume moves a store up or down: a store doing twice the company average gets
// about 1.7x its need, half the average gets about 0.6x (capped at 0.5x to 1.75x).
const VOL_POW = 0.75, VOL_MIN = 0.5, VOL_MAX = 1.75;
function scoresFor(daily) {
  const out = {};
  if (!daily) return out;
  const sales = STORES.map(st => daily.stores[st.name]?.mtd?.k?.netSales).filter(v => v > 0);
  const avg = sales.length ? sales.reduce((a, b) => a + b, 0) / sales.length : null;
  const ranked = STORES.filter(st => daily.stores[st.name]?.mtd?.k?.netSales > 0).sort((a, b) => daily.stores[b.name].mtd.k.netSales - daily.stores[a.name].mtd.k.netSales).map(st => st.name);
  const money = n => '$' + Math.round(n).toLocaleString('en-US');
  for (const st of STORES) {
    const snap = daily.stores[st.name]; if (!snap) continue;
    const raw = needScore(snap, { lastVisit: S.lastVisit[st.name], today: today(), team: S.teams?.[st.name] });
    const m = snap.mtd, ns = m?.k?.netSales, bud = m?.budget?.netSales;
    const factor = avg && ns > 0 ? Math.max(VOL_MIN, Math.min(VOL_MAX, Math.pow(ns / avg, VOL_POW))) : 1;
    const behind = ns != null && bud != null && bud > ns ? bud - ns : 0;
    const lead = [];
    if (behind > 0) lead.push({ key: 'dollars', pts: 0, text: `${money(behind)} behind budget this month` });
    if (ns > 0 && factor >= 1.1) lead.push({ key: 'volume', pts: 0, text: `High-volume store: ${money(ns)} this month, #${ranked.indexOf(st.name) + 1} of ${ranked.length}` });
    out[st.name] = { score: Math.min(100, Math.round(raw.score * factor)), need: raw.score, factor, behind, sales: ns ?? null, parts: [...lead, ...raw.parts] };
  }
  return out;
}
const isAdmin = () => S.user?.role === 'admin';
const hasFliq = u => (u?.pilots || []).includes('frontlineiq');
const activeOffer = () => offerActive(S.offer, today()) ? S.offer : null;
const seesAll = () => ['admin', 'exec'].includes(S.user?.role);
// Field leaders = Market Leaders and directors. Each store has one Market Leader; directors can overlap.
const leaders = () => S.users.filter(u => FIELD.includes(u.role) && (u.stores || []).length).sort((a, b) => (a.role === 'leader' ? 0 : 1) - (b.role === 'leader' ? 0 : 1) || (a.name || a.email).localeCompare(b.name || b.email));
const leaderOf = store => S.users.find(l => l.role === 'leader' && (l.stores || []).includes(store));
const marketOf = store => (S.markets || []).find(m => (m.stores || []).includes(store));
const marketsOf = email => (S.markets || []).filter(m => m.leader === email || m.director === email);
const whoLabel = l => `${l.name || l.email}${l.role === 'director' ? ' (Director)' : ''}`;

// ---------------------------------------------------------------- auth screens
function renderSignIn(msg = '') {
  $('#who').innerHTML = '';
  $('#app').innerHTML = `
  <form class="panel narrow" id="signin">
    <h2>Sign in</h2>
    <p>Use your @${esc(EMAIL_DOMAIN)} email. First time here? Enter your email, choose a password, and tap Create account. You will get a verification email.</p>
    <label for="em">Email<input type="email" id="em" autocomplete="username" required></label>
    <label for="pw">Password<input type="password" id="pw" autocomplete="current-password" minlength="8" required></label>
    ${msg ? `<p class="err">${esc(msg)}</p>` : ''}
    <div class="row">
      <button class="btn primary" type="submit">Sign in</button>
      <button class="btn" type="button" id="reg">Create account</button>
      <button class="link" type="button" id="forgot">Forgot password</button>
    </div>
  </form>`;
  const em = () => $('#em').value.trim().toLowerCase(), pw = () => $('#pw').value;
  $('#signin').onsubmit = async e => { e.preventDefault(); try { await S.be.signIn(em(), pw()); } catch (x) { renderSignIn(friendly(x)); } };
  $('#reg').onclick = async () => {
    if (!em().endsWith('@' + EMAIL_DOMAIN)) return renderSignIn(`Use your @${EMAIL_DOMAIN} email to create an account.`);
    if (pw().length < 8) return renderSignIn('Choose a password of at least 8 characters.');
    try { await S.be.register(em(), pw()); } catch (x) { renderSignIn(friendly(x)); }
  };
  $('#forgot').onclick = async () => {
    if (!em()) return renderSignIn('Type your email first, then tap Forgot password.');
    try { await S.be.reset(em()); toast('Password reset email sent.'); } catch (x) { renderSignIn(friendly(x)); }
  };
}
function friendly(x) {
  const c = x?.code || '';
  if (c.includes('invalid-credential') || c.includes('wrong-password') || c.includes('user-not-found')) return 'Email or password is not right. New here? Tap Create account.';
  if (c.includes('email-already-in-use')) return 'That email already has an account. Sign in, or tap Forgot password.';
  if (c.includes('too-many-requests')) return 'Too many tries. Wait a few minutes and try again.';
  if (c.includes('permission-denied')) return 'You do not have access to that. Ask Frank to check your login.';
  return x?.message || 'Something went wrong.';
}
function renderVerify(email) {
  $('#who').innerHTML = signOutBtn(); wireSignOut();
  $('#app').innerHTML = `<div class="panel narrow"><h2>Check your email</h2>
    <p>We sent a verification link to <b>${esc(email)}</b>. Open it, then come back and tap Continue. Check junk mail if you do not see it.</p>
    <div class="row"><button class="btn primary" id="cont">Continue</button><button class="btn" id="again">Send it again</button></div></div>`;
  $('#cont').onclick = async () => { if (await S.be.refresh()) location.reload(); else toast('Not verified yet. Open the link in the email first.', true); };
  $('#again').onclick = async () => { try { await S.be.resendVerify(); toast('Sent.'); } catch (x) { toast(friendly(x), true); } };
}
function renderNotRostered(email) {
  $('#who').innerHTML = signOutBtn(); wireSignOut();
  $('#app').innerHTML = `<div class="panel narrow"><h2>You are signed in, but not set up yet</h2>
    <p><b>${esc(email)}</b> does not have stores assigned. Ask Frank Pina to add you, then refresh this page.</p></div>`;
}
const signOutBtn = () => `<button class="link light" id="so">Sign out</button>`;
function wireSignOut() { const b = $('#so'); if (b) b.onclick = () => S.be.signOut(); }

// ---------------------------------------------------------------- shell
function renderShell() {
  const u = S.user;
  $('#who').innerHTML = `<span>${esc(u.name || u.email)} <small>${esc(roleLabel(u.role))}</small></span>${DEMO ? '' : signOutBtn()}`;
  wireSignOut();
  const tabs = [];
  tabs.push(['brief', 'Daily brief']);
  if (seesAll()) tabs.push(['leaders', 'Leaders']);
  tabs.push(['one', seesAll() ? '1 on 1s' : 'My 1 on 1']);
  tabs.push(['week', seesAll() ? 'Weekly plans' : 'My week'], ['stores', seesAll() ? 'Stores' : 'My stores'], ['messages', 'Team messages'], ['visits', 'Visit log']);
  if (isAdmin()) tabs.push(['upload', 'Upload'], ['setup', 'Setup']);
  tabs.push(['guide', 'How it works']);
  if (!tabs.some(t => t[0] === S.tab)) S.tab = tabs[0][0];
  $('#app').innerHTML = `
    <nav class="tabs" aria-label="Sections">${tabs.map(([k, l]) => `<button data-tab="${k}" class="${S.tab === k ? 'on' : ''}">${l}</button>`).join('')}</nav>
    ${alertBar()}
    <div id="view"></div>`;
  const ab = $('#alertgo'); if (ab) ab.onclick = () => { S.tab = 'brief'; S.visit = null; renderShell(); setTimeout(() => $('#alerts')?.scrollIntoView({ behavior: 'smooth' }), 300); };
  document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => { S.tab = b.dataset.tab; S.visit = null; renderShell(); });
  if (S.visit) return viewVisit();
  ({ brief: viewBrief, one: viewOne, leaders: viewLeaders, week: viewWeek, stores: viewStores, messages: viewMessages, visits: viewVisits, upload: viewUpload, setup: viewSetup, guide: viewGuide })[S.tab]();
}
const dataLine = () => {
  const d = S.meta.latestDaily;
  if (!d) return `<span class="warn">No daily report uploaded yet.</span>`;
  const age = daysApart(d, today());
  return `Numbers through <b>${esc(longDate(d))}</b>${age > 2 ? ` <span class="warn">(${age} days old)</span>` : ''}${S.rsa?.to ? ` · Consultants through ${esc(shortDate(S.rsa.to))}` : ''}`;
};
const needChip = s => s == null ? '<span class="need low">--</span>' : `<span class="need ${band(s)}" title="Priority: need weighted by revenue, 0 to 100">${s}</span>`;


// ---------------------------------------------------------------- schedule change alerts
// When a Market Leader or director changes a schedule that is already set, they give a reason and
// the VP (every admin) gets an alert in the app. Admin changes are not flagged.
const CHANGE_REASONS = ['Called into a store (issue or emergency)', 'Store numbers changed', 'Training, meeting or company event', 'Time off or personal', 'Travel or weather', 'Other'];
const vpNames = () => { const a = S.users.filter(u => u.role === 'admin'); return a.length ? a.map(u => firstOf(u.name) || u.email).join(' and ') : 'your VP'; };
const needsReason = () => !isAdmin();
function alertBar() {
  if (!isAdmin()) return '';
  const n = (S.alerts || []).filter(a => !a.seen).length;
  return n ? `<div class="warnbox" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:0 0 14px"><b>${n} schedule change${n > 1 ? 's' : ''} to review.</b><button class="btn tiny" id="alertgo" type="button">Review</button></div>` : '';
}
// Day by day, what moved: store to store, store to off, off to store.
function scheduleDiff(beforeDays, afterDays) {
  const b = Object.fromEntries((beforeDays || []).map(d => [d.date, dayText(d)]));
  const a = Object.fromEntries((afterDays || []).map(d => [d.date, dayText(d)]));
  const dates = [...new Set([...Object.keys(b), ...Object.keys(a)])].sort();
  return dates.map(date => ({ date, from: b[date] || 'Off', to: a[date] || 'Off' })).filter(x => x.from !== x.to);
}
// Inline form (no pop-up): pick a reason, add details, then the change saves and the alert goes out.
function askReason(boxSel, title, onOk, onCancel) {
  const box = $(boxSel); if (!box) return;
  box.innerHTML = `<div class="reasonp">
    <h3 style="margin:0 0 4px">${esc(title)}</h3>
    <p class="small" style="margin:0 0 10px">${esc(vpNames())} gets an alert with your reason and what moved.</p>
    <label for="rsn">Reason<select id="rsn"><option value="">Pick one</option>${CHANGE_REASONS.map(r => `<option>${esc(r)}</option>`).join('')}</select></label>
    ${fieldBox('rsnote', 'Details', '', 2, 'What happened, and how the stores you moved will get covered.')}
    <div class="row" style="margin-top:10px"><button class="btn primary" id="rsok" type="button">Save and alert ${esc(vpNames())}</button><button class="link" id="rscancel" type="button">Cancel</button></div></div>`;
  box.hidden = false; wireMics(box); box.scrollIntoView({ behavior: 'smooth', block: 'center' });
  $('#rscancel').onclick = () => { box.innerHTML = ''; box.hidden = true; (onCancel || viewWeek)(); };
  $('#rsok').onclick = async () => {
    const reason = $('#rsn').value, note = $('#rsnote').value.trim();
    if (!reason) return toast('Pick a reason first.', true);
    if (reason === 'Other' && note.length < 5) return toast('Add a few words on why.', true);
    if (note.length < 5 && !confirmNoNote(box)) return;
    $('#rsok').disabled = true;
    try { await onOk(reason, note); } catch (e) { $('#rsok').disabled = false; toast(friendly(e), true); }
  };
}
// One nudge for details before saving without any.
function confirmNoNote(box) { if (box.dataset.nudged) return true; box.dataset.nudged = '1'; toast('Add details so your VP knows what happened, or tap Save again to send without them.'); return false; }
async function sendScheduleAlert(who, week, reason, note, changes, kind) {
  if (!needsReason()) return null;
  const a = { id: `${who.email}_${Date.now()}`, type: 'schedule', kind, email: S.user.email, name: S.user.name || S.user.email, forEmail: who.email, weekStart: week,
    reason, note, changes, at: new Date().toISOString(), seen: false };
  await S.be.saveAlert(a);
  S.lastAlert = a;
  return a;
}
function alertMailto(a) {
  const to = S.users.filter(u => u.role === 'admin').map(u => u.email).join(',');
  const body = [`Schedule change for the week of ${shortDate(a.weekStart)}.`, '', `Reason: ${a.reason}`, a.note ? `Details: ${a.note}` : '', '', ...a.changes.map(c => `${dayLabel(c.date)}: ${c.from} to ${c.to}`), '', firstOf(a.name)].filter(x => x !== null).join('\n');
  return `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(`Schedule change: ${a.name}, week of ${shortDate(a.weekStart)}`)}&body=${encodeURIComponent(body)}`;
}
const changeLines = a => (a.changes || []).map(c => `${dayLabel(c.date)}: ${c.from} to ${c.to}`);

// ---------------------------------------------------------------- the week
async function viewWeek() {
  const v = $('#view');
  const week = S.week || (S.week = weekStartOf(today()));
  const thisWeek = weekStartOf(today());
  if (seesAll() && !S.viewEmail) S.viewEmail = leaders()[0]?.email || null;
  const email = seesAll() ? S.viewEmail : S.user.email;
  const who = S.users.find(u => u.email === email) || (email === S.user.email ? S.user : null);
  const picker = seesAll() ? `<label for="lp" style="margin:0">Field leader<select id="lp">${leaders().map(l => `<option value="${esc(l.email)}" ${l.email === email ? 'selected' : ''}>${esc(whoLabel(l))}</option>`).join('')}</select></label>` : '';
  if (!who || !(who.stores || []).length) {
    v.innerHTML = `<div class="spread">${picker}</div><div class="panel"><h2>No stores assigned</h2><p>${seesAll() ? 'Assign stores to a Market Leader in Setup.' : 'Ask Frank to assign your stores.'}</p></div>`;
    wirePicker(); return;
  }
  v.innerHTML = '<p class="loading">Loading the plan…</p>';
  let [plan, to] = await Promise.all([S.be.plan(email, week), S.be.timeOff(email, week)]);
  const canEdit = email === S.user.email || isAdmin();
  if (!plan && week === thisWeek && canEdit) {
    const prevPlan = await S.be.plan(email, addDays(week, -7)).catch(() => null);
    const carry = prevPlan?.anchorChoice && prevPlan.anchorChoice.mode !== 'auto' ? { ...prevPlan.anchorChoice, carried: true } : undefined;
    plan = newPlan(who, week, to?.off, carry);
    await S.be.savePlan(plan);
  }
  // A plan built early (when days off were picked) refreshes with the newest numbers once its week starts.
  // Next week's plan keeps updating with each new daily report until the week starts, so the one the
  // leader plans from on Sunday is built on Saturday's numbers. Days set by hand are never overwritten.
  if (plan && week > thisWeek && canEdit && S.daily && plan.preview && S.meta.latestDaily > (plan.basisDate || '') && !handSet(plan) && !plan.days.some(d => d.status === 'done')) {
    plan = { ...newPlan(who, week, plan.off, plan.anchorChoice), preview: true, pivots: plan.pivots || [], dismissed: plan.dismissed || [] };
    await S.be.savePlan(plan);
  }
  // Monday: only rebuild if nobody refreshed it over the weekend (built before Saturday's numbers).
  if (plan && week === thisWeek && canEdit && S.daily && (plan.preview ? S.meta.latestDaily > plan.basisDate && (plan.basisDate || '') < addDays(week, -2) && !handSet(plan) : !plan.basisDate) && !plan.days.some(d => d.status === 'done')) {
    plan = { ...newPlan(who, week, plan.off, plan.anchorChoice), pivots: plan.pivots || [], dismissed: plan.dismissed || [] };
    await S.be.savePlan(plan);
  }
  const t = today();
  const sugg = plan && week === thisWeek && canEdit && S.meta.latestDaily > plan.basisDate
    ? pivotSuggestion({ plan, scores: S.scores, today: t, dismissed: plan.dismissed || [] }) : null;
  const off = safeOff(plan?.off || to?.off || who.off, who.role);
  const doneCount = plan ? plan.days.filter(d => d.status === 'done').length : 0;

  v.innerHTML = `
  <div class="spread">
    <div>
      <p class="eyebrow">${week === thisWeek ? 'This week' : week < thisWeek ? 'Past week' : 'Next week'}${seesAll() ? ' · ' + esc(who.name || who.email) : ''}${marketsOf(who.email).length ? ' · ' + marketsOf(who.email).map(m => esc(m.name)).join(', ') : ''}</p>
      <div class="weekhead">
        <button class="btn tiny" id="prevw" aria-label="Previous week">‹</button>
        <h2 class="big">${esc(dayLabel(week))} to ${esc(dayLabel(addDays(week, 6)))}</h2>
        <button class="btn tiny" id="nextw" aria-label="Next week" ${week > thisWeek ? 'disabled' : ''}>›</button>
      </div>
      <p class="small" style="margin:4px 0 0">${dataLine()}</p>
    </div>
    <div class="row">${picker}</div>
  </div>
  ${!S.daily ? `<div class="warnbox"><b>No numbers yet.</b> Every store is on the plan in order for now. Once the first daily report is uploaded, the week re-ranks so the stores that need you most come first.</div>` : ''}
  ${!plan && week > thisWeek ? `<div class="panel"><h2>Pick your days off to build next week</h2><p>Tap your 2 days below and save. Your schedule builds right away from the latest numbers and refreshes on Sunday with Saturday's.</p></div>` : ''}
  ${plan?.preview && week > thisWeek ? `<div class="warnbox">${handSet(plan) && S.meta.latestDaily > (plan.basisDate || '')
    ? `Built from numbers through ${esc(longDate(plan.basisDate))}. Newer numbers are in (through ${esc(longDate(S.meta.latestDaily))}). You set some days by hand, so it won't change on its own. Tap <b>Rebuild from the numbers</b> to redo the week from the latest numbers.`
    : `Built from numbers through ${esc(longDate(plan.basisDate))}. It updates on its own with each new daily report until the week starts Monday, keeping your days off. Plan from Sunday's version, built on Saturday's numbers.`}</div>` : ''}
  ${!plan && S.daily && week < thisWeek ? `<div class="panel"><h2>No plan for this week</h2><p>Nothing was planned or logged.</p></div>` : ''}
  ${S.lastAlert && S.lastAlert.weekStart === week ? `<div class="warnbox" style="border-left-color:var(--green)"><b>${esc(vpNames())} was alerted in the app.</b> <a href="${alertMailto(S.lastAlert)}">Email them too</a></div>` : ''}
  <div id="swapreason" class="panel" hidden></div>
  ${plan && week >= thisWeek ? anchorPanel(plan, who, canEdit) : ''}
  ${sugg ? pivotCard(sugg) : ''}
  ${week >= thisWeek && !plan ? offPanel(week, off, to, who, canEdit, plan, week === thisWeek) : ''}
  ${plan ? `
  <div class="cols">
    <div style="min-width:0">
      ${canEdit && week >= thisWeek ? `<div class="row" style="margin:0 0 12px"><button class="btn" type="button" id="wkedit">${S.editWeek ? 'Close editor' : 'Edit my week'}</button><button class="btn" type="button" id="wkrebuild">Rebuild from the numbers</button><span class="small muted">Set any day by hand, or let the app redo the open days.</span></div>` : ''}
      ${S.editWeek && canEdit ? weekEditor(plan, who) : `<div class="days">${plan.days.map((d, i) => dayCard(d, i, plan, canEdit)).join('')}</div>`}
      <div class="offrow"><span>${doneCount} of ${plan.days.length} visits done</span><span>· Days off: <b>${off.map(x => DAY_LONG[x]).join(' and ')}</b></span></div>
      ${week === thisWeek ? remotePanel(plan, who, canEdit) : ''}
      ${(plan.pivots || []).length ? `<div class="panel"><h3>Changes made this week</h3><ul class="small">${plan.pivots.map(p => `<li>${esc(dayLabel(p.date))}: ${esc(p.from || 'open')} to <b>${esc(p.to)}</b>${p.reason ? '. ' + esc(p.reason) : ''}</li>`).join('')}</ul></div>` : ''}
    </div>
    <div class="panel" style="align-self:start">
      <h3>Where you are needed most</h3>
      <p class="small">Priority from the store numbers (sales and SPG with cancellations against budget and LY, close rate, cancellations, protection, finance), the consultants (below minimum or slipping this week) and days since the last visit, weighted by revenue. High-volume stores move up because that's where the dollars are. Higher needs you more.</p>
      <ul class="rank">${who.stores.slice().sort((a, b) => (S.scores[b]?.score ?? -1) - (S.scores[a]?.score ?? -1)).map(s => `
        <li>${needChip(S.scores[s]?.score)}<span class="nm">${esc(s)}<small>${esc(S.scores[s]?.parts?.[0]?.text || 'No flags')}</small></span><button class="link" data-open="${esc(s)}">Open</button></li>`).join('')}</ul>
    </div>
  </div>
  ${week >= thisWeek ? offPanel(week, off, to, who, canEdit, plan, week === thisWeek) : ''}` : ''}`;
  wirePicker();
  $('#prevw').onclick = () => { S.week = addDays(week, -7); S.editWeek = false; viewWeek(); };
  $('#nextw').onclick = () => { S.week = addDays(week, 7); S.editWeek = false; viewWeek(); };
  v.querySelectorAll('[data-open]').forEach(b => b.onclick = () => openVisit({ store: b.dataset.open, date: today(), email, kind: 'drop-in' }));
  v.querySelectorAll('[data-remote]').forEach(b => b.onclick = () => openVisit({ store: b.dataset.remote, date: today(), email, kind: 'remote', remote: true }));
  wireOffPanel(week, who, plan, week === thisWeek);
  if (!plan) return;
  wireAnchor(plan, who, week, week === thisWeek);
  wireWeekTools(plan, who, week, week === thisWeek);
  v.querySelectorAll('[data-go]').forEach(b => b.onclick = () => { const d = plan.days[+b.dataset.go]; openVisit({ store: d.store, date: d.date, email, kind: d.kind, dayIndex: +b.dataset.go }); });
  v.querySelectorAll('[data-gostop]').forEach(b => b.onclick = () => { const [i, j] = b.dataset.gostop.split(':').map(Number); const d = plan.days[i], x = d.stops[j]; openVisit({ store: x.store, date: d.date, email, kind: x.kind || 'first', dayIndex: i, stop: j }); });
  // Adding a stop is more coverage, so no reason needed. The day's first store becomes the morning.
  v.querySelectorAll('[data-addstop]').forEach(sel => sel.onchange = async () => {
    const i = +sel.dataset.addstop, d = plan.days[i], st = sel.value;
    const dm = driveMin(dayStores(d).slice(-1)[0], st);
    d.stops = [...(d.stops || []), { drive: dm, store: st, part: d.stops?.length ? 'Stop ' + (d.stops.length + 2) : 'PM', kind: plan.days.some((x, j) => j < i && dayStores(x).includes(st)) ? 'second' : 'first', status: 'planned' }];
    if (!d.part) d.part = 'AM';
    plan.calls = (plan.calls || []).filter(c => c !== st);
    plan.pivots = [...(plan.pivots || []), { date: d.date, from: d.store, to: dayText(d), reason: 'Added a stop', at: new Date().toISOString(), by: S.user.email }];
    await S.be.savePlan(plan); toast(`${st} added to ${dayLabel(d.date)}.${dm > MAX_SPLIT_MIN ? ` Heads up: that's about a ${driveText(dm)} drive.` : ''}`); viewWeek();
  });
  // Taking a stop off is a schedule change: reason and alert, same as a swap.
  v.querySelectorAll('[data-delstop]').forEach(b => b.onclick = () => {
    const [i, j] = b.dataset.delstop.split(':').map(Number); const d = plan.days[i], x = d.stops[j];
    const before = dayText(d);
    const apply = async (reason, note) => {
      d.stops.splice(j, 1); if (!d.stops.length && !d.anchor) delete d.part;
      plan.pivots = [...(plan.pivots || []), { date: d.date, from: before, to: dayText(d), reason: reason ? `${reason}${note ? ': ' + note : ''}` : 'Removed a stop', at: new Date().toISOString(), by: S.user.email }];
      await S.be.savePlan(plan);
      if (reason) await sendScheduleAlert(who, week, reason, note, [{ date: d.date, from: before, to: dayText(d) }], 'stop');
      toast(`${x.store} taken off ${dayLabel(d.date)}.${reason ? ` ${vpNames()} was alerted.` : ''}`); viewWeek();
    };
    if (!needsReason()) return apply(null, '');
    askReason('#swapreason', `Taking ${x.store} off ${dayLabel(d.date)}`, apply);
  });
  v.querySelectorAll('[data-swap]').forEach(sel => sel.onchange = async () => {
    const i = +sel.dataset.swap, d = plan.days[i], to = sel.value, from = d.store;
    const apply = async (reason, note) => {
      plan.pivots = [...(plan.pivots || []), { date: d.date, from, to, reason: reason ? `${reason}${note ? ': ' + note : ''}` : 'Changed by hand', at: new Date().toISOString(), by: S.user.email }];
      Object.assign(d, { store: to, kind: plan.days.some((x, j) => j !== i && x.store === to && x.date < d.date) ? 'second' : 'first', score: S.scores[to]?.score ?? null });
      plan.calls = (plan.calls || []).filter(c => c !== to);
      await S.be.savePlan(plan);
      if (reason) await sendScheduleAlert(who, week, reason, note, [{ date: d.date, from: from || 'Open day', to }], 'swap');
      toast(`${dayLabel(d.date)} is now ${to}.${reason ? ` ${vpNames()} was alerted.` : ''}`); viewWeek();
    };
    if (!needsReason()) return apply(null, '');
    v.querySelectorAll('[data-swap]').forEach(x => x.disabled = true);
    askReason('#swapreason', `Moving ${dayLabel(d.date)} from ${from || 'an open day'} to ${to}`, apply);
  });
  if (sugg) {
    $('#pvyes').onclick = async () => {
      // Drop the least-opportunity visit, add the new store, run the rest of the week in priority order.
      sugg.reorder.forEach(r => Object.assign(plan.days[r.i], { store: r.store, kind: r.kind, score: r.score }));
      plan.pivots = [...(plan.pivots || []), { date: sugg.date, from: sugg.from, to: sugg.to, reason: sugg.reason, at: new Date().toISOString() }];
      if (sugg.loses) plan.calls = [...new Set([...(plan.calls || []), sugg.from])];
      plan.calls = plan.calls.filter(c => c !== sugg.to);
      await S.be.savePlan(plan);
      await sendScheduleAlert(who, week, 'Store numbers changed', `Took the app's suggested swap. ${sugg.reason || ''}`.trim(), [{ date: sugg.date, from: sugg.from, to: sugg.to }], 'pivot').catch(() => null);
      toast(`Swapped. ${sugg.to} is in, ${sugg.from} is out.${needsReason() ? ` ${vpNames()} was told.` : ''}`); viewWeek();
    };
    $('#pvno').onclick = async () => {
      plan.dismissed = [...(plan.dismissed || []), sugg.key];
      await S.be.savePlan(plan); toast('Kept your plan.'); viewWeek();
    };
  }
}
// ---------------------------------------------------------------- edit the week by hand
// Every open day: a morning (or full-day) store and an optional afternoon store, with drive times.
// Logged or past days are locked. Saving a hand-built week asks a Market Leader for a reason.
function weekEditor(plan, who) {
  const stores = who.stores || [];
  const t = today();
  const opt = (sel, list, none) => `${none ? `<option value="">${none}</option>` : ''}${list.map(o => `<option value="${esc(o.s)}" ${o.s === sel ? 'selected' : ''}>${esc(o.s)}${o.m != null ? ` (~${driveText(o.m)})` : ''}${S.scores[o.s]?.score != null ? ` · ${S.scores[o.s].score}` : ''}</option>`).join('')}`;
  return `<div class="panel"><h3 style="margin:0 0 4px">Edit my week</h3>
    <p class="small" style="margin:0 0 10px">Pick the morning store for each day, and an afternoon store if you're splitting the day. Afternoon choices are sorted by drive time from the morning store; the number after it is the store's priority. Days already logged or past are locked.</p>
    <div class="scroller"><table class="grid"><thead><tr><th>Day</th><th>Morning / full day</th><th>Afternoon</th><th></th></tr></thead><tbody>
    ${plan.days.map((d, i) => {
      const lock = d.status === 'done' || d.date < t;
      const pm = d.stops?.[0]?.store || '';
      const amList = stores.map(s => ({ s, m: null })).sort((a, b) => (S.scores[b.s]?.score ?? 0) - (S.scores[a.s]?.score ?? 0));
      const pmList = stores.filter(s => s !== d.store).map(s => ({ s, m: driveMin(d.store, s) })).sort((a, b) => (a.m ?? 999) - (b.m ?? 999));
      const m = pm ? driveMin(d.store, pm) : null;
      return `<tr><td class="nm">${esc(DAY_LONG[dow(d.date)])} ${esc(shortDate(d.date))}</td>
        <td>${lock ? esc(d.store || '') : `<select data-wam="${i}">${opt(d.store, amList)}</select>`}</td>
        <td>${lock ? esc((d.stops || []).map(x => x.store).join(', ') || '--') : `<select data-wpm="${i}">${opt(pm, pmList, 'None (full day)')}</select>`}</td>
        <td class="small ${m > MAX_SPLIT_MIN ? 'warn' : 'muted'}" data-wdrive="${i}">${lock ? (d.status === 'done' ? 'Logged' : 'Past') : m != null ? `~${driveText(m)} drive` : ''}</td></tr>`;
    }).join('')}
    </tbody></table></div>
    ${(d => d.length ? `<p class="small" style="margin:8px 0 0">Not on any day: ${d.map(esc).join(', ')}. Cover them with a remote call or add them to a day.</p>` : '')(stores.filter(s => !plan.days.some(d => dayStores(d).includes(s))))}
    <div id="wkreason" style="margin-top:10px"></div>
    <div class="row" style="margin-top:12px"><button class="btn primary" type="button" id="wksave">Save my week</button><button class="link" type="button" id="wkcancel">Cancel</button></div>
  </div>`;
}
function wireWeekTools(plan, who, week, isCurrent) {
  const ed = $('#wkedit'); if (!ed) return;
  ed.onclick = () => { S.editWeek = !S.editWeek; viewWeek(); };
  $('#wkrebuild').onclick = async () => {
    // Back to the app's plan for the open days. Following the numbers needs no reason.
    const fresh = isCurrent ? rebuildPlan(plan, who, plan.off, week) : { ...newPlan(who, week, plan.off, plan.anchorChoice), preview: plan.preview, pivots: plan.pivots || [], dismissed: plan.dismissed || [] };
    fresh.pivots = [...(fresh.pivots || []), { date: today(), from: 'Week', to: 'Rebuilt from the numbers', reason: 'Rebuilt from the numbers', at: new Date().toISOString(), by: S.user.email }];
    try { await S.be.savePlan(fresh); S.editWeek = false; toast('Open days rebuilt from the latest numbers.'); viewWeek(); } catch (e) { toast(friendly(e), true); }
  };
  if (!S.editWeek) return;
  const v = $('#view');
  // Live drive time as the afternoon changes; afternoon list re-sorts when the morning changes.
  const redraw = i => {
    const am = v.querySelector(`[data-wam="${i}"]`)?.value, pmSel = v.querySelector(`[data-wpm="${i}"]`), cell = v.querySelector(`[data-wdrive="${i}"]`);
    if (!pmSel || !cell) return;
    const m = pmSel.value ? driveMin(am, pmSel.value) : null;
    cell.textContent = m != null ? `~${driveText(m)} drive` : ''; cell.className = `small ${m > MAX_SPLIT_MIN ? 'warn' : 'muted'}`;
  };
  v.querySelectorAll('[data-wam]').forEach(sel => sel.onchange = () => {
    const i = sel.dataset.wam, pmSel = v.querySelector(`[data-wpm="${i}"]`), keep = pmSel.value;
    const list = (who.stores || []).filter(s => s !== sel.value).map(s => ({ s, m: driveMin(sel.value, s) })).sort((a, b) => (a.m ?? 999) - (b.m ?? 999));
    pmSel.innerHTML = `<option value="">None (full day)</option>` + list.map(o => `<option value="${esc(o.s)}" ${o.s === keep ? 'selected' : ''}>${esc(o.s)} (~${driveText(o.m)})${S.scores[o.s]?.score != null ? ` · ${S.scores[o.s].score}` : ''}</option>`).join('');
    redraw(i);
  });
  v.querySelectorAll('[data-wpm]').forEach(sel => sel.onchange = () => redraw(sel.dataset.wpm));
  $('#wkcancel').onclick = () => { S.editWeek = false; viewWeek(); };
  $('#wksave').onclick = () => {
    const days = plan.days.map((d, i) => {
      const am = v.querySelector(`[data-wam="${i}"]`); if (!am) return d;
      const pm = v.querySelector(`[data-wpm="${i}"]`).value;
      const nd = { date: d.date, store: am.value, status: 'planned', score: S.scores[am.value]?.score ?? null, stops: [] };
      if (pm && pm !== am.value) { nd.part = 'AM'; nd.stops.push({ store: pm, part: 'PM', status: 'planned', drive: driveMin(am.value, pm) }); }
      if (plan.anchor && am.value === plan.anchor && nd.stops.length) nd.anchor = true;
      return nd;
    });
    const seen = new Set();
    days.forEach(d => { if (d.status === 'done' || d.date < today()) { dayStores(d).forEach(s => seen.add(s)); return; } d.kind = seen.has(d.store) ? 'second' : 'first'; seen.add(d.store); d.stops.forEach(x => { x.kind = seen.has(x.store) ? 'second' : 'first'; seen.add(x.store); }); });
    const diff = scheduleDiff(plan.days, days);
    if (!diff.length) { S.editWeek = false; toast('No changes.'); return viewWeek(); }
    const far = days.filter(d => d.stops?.[0]?.drive > MAX_SPLIT_MIN);
    const commit = async (reason, note) => {
      const fresh = { ...plan, days, calls: (who.stores || []).filter(s => !days.some(d => dayStores(d).includes(s))) };
      fresh.pivots = [...(plan.pivots || []), ...diff.map(c => ({ date: c.date, from: c.from, to: c.to, reason: reason ? `${reason}${note ? ': ' + note : ''}` : 'Edited by hand', at: new Date().toISOString(), by: S.user.email }))];
      await S.be.savePlan(fresh);
      if (reason) await sendScheduleAlert(who, week, reason, note, diff, 'edit');
      S.editWeek = false;
      toast(`Week saved.${far.length ? ` Heads up: ${far.length} day${far.length > 1 ? 's have' : ' has'} a long afternoon drive.` : ''}${reason ? ` ${vpNames()} was alerted.` : ''}`); viewWeek();
    };
    if (needsReason()) { $('#wksave').disabled = true; return askReason('#wkreason', 'Why are you changing your week?', commit, () => { S.editWeek = false; viewWeek(); }); }
    commit(null, '').catch(e => toast(friendly(e), true));
  };
}
// ---------------------------------------------------------------- anchor store
function anchorPanel(plan, who, canEdit) {
  const aw = plan.anchorWhy || {};
  const ch = plan.anchorChoice || { mode: 'auto' };
  const auto = autoAnchor(who.stores || []);
  const src = { leader: `${who.email === S.user.email ? 'You' : esc(firstOf(who.name) || 'They')} picked it${ch.carried ? ' (carried from last week)' : ''}`, admin: 'Set on the market in Setup', app: 'The app picked it from the numbers' }[aw.source] || '';
  const head = aw.store ? `<b>${esc(aw.store)}</b>, ${aw.days === plan.days.length ? 'every work day' : aw.days + ' mornings'} this week. <span class="small muted">${src}.</span>`
    : `<b>No anchor store this week.</b> <span class="small muted">${ch.mode === 'none' ? (who.email === S.user.email ? 'You turned it off.' : 'Turned off for this week.') : esc(aw.why || '')}</span>`;
  const why = aw.store && (aw.why || []).length ? `<ul class="blist small" style="margin:6px 0 0">${aw.why.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : '';
  const appSays = (auto.store || '') !== (aw.store || '') ? `<p class="small" style="margin:6px 0 0"><b>The numbers say:</b> ${auto.store ? `${esc(auto.store)}, ${auto.days} mornings.` : 'no anchor needed.'}</p>` : '';
  const stores = who.stores || [];
  return `<section class="panel" style="border-left:6px solid #F68C2C">
    <div class="spread" style="margin:0"><div><p class="eyebrow">Anchor store</p><p style="margin:0">${head}</p>${why}${appSays}</div>
    ${canEdit && who.role === 'leader' ? `<button class="btn tiny" type="button" id="anchoredit">Change</button>` : ''}</div>
    <div id="anchorform" hidden style="margin-top:12px">
      <p class="small" style="margin:0 0 8px">An anchor store gets your mornings to set the tone, then you go to a second store for the afternoon. Use it for a store that's struggling, a store without a GM, or a new leader who needs you there. You know things the numbers don't.</p>
      <div class="two">
        <label for="anmode">Anchor<select id="anmode"><option value="auto" ${ch.mode === 'auto' ? 'selected' : ''}>Let the app decide${auto.store ? ` (${esc(auto.store)})` : ' (none right now)'}</option>${stores.map(s => `<option value="store:${esc(s)}" ${ch.mode === 'store' && ch.store === s ? 'selected' : ''}>${esc(s)}</option>`).join('')}<option value="none" ${ch.mode === 'none' ? 'selected' : ''}>No anchor this week</option></select></label>
        <label for="andays">Mornings this week<select id="andays">${[5, 4, 3, 2, 1].filter(n => n <= plan.days.length).map(n => `<option value="${n}" ${(+ch.days || aw.days || 5) === n ? 'selected' : ''}>${n === plan.days.length ? `Every work day (${n})` : n}</option>`).join('')}</select></label>
      </div>
      <div id="anreason" style="margin-top:10px"></div>
      <div class="row" style="margin-top:10px"><button class="btn primary" type="button" id="ansave">Save and rebuild my week</button><button class="link" type="button" id="ancancel">Cancel</button><span class="small muted">Visits already logged stay put.</span></div>
    </div>
  </section>`;
}
function wireAnchor(plan, who, week, isCurrent) {
  const ed = $('#anchoredit'); if (!ed) return;
  const form = $('#anchorform');
  ed.onclick = () => { form.hidden = false; ed.hidden = true; };
  $('#ancancel').onclick = () => viewWeek();
  $('#ansave').onclick = async () => {
    const v = $('#anmode').value, days = +$('#andays').value;
    const choice = v === 'auto' ? { mode: 'auto' } : v === 'none' ? { mode: 'none' } : { mode: 'store', store: v.slice(6), days };
    if (choice.mode === 'auto') choice.days = null;
    const auto = autoAnchor(who.stores || []);
    const save = async (reason, note) => {
      if (reason) choice.reason = `${reason}${note ? ': ' + note : ''}`;
      choice.by = S.user.email; choice.at = new Date().toISOString();
      const fresh = isCurrent ? rebuildPlan(plan, who, plan.off, week, choice) : { ...newPlan(who, week, plan.off, choice), preview: plan.preview, pivots: plan.pivots || [], dismissed: plan.dismissed || [] };
      fresh.anchorChoice = choice;
      const label = choice.mode === 'auto' ? `the app's pick (${auto.store || 'none'})` : choice.mode === 'none' ? 'no anchor' : `${choice.store}, ${days} mornings`;
      fresh.pivots = [...(fresh.pivots || []), { date: today(), from: 'Anchor', to: label, reason: choice.reason || 'Anchor changed', at: choice.at, by: S.user.email }];
      await S.be.savePlan(fresh);
      if (reason) await sendScheduleAlert(who, week, reason, `Anchor store set to ${label}. The numbers say ${auto.store ? auto.store + ', ' + auto.days + ' mornings' : 'no anchor'}.${note ? ' ' + note : ''}`, scheduleDiff(plan.days, fresh.days), 'anchor');
      toast(`Anchor: ${label}. Week rebuilt.${reason ? ` ${vpNames()} was alerted.` : ''}`); viewWeek();
    };
    // Going with the numbers needs no reason. Overriding them does, and the VP hears about it.
    const overrides = choice.mode !== 'auto' && !(choice.mode === 'store' && choice.store === auto.store);
    if (overrides && needsReason()) { $('#ansave').disabled = true; return askReason('#anreason', 'Why this anchor? Tell your VP what you know.', save); }
    try { await save(null, ''); } catch (e) { toast(friendly(e), true); }
  };
}
// ---------------------------------------------------------------- days off
// Leaders pick any 2 days off for each week. Saving builds (or rebuilds) that week's plan right away.
function offPanel(week, off, to, who, canEdit, plan, isCurrent) {
  const days = [0, 1, 2, 3, 4, 5, 6].map(i => addDays(week, i));
  const storeOn = offStores(week, off, who, plan);
  const set = to ? `Set ${to.by && to.by !== who.email ? 'by ' + esc(to.by) + ' ' : ''}on ${esc(shortDate(to.at.slice(0, 10)))}` : `Not set yet. Using ${esc(who.email === S.user.email ? 'your' : 'the')} default: ${off.map(x => DAY_NAMES[x]).join(' and ')}.`;
  return `<section class="panel" aria-labelledby="offh">
    <div class="spread" style="margin:0 0 10px"><div><h3 id="offh" style="margin:0">Days off, week of ${esc(dayLabel(week))}</h3>
      <p class="small" style="margin:2px 0 0">Tap any 2 days. Tuesday, Wednesday or Thursday works best. Your visit schedule builds as soon as you save.</p></div>
      <span class="small ${to ? 'good' : 'warn'}">${set}</span></div>
    <div class="offstrip">${days.map(d => {
      const w = dow(d), isOff = off.includes(w), suggested = [2, 3, 4].includes(w);
      const doneHere = plan?.days.some(x => x.date === d && x.status === 'done');
      return `<button type="button" class="offday ${isOff ? 'isoff' : ''} ${suggested ? 'sugg' : ''}" data-offday="${w}" aria-pressed="${isOff}" ${canEdit && !(isCurrent && d < today()) && !doneHere ? '' : 'disabled'}><b>${DAY_NAMES[w]}</b><span>${esc(shortDate(d))}</span><em>${isOff ? 'Off' : esc(storeOn[d] || 'Working')}</em></button>`;
    }).join('')}</div>
    <div id="offreason" hidden style="margin-top:12px"></div>
    ${canEdit ? `<div class="row" style="margin-top:12px"><button class="btn primary" id="offsave">${plan ? 'Save and rebuild my schedule' : 'Save and build my schedule'}</button>
      <span class="small">${isCurrent && plan ? 'Visits already logged stay on their days.' : !isCurrent ? 'It updates with each new daily report until the week starts Monday.' : ''}</span></div>` : ''}
  </section>`;
}
// Which store each working day goes to: the saved plan's days, or what the plan would be with these days off.
function offStores(week, off, who, plan) {
  const out = {};
  const same = plan && [...(plan.off || [])].sort().join() === [...off].sort().join();
  const src = same ? plan.days : validOff(off) ? newPlan(who, week, off, plan?.anchorChoice).days : [];
  src.forEach(d => { if (d.store) out[d.date] = dayText(d); });
  return out;
}
function wireOffPanel(week, who, plan, isCurrent) {
  const btn = $('#offsave'); if (!btn) return;
  const all = [...$('#view').querySelectorAll('[data-offday]')];
  const relabel = () => {
    const picked = all.filter(b => b.classList.contains('isoff')).map(b => +b.dataset.offday);
    const map = validOff(picked) ? offStores(week, picked, who, plan) : {};
    all.forEach(b => { const d = addDays(week, +b.dataset.offday); b.querySelector('em').textContent = b.classList.contains('isoff') ? 'Off' : (map[d] || 'Working'); });
  };
  all.forEach(b => b.onclick = () => {
    const on = !b.classList.contains('isoff');
    b.classList.toggle('isoff', on); b.setAttribute('aria-pressed', on); relabel();
  });
  btn.onclick = async () => {
    const picked = [...$('#view').querySelectorAll('.offday.isoff')].map(b => +b.dataset.offday).sort();
    if (!validOff(picked)) return toast('Pick exactly 2 days off.', true);
    const to0 = await S.be.timeOff(who.email, week).catch(() => null);
    const build = () => (isCurrent && plan) ? rebuildPlan(plan, who, picked, week) : { ...newPlan(who, week, picked, plan?.anchorChoice), preview: !isCurrent, pivots: plan?.pivots || [], dismissed: plan?.dismissed || [] };
    const commit = async (reason, note) => {
      const fresh = build();
      await S.be.saveTimeOff({ email: who.email, weekStart: week, off: picked, at: new Date().toISOString(), by: S.user.email, ...(reason ? { reason, note } : {}) });
      if (reason) fresh.pivots = [...(fresh.pivots || []), { date: today(), from: `Days off ${(plan.off || []).map(x => DAY_NAMES[x]).join('/')}`, to: `Days off ${picked.map(x => DAY_NAMES[x]).join('/')}`, reason: `${reason}${note ? ': ' + note : ''}`, at: new Date().toISOString(), by: S.user.email }];
      await S.be.savePlan(fresh);
      if (reason) await sendScheduleAlert(who, week, reason, note, scheduleDiff(plan.days, fresh.days), 'rebuild');
      toast(`Days off: ${picked.map(x => DAY_LONG[x]).join(' and ')}. Schedule ${plan ? 'rebuilt' : 'built'}.${reason ? ` ${vpNames()} was alerted.` : ''}`);
      viewWeek();
    };
    // Changing a schedule that's already set (days off saved before, or the week has started) needs a reason.
    const changing = plan && (to0 || today() > week) && [...(plan.off || [])].sort().join() !== picked.join();
    if (changing && needsReason()) {
      if (!scheduleDiff(plan.days, build().days).length) return commit(null, '');
      btn.disabled = true;
      return askReason('#offreason', 'Why are you rebuilding your schedule?', commit);
    }
    try { await commit(null, ''); } catch (e) { toast(friendly(e), true); }
  };
}
// Keep logged visits on their days, then fill the other work days from today on: stores not seen yet
// first (highest need first), then second visits to the highest-need stores.
function rebuildPlan(plan, who, off, week, choice) {
  const fresh = newPlan(who, week, off, choice === undefined ? plan.anchorChoice : choice);
  // Logged visits and days already past stay as they were.
  const done = plan.days.filter(d => d.status === 'done' || d.date < today());
  const open = fresh.days.map(d => d.date).filter(dt => dt >= today() && !done.some(k => k.date === dt));
  if (fresh.anchor) {
    fresh.days = [...done, ...fresh.days.filter(d => open.includes(d.date))].sort((a, b) => a.date.localeCompare(b.date));
    fresh.pivots = plan.pivots || []; fresh.dismissed = plan.dismissed || []; fresh.builtAt = plan.builtAt; fresh.basisDate = plan.basisDate;
    return fresh;
  }
  const byNeed = who.stores.slice().sort((a, b) => (S.scores[b]?.score ?? 0) - (S.scores[a]?.score ?? 0));
  const seen = new Set(done.filter(d => d.status === 'done').map(d => d.store));
  const order = [...byNeed.filter(s => !seen.has(s)), ...byNeed];
  fresh.days = [...done, ...open.map((date, i) => ({ date, store: order[i] || null, kind: seen.has(order[i]) || order.indexOf(order[i]) < i ? 'second' : 'first', status: 'planned', score: S.scores[order[i]]?.score ?? null }))]
    .sort((a, b) => a.date.localeCompare(b.date));
  fresh.calls = byNeed.filter(s => !fresh.days.some(d => d.store === s));
  fresh.pivots = plan.pivots || []; fresh.dismissed = plan.dismissed || [];
  fresh.builtAt = plan.builtAt; fresh.basisDate = plan.basisDate;
  return fresh;
}
// Remote coaching: while on a full-day visit, coach the other stores by phone or video and log it.
function remotePanel(plan, who, canEdit) {
  const t = today(), here = plan.days.find(d => d.date === t)?.store;
  const wk = plan.weekStart, end = addDays(wk, 6);
  const remote = S.visits.filter(v => v.remote && v.email === who.email && v.date >= wk && v.date <= end);
  const stores = (who.stores || []).filter(s => s !== here).sort((a, b) => ((plan.calls || []).includes(b) ? 1 : 0) - ((plan.calls || []).includes(a) ? 1 : 0) || (S.scores[b]?.score ?? 0) - (S.scores[a]?.score ?? 0));
  return `<div class="panel">
    <div class="spread" style="margin:0 0 6px"><h3 style="margin:0">Remote coaching</h3><span class="small">${remote.length} logged this week</span></div>
    <p class="small">Visits are full days. While you're in one store, call or video your other stores and log the coaching here. ${plan.calls?.length ? 'Stores marked Call have no visit this week.' : ''}</p>
    <div class="remotes">${stores.map(s => {
      const done = remote.filter(v => v.store === s);
      return `<div class="rstore"><span class="nm">${needChip(S.scores[s]?.score)} <b>${esc(s)}</b>${(plan.calls || []).includes(s) ? ' <span class="pill check">Call</span>' : ''}</span>
        <span class="small">${done.length ? `<span class="good">Coached ${done.map(v => DAY_NAMES[dow(v.date)]).join(', ')}</span>` : '<span class="muted">Not yet this week</span>'}</span>
        ${canEdit ? `<button type="button" class="btn tiny" data-remote="${esc(s)}">${done.some(v => v.date === t) ? 'Open today\'s' : 'Log remote coaching'}</button>` : ''}</div>`;
    }).join('')}</div></div>`;
}
function newPlan(who, week, off, choice) {
  const plan = { email: who.email, name: who.name || who.email,
    ...buildPlan({ weekStart: week, stores: who.stores, scores: S.scores, off: safeOff(off || who.off, who.role), role: who.role }),
    builtAt: new Date().toISOString(), basisDate: S.meta.latestDaily, pivots: [], dismissed: [] };
  if (choice) plan.anchorChoice = choice;
  const a = anchorFor(who, choice);
  plan.anchorWhy = a ? { store: a.store, days: a.days, source: a.source, why: a.why } : { none: true, why: autoAnchor(who.stores || []).whyNot || '' };
  return a ? anchorPlan(plan, who, a.store, a.days, a.source) : plan;
}
// Who decides the anchor, in order: the Market Leader's choice for the week (they may know something
// the numbers don't), then an anchor Frank set on the market in Setup, then the app from the numbers.
function anchorFor(who, choice) {
  const mine = who.stores || [];
  if (choice?.mode === 'none') return null;
  if (choice?.mode === 'store' && mine.includes(choice.store)) return { store: choice.store, days: +choice.days || 5, source: 'leader', why: choice.reason ? [choice.reason] : [] };
  const m = (S.markets || []).find(x => x.leader === who.email);
  if (m?.anchorMode === 'none') return null;
  if (m?.anchor && mine.includes(m.anchor)) return { store: m.anchor, days: +m.anchorDays || 5, source: 'admin', why: ['Set on the market in Setup'] };
  if (who.role !== 'leader') return null;
  const auto = autoAnchor(mine);
  return auto.store ? auto : null;
}
// The app's call: one store well below the rest of the market, short on sales or SPG with
// cancellations, gets anchor mornings. No store leader on file makes the case stronger.
// More need, more mornings: 75+ gets every work day, 65+ gets 4, otherwise 3.
function autoAnchor(stores) {
  const rows = stores.map(s => ({ s, sc: S.scores[s]?.score, m: S.daily?.stores?.[s]?.mtd })).filter(r => r.sc != null).sort((a, b) => b.sc - a.sc);
  if (rows.length < 2) return { whyNot: rows.length ? 'Only one store in the market.' : 'No numbers yet.' };
  const top = rows[0], rest = rows.slice(1), avg = rest.reduce((t, r) => t + r.sc, 0) / rest.length;
  const bud = top.m?.vsBud?.netSales, spg = top.m?.vsLy?.spg;
  const listOnFile = (S.storeLeaders || []).length > 0;
  const withRoles = (S.storeLeaders || []).some(l => isGM(l.role));
  const noLeader = listOnFile && !(S.storeLeaders || []).some(l => l.store === top.s && (!withRoles || isGM(l.role)));
  const gap = top.sc - avg, short = (bud != null && bud <= -10) || (spg != null && spg <= -10);
  const yes = (top.sc >= 55 && short && gap >= 15) || (top.sc >= 45 && noLeader && gap >= 10);
  if (!yes) return { whyNot: `No store stands out enough. ${top.s} has the highest priority at ${top.sc}, ${Math.round(gap)} points above the rest of the market. The app anchors a store at 55+ priority (need weighted by revenue), 15+ points above the rest, and 10% or more short on sales or SPG with cancellations.` };
  const why = [`Priority ${top.sc}, ${Math.round(gap)} points above the rest of the market`];
  const sc0 = S.scores[top.s];
  if (sc0?.behind > 0) why.push(`$${Math.round(sc0.behind).toLocaleString('en-US')} behind budget this month${sc0.factor >= 1.1 ? ', one of the bigger stores' : ''}`);
  if (bud != null) why.push(`Sales ${pct(bud)} to budget this month`);
  if (spg != null) why.push(`SPG with cancellations ${pct(spg)} vs LY`);
  const t = S.teams?.[top.s]; if (t?.below?.length) why.push(`${t.below.length} consultant${t.below.length > 1 ? 's' : ''} below the minimum`);
  if (noLeader) why.push(withRoles ? 'No GM in the store' : 'No store leader on file');
  return { store: top.s, days: top.sc >= 75 ? 5 : top.sc >= 65 ? 4 : 3, source: 'app', why };
}
// Anchor store: an underperforming store with no GM. The Market Leader spends the morning there on
// anchor days to set the tone, then travels to a second store for the afternoon. The other stores
// rotate through the afternoons and any full days left, highest need first.
function anchorPlan(plan, who, anchor, n, source) {
  n = Math.min(n, plan.days.length);
  // When the app picks the anchor, stores too far for an afternoon still get a full day each,
  // so it gives back anchor mornings (never fewer than 2). A leader's own pick is kept as set.
  if (source === 'app') {
    const farN = (who.stores || []).filter(s => s !== anchor && (driveMin(anchor, s) ?? 999) > MAX_SPLIT_MIN).length;
    n = Math.max(Math.min(2, n), Math.min(n, plan.days.length - farN));
  }
  // Mid-week rebuilds only fill days from today on, so the anchor mornings go there first.
  const t0 = today(), order = plan.days.map((d, i) => i).sort((a, b) => (plan.days[a].date < t0) - (plan.days[b].date < t0) || a - b);
  const anchorIdx = new Set(order.slice(0, n));
  const others = (who.stores || []).filter(s => s !== anchor).sort((a, b) => (S.scores[b]?.score ?? 0) - (S.scores[a]?.score ?? 0) || a.localeCompare(b));
  // Afternoons only go to stores within a short drive of the anchor. Far stores get full days.
  const near = others.filter(s => (driveMin(anchor, s) ?? 999) <= MAX_SPLIT_MIN), far = others.filter(s => !near.includes(s));
  let kn = 0, kf = 0, kAll = 0;
  const nextNear = () => near.length ? near[kn++ % near.length] : null;
  const nextFull = () => far.length && kf < far.length ? far[kf++] : others.length ? others[kAll++ % others.length] : null;
  const seen = new Set();
  const kindOf = st => { const kd = seen.has(st) ? 'second' : 'first'; seen.add(st); return kd; };
  plan.days = plan.days.map((d, i) => {
    if (anchorIdx.has(i)) {
      const pm = nextNear();
      const day = { date: d.date, store: anchor, kind: kindOf(anchor), anchor: true, part: pm ? 'AM' : null, status: 'planned', score: S.scores[anchor]?.score ?? null, stops: [] };
      if (!pm) delete day.part;
      if (pm) day.stops.push({ store: pm, part: 'PM', kind: kindOf(pm), status: 'planned', drive: driveMin(anchor, pm) });
      return day;
    }
    const st = nextFull() || anchor;
    return { date: d.date, store: st, kind: kindOf(st), status: 'planned', score: S.scores[st]?.score ?? null };
  });
  plan.anchor = anchor; plan.anchorDays = n;
  plan.calls = others.filter(s => !plan.days.some(d => d.store === s || (d.stops || []).some(x => x.store === s)));
  return plan;
}
const dayStores = d => d ? [d.store, ...(d.stops || []).map(x => x.store)].filter(Boolean) : [];
const dayText = d => dayStores(d).join(' + ') || 'Open day';
function wirePicker() { const lp = $('#lp'); if (lp) lp.onchange = () => { S.viewEmail = lp.value; S.editWeek = false; viewWeek(); }; }
function pivotCard(s) {
  return `<div class="pivot" role="region" aria-label="Suggested change">
    <p class="eyebrow">New numbers since your week was built</p>
    <h3>Suggested: add ${esc(s.to)} and drop ${esc(s.from)}</h3>
    <p>${esc(s.reasonTo)}${s.why.length ? ' ' + s.why.map(esc).join('. ') + '.' : ''} ${esc(s.reasonFrom)}${s.loses ? ` ${esc(s.from)} moves to a phone check-in this week.` : ''}</p>
    <p class="small" style="margin:0 0 4px"><b>Rest of the week, highest priority first:</b></p>
    <ol class="small" style="margin:0 0 12px;padding-left:20px;color:var(--ink)">${s.reorder.map(r => `<li>${esc(DAY_LONG[dow(r.date)])} ${esc(shortDate(r.date))}: <b>${esc(r.store)}</b> (need ${r.score})${r.store === s.to ? ' <span class="pill set">New</span>' : ''}</li>`).join('')}</ol>
    <div class="row"><button class="btn accent" id="pvyes">Make the swap</button><button class="btn" id="pvno">Keep my plan</button></div>
  </div>`;
}
function dayCard(d, i, plan, canEdit) {
  const t = today(), past = d.date < t, isToday = d.date === t;
  const stores = S.users.find(u => u.email === plan.email)?.stores || (plan.email === S.user.email ? S.user.stores : Object.keys(plan.basis));
  const multi = (d.stops || []).length > 0;
  const kindPill = k => k === 'second' ? '<span class="pill check">Check the plan</span>' : k === 'first' ? '<span class="pill set">Set the plan</span>' : '';
  const statePill = st => st === 'done' ? '<span class="pill done">Visited</span>' : past ? '<span class="pill off">Not logged</span>' : isToday ? '<span class="pill">Today</span>' : '';
  // One block per store on the day: a preview of why you're going and who to see, and its own visit button.
  const block = (store, part, kind, status, btn, extra = '') => {
    const sc = S.scores[store], team = S.teams?.[store];
    const see = team ? [...team.below.map(r => titleName(r.name) + ' (below min)'), ...team.slipping.map(r => titleName(r.name) + ' (slipping)')].slice(0, 3) : [];
    const lv = suggestLever(S.daily?.stores?.[store]?.mtd), lvL = lv ? LEVERS.find(l => l.key === lv)?.label : null;
    return `<div class="dstop ${status === 'done' ? 'isdone' : ''}">
      <div class="row" style="justify-content:space-between;flex-wrap:nowrap"><span class="store">${part ? `<span class="part">${esc(part)}</span> ` : ''}${esc(store)}</span>${needChip(sc?.score)}</div>
      ${extra}
      <div class="row">${kindPill(kind)}${statePill(status)}</div>
      ${status !== 'done' && !past ? intentHtml(visitIntent(store, { anchor: part === 'AM' && d.anchor, kind, date: d.date }), true) : ''}
      <div class="row" style="gap:6px">${btn}</div>
    </div>`;
  };
  const primaryBtn = d.store ? `<button class="btn tiny ${isToday ? 'primary' : ''}" data-go="${i}">${d.status === 'done' ? 'See visit' : `Open ${multi ? (d.part || 'AM') + ' ' : ''}visit`}</button>` : '';
  return `<article class="day ${isToday ? 'today' : ''} ${d.status === 'done' && !multi ? 'isdone' : ''}">
    <div class="row" style="justify-content:space-between"><span class="dname">${esc(DAY_LONG[dow(d.date)])} ${esc(shortDate(d.date))}</span>${multi ? `<span class="small muted">${dayStores(d).length} stores</span>` : ''}</div>
    ${d.store ? block(d.store, multi ? (d.part || 'AM') : '', d.kind, d.status, primaryBtn, d.anchor ? '<div class="row"><span class="pill anchor">Anchor store</span></div>' : '') : '<div class="store">Open day</div>'}
    ${(d.stops || []).map((x, j) => {
      const m = driveMin(j ? d.stops[j - 1].store : d.store, x.store);
      const drive = m != null ? `<p class="small ${m > MAX_SPLIT_MIN ? 'warn' : 'muted'}" style="margin:0">~${driveText(m)} drive from ${esc(j ? d.stops[j - 1].store : d.store)}</p>` : '';
      const btn = `<button class="btn tiny ${isToday ? 'primary' : ''}" type="button" data-gostop="${i}:${j}">${x.status === 'done' ? 'See visit' : `Open ${esc(x.part || 'stop')} visit`}</button>${canEdit && x.status !== 'done' && !past ? `<button class="link" type="button" data-delstop="${i}:${j}" aria-label="Remove ${esc(x.store)}">Remove</button>` : ''}`;
      return block(x.store, x.part || 'Stop', x.kind, x.status, btn, drive);
    }).join('')}
    <div class="foot">
      ${canEdit && !past && d.store ? `<select data-addstop="${i}" aria-label="Add a stop on ${esc(DAY_LONG[dow(d.date)])}"><option value="" disabled selected>+ Add a stop</option>${(() => { const from = dayStores(d).slice(-1)[0]; return stores.filter(s => !dayStores(d).includes(s)).map(s => ({ s, m: driveMin(from, s) })).sort((a, b) => (a.m ?? 999) - (b.m ?? 999)).map(o => `<option value="${esc(o.s)}">${esc(o.s)}${o.m != null ? ` (~${driveText(o.m)})` : ''}</option>`).join(''); })()}</select>` : ''}
      ${canEdit && d.status !== 'done' && !past ? `<select data-swap="${i}" aria-label="Change ${multi ? 'the morning store' : 'store'} for ${esc(DAY_LONG[dow(d.date)])}"><option value="" disabled selected>Change ${multi ? 'AM store' : 'store'}</option>${stores.filter(s => s !== d.store).map(s => `<option>${esc(s)}</option>`).join('')}</select>` : ''}
    </div>
  </article>`;
}

// ---------------------------------------------------------------- a visit (also the store detail page)
// ---------------------------------------------------------------- Smart Scheduler: who's working
// Reads the posted schedules straight from the Smart Scheduler. Each Market Leader connects once with
// their Smart Scheduler sign-in (same email and password as Order Verification); it stays connected.
const SCHED_CFG = { apiKey: 'AIzaSyDLeBfi4LrYtkXxS9fh9BPf40NcPIsIqQA', authDomain: 'smart-scheduler-1915.firebaseapp.com', projectId: 'smart-scheduler-1915',
  storageBucket: 'smart-scheduler-1915.firebasestorage.app', messagingSenderId: '1678890298', appId: '1:1678890298:web:483e73dbcf5b7f7875ac03' };
const SCH = { ready: null, user: null, weeks: {}, hours: null };
const SCH_ROLES = { C: 'Consultant', PT: 'Part-time Consultant', GM: 'General Manager', L: 'Assistant General Manager', LSL: 'Lead Selling Leader', ASL: 'Assistant Selling Leader', KH: 'Key Holder', CSR: 'Guest Solutions', CSRK: 'Guest Solutions key holder', MM: 'Market Manager' };
const schedSlug = st => String(st).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
const schedStore = n => { let c = canonicalStore(n); if (!isKnownStore(c)) c = canonicalStore(String(n).replace(/^Jacksonville\s+/i, '')); return c; };
function schedInit() {
  if (SCH.ready) return SCH.ready;
  SCH.ready = (async () => {
    if (DEMO) { SCH.user = { email: 'demo' }; return SCH; }
    const [{ initializeApp, getApps }, A, F2] = await Promise.all([import(FB + 'firebase-app.js'), import(FB + 'firebase-auth.js'), import(FB + 'firebase-firestore.js')]);
    const app = getApps().find(a => a.name === 'sched') || initializeApp(SCHED_CFG, 'sched');
    Object.assign(SCH, { A, F: F2, auth: A.getAuth(app), db: F2.getFirestore(app) });
    await new Promise(res => { const un = A.onAuthStateChanged(SCH.auth, u => { SCH.user = u; un(); res(); }); });
    return SCH;
  })().catch(e => { console.warn('scheduler', e); return SCH; });
  return SCH.ready;
}
async function schedSignIn(email, pw) { await schedInit(); await SCH.A.signInWithEmailAndPassword(SCH.auth, email, pw); SCH.user = SCH.auth.currentUser; SCH.weeks = {}; }
// The week's schedules by store (Monday weeks, same as the Smart Scheduler). null when not connected.
async function schedWeek(week) {
  await schedInit(); if (!SCH.user) return null;
  if (SCH.weeks[week]) return SCH.weeks[week];
  if (DEMO) { const out = {}; const keys = ['O', 'C', '', 'O', 'C', 'O', 'S']; STORES.forEach(st => { const ppl = (S.roster || []).filter(r => r.store === st.name); if (ppl.length) out[st.name] = { posted: true, people: ppl.map((r, i) => ({ name: r.name, role: r.title === 'ASM' ? 'L' : 'C', days: keys.map((k, d) => (i + d) % 4 === 3 ? '' : k) })) }; }); SCH.hours = {}; return (SCH.weeks[week] = out); }
  const F2 = SCH.F, out = {};
  try {
    const [q, h] = await Promise.all([F2.getDocs(F2.query(F2.collection(SCH.db, 'rosters'), F2.where('week', '==', week))), SCH.hours ? null : F2.getDocs(F2.collection(SCH.db, 'stores'))]);
    if (h) { SCH.hours = {}; h.docs.forEach(d => { SCH.hours[d.id] = d.data(); }); }
    q.docs.forEach(d => { const r = d.data(), st = schedStore(r.store); if (isKnownStore(st)) out[st] = { ...r, slug: schedSlug(r.store) }; });
  } catch (e) { console.warn('scheduler read', e); return null; }
  return (SCH.weeks[week] = out);
}
// Shift times from the store's hours, the same templates the Smart Scheduler uses on a normal day.
function schedShift(r, di, k, iso) {
  if (!k) return null; if (k === 'PTO') return { pto: true };
  const sd = SCH.hours?.[r.slug]; let day = sd?.days?.[di] || {};
  if (sd?.cur && sd.switchOn && iso < sd.switchOn) day = { ...day, ...(sd.cur[di] || {}) };
  if (day.closed) return null;
  if (k === 'P') return { name: 'Part-time', in: 11.5, out: 16.5 };
  const m = /^X:([\d.]+)-([\d.]+)$/.exec(k); if (m) return { name: 'Custom', in: +m[1], out: +m[2] };
  const o = day.open ?? 10, c = day.close ?? 19, O = { k: 'O', name: 'Opener', in: 9.5, out: 18 };
  const T = di === 6 ? (c <= 19 ? [{ k: 'S', name: 'Sunday', in: o - 0.5, out: c }] : [{ k: 'S', name: 'Sun open', in: o - 0.5, out: 19 }, { k: 'S2', name: 'Sun close', in: c - 8.5, out: c }])
    : c <= 19 ? [O, { k: 'C', name: 'Closer', in: 10.5, out: 19 }, { k: 'F', name: 'Full day', in: 9.5, out: 19 }]
    : c === 20 ? [O, { k: 'C', name: 'Closer', in: 11.5, out: 20 }] : [O, { k: 'M', name: 'Mid', in: 11.5, out: 20 }, { k: 'C', name: 'Closer', in: 12.5, out: 21 }];
  return T.find(t => t.k === k) || { name: k };
}
const clk = t => { const h = Math.floor(t), m = Math.round((t - h) * 60); return `${((h + 11) % 12) + 1}${m ? ':' + String(m).padStart(2, '0') : ''}${h < 12 ? 'a' : 'p'}`; };
const shiftText = sh => sh?.pto ? 'PTO' : sh?.in != null ? `${clk(sh.in)} to ${clk(sh.out)}` : sh?.name || '';
// Who is on the schedule at a store on a date: { posted, on: [...], pto: [...], off: [...] }, or null.
function whoOn(store, date) {
  const r = SCH.weeks[weekStartOf(date)]?.[store]; if (!r) return null;
  const di = (dow(date) + 6) % 7, on = [], pto = [], off = [];
  (r.people || []).forEach(p => { const sh = schedShift(r, di, p.days?.[di], date); const x = { name: p.name, cid: cidOf(p.name), role: p.role, roleLabel: SCH_ROLES[p.role] || p.role, lead: !['C', 'PT', 'CSR'].includes(p.role), sh };
    (sh?.pto ? pto : sh ? on : off).push(x); });
  on.sort((a, b) => (b.lead - a.lead) || (a.sh.in ?? 99) - (b.sh.in ?? 99));
  return { posted: r.posted !== false, on, pto, off };
}
function whoPanel(store, date) {
  if (!SCH.user) return `<div class="budgetline whoon"><b>Who's working:</b> connect the Smart Scheduler to see who's on the schedule before you walk in. <button type="button" class="btn tiny" data-schedconnect>Connect</button>
    <form id="schedform" hidden style="margin-top:8px"><p class="small" style="margin:0 0 6px">Use your Smart Scheduler sign-in (the same email and password as Order Verification). You only do this once on this device.</p>
    <div class="row" style="gap:6px;flex-wrap:wrap"><input type="email" id="schedem" value="${esc(S.user?.email || '')}" autocomplete="username" style="flex:1;min-width:200px"><input type="password" id="schedpw" placeholder="Smart Scheduler password" autocomplete="current-password" style="flex:1;min-width:160px"><button class="btn tiny primary" type="submit">Connect</button></div><p class="small err" id="schederr"></p></form></div>`;
  const w = whoOn(store, date);
  if (!w) return `<div class="budgetline whoon"><b>Who's working:</b> no schedule in the Smart Scheduler for ${esc(store)} this week yet.</div>`;
  const line = list => list.map(p => `${esc(titleName(p.name))} <span class="small muted">${esc(shiftText(p.sh))}</span>`).join(' · ');
  const leads = w.on.filter(p => p.lead), team = w.on.filter(p => !p.lead);
  return `<div class="budgetline whoon"><b>Who's working ${date === today() ? 'today' : dayLabel(date)} at ${esc(store)}</b>${w.posted ? '' : ' <span class="pill off">Draft schedule</span>'}
    <p class="small" style="margin:4px 0 0"><b>Leaders and key holders (${leads.length}):</b> ${leads.length ? line(leads) : '<span class="warn">none scheduled</span>'}</p>
    <p class="small" style="margin:2px 0 0"><b>Consultants (${team.length}):</b> ${team.length ? line(team) : '<span class="warn">none scheduled</span>'}</p>
    ${w.pto.length ? `<p class="small muted" style="margin:2px 0 0">PTO: ${w.pto.map(p => esc(titleName(p.name))).join(', ')}</p>` : ''}</div>`;
}
// One line for the daily brief: who leads the floor and how many consultants are on.
function whoLine(store, date) {
  if (!SCH.user) return '';
  const w = whoOn(store, date); if (!w) return `<p class="small muted" style="margin:4px 0 0">No schedule in the Smart Scheduler yet.</p>`;
  const leads = w.on.filter(p => p.lead), team = w.on.filter(p => !p.lead);
  return `<p class="small" style="margin:4px 0 0"><b>On today:</b> ${leads.length ? leads.map(p => `${esc(titleName(p.name).split(' ')[0])} (${esc(p.role === 'MM' ? 'MM' : p.roleLabel)}, ${esc(shiftText(p.sh))})`).join(', ') : '<span class="warn">no leader scheduled</span>'} · ${team.length} consultant${team.length === 1 ? '' : 's'}${w.posted ? '' : ' · draft'}</p>`;
}
function wireWho(root, again) {
  root.querySelectorAll('[data-schedconnect]').forEach(b => b.onclick = () => { const f = root.querySelector('#schedform'); if (f) { f.hidden = false; root.querySelector('#schedpw')?.focus(); } });
  const f = root.querySelector('#schedform'); if (!f) return;
  f.onsubmit = async e => { e.preventDefault(); const er = root.querySelector('#schederr'); er.textContent = 'Connecting…';
    try { await schedSignIn(root.querySelector('#schedem').value.trim().toLowerCase(), root.querySelector('#schedpw').value); toast('Smart Scheduler connected.'); again(); }
    catch (x) { er.textContent = /invalid|wrong|not-found/.test(x?.code || '') ? "That email and password don't match the Smart Scheduler. Use the same sign-in as Order Verification, or reset it in the Smart Scheduler." : (x?.message || 'Could not connect.'); } };
}
async function openVisit(x) {
  S.visit = x; S.V = null; S.vPhotos = []; S.vPlans = [];
  try { S.vPlans = (await Promise.all([S.be.plan(x.email, weekStartOf(x.date)), S.be.plan(x.email, addDays(weekStartOf(x.date), 7))])).filter(Boolean); } catch (e) {}
  // In person at a store that isn't the one planned for that day (or on a day off): ask why, move the
  // plan, and alert the VP. Opening a visit that's already started never asks again.
  if (!x.remote && !x.reasonOk && needsReason() && x.email === S.user.email) {
    const plan = S.vPlans.find(p => p.weekStart === weekStartOf(x.date));
    const idx = plan ? plan.days.findIndex(d => d.date === x.date) : -1;
    const day = idx >= 0 ? plan.days[idx] : null;
    const vid = `${x.email}_${x.date}_${slug(x.store)}`;
    const started = S.visits.some(v => v.id === vid) || localGet(vid);
    const planned = day?.store || (plan && idx < 0 ? 'Off' : null);
    if (plan && planned && !dayStores(day).includes(x.store) && !(day?.status === 'done' && !day.stops?.length) && !started) {
      S.visit = null; window.scrollTo(0, 0);
      const v = $('#view');
      v.innerHTML = `<div class="panel"><p class="eyebrow">${esc(dayLabel(x.date))}</p>
        <h2 style="margin:0 0 6px">${planned === 'Off' ? `This is your day off` : `Your plan has you at ${esc(planned)}`}</h2>
        <p style="margin:0 0 12px">You're opening a visit at <b>${esc(x.store)}</b>.${planned !== 'Off' ? ` Going to ${esc(planned)} instead? ` : ''}</p>
        ${planned !== 'Off' ? `<div class="row" style="margin:0 0 8px"><button class="btn" id="goplan" type="button">Open ${esc(planned)} instead</button><button class="btn primary" id="addplan" type="button">Add ${esc(x.store)} as a stop today${(() => { const m = driveMin(dayStores(day).slice(-1)[0], x.store); return m != null ? ` (~${driveText(m)} drive)` : ''; })()}</button></div><p class="small" style="margin:0 0 12px">Adding a stop keeps ${esc(dayText(day))} on your day. Replacing it takes a reason below.</p>` : ''}
        <div id="visitreason"></div></div>`;
      const gp = $('#goplan'); if (gp) gp.onclick = () => openVisit({ ...x, store: planned, kind: day.kind, dayIndex: idx });
      const ap = $('#addplan'); if (ap) ap.onclick = async () => {
        day.stops = [...(day.stops || []), { store: x.store, part: day.stops?.length ? 'Stop ' + (day.stops.length + 2) : 'PM', kind: 'first', status: 'planned' }];
        if (!day.part) day.part = 'AM';
        plan.pivots = [...(plan.pivots || []), { date: x.date, from: planned, to: dayText(day), reason: 'Added a stop', at: new Date().toISOString(), by: S.user.email }];
        await S.be.savePlan(plan); toast(`${x.store} added as a stop today.`);
        openVisit({ ...x, reasonOk: true, kind: 'first', dayIndex: idx });
      };
      askReason('#visitreason', `Why ${x.store} instead${planned === 'Off' ? ' of a day off' : ` of ${dayText(day)}`}?`, async (reason, note) => {
        if (day) {
          plan.pivots = [...(plan.pivots || []), { date: x.date, from: day.store, to: x.store, reason: `${reason}${note ? ': ' + note : ''}`, at: new Date().toISOString(), by: S.user.email }];
          Object.assign(day, { store: x.store, kind: plan.days.some((d, j) => j !== idx && d.store === x.store && d.date < x.date) ? 'second' : 'first', score: S.scores[x.store]?.score ?? null, stops: [], anchor: false }); delete day.part;
          plan.calls = (plan.calls || []).filter(c => c !== x.store);
          await S.be.savePlan(plan);
        }
        const who = S.users.find(u => u.email === x.email) || S.user;
        await sendScheduleAlert(who, plan.weekStart, reason, note, [{ date: x.date, from: planned, to: x.store }], 'visit');
        toast(`${vpNames()} was alerted. Opening ${x.store}.`);
        openVisit({ ...x, reasonOk: true, kind: day ? day.kind : 'drop-in', dayIndex: day ? idx : undefined });
      }, () => { S.tab = 'week'; renderShell(); });
      return;
    }
  }
  renderShell(); window.scrollTo(0, 0);
  try { if (x.remote) return; S.vPhotos = await S.be.photos(`${x.email}_${x.date}_${slug(x.store)}${x.remote ? '_remote' : ''}`); if (S.vPhotos.length && S.visit === x) viewVisit(); } catch (e) {}
}
function tileFor(label, value, sub, cls) { return `<div class="tile ${cls || ''}"><div class="tl">${esc(label)}</div><div class="tv">${value}</div><div class="tg">${sub || ''}</div></div>`; }
const PERIODS = [['day', 'Prior day'], ['wtd', 'Week to date'], ['mtd', 'Month to date']];
function storeTiles(snap, period = 'mtd') {
  const m = snap?.[period] || snap?.mtd;
  if (!m) return '<p class="muted">No numbers for this store in the latest daily report.</p>';
  const vsCls = v => v == null ? '' : v >= 0 ? 'green' : v >= -10 ? 'amber' : 'red';
  const goalCls = (v, g, lower) => v == null ? '' : status({ lower }, v, g);
  const money = v => v == null ? '--' : '$' + Math.round(v).toLocaleString('en-US');
  const p1 = v => v == null ? '--' : v.toFixed(1) + '%';
  const tiles = [];
  tiles.push(tileFor('Sales', money(m.k.netSales), m.vsBud.netSales != null ? `${pct(m.vsBud.netSales)} to budget` : '', vsCls(m.vsBud.netSales)));
  tiles.push(tileFor('SPG w/ cancellations', m.k.spg != null ? '$' + m.k.spg.toFixed(2) : '--', m.vsLy.spg != null ? `${pct(m.vsLy.spg)} vs LY` : m.vsBud.spg != null ? `${pct(m.vsBud.spg)} to budget` : '', vsCls(m.vsLy.spg ?? m.vsBud.spg)));
  tiles.push(tileFor('Close rate', p1(m.k.closeRate), m.vsBud.closeRate != null ? `${m.vsBud.closeRate > 0 ? '+' : ''}${Math.round(m.vsBud.closeRate)} bps to budget` : '', m.vsBud.closeRate == null ? '' : m.vsBud.closeRate >= 0 ? 'green' : m.vsBud.closeRate > -200 ? 'amber' : 'red'));
  tiles.push(tileFor('Traffic', m.k.traffic != null ? Math.round(m.k.traffic).toLocaleString('en-US') : '--', m.vsLy.traffic != null ? `${pct(m.vsLy.traffic)} vs LY` : '', ''));
  tiles.push(tileFor('Sales / hour', money(m.k.sph), `goal $${STORE_GOALS.sph}`, goalCls(m.k.sph, STORE_GOALS.sph)));
  tiles.push(tileFor('Avg ticket', money(m.k.avgTicket), '', ''));
  tiles.push(tileFor('Finance % of sales', p1(m.k.financePct), `goal ${STORE_GOALS.financePct}%`, goalCls(m.k.financePct, STORE_GOALS.financePct)));
  tiles.push(tileFor('Apps to traffic', p1(m.k.appsToTraffic), `goal ${STORE_GOALS.appsToTraffic}%`, goalCls(m.k.appsToTraffic, STORE_GOALS.appsToTraffic)));
  tiles.push(tileFor('Bedding % of sales', p1(m.k.beddingPct), `goal ${STORE_GOALS.beddingPct}%`, goalCls(m.k.beddingPct, STORE_GOALS.beddingPct)));
  tiles.push(tileFor('Protection attach', p1(m.k.protectionAttach), `goal ${STORE_GOALS.protectionAttach}%`, goalCls(m.k.protectionAttach, STORE_GOALS.protectionAttach)));
  tiles.push(tileFor('Delivery % of sales', p1(m.k.deliveryPct), `goal ${STORE_GOALS.deliveryPct}%`, goalCls(m.k.deliveryPct, STORE_GOALS.deliveryPct)));
  tiles.push(tileFor('Cancellations', p1(m.k.cancelPct), `of gross, goal ${STORE_GOALS.cancelPct}% or less`, goalCls(m.k.cancelPct, STORE_GOALS.cancelPct, true)));
  // Open carts are a snapshot (as of the cart file), so they read the same in every period.
  const cx = S.carts?.stores?.[m.store];
  const cartTiles = cx ? [
    tileFor('Open carts', cx.n.toLocaleString('en-US'), `as of ${shortDate(S.carts.date)}`, ''),
    tileFor('Cart value (est.)', money(cx.value), 'money already in the building', ''),
    tileFor('Started this week', String(cx.wk), money(cx.wkValue) + ' est.', ''),
    tileFor('Due a call today', String(cx.due), 'day 1, 3 or 7 follow-up', cx.due ? 'amber' : 'green'),
    tileFor('Over 2 weeks old', String(cx.old), 'call or close out', cx.old ? 'amber' : '')
  ] : [];
  return `<div class="tiles">${tiles.join('')}</div>${cartTiles.length ? `<p class="eyebrow" style="margin:12px 0 6px">Open carts</p><div class="tiles">${cartTiles.join('')}</div>` : ''}`;
}
const Speech = window.SpeechRecognition || window.webkitSpeechRecognition;
let rec = null, recBtn = null;
const micBtn = id => Speech ? `<button type="button" class="mic" data-mic="${id}" aria-label="Talk to text"><svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><path fill="currentColor" d="M12 14a3 3 0 0 0 3-3V5a3 3 0 1 0-6 0v6a3 3 0 0 0 3 3zm5-3a5 5 0 0 1-10 0H5a7 7 0 0 0 6 6.92V21h2v-3.08A7 7 0 0 0 19 11h-2z"/></svg><span>Talk</span></button><span class="live"></span>` : '';
function stopMic() { if (rec) { try { rec.stop(); } catch (e) {} } }
function wireMics(root) {
  root.querySelectorAll('[data-mic]').forEach(b => b.onclick = () => {
    if (rec && recBtn === b) return stopMic();
    stopMic();
    const box = $('#' + b.dataset.mic), live = b.parentElement.querySelector('.live');
    rec = new Speech(); recBtn = b;
    rec.lang = 'en-US'; rec.continuous = true; rec.interimResults = true;
    rec.onresult = e => {
      let interim = '';
      for (let i = e.resultIndex; i < e.results.length; i++) {
        const t = e.results[i][0].transcript.trim();
        if (e.results[i].isFinal) { const cur = box.value.replace(/\s+$/, ''); box.value = box.tagName === 'INPUT' ? (cur ? cur + ' ' : '') + t.replace(/[.]$/, '') : (cur ? cur + (/[.!?]$/.test(cur) ? ' ' : '. ') : '') + t.charAt(0).toUpperCase() + t.slice(1); box.dispatchEvent(new Event('input')); }
        else interim += t + ' ';
      }
      if (live) live.textContent = interim;
    };
    rec.onerror = e => {
      if (e.error === 'not-allowed' || e.error === 'service-not-allowed') toast('Microphone is blocked. Allow it for this site in your browser settings, then tap Talk again.', true);
      else if (e.error === 'no-speech') toast('Did not hear anything. Tap Talk and try again.', true);
      else if (e.error !== 'aborted') toast('Talk to text stopped: ' + e.error, true);
    };
    rec.onend = () => { b.classList.remove('on'); b.querySelector('span').textContent = 'Talk'; if (live) live.textContent = ''; rec = null; recBtn = null; };
    try { rec.start(); b.classList.add('on'); b.querySelector('span').textContent = 'Stop'; box.focus(); } catch (x) { toast('Could not start talk to text.', true); rec = null; }
  });
}
// Store leader picker for a visit: the store's leaders from the store leader list (GM first), plus
// ASMs and Sales Leads on the sales roster, then "Someone else" to type a name.
function leaderOptions(store) {
  const list = [], seen = new Set();
  const add = (name, role) => { const n = titleName(String(name || '').trim()); if (!n || seen.has(n.toLowerCase())) return; seen.add(n.toLowerCase()); list.push({ n, r: role }); };
  (S.storeLeaders || []).filter(l => l.store === store).forEach(l => add(l.name, l.role || 'Leader'));
  (S.roster || []).filter(r => r.store === store && ['ASM', 'Sales Lead'].includes(r.title)).forEach(r => add(r.name, r.title === 'ASM' ? 'Assistant Selling Manager' : r.title));
  return list;
}
function leaderField(id, label, value, field, dis, store) {
  const opts = leaderOptions(store), v = String(value || '').trim();
  if (!opts.length) return fieldInput(id, label, v, field, dis, 'placeholder="Leader name" style="width:100%"');
  const known = opts.some(o => o.n.toLowerCase() === v.toLowerCase());
  return `<label for="${id}_sel" style="margin:0 0 4px">${esc(label)}<select id="${id}_sel" data-leadsel="${id}" ${dis}>
      ${v ? '' : '<option value="" selected disabled>Pick the leader</option>'}${opts.map(o => `<option value="${esc(o.n)}" ${o.n.toLowerCase() === v.toLowerCase() ? 'selected' : ''}>${esc(o.n)} · ${esc(o.r)}</option>`).join('')}
      <option value="__other" ${v && !known ? 'selected' : ''}>Someone else…</option></select></label>
    <div data-leadother="${id}" ${v && !known ? '' : 'hidden'} style="margin-top:6px">${fieldInput(id, 'Their name', v, field, dis, 'placeholder="Leader name" style="width:100%"')}</div>`;
}
function wireLeaderPicks(root) {
  root.querySelectorAll('[data-leadsel]').forEach(sel => sel.onchange = () => {
    const id = sel.dataset.leadsel, inp = $('#' + id), other = root.querySelector(`[data-leadother="${id}"]`);
    if (sel.value === '__other') { other.hidden = false; inp.value = ''; inp.focus(); }
    else { other.hidden = true; inp.value = sel.value; }
    inp.dispatchEvent(new Event('input'));
    // The store leader notes follow the leader picked for the win, until they're set on their own.
    if (id === 'lwname' && sel.value !== '__other') { const lc = $('#lcname_sel'); if (lc && !lc.dataset.touched) { lc.value = sel.value; lc.dispatchEvent(new Event('change')); delete lc.dataset.touched; } }
    if (id === 'lcname') sel.dataset.touched = '1';
  });
}
// A one-line field with talk to text: type it or say it.
const fieldInput = (id, label, value, field = '', dis = '', extra = '') => `
  <div class="fieldhead"><label for="${id}">${esc(label)}</label>${dis ? '' : micBtn(id)}</div>
  <input id="${id}" value="${esc(value || '')}" ${field ? `data-field="${field}"` : ''} ${dis} ${extra}>`;
const fieldBox = (id, label, value, rows = 3, hint = '', field = '', dis = '') => `
  <div class="fieldhead"><label for="${id}">${esc(label)}</label>${dis ? '' : micBtn(id)}</div>
  ${hint ? `<p class="small" style="margin:0 0 4px">${esc(hint)}</p>` : ''}
  <textarea id="${id}" rows="${rows}" ${field ? `data-field="${field}"` : ''} ${dis}>${esc(value || '')}</textarea>`;

// ---------------------------------------------------------------- remote: does the leader know the floor?
// The leader answers from memory; where the daily report has the number, it shows next to it.
const FLOOR_Q = [
  { k: 'guests', l: 'Guests yesterday', rep: snap => snap?.day?.k?.traffic },
  { k: 'carts', l: 'Carts started yesterday', rep: () => null },
  { k: 'apps', l: 'Finance apps yesterday', rep: snap => snap?.day?.k?.traffic != null && snap?.day?.k?.appsToTraffic != null ? Math.round(snap.day.k.traffic * snap.day.k.appsToTraffic / 100) : null },
  { k: 'bundles', l: 'Sales with the full bundle yesterday', rep: () => null },
  { k: 'coachings', l: '1 on 1 coachings yesterday (and with who)', rep: () => null }
];
function floorBlock(V, x, snap, dis, tri) {
  const f = V.floor || {};
  return `<p class="small" style="margin:0 0 8px">Ask each one and type what the leader says. Names count: "who got the carts" tells you more than "a few".</p>
    <div class="scroller"><table class="grid"><thead><tr><th>Ask</th><th>Leader says</th><th>Report</th></tr></thead><tbody>
    ${FLOOR_Q.map(q => { const r = q.rep(snap), said = numOf(f[q.k]); const off = r != null && said != null && Math.abs(said - r) > Math.max(2, r * 0.15);
      return `<tr><td>${esc(q.l)}</td><td><input id="fl_${q.k}" data-field="floor.${q.k}" value="${esc(f[q.k] || '')}" ${dis} style="width:100%;min-width:120px" placeholder="${q.k === 'coachings' ? 'e.g. 2: Maria, Tyler' : 'e.g. 12'}"></td><td class="num ${off ? 'bad' : ''}">${r != null ? Math.round(r) : '<span class="muted">not in report</span>'}</td></tr>`; }).join('')}
    </tbody></table></div>
    <div class="item"><div class="txt">Answered without looking it up</div>${tri('floor.knew', f.knew)}</div>
    ${fieldBox('floornotes', 'What this tells you about the floor', f.notes, 2, '', 'floor.notes', dis)}`;
}

// ---------------------------------------------------------------- the intent of a visit
// Why the Market Leader is going, in plain words: the purpose, the lever, the people behind the gap,
// the behaviors to coach, and who models it. We change the outcome through people and behaviors.
function visitIntent(store, { anchor, kind, date, remote } = {}) {
  const snap = S.daily?.stores?.[store], m = snap?.mtd, sc = S.scores[store];
  const lk = suggestLever(m), L = lk ? leverStatus(m).find(l => l.key === lk) : null;
  const weak = L ? L.inputs.filter(i => i.metric && i.ratio != null && i.ratio < 1).sort((a, b) => a.ratio - b.ratio)[0] : null;
  const floorIn = L ? L.inputs.find(i => !i.metric) : null;
  const people = S.rsa?.people || [], pace = paceFactor(S.rsa?.to);
  const drag = weak && STORE_TO_RSA[weak.metric] ? draggers(people, store, weak.metric, DEFAULT_GOALS, pace, 2) : [];
  const help = weak && STORE_TO_RSA[weak.metric] ? helpers(people, store, weak.metric, DEFAULT_GOALS, pace, 1) : [];
  const team = S.teams?.[store];
  const see = [...drag.map(d => ({ cid: d.cid, n: titleName(d.name), why: `${(PLAIN[d.key] || d.key)} ${fmtMetric(d.key, d.value)} vs ${fmtMetric(d.key, d.goal)}` })),
    ...(team?.below || []).map(r => ({ cid: r.cid, n: titleName(r.name), why: `$${Math.round(r.sph)} an hour, under the minimum` })),
    ...(team?.slipping || []).map(r => ({ cid: r.cid, n: titleName(r.name), why: `slipping this week` }))]
    .filter((x, i, a) => a.findIndex(y => y.cid === x.cid) === i).slice(0, 3);
  const prior = S.visits.filter(v => v.store === store && v.status !== 'draft' && (!date || v.date < date)).sort((a, b) => b.date.localeCompare(a.date))[0];
  const open = prior ? (prior.actions || []).filter(hasCommitment).filter(a => autoFollow(a, snap, {})?.v !== 'yes').length : 0;
  const purpose = remote ? `Remote coaching: go over the numbers with the leader, coach the people behind the gap through the leader, and role-play the play with the leader or an associate on video.${open ? ` Check in on the ${open} open commitment${open > 1 ? 's' : ''} from the last visit.` : ''}`
    : anchor ? 'Anchor morning: be on the floor at open, set the tone, and run the play with the team.'
    : open ? `Follow up and inspect: ${open} commitment${open > 1 ? 's' : ''} from the last visit ${open > 1 ? "aren't" : "isn't"} there yet.`
    : kind === 'second' ? 'Second visit this week: inspect the plan you set and the behaviors behind it.'
    : 'Set the plan: find the gap, coach the people behind it, leave with commitments.';
  const behaviors = [weak && { l: weak.label, b: weak.behavior.split('. ')[0] }, floorIn && { l: floorIn.label, b: floorIn.behavior.split('. ')[0] }].filter(Boolean);
  return { purpose, L, weak, why: (sc?.parts || []).slice(0, 2).map(p => p.text), see, model: help[0] ? titleName(help[0].name) : null, behaviors };
}
function intentHtml(it, compact) {
  if (!it) return '';
  return `<div class="intent ${compact ? 'compact' : ''}">
    ${compact ? '' : '<p class="eyebrow" style="margin:0 0 2px">Intent of this visit</p>'}
    <p class="ipurpose">${esc(it.purpose)}</p>
    ${it.why.length && !compact ? `<ul class="whylist">${it.why.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : it.why.length ? `<p class="small" style="margin:0">${esc(it.why[0])}</p>` : ''}
    ${it.L ? `<p class="small" style="margin:0"><b>Lever:</b> ${esc(it.L.label)} (${esc(fmtMetric(it.L.key, it.L.value))} vs ${esc(fmtMetric(it.L.key, it.L.goal))} goal).${compact ? '' : ' The number moves through people and behaviors, not by talking about the number.'}</p>` : ''}
    ${it.see.length ? `<p class="small" style="margin:0"><b>People:</b> ${it.see.map(p => `${esc(p.n)} <span class="muted">(${esc(p.why)})</span>`).join(', ')}</p>` : ''}
    ${it.behaviors.length ? `<p class="small" style="margin:0"><b>Behaviors to coach:</b> ${it.behaviors.map(b => esc(String(b.b).replace(/[.\s]+$/, ''))).join('. ')}.</p>` : ''}
    ${it.model ? `<p class="small" style="margin:0"><b>Model:</b> ${esc(it.model)} shows the team how it's done.</p>` : ''}
    ${compact ? '' : '<p class="small" style="margin:4px 0 0"><b>Leave with:</b> a from X to Y commitment from the leader and from each person you coach.</p>'}
  </div>`;
}

// ---------------------------------------------------------------- run the play, and FrontLine IQ, on a visit
function playBlock(V, canLog, dis, tri) {
  const o = activeOffer();
  const ex = offerMath(3000);
  return `<div class="play">${PLAY.map((p, i) => `<div class="pstep"><span class="pnum">${i + 1}</span><div><b>${esc(p.t)}</b><p class="small" style="margin:2px 0 0">${esc(p.d)}</p></div></div>`).join('')}</div>
    ${o ? `<div class="offer"><p class="eyebrow" style="margin:0">${esc(o.name)}${o.end ? ` · through ${esc(shortDate(o.end))}` : ''}</p>
      <ul class="blist small" style="margin:6px 0">${(o.lines || []).map(l => `<li>${esc(l)}</li>`).join('')}</ul>
      <p class="small" style="margin:0"><b>Say it in dollars:</b> on a $3,000 room, financing or the bundle saves $${ex.one}. Financing AND the bundle saves $${ex.both}.</p>
      ${o.fine ? `<p class="small muted" style="margin:4px 0 0">${esc(o.fine)}</p>` : ''}</div>` : ''}
    ${V.remote ? `<p class="small" style="margin:12px 0 6px"><b>Remote: you can't watch the floor.</b> Ask the leader for specifics (names and counts, not "we're good"), and role-play the play on video, with the leader or with an associate. Connection comes first.</p>`
      : `<p class="small" style="margin:12px 0 6px"><b>Watch one live guest, or run it as a practice.</b> Connection comes first.</p>`}
    ${(V.remote ? PLAY_CHECKS_REMOTE : PLAY_CHECKS).map((t, i) => `<div class="item"><div class="txt">${esc(t)}</div>${tri(`play.${i}`, V.play?.[i])}</div>`).join('')}
    ${fieldBox('playnotes', V.remote ? 'What the leader told you, and how the role-play went' : 'What you saw', V.playNotes, 2, '', 'playNotes', dis)}`;
}
function fliqBlock(V, canLog, dis, tri) {
  const f = V.fliq || {};
  return `<p class="small" style="margin:0 0 8px"><b>Today's focus:</b> ${esc(FLIQ_DAILY[dow(V.date)])}</p>
    <div class="two">
      <div>${fieldInput('fliqUsing', 'Associates who used it today', f.using, 'fliq.using', dis, 'inputmode="numeric" placeholder="For example 4" style="width:100%"')}</div>
      <div>${fieldInput('fliqFloor', 'Associates on the floor', f.floor, 'fliq.floor', dis, 'inputmode="numeric" placeholder="For example 6" style="width:100%"')}</div>
    </div>
    ${(V.remote ? FLIQ_CHECKS_REMOTE : FLIQ_CHECKS).map((t, i) => `<div class="item"><div class="txt">${esc(t)}</div>${tri(`fliq.c.${i}`, f.c?.[i])}</div>`).join('')}
    ${fieldBox('fliqnotes', 'What FrontLine IQ flagged, and what you did with it', f.notes, 3, 'Which associates, what it coached them on, and what you saw on the floor.', 'fliq.notes', dis)}`;
}

// ---------------------------------------------------------------- levers on a visit
// Targets for inputs counted on the floor (no number in the daily report).
const COUNT_TARGET = { value: 'Every guest, before price', options: 'Every guest', bundle: 'Every sale', cart: '8 of 10 guests', pieces: 'One more piece per ticket', quality: 'Every guest', price: 'No discount without a leader' };
// The inputs to coach by default once a lever is picked: the ones below goal first, then a floor input.
function defaultInputs(L) {
  const below = L.inputs.filter(i => i.ratio != null && i.ratio < 1).sort((a, b) => a.ratio - b.ratio).map(i => i.key);
  const floor = L.inputs.filter(i => !i.metric).map(i => i.key);
  return [...new Set([...below, ...floor, ...L.inputs.map(i => i.key)])].slice(0, 2);
}
function leverBlock(V, x, snap, people, canLog, dis) {
  const pL = snap?.mtd;
  const LS = leverStatus(pL), sugg = suggestLever(pL);
  if (!LS.length) return '<p class="small muted">No store numbers yet.</p>';
  const fv = (k, v) => v == null ? '--' : fmtMetric(k, v);
  const cards = LS.map(L => { const on = V.lever === L.key, below = L.ratio != null && L.ratio < 1;
    return `<div class="fcard pick ${on ? 'on' : ''}"><span class="row" style="justify-content:space-between"><span class="eyebrow">${L.key === sugg ? 'Suggested from the numbers' : 'Lever'}</span>${canLog ? `<button type="button" class="btn tiny ${on ? 'primary' : ''}" data-lpick="${L.key}" ${dis}>${on ? 'Pulling this ✓' : 'Pull this lever'}</button>` : ''}</span>
      <b class="fl">${esc(L.label)}</b><span class="small">Now <b class="${below ? 'bad' : 'good'}">${esc(fv(L.key, L.value))}</b>${L.goal != null ? `, goal ${esc(fv(L.key, L.goal))}` : ''} this month. ${esc(L.why)}</span></div>`; }).join('');
  const L = LS.find(l => l.key === V.lever);
  const chosen = V.leverInputs || [];
  const inputs = L ? `<h3 id="lvinputs" style="margin:16px 0 4px">Coach this: what drives ${esc(L.label.toLowerCase())}</h3>
    <p class="small" style="margin:0 0 10px">Coach the inputs, not the outcome. Two are picked from the numbers. Change them if you see something different on the floor.</p>
    <div class="focus">${L.inputs.map(inp => {
      const on = chosen.includes(inp.key), d = drillFor(inp.drill), below = inp.ratio != null && inp.ratio < 1;
      const dr = inp.metric && STORE_TO_RSA[inp.metric] ? draggers(people, x.store, inp.metric, DEFAULT_GOALS, paceFactor(S.rsa?.to), 2) : [];
      const hp = inp.metric && STORE_TO_RSA[inp.metric] ? helpers(people, x.store, inp.metric, DEFAULT_GOALS, paceFactor(S.rsa?.to), 2) : [];
      return `<div class="fcard pick ${on ? 'on' : ''}">
        <span class="row" style="justify-content:space-between"><span class="eyebrow">${inp.metric ? (below ? 'Below goal' : 'At goal') : (V.remote ? 'Ask the leader for the count' : 'Count it on the floor')}</span>${canLog ? `<button type="button" class="btn tiny ${on ? 'primary' : ''}" data-linput="${inp.key}" ${dis}>${on ? 'Coaching this ✓' : 'Coach this'}</button>` : ''}</span>
        <b class="fl">${esc(inp.label)}</b>
        <span class="small">${inp.metric ? `Now <b class="${below ? 'bad' : 'good'}">${esc(fv(inp.metric, inp.value))}</b>, goal ${esc(fv(inp.metric, inp.goal))}.` : `Not in the daily report. ${V.remote ? 'Ask the leader for it' : 'Count it today'}: <b>${esc(inp.count)}</b>, goal ${esc(COUNT_TARGET[inp.key] || '')}.`}${inp.also ? ` ${esc(inp.also)}` : ''}</span>
        <span class="do"><b>Do this:</b> ${esc(inp.behavior)}</span>
        ${inp.fact ? `<span class="small"><b>Why it works:</b> ${esc(inp.fact)}</span>` : ''}
        ${inp.examples ? `<div class="vex">${inp.examples.map(e2 => `<p class="small" style="margin:6px 0 0"><b>${esc(e2.t)}.</b> Say it like: "${esc(e2.say)}" <span class="muted">Why: ${esc(e2.why)}</span></p>`).join('')}</div>` : ''}
        <span class="asks"><span>• ${esc(inp.ask)}</span></span>
        <span class="small"><b>Practice it standing up:</b> ${esc(d.title)}. ${esc(d.guest)}</span>
        ${dr.length ? `<div class="drag"><p class="eyebrow" style="margin:0 0 4px">Who's pulling this down</p>${dr.map(dd => { const added = V.consultants.some(c => c.cid === dd.cid);
          return `<div class="dragrow"><span><b>${esc(titleName(dd.name))}</b> <span class="small">${esc(fmtMetric(dd.key, dd.value))} vs ${esc(fmtMetric(dd.key, dd.goal))} goal${dd.impactText ? ', ' + esc(dd.impactText) : ''}</span></span>${canLog ? `<button type="button" class="btn tiny" data-dragc="${esc(dd.cid)}" data-lever="${inp.metric}" ${added ? 'disabled' : ''}>${added ? 'On this visit' : 'Coach'}</button>` : ''}</div>`; }).join('')}</div>` : inp.metric && S.rsa ? `<p class="small muted" style="margin:0">Nobody on the team is below goal on this one. Coach it with the whole team in the huddle.</p>` : ''}
        ${hp.length ? `<div class="drag help"><p class="eyebrow" style="margin:0 0 4px">Who's carrying this</p>${hp.map(dd => { const added = V.consultants.some(c => c.cid === dd.cid);
          return `<div class="dragrow"><span><b>${esc(titleName(dd.name))}</b> <span class="small">${esc(fmtMetric(dd.key, dd.value))} vs ${esc(fmtMetric(dd.key, dd.goal))} goal</span></span>${canLog ? `<button type="button" class="btn tiny" data-helpc="${esc(dd.cid)}" data-lever="${inp.metric}" ${added ? 'disabled' : ''}>${added ? 'On this visit' : 'Recognize'}</button>` : ''}</div>`; }).join('')}
          <p class="small muted" style="margin:4px 0 0">Have them show the team how they do it in the huddle.</p></div>` : ''}
      </div>`; }).join('')}</div>` : `<p class="small" style="margin:12px 0 0">${sugg ? `The numbers point to <b>${esc(LS.find(l => l.key === sugg).label)}</b>. You know the store: pick the lever you're pulling today.` : 'All three are at goal. Pick the one to push.'}</p>`;
  return `<div class="focus">${cards}</div>${inputs}`;
}

// ---------------------------------------------------------------- the visit
// One page the leader works through in the store, top to bottom. It opens already knowing why they
// are there, what to coach and who to see. Everything autosaves as a draft; Submit closes it out.
const V_OPEN = new Set(['why', 'win', 'follow', 'floor', 'play', 'fliq', 'lever', 'focus', 'people', 'photos', 'action', 'leadercommit', 'el2', 'el3']);
const PHOTO_MAX = 30;
const PHOTO_ELS = ['visual', 'facilities'];   // every area in these gets a photo, asked for right where it's scored
// The areas that need a photo on every in-person visit, so every Market Leader walks the same store.
// Items that are the same physical area as a Visual walk area share its photo, so nobody takes it twice.
const PHOTO_SAME = {
  'Bathrooms spotless': 'Restrooms', 'Backroom clean and safe': 'Backroom / Warehouse',
  'Lighting fully functional, storefront and windows clean': 'Front Entrance and Windows',
  'Markdown and clearance strategy: right product, right time, right place': 'Clearance / Outlet',
  'Accessories and attachments available to complete the sale': 'Occasional / Accents / Accessories'
};
const photoKey = (el, item) => PHOTO_SAME[item] ? { el: 'visual', item: PHOTO_SAME[item] } : { el, item };
// Unique photo asks per element: shared items count only where the area is walked (Visual).
const photoAreas = e => PHOTO_ELS.includes(e.key) ? (e.aor ? AORS : e.items.filter(t => !PHOTO_SAME[t])) : [];
const REMOTE_TYPES = ['Phone call', 'Video call', 'Teams or text'];
const DRILL_KEYS = ['sph', 'closeRate', 'avgTicket', 'effMargin', 'financePct', 'appsToTraffic', 'beddingPct', 'protectionPct', 'deliveryPct', 'cancelPct'];
function blankVisit(x, who) {
  return { id: `${x.email}_${x.date}_${slug(x.store)}${x.remote ? '_remote' : ''}`, remote: !!x.remote, email: x.email, name: who.name || x.email, role: who.role, store: x.store, date: x.date,
    kind: x.remote ? 'remote' : x.kind || 'drop-in', vtype: x.remote ? 'Phone call' : kindToType(x.kind), status: 'draft', leaderWin: { name: '', text: '' }, follow: {}, focus: null,
    consultants: [], checks: {}, segs: {}, aor: {}, elNotes: {}, actions: [{}, {}, {}], reflection: '', working: '', notes: '' };
}
const localKey = id => `1915fl_draft_${id}`;
function localGet(id) { try { return JSON.parse(localStorage.getItem(localKey(id)) || 'null'); } catch (e) { return null; } }
function localSet(v) { try { localStorage.setItem(localKey(v.id), JSON.stringify(v)); } catch (e) {} }
function localDrop(id) { try { localStorage.removeItem(localKey(id)); } catch (e) {} }

async function viewVisit() {
  const x = S.visit, v = $('#view');
  const who = S.users.find(u => u.email === x.email) || S.user;
  const id = `${x.email}_${x.date}_${slug(x.store)}${x.remote ? '_remote' : ''}`;
  const saved = S.visits.find(y => y.id === id);
  const local = localGet(id);
  // Take whichever copy is newer: the one on this device (spotty store wifi) or the saved one.
  const V = S.V && S.V.id === id ? S.V : structuredClone((local && (!saved || (local.at || '') > (saved.at || ''))) ? local : saved || blankVisit(x, who));
  S.V = V;
  if (!V.actions?.length) V.actions = [{}, {}, {}];
  for (const k of ['consultants']) if (!Array.isArray(V[k])) V[k] = [];
  for (const k of ['checks', 'segs', 'aor', 'elNotes', 'follow']) if (!V[k] || typeof V[k] !== 'object') V[k] = {};
  const canLog = (x.email === S.user.email || isAdmin()) && x.date <= today();
  const later = x.date > today();
  const snap = S.daily?.stores?.[x.store];
  const sc = S.scores[x.store];
  const env = environment(snap);
  const period = S.vPeriod || (snap?.day ? 'day' : 'mtd');
  const prior = S.visits.filter(y => y.store === x.store && y.date < x.date && y.id !== id && y.status !== 'draft').sort((a, b) => b.date.localeCompare(a.date))[0]
    || S.visits.filter(y => y.store === x.store && y.date < x.date && y.id !== id).sort((a, b) => b.date.localeCompare(a.date))[0];
  const priorSum = prior ? visitSummary(prior) : null;
  const priorRaw = prior ? (prior.actions || []).filter(hasCommitment) : [];
  const focusAll = storeFocus(snap, 4);
  // One place to coach: pick the lever, and the two inputs under it are the two things the team works on.
  { const LSx = leverStatus(snap?.mtd);
    const inpKey = inp => inp.metric || inp.key, matches = (inp, k) => inp.metric === k || inp.alt === k;
    // Opened from the daily brief on a store number: open the lever that number sits under.
    if (x.focusKey && canLog) { const Lf = LSx.find(L => L.inputs.some(i => matches(i, x.focusKey))); if (Lf) { V.lever = Lf.key; const ik = Lf.inputs.find(i => matches(i, x.focusKey)).key; V.leverInputs = [ik, ...(V.leverInputs || []).filter(k => k !== ik && Lf.inputs.some(i => i.key === k))].slice(0, 2); } }
    if (!V.lever && canLog && V.status !== 'done' && LSx.length) { V.lever = suggestLever(snap?.mtd) || [...LSx].filter(l => l.ratio != null && l.key !== 'effMargin').sort((a, b) => a.ratio - b.ratio)[0]?.key || null; }
    const Lc = LSx.find(l => l.key === V.lever);
    if (Lc && !(V.leverInputs || []).length && V.status !== 'done') V.leverInputs = defaultInputs(Lc);
    if (Lc && (V.leverInputs || []).length) V.focus = Lc.inputs.filter(i => V.leverInputs.includes(i.key)).map(inpKey);
    if (!V.focus) V.focus = focusAll.slice(0, 2).map(f => f.key); }
  if (x.addCid && canLog && !V.consultants.some(c => c.cid === x.addCid)) {
    const ap = (S.rsa?.people || []).find(q => q.cid === x.addCid);
    if (ap) { const rk = x.lever && STORE_TO_RSA[x.lever]; const dk = { creditApps: 'appsToTraffic', protectionSph: 'protectionPct', beddingSph: 'beddingPct' }[rk] || rk;
      V.consultants.push({ cid: ap.cid, name: ap.name, why: x.why || 'added', lever: x.lever || null, notes: '', practice: {}, ...(dk ? { drill: { key: dk, title: drillFor(dk).title } } : {}) }); }
    x.addCid = null;
  }
  S.vSnap = snap; S.vPriorRaw = priorRaw;
  if (x.remote && canLog) (V.consultants || []).forEach(c => { if (!c.mode) c.mode = 'leader'; });
  if (canLog && !later && !V.actions.some(hasCommitment)) fillCommitments(V, focusAll);
  // Wins to celebrate, built from the numbers: store results, last visit's commitments that moved,
  // and the people carrying the store. The leader win box starts with these; change anything.
  const winList = (() => {
    const out = [];
    if (snap?.mtd) winsFor(snap.mtd).forEach(w => out.push(`${w} this month`));
    if (snap?.wtd?.vsBud?.netSales >= 0) out.push(`Sales ${pct(snap.wtd.vsBud.netSales)} to budget this week`);
    priorRaw.forEach(a => { const au = autoFollow(a, snap, V); if (au?.v === 'yes') out.push(`Hit last visit's commitment: ${commitmentText(a).split('. ')[0]}`); else if (au?.v === 'partial') out.push(`Moving on last visit's commitment: ${String(a.what || '').trim()}, ${au.text.charAt(0).toLowerCase() + au.text.slice(1).replace(/\.$/, '')}`); });
    const seen = new Set();
    focusAll.concat(storeFocus(snap, 6)).forEach(f => helpers(S.rsa?.people || [], x.store, f.key, DEFAULT_GOALS, paceFactor(S.rsa?.to), 1).forEach(h => {
      if (seen.has(h.cid) || seen.size >= 3) return; seen.add(h.cid);
      const nm = titleName(h.name), val = fmtMetric(h.key, h.value);
      out.push(f.key === 'cancelPct' ? `${nm} has the fewest cancels in the store at ${val}` : h.key === 'creditApps' ? `${nm} leads the store in credit apps with ${val}` : h.key === 'sph' ? `${nm} leads the store at ${val} an hour` : `${nm} leads the store in ${(PLAIN_LABELS[f.key] || f.label).toLowerCase()} at ${val}`); }));
    (S.teams?.[x.store]?.rising || []).slice(0, 2).forEach(r => out.push(`${titleName(r.name)} is up to $${Math.round(r.wk.sph)} an hour this week, from $${Math.round(r.wk.priorSph)}`));
    return [...new Set(out)].slice(0, 6);
  })();
  if (canLog && !later) {
    V.leaderWin = V.leaderWin || {};
    if (!V.leaderWin.name) {
      const lastName = S.visits.filter(y => y.store === x.store && y.id !== id).sort((a, b) => b.date.localeCompare(a.date)).map(y => y.leaderCommit?.name || y.leaderWin?.name).find(n => n && n !== 'Store leader');
      const asm = (S.roster || []).find(r => r.store === x.store && r.title === 'ASM')?.name;
      const listed = (S.storeLeaders || []).find(l => l.store === x.store)?.name;
      V.leaderWin.name = listed || lastName || (asm ? titleName(asm) : '');
    }
    if (!V.leaderWin.text && !V.leaderWin.touched && winList.length) { V.leaderWin.text = winList.map(w => `- ${w}.`).join('\n'); V.leaderWin.suggested = true; }
  }
  if (canLog && !later && !V.leaderCommit?.what) {
    const f = focusAll.find(q => V.focus.includes(q.key));
    const nextDay = (S.vPlans || []).flatMap(p => p.days || []).filter(d => d.store === x.store && d.date > x.date).map(d => d.date).sort()[0] || addDays(x.date, 7);
    if (f) V.leaderCommit = { ...(V.leaderCommit || {}), key: f.key, what: `Team ${f.label.toLowerCase()}`, from: fmtMetric(f.key, f.value), to: fmtMetric(f.key, f.target ?? f.goal), by: nextDay, supportBy: V.leaderCommit?.supportBy || nextDay };
  }
  const people = S.rsa?.people || [];
  // Who's on the Smart Scheduler today. Coaching picks come from people actually working, when we know.
  try { await Promise.race([schedWeek(weekStartOf(x.date)), new Promise(r => setTimeout(r, 2500))]); } catch (e) {}
  const onDay = whoOn(x.store, x.date), onSet = onDay ? new Set(onDay.on.map(p => p.cid)) : null;
  const pickPool = onSet && people.some(p => p.store === x.store && onSet.has(p.cid)) ? people.filter(p => p.store !== x.store || onSet.has(p.cid)) : people;
  const picks = rsaPicks(pickPool, x.store, DEFAULT_GOALS, paceFactor(S.rsa?.to), 3, S.weeks || {});
  if (!V.consultants.length && picks.length && V.status === 'draft' && !saved) V.consultants = picks.map(p => ({ cid: p.cid, name: p.name, why: p.why, notes: '', practice: {} }));
  const storePeople = people.filter(p => p.store === x.store).sort((a, b) => b.k.sph - a.k.sph);
  // Everyone on this store's sales team roster, even with no RSA numbers yet (new hires, early in the month).
  const rosterOnly = (S.roster || []).filter(r => r.store === x.store && r.title !== '__skip' && !people.some(p => p.cid === r.cid)).sort((a, b) => a.name.localeCompare(b.name));
  const team = S.teams?.[x.store];
  const leader = leaderOf(x.store);
  const tagText = { below: 'Below minimum', slipping: 'Slipping this week', gap: 'Biggest gap', model: 'Recognize and model', added: 'Added', drag: 'Pulling a store number down' };
  const dis = canLog ? '' : 'disabled';
  const tri = (path, val, opts = [['yes', 'Yes'], ['partial', 'Partial'], ['no', 'No']]) =>
    `<div class="tri" role="group">${opts.map(([k, l]) => `<button type="button" data-path="${path}" data-v="${k}" class="${val === k ? 'on' : ''}" ${dis}>${l}</button>`).join('')}</div>`;
  const sec = (key, num, title, q, body, accent) => `<section class="vsec ${V_OPEN.has(key) ? 'open' : ''}" data-sec="${key}">
    <button type="button" class="vsec-hd" aria-expanded="${V_OPEN.has(key)}"><span class="vnum ${accent ? 'acc' : ''}">${num}</span><span class="vtl"><b>${title}</b><small>${q}</small></span><span class="chev" aria-hidden="true">›</span></button>
    <div class="vsec-bd">${body}</div></section>`;
  const photoCard = ph => `<figure class="pcard">
      <img src="${ph.data}" alt="${esc(ph.caption || 'Visit photo')}">
      <figcaption><span class="pill">${esc(ph.el === 'general' ? 'General' : ELEMENTS.find(e => e.key === ph.el)?.t || ph.el)}${ph.item ? ' · ' + esc(ph.item) : ''}</span>
        ${canLog ? fieldBox(`pc_${ph.id}`, 'Comments', ph.caption, 2, '', '', '').replace('<textarea ', `<textarea data-pcap="${esc(ph.id)}" placeholder="What does this show? What needs to change?" `) : ph.caption ? `<p class="small">${esc(ph.caption)}</p>` : ''}
        ${canLog ? `<button type="button" class="link" data-delphoto="${esc(ph.id)}" style="color:var(--red);padding-left:0">Remove</button>` : ''}</figcaption></figure>`;
  const photoStrip = (el, extraOnly) => { const areas = photoAreas(ELEMENTS.find(z => z.key === el) || {}); const list = (S.vPhotos || []).filter(p => p.el === el && (!extraOnly || !areas.includes(p.item))); return list.length ? `<div class="gallery" style="margin-top:10px">${list.map(photoCard).join('')}</div>` : ''; };
  // The photo ask for one area, under its score: the shots taken so far with comments, or the buttons.
  const areaPhoto = (el0, area0, val) => {
    const { el, item: area } = photoKey(el0, area0), shared = area !== area0;
    const list = (S.vPhotos || []).filter(p => p.el === el && p.item === area);
    if (shared) return `<div class="aphoto ${list.length ? 'done' : ''}"><span class="small muted">Same photo as <b>${esc(area)}</b> in the Visual walk${list.length ? '' : ', take it there'}.</span>${list.map(ph => `<img class="pmini" src="${ph.data}" alt="${esc(ph.caption || area)}">`).join('')}</div>`;
    if (list.length) return `<div class="aphoto done"><div class="gallery">${list.map(photoCard).join('')}</div>${canLog ? `<div class="row" style="gap:6px"><span class="small good">Photo taken</span>${photoBtns(el, area).replace(/Take a photo/, 'Add another')}</div>` : ''}</div>`;
    if (!canLog) return '';
    return `<div class="aphoto"><span class="small ${flagged(val) ? 'warn' : 'muted'}">${flagged(val) ? 'Needs work: get a photo and say what has to change.' : 'Photo of this area'}</span>${photoBtns(el, area)}</div>`;
  };
  const photoBtns = (el, item = '') => canLog ? `<span class="row" style="gap:6px"><button type="button" class="btn tiny primary" data-photo="${el}" data-item="${esc(item)}" data-src="cam">Take a photo</button><button type="button" class="btn tiny" data-photo="${el}" data-item="${esc(item)}" data-src="lib">From library</button></span>` : '';
  const flagged = v => v === 'no' || v === 'partial' || v === 'needs';
  const kpiCell = (p, key, label, fmtFn, lower) => {
    const g = goalsFor(DEFAULT_GOALS, x.store)[key], val = p.k?.[key];
    const cls = val == null || g == null ? '' : status({ lower }, val, g);
    return `<div class="kc ${cls}"><span>${label}</span><b>${val == null ? '--' : fmtFn(val)}</b></div>`;
  };
  const money = n => '$' + Math.round(n).toLocaleString('en-US');
  const p1 = n => n.toFixed(1) + '%';
  // Year to date, last month, this month and this week side by side, with which way each is moving.
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const lmLabel = S.trendInfo?.lastMonth ? MON[+S.trendInfo.lastMonth.slice(5, 7) - 1] : 'Last mo';
  const wkNote = S.trendInfo ? ` This week runs Monday to Sunday (${esc(shortDate(S.trendInfo.week.from))} on)${S.trendInfo.wkOk ? '' : ', and needs the RSA copy through last Sunday to show'}.` : '';
  const trendBlock = (cid, p) => {
    const t = S.trends?.[cid];
    if (!t || !(t.mtd || t.ytd || t.lastMonth)) return '';
    const rd = trendRead(t), g = goalsFor(DEFAULT_GOALS, x.store);
    const ti = S.trendInfo || {};
    const cols = [['ytd', 'YTD'], ['lastMonth', lmLabel], ['mtd', ti.cur ? MON[+ti.cur.slice(5, 7) - 1] + ' MTD' : 'MTD'], ['wtd', 'This wk'], ['day', ti.day ? shortDate(ti.day) : 'Prev day']];
    const cell = (per, r) => { const v = t[per]?.k?.[r.key]; const gl = r.total ? (per === 'lastMonth' ? g[r.key] : null) : g[r.key]; const cls = v == null || gl == null ? '' : status({ lower: r.lower }, v, gl) === 'green' ? 'good' : status({ lower: r.lower }, v, gl) === 'red' ? 'bad' : ''; return `<td class="num ${cls}">${trendFmt(r, v)}</td>`; };
    const arrow = k => rd.rows[k] === 'up' ? '<td class="ar good" title="Getting better">▲</td>' : rd.rows[k] === 'down' ? '<td class="ar bad" title="Slipping">▼</td>' : '<td class="ar muted">·</td>';
    const hrs = per => t[per]?.hours != null ? Math.round(t[per].hours).toLocaleString('en-US') : '--';
    const ytdNote = S.trendInfo?.ytd && !S.trendInfo.ytd.uploaded ? ` YTD is built from the monthly copies on file since ${esc(shortDate(S.trendInfo.ytd.from))}.` : !S.trendInfo?.ytd ? ' Upload a year-to-date RSA report to fill in YTD.' : '';
    const say = (r, a, b) => `${r.say} went from ${trendFmt(r, a)} ${TREND_LABEL[rd.baseP]} to ${trendFmt(r, b)} ${TREND_LABEL[r.per]}`;
    const line = [rd.up ? `<span class="good">Getting better:</span> ${esc(say(rd.up, rd.up.from, rd.up.to))}.` : '', rd.down ? `<span class="bad">Slipping:</span> ${esc(say(rd.down, rd.down.from, rd.down.to))}.` : ''].filter(Boolean).join(' ');
    return `<div class="trend">
      ${line ? `<p class="small" style="margin:0 0 6px">${line}</p>` : ''}
      <div class="scroller"><table class="grid tgrid"><thead><tr><th></th>${cols.map(([, l]) => `<th class="num">${l}</th>`).join('')}<th></th></tr></thead><tbody>
        <tr class="hrs"><td>Hours</td>${cols.map(([k]) => `<td class="num">${hrs(k)}</td>`).join('')}<td></td></tr>
        ${TREND_ROWS.map(r => `<tr><td>${r.label}</td>${cols.map(([k]) => cell(k, r)).join('')}${arrow(r.key)}</tr>`).join('')}
      </tbody></table></div>
      <p class="small muted" style="margin:4px 0 0">Green is at goal, red is well off it. ▲ ▼ compare ${rd.recentP === 'wtd' ? 'this week' : TREND_LABEL[rd.recentP] || 'the latest numbers'} with ${TREND_LABEL[rd.baseP] || 'the longer run'}.${wkNote}${ytdNote}</p>
    </div>`;
  };
  // A question to open with, from the trend: ask about what changed instead of reading the number at them.
  const trendAsk = (cid, first) => {
    const rd = trendRead(S.trends?.[cid]); if (!rd.baseP) return '';
    if (rd.down) return `\nOpen with the trend: "Your ${rd.down.say} was ${trendFmt(rd.down, rd.down.from)} ${TREND_LABEL[rd.baseP]} and it's ${trendFmt(rd.down, rd.down.to)} ${TREND_LABEL[rd.down.per]}. What's changed?" Let ${first} answer before you coach.`;
    if (rd.up) return `\nRecognize the trend: "Your ${rd.up.say} went from ${trendFmt(rd.up, rd.up.from)} ${TREND_LABEL[rd.baseP]} to ${trendFmt(rd.up, rd.up.to)} ${TREND_LABEL[rd.up.per]}. What are you doing differently?" Then ask ${first} to show the team.`;
    return '';
  };

  const html = `
  <div class="spread">
    <div>
      <button class="link" id="back" style="padding-left:0">‹ Back</button>
      <p class="eyebrow">${x.remote ? 'Remote coaching · ' : (() => { const pd = (S.vPlans || []).flatMap(p => p.days || []).find(d => d.date === x.date && dayStores(d).includes(x.store)); if (!pd || dayStores(pd).length < 2) return 'Full-day visit · '; const part = pd.store === x.store ? (pd.part || 'AM') : (pd.stops.find(y => y.store === x.store)?.part || 'Stop'); return `${pd.anchor && pd.store === x.store ? 'Anchor store, ' : ''}${part} visit · `; })()}${esc(longDate(x.date))} · ${esc(who.name || who.email)}${leader && leader.email !== who.email ? ' · Market Leader: ' + esc(leader.name || leader.email) : ''}</p>
      <h2 class="big" style="margin:0">${esc(x.store)}</h2>
      <p class="small" style="margin:4px 0 0">${dataLine()}</p>
    </div>
    <div class="row">
      <label for="vtype" style="margin:0">Visit type<select id="vtype" ${dis}>${(x.remote ? REMOTE_TYPES : VISIT_TYPES).map(t => `<option ${V.vtype === t ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
      ${needChip(sc?.score)}
    </div>
  </div>
  ${later ? `<div class="warnbox">This visit is on ${esc(longDate(x.date))}. Use it to prep; you can fill it in that day.</div>` : ''}
  ${(() => { const b = budgetFor(x.store, x.date), y = S.meta.latestDaily, mb = y ? budgetToDate(x.store, y) : null, m = S.daily?.stores?.[x.store]?.mtd?.k;
    return b ? `<div class="budgetline"><b>Today's budget at ${esc(x.store)}:</b> ${$k(b.sales)} revenue · SPG $${Math.round(b.spg)} · about ${Math.round(b.traffic)} guests.${mb && m?.netSales != null ? ` Month to date ${$k(m.netSales)} vs ${$k(mb.sales)} budget (${vsTag(vsPct(m.netSales, mb.sales))}).` : ''} <span class="small muted">Make sure the leader knows both numbers.</span></div>` : ''; })()}
  ${(() => { const sc = S.carts?.stores?.[x.store]; if (!sc) return ''; const tops = Object.entries(S.carts.people || {}).filter(([, p]) => p.store === x.store).sort((a, b) => b[1].value - a[1].value).slice(0, 3);
    return `<div class="budgetline cartsline"><b>Open carts at ${esc(x.store)}:</b> ${sc.n} carts · about ${$k(sc.value)} estimated${sc.due ? ` · <b>${sc.due} due a follow-up today</b>` : ''}${sc.old ? ` · ${sc.old} older than 2 weeks` : ''}. Most to follow up: ${tops.map(([, p]) => `${esc(titleName(p.name))} (${p.n}, ${$k(p.value)})`).join(', ')}. <span class="small muted">This is money already in the building: inspect the follow-up plan with the leader.</span></div>`; })()}
  ${whoPanel(x.store, x.date)}
  ${V.status === 'done' && !x.remote && V.photosMissing ? `<div class="editbox"><b>Submitted with ${V.photosMissing} photo${V.photosMissing > 1 ? 's' : ''} missing.</b> <span class="small">${esc((V.photosMissingList || []).join(', '))}</span><br><span class="small"><b>Reason:</b> ${esc(V.photoReason || 'none given')}. Add them any time and this clears.</span></div>` : ''}
  ${V.edits?.length ? `<div class="editbox"><b>Edited after it was submitted</b> <span class="small">(submitted ${esc(dayLabel((V.submittedAt || V.date).slice(0, 10)))})</span><ul class="small">${editedText(V).map(t => `<li>${esc(t)}</li>`).join('')}</ul></div>` : ''}
  ${env ? `<div class="env ${env.kind}"><b>${env.kind === 'headwind' ? 'Headwind.' : env.kind === 'tailwind' ? 'Tailwind.' : 'Normal traffic.'}</b> ${esc(env.text)}</div>` : ''}

  ${sec('why', '1', 'Why you are here', 'The intent of the visit. We change the outcome through people and behaviors.', `
    ${(() => { const pd = (S.vPlans || []).flatMap(p => p.days || []).find(d => d.date === x.date && dayStores(d).includes(x.store)); return intentHtml(visitIntent(x.store, { remote: x.remote, anchor: pd?.anchor && pd.store === x.store && !x.remote, kind: pd ? (pd.store === x.store ? pd.kind : pd.stops.find(y => y.store === x.store)?.kind) : x.kind, date: x.date })); })()}
    ${sc?.parts?.length > 2 ? `<details style="margin:8px 0"><summary class="small" style="cursor:pointer">All the reasons from the numbers</summary><ul class="whylist">${sc.parts.slice(0, 6).map(p => `<li>${esc(p.text)}</li>`).join('')}</ul></details>` : sc?.parts?.length ? '' : '<p class="small">No flags. Use the visit to lock in what is working.</p>'}
    <div class="ptog" role="group" aria-label="Period">${PERIODS.filter(([k]) => snap?.[k]).map(([k, l]) => `<button type="button" data-period="${k}" class="${period === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    ${storeTiles(snap, period)}`)}

  ${sec('win', '★', 'Leader win', 'Start here. Celebrate the leader before anything else.', `
    ${winList.length ? `<div class="wins"><p class="eyebrow" style="margin:0 0 4px">Wins to celebrate</p><ul>${winList.map(w => `<li>${esc(w)}</li>`).join('')}</ul></div>` : '<p class="small muted">No wins in the numbers yet. Find one on the floor and call it out.</p>'}
    <div style="margin-bottom:10px">${leaderField('lwname', 'Leader', V.leaderWin?.name, 'leaderWin.name', dis, x.store)}</div>
    ${fieldBox('lwtext', 'What you will celebrate with the leader and the team', V.leaderWin?.text, 4, V.leaderWin?.suggested ? 'Started from the wins above. Add what you saw on the floor.' : '', 'leaderWin.text', dis)}`, true)}

  ${sec('follow', '↻', 'Inspect and follow up', prior ? `Last visit ${esc(dayLabel(prior.date))}${prior.email !== x.email ? ' (' + esc(prior.name) + ')' : ''}. Each commitment is checked against today's numbers. Watch for the behavior on the floor.` : 'No earlier visit to this store.', (priorSum?.commitments.length || priorSum?.leaderCommit || priorSum?.support) ? `
    ${priorRaw.length ? `<p class="small" style="margin:0 0 8px"><span class="pill check">Inspection visit</span> ${priorRaw.filter(a => autoFollow(a, snap, V)?.v === 'yes').length} of ${priorRaw.length} commitments hit. Anything not there yet is ready to carry forward in today's action plan.</p>` : ''}
    ${priorSum.commitments.map((c, i) => { const a = priorRaw[i] || {}; const au = autoFollow(a, snap, V); if (canLog) setPath(V, `follow.${i}`, au?.v ?? null);
      return `<div class="item"><div class="txt">${esc(c)}${a.how ? `<p class="small" style="margin:4px 0 0"><b>${x.remote ? 'Ask the leader about' : 'Inspect on the floor'}:</b> ${esc(a.how)}</p>` : ''}</div>${followBadge(au?.v, au)}</div>`; }).join('')}
    ${priorSum.fixes.length ? `<p class="small" style="margin-top:10px">6 Elements flagged last time: ${priorSum.fixes.map(esc).join(', ')}</p>` : ''}
    ${(() => { const au = priorSum.leaderCommit ? autoFollow(prior.leaderCommit, snap, V) : null; if (canLog) setPath(V, 'follow.lc', au?.v ?? null); S.lcAuto = au; return ''; })()}
    ${priorSum.leaderCommit ? `<div class="item"><div class="txt"><b>${esc(priorSum.leaderName || 'Store leader')} committed to:</b> ${esc(priorSum.leaderCommit)}</div>${followBadge(S.lcAuto?.v, S.lcAuto)}</div>` : ''}
    ${priorSum.support ? `<div class="item"><div class="txt"><b>Support promised by ${esc(prior.name)}:</b> ${esc(priorSum.support)}</div><p class="small muted" style="margin:4px 0 0">Review it with the leader.</p></div>` : ''}`
    : '<p class="small">Nothing to check. Your action plan today becomes the start of the next visit.</p>', true)}

  ${sec('play', '▶', 'Connect, build value, run the play', activeOffer() ? `How we sell, and the ${esc(S.offer.name)} offer that rewards the bundle.` : 'How we sell, on every guest.', playBlock(V, canLog, dis, tri), true)}
  ${x.remote ? sec('floor', '◎', 'Does the leader know their floor?', 'Ask for yesterday, by the numbers. A leader who knows the floor can coach it.', floorBlock(V, x, snap, dis, tri), true) : ''}
  ${hasFliq(who) ? sec('fliq', '◆', 'FrontLine IQ', 'Pilot: the AI sales coach for your associates. Check it on every visit.', fliqBlock(V, canLog, dis, tri), true) : ''}

  ${sec('lever', '2', 'Pull a lever, coach two things', 'Pick the outcome: close rate, average ticket or effective margin. The two inputs you pick under it are what the team works on today, with the people behind each one.', leverBlock(V, x, snap, people, canLog, dis), true)}

  ${sec('people', '3', x.remote ? 'Consultants: coach them through the leader' : 'Consultants coached', x.remote ? `Help the leader coach each person. Use the data and the trend to open a conversation about behavior, role-play it with the leader, and agree on when they coach it. Each card shows year to date, last month, month to date and this week${S.rsa?.to ? ' through ' + esc(shortDate(S.rsa.to)) : ''}.` : `Who to see first. Coach one thing with each person. Each card shows year to date, last month, month to date and this week${S.rsa?.to ? ' through ' + esc(shortDate(S.rsa.to)) : ''}.`, `
    ${V.consultants.map((c, ci) => {
      const p = people.find(q => q.cid === c.cid) || { name: c.name, k: {} };
      const pk = picks.find(q => q.cid === c.cid);
      const wk = S.weeks?.[c.cid];
      return `<div class="ccard">
        <div class="row" style="justify-content:space-between"><h3 style="margin:0">${esc(titleName(c.name))}${p.store && p.store !== x.store ? ` <span class="small muted">(${esc(p.store)})</span>` : ''}${(() => { if (!onDay) return ''; const on = onDay.on.find(q => q.cid === c.cid); return on ? ` <span class="pill set">On ${x.date === today() ? 'today' : dayLabel(x.date)} ${esc(shiftText(on.sh))}</span>` : onDay.pto.some(q => q.cid === c.cid) ? ' <span class="pill off">PTO</span>' : ' <span class="pill off">Not on the schedule</span>'; })()}
          <a class="small" href="https://fpina-1915south.github.io/consultant-scorecard/?store=${encodeURIComponent(p.store || x.store)}&c=${encodeURIComponent(c.cid)}" target="_blank" rel="noopener" style="margin-left:6px">Open in Scorecard</a></h3><span class="row" style="gap:6px"><span class="tag ${c.why}">${tagText[c.why] || ''}</span>${canLog ? `<button type="button" class="link" data-rmc="${ci}" aria-label="Remove ${esc(titleName(c.name))}">Remove</button>` : ''}</span></div>
        ${trendBlock(c.cid, p) || (p.k?.sph != null ? `<div class="kpis">
          ${kpiCell(p, 'sph', 'SPH', money)}${wk?.hours >= 1 && wk.sph != null ? `<div class="kc ${wk.priorSph && wk.sph < wk.priorSph * 0.75 ? 'red' : wk.priorSph && wk.sph > wk.priorSph * 1.25 ? 'green' : ''}"><span>This week</span><b>${money(wk.sph)}</b></div>` : ''}
          ${kpiCell(p, 'financePct', 'Finance', p1)}${kpiCell(p, 'beddingPct', 'Bedding', p1)}${kpiCell(p, 'protectionPct', 'Protection', p1)}${kpiCell(p, 'creditApps', 'Apps', n => String(Math.round(n)))}${kpiCell(p, 'cancelPct', 'Cancel', p1, true)}
        </div>` : '')}
        ${(() => { const cr = cartsFor(c.cid); if (!cr) return ''; const fn = titleName(c.name).split(' ')[0], via = (c.mode || (x.remote ? 'leader' : 'direct')) === 'leader';
          const L = V.leaderWin?.name ? titleName(V.leaderWin.name).split(' ')[0] : 'the leader';
          const tip = cr.due ? `${via ? `Have ${esc(L)} sit with ${esc(fn)}` : `Sit with ${esc(fn)}`} for 10 minutes and make the ${cr.due} follow-up call${cr.due > 1 ? 's' : ''} due today together.` : `Ask ${esc(fn)} to walk ${via ? esc(L) : 'you'} through their biggest carts and when each one gets its next call (day 1, 3 and 7).`;
          return `<div class="cartbox"><p style="margin:0"><b>Open carts:</b> ${esc(cartLine(cr))}${cr.old ? ` · <span class="warn">${cr.old} older than 2 weeks</span>` : ''}</p>
            ${cr.top?.length ? `<ul class="small">${cr.top.map(t => `<li>${esc(t.g)} · ${$k(t.v)} · ${t.a} day${t.a === 1 ? '' : 's'} old${CART_DUE.includes(t.a) ? ' <b class="warn">call today</b>' : ''}</li>`).join('')}</ul>` : ''}
            <p class="small" style="margin:4px 0 0"><b>Follow-up:</b> ${tip}</p></div>`; })()}
        <div class="segwho"><p class="small" style="margin:0 0 6px"><b>How are you coaching ${esc(titleName(c.name).split(' ')[0])}?</b> <span class="muted">Through the leader: you give the leader the plan, role-play it with them, and they coach ${esc(titleName(c.name).split(' ')[0])}.</span></p>${tri(`consultants.${ci}.mode`, c.mode || (x.remote ? 'leader' : 'direct'), [['direct', 'Directly with them'], ['leader', 'Through the store leader']])}</div>
        ${(() => { const cc = consultantCoaching({ p, store: p.store || x.store, why: c.why, wk, goals: DEFAULT_GOALS, pace: paceFactor(S.rsa?.to), teamFocus: focusAll.find(f => V.focus.includes(f.key))?.label, lever: c.lever });
          const viaLeader = (c.mode || (x.remote ? 'leader' : 'direct')) === 'leader';
          const base = (viaLeader ? leaderCoachText(cc, titleName(c.name).split(' ')[0], V.leaderWin?.name || 'the store leader') : cc.text).split('\n');
          const ask = trendAsk(c.cid, titleName(c.name).split(' ')[0]).trim();
          if (ask) base.splice(Math.min(2, base.length), 0, ask);
          const text = base.join('\n');
          return `<div class="suggest"><p class="eyebrow">${viaLeader ? `Help ${esc(V.leaderWin?.name ? titleName(V.leaderWin.name).split(' ')[0] : 'the leader')} coach ${esc(titleName(c.name).split(' ')[0])}` : 'Suggested coaching'}</p><div class="stext">${esc(text)}</div>
          ${canLog ? `<div class="row" style="margin-top:8px"><button type="button" class="btn tiny primary" data-usec="${ci}">Use this in my notes</button><span class="small muted">or write or say your own below</span></div>` : ''}</div>`; })()}
        ${(() => {
          const cc = consultantCoaching({ p, store: p.store || x.store, why: c.why, wk, goals: DEFAULT_GOALS, pace: paceFactor(S.rsa?.to), lever: c.lever });
          // The practice follows what this person is being coached on, unless the leader picked a different one or already scored it.
          const scored = Object.keys(c.drill?.scored || {}).length;
          if (!c.drill?.key || (!c.drill.manual && !scored && c.drill.key !== cc.drillKey)) c.drill = { ...(c.drill || {}), key: cc.drillKey, title: cc.drill.title };
          const firstName = titleName(c.name).split(' ')[0];
          const onWhat = cc.items.find(f => !f.stretch) || cc.items[0];
          const dr = drillFor(c.drill.key), d = c.drill, n = Object.keys(d.scored || {}).length;
          const opts = DRILL_KEYS.map(k => [k, drillFor(k).title]);
          if ((c.mode || (x.remote ? 'leader' : 'direct')) === 'leader') {
            const L = V.leaderWin?.name ? titleName(V.leaderWin.name).split(' ')[0] : 'the leader', ld = c.lead || {};
            return `<div class="drill">
              <p class="eyebrow" style="margin:0">Practice the coaching conversation with ${esc(L)}</p>
              <h4>You play ${esc(firstName)}. ${esc(L)} coaches you.</h4>
              <p class="small">Act like ${esc(firstName)} would. Push back a little. Then have ${esc(L)} run the ${esc(dr.title.toLowerCase())} practice with ${esc(firstName)} on the floor.</p>
              <p class="small" style="margin:6px 0 0"><b>Score ${esc(L)}'s coaching</b></p>
              ${COACH_WATCH.map((t, i) => `<div class="item"><div class="txt">${esc(t)}</div>${tri(`consultants.${ci}.lead.scored.${i}`, ld.scored?.[i])}</div>`).join('')}
              <div class="two">
                <div>${fieldBox(`lw${ci}`, 'What the leader did well', ld.well, 2, '', `consultants.${ci}.lead.well`, dis)}</div>
                <div>${fieldBox(`la${ci}`, 'One adjustment for the leader', ld.adjust, 2, '', `consultants.${ci}.lead.adjust`, dis)}</div>
              </div>
              <label for="lwhen${ci}" style="max-width:260px;margin-top:8px">${esc(L)} coaches ${esc(firstName)} by<input id="lwhen${ci}" type="date" data-field="consultants.${ci}.lead.by" value="${esc(ld.by || addDays(x.date, 1))}" ${dis}></label>
              <p class="small muted" style="margin:4px 0 0">Follow up after to hear how it went.</p>
            </div>`;
          }
          return `<div class="drill">
            <div class="row" style="justify-content:space-between"><p class="eyebrow" style="margin:0">${esc(firstName)}'s stand-up practice${onWhat ? ` on ${esc(onWhat.label.toLowerCase())}` : ''}</p>
              <select id="dk${ci}" data-drill="${ci}" aria-label="Practice drill" ${dis}>${opts.map(([k, t]) => `<option value="${k}" ${drillFor(k).title === dr.title ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select></div>
            <h4>${esc(dr.title)}</h4>
            <p class="small"><b>You play the guest:</b> ${esc(dr.guest)}</p>
            <p class="small" style="margin:6px 0 0"><b>Rep 1. Score what you see</b> <span class="muted dcount" data-total="${dr.watch.length}">(${n} of ${dr.watch.length})</span></p>
            ${dr.watch.map((t, i) => `<div class="item"><div class="txt">${esc(t)}</div>${tri(`consultants.${ci}.drill.scored.${i}`, d.scored?.[i])}</div>`).join('')}
            <div class="two">
              <div>${fieldBox(`dw${ci}`, 'What they did well', d.well, 2, '', `consultants.${ci}.drill.well`, dis)}</div>
              <div>${fieldBox(`da${ci}`, 'One adjustment', d.adjust, 2, '', `consultants.${ci}.drill.adjust`, dis)}</div>
            </div>
            <p class="small" style="margin:8px 0 6px"><b>Rep 2. Run it again with the adjustment</b></p>
            ${tri(`consultants.${ci}.drill.rerun`, d.rerun, [['better', 'Better'], ['same', 'Same'], ['none', 'Did not rerun']])}
            <details class="prac" ${Object.keys(c.practice || {}).length ? 'open' : ''} style="margin-top:10px"><summary>Full Core 4 run (optional) <span class="small muted">${Object.keys(c.practice || {}).length} of ${PRACTICE.length} scored</span></summary>
              ${PRACTICE.map((t, i) => `<div class="item"><div class="txt">${esc(t)}</div>${tri(`consultants.${ci}.practice.${i}`, c.practice?.[i])}</div>`).join('')}
            </details>
          </div>`; })()}
        <div class="cnotes">${fieldBox(`cn${ci}`, `Notes on ${titleName(c.name).split(' ')[0]}`, c.notes, 4, 'What you saw, what you talked about, what they said. Type it or tap Talk.', `consultants.${ci}.notes`, dis)}</div>
      </div>`;
    }).join('') || `<p class="small muted">${S.rsa ? 'No consultants matched to this store in the RSA report.' : 'Upload the RSA report to get consultant picks.'}</p>`}
    ${canLog ? `<div class="addc"><div style="flex:1;min-width:220px">${fieldInput('addc', 'Coach any consultant', '', '', '', 'list="addcList" placeholder="Start typing a name" autocomplete="off" style="width:100%"')}</div>
      <button type="button" class="btn" id="addcgo">Add</button></div>
      <datalist id="addcList">${[...storePeople, ...rosterOnly, ...people.filter(p => p.store !== x.store).sort((a, b) => a.name.localeCompare(b.name))].filter(p => !V.consultants.some(c => c.cid === p.cid)).map(p => `<option value="${esc(titleName(p.name))}">${esc(p.store || 'No store')} · ${p.k ? '$' + Math.round(p.k.sph) + ' SPH' : 'no RSA numbers yet'}</option>`).join('')}</datalist>
      <p class="small muted" style="margin:4px 0 0">This store's whole team shows first, including anyone without RSA numbers yet. Anyone in the RSA report works, and a name not on either list can be coached too.${rosterOnly.length ? ` ${storePeople.length} of ${storePeople.length + rosterOnly.length} on the ${esc(x.store)} team have RSA numbers${S.rsa?.to ? ' through ' + esc(shortDate(S.rsa.to)) : ''}.` : ''}</p>` : ''}
    <div class="cnotes" style="margin-top:12px">${fieldBox('teamnotes', 'Notes on the team', V.teamNotes, 3, 'Anything about the sales team as a whole: energy, staffing, who is ready for more.', 'teamNotes', dis)}</div>
    ${team?.rows.length ? `<details style="margin-top:10px"><summary class="small" style="cursor:pointer;font-weight:600;color:var(--navy)">Whole team: year, month and this week (${team.rows.length})</summary>
      <div class="scroller" style="margin-top:8px"><table class="grid"><thead><tr><th>Consultant</th><th class="num">Revenue MTD</th><th class="num">Open carts</th><th class="num">YTD SPH</th><th class="num">Month SPH</th><th class="num">Before this week</th><th class="num">This week</th><th class="num">Hrs</th><th>Flag</th></tr></thead><tbody>
      ${team.rows.map(r => `<tr><td class="nm">${esc(titleName(r.name))}</td><td class="num">${(() => { const v = people.find(q => q.cid === r.cid)?.k?.netSales; return v != null ? '$' + Math.round(v).toLocaleString('en-US') : '--'; })()}</td><td class="num">${(() => { const cr = cartsFor(r.cid); return cr ? `${cr.n} · ${$k(cr.value)}` : '--'; })()}</td><td class="num">${S.trends?.[r.cid]?.ytd ? '$' + Math.round(S.trends[r.cid].ytd.k.sph) : '--'}</td><td class="num ${r.below ? 'bad' : ''}">$${Math.round(r.sph)}</td><td class="num">${r.wk.priorSph ? '$' + Math.round(r.wk.priorSph) : '--'}</td><td class="num ${r.slipping ? 'bad' : r.rising ? 'good' : ''}">${r.wk.sph != null && r.wk.hours >= 1 ? '$' + Math.round(r.wk.sph) : '--'}</td><td class="num">${r.wk.hours >= 1 ? Math.round(r.wk.hours) : '--'}</td><td>${r.below ? '<span class="tag below">Below min</span>' : r.slipping ? '<span class="tag slipping">Slipping</span>' : r.rising ? '<span class="tag model">Rising</span>' : ''}</td></tr>`).join('')}
      </tbody></table></div></details>` : ''}`)}

  ${x.remote ? '' : ELEMENTS.map(e => sec('el' + e.n, e.n, e.t, PHOTO_ELS.includes(e.key) ? `${e.q} Photos: ${photoAreas(e).filter(a => (S.vPhotos || []).some(p => p.el === e.key && p.item === a)).length} of ${photoAreas(e).length}` : e.q, `
    ${PHOTO_ELS.includes(e.key) && canLog ? `<p class="small" style="margin:0 0 8px">Score each area and take a photo of it as you go, good or bad. Every visit gets the same photos so we can compare stores and visits.</p>` : ''}
    ${e.key === 'culture' ? `<p class="small">Score the value segments from something you saw today: watch a team member with a live guest, or run it as a practice with them.</p>` + SEGMENTS.map((g, gi) => { const sm = V.segMeta?.[gi] || {};
      return `<div class="seggrp"><p class="eyebrow">${esc(g.name)}</p><p class="small">${esc(g.must)}</p>
      <div class="segwho">
        ${tri(`segMeta.${gi}.how`, sm.how, [['observed', 'Watched a live guest'], ['practice', 'Practiced with them']])}
        <label for="segp${gi}" style="margin:8px 0 0">Team member<select id="segp${gi}" data-segp="${gi}" ${dis}><option value="">Choose…</option>${storePeople.map(p => `<option value="${esc(p.cid)}" ${sm.cid === p.cid ? 'selected' : ''}>${esc(titleName(p.name))}</option>`).join('')}${V.consultants.filter(c => !storePeople.some(p => p.cid === c.cid)).map(c => `<option value="${esc(c.cid)}" ${sm.cid === c.cid ? 'selected' : ''}>${esc(titleName(c.name))}</option>`).join('')}</select></label>
      </div>
      ${g.items.map((t, i) => `<div class="item"><div class="txt">${esc(t)}</div>${tri(`segs.${gi}.${i}`, V.segs?.[gi]?.[i])}</div>`).join('')}
      ${fieldBox(`segn${gi}`, 'What you saw', sm.notes, 2, '', `segMeta.${gi}.notes`, dis)}</div>`; }).join('') : ''}
    ${e.items.map((t, i) => `<div class="item"><div class="txt">${esc(t)}</div>${tri(`checks.${e.key}.${i}`, V.checks?.[e.key]?.[i])}
      ${PHOTO_ELS.includes(e.key) ? areaPhoto(e.key, t, V.checks?.[e.key]?.[i]) : ''}</div>`).join('')}
    ${e.aor ? `<p class="small">Walk every area of responsibility. Heroes leading, clean displays, pricing and POP right.</p>${AORS.map(a => `<div class="item"><div class="txt">${esc(a)}</div>${tri(`aor.${a}`, V.aor?.[a], [['pass', 'Pass'], ['needs', 'Needs work']])}
      ${areaPhoto(e.key, a, V.aor?.[a])}</div>`).join('')}` : ''}
    ${PHOTO_ELS.includes(e.key) && photoStrip(e.key, true) ? `<div class="elphotos"><b class="small">Other photos for ${esc(e.t)}</b>${photoStrip(e.key, true)}</div>` : ''}
    ${fieldBox(`eln${e.n}`, 'Notes', V.elNotes?.[e.key], 2, '', `elNotes.${e.key}`, dis)}
    ${canLog ? `<div class="row" style="margin-top:8px"><span class="small muted">Something else worth a picture?</span>${photoBtns(e.key)}</div>` : ''}${!PHOTO_ELS.includes(e.key) ? photoStrip(e.key) : ''}`)).join('')}

  ${x.remote ? `<div class="warnbox">This is a remote visit, so there's no 6 Elements walk or photos. Go over the numbers with the leader, coach the focus items and the consultants, and set commitments.</div>` : ''}

  ${sec('action', '✓', 'Action plan', 'Up to 3 commitments, each from X to Y by a date. We fill in suggestions. Change anything.', `
    ${(() => { S.cOpts = canLog ? commitmentOptions(V, snap, priorRaw) : []; return ''; })()}
    ${[0, 1, 2].map(i => { const a = V.actions[i] || {}; return `<div class="ap"><div class="row" style="justify-content:space-between"><p class="eyebrow" style="margin:0">Commitment ${i + 1}</p>${a.carried ? '<span class="pill">Carried forward</span>' : ''}${a.suggested ? '<span class="pill check">Suggested</span>' : ''}</div>
      ${canLog && S.cOpts.length ? `<select data-apick="${i}" aria-label="Choose commitment ${i + 1}" style="width:100%;margin:6px 0"><option value="">Choose from today's opportunities and coaching…</option>${[...new Set(S.cOpts.map(o => o.group))].map(gr => `<optgroup label="${esc(gr)}">${S.cOpts.map((o, k) => o.group === gr ? `<option value="${k}">${esc(o.what)}: ${esc(o.from)} to ${esc(o.to)}</option>` : '').join('')}</optgroup>`).join('')}</select>` : ''}
      ${fieldInput(`apw${i}`, 'What', a.what || a.behavior, `actions.${i}.what`, dis, 'placeholder="The behavior or number" style="width:100%"')}
      <div class="two">
        <div>${fieldInput(`apf${i}`, 'From (today)', a.from, `actions.${i}.from`, dis, 'placeholder="Where it is now" style="width:100%"')}</div>
        <div>${fieldInput(`apt${i}`, 'To', a.to, `actions.${i}.to`, dis, 'placeholder="Where it will be" style="width:100%"')}</div>
      </div>
      ${fieldBox(`aph${i}`, 'How', a.how, 2, '', `actions.${i}.how`, dis)}
      <div class="row"><div style="flex:1;min-width:180px">${fieldInput(`apo${i}`, 'Owner', a.owner, `actions.${i}.owner`, dis, 'style="width:100%"')}</div>
      <label for="apd${i}" style="margin:0">By<input id="apd${i}" type="date" data-field="actions.${i}.due" value="${esc(a.due || '')}" ${dis}></label></div>
      <p class="small preview" id="apv${i}" ${hasCommitment(a) ? '' : 'hidden'}>${hasCommitment(a) ? esc(commitmentText(a)) : ''}</p></div>`; }).join('')}
    ${canLog ? `<button type="button" class="btn" id="apsugg">Suggest commitments</button> <span class="small muted">Fills any empty ones from your focus items, 6 Elements fixes and practice results.</span>` : ''}
    <div style="margin-top:12px">${fieldBox('vwork', 'What is working (goes in the recap message)', V.working, 2, '', 'working', dis)}</div>`, true)}

  ${sec('leadercommit', '✓', 'Store leader notes', `Notes on what the leader commits to, and the support they need from ${who.role === 'director' ? 'their director' : 'their Market Leader'}.`, `
    <div style="margin-bottom:10px">${leaderField('lcname', 'Store leader', V.leaderCommit?.name || V.leaderWin?.name, 'leaderCommit.name', dis, x.store)}</div>
    <p class="small" style="margin:0 0 6px"><b>I commit to</b> <span class="muted">(in their words, from X to Y by a date. We suggest one; change anything.)</span></p>
    ${fieldInput('lcwhat', 'What', V.leaderCommit?.what, 'leaderCommit.what', dis, 'placeholder="The behavior or number" style="width:100%"')}
    <div class="two">
      <div>${fieldInput('lcfrom', 'From', V.leaderCommit?.from, 'leaderCommit.from', dis, 'placeholder="Where it is now" style="width:100%"')}</div>
      <div>${fieldInput('lcto', 'To', V.leaderCommit?.to, 'leaderCommit.to', dis, 'placeholder="Where it will be" style="width:100%"')}</div>
    </div>
    <label for="lcby" style="max-width:220px">By<input id="lcby" type="date" data-field="leaderCommit.by" value="${esc(V.leaderCommit?.by || '')}" ${dis}></label>
    ${fieldBox('lcsupport', `Support I need from ${esc(who.name || 'you')}`, V.leaderCommit?.support, 2, 'People, schedule, product, training, a call with someone. Be specific.', 'leaderCommit.support', dis)}
    <label for="lcsby" style="max-width:220px">Support by<input id="lcsby" type="date" data-field="leaderCommit.supportBy" value="${esc(V.leaderCommit?.supportBy || '')}" ${dis}></label>
    ${fieldBox('lcnotes', 'Notes', V.leaderCommit?.notes, 3, 'Anything else from the conversation with the leader.', 'leaderCommit.notes', dis)}`, true)}

  ${sec('reflect', '★', 'Your reflection', 'One line before your next stop.', `
    ${fieldBox('vref', 'What I will coach next visit', V.reflection, 2, '', 'reflection', dis)}
    ${fieldBox('vnotes', 'Other notes', V.notes, 2, '', 'notes', dis)}`, true)}

  <div class="vscore" id="vscore"></div>
  ${canLog ? `<div class="vbar"><span class="small" id="vsaved">${V.status === 'done' ? 'Submitted ' + esc(dayLabel((V.submittedAt || V.date).slice(0, 10))) + (V.lastEdit ? ` · edited ${esc(dayLabel(V.lastEdit.at.slice(0, 10)))}` : '') + '. Changes now are stamped and sent to your VP.' : 'Draft saves as you go'}</span>
    <button class="btn" id="vprint" type="button">Print / PDF</button>
    <button class="btn primary" id="vsubmit" type="button">${V.status === 'done' ? 'Update visit' : 'Submit visit'}</button></div>` : ''}
  <input type="file" id="photoIn" accept="image/*" capture="environment" hidden>
  <input type="file" id="photoLib" accept="image/*" multiple hidden>`;
  v.innerHTML = html;
  drawScore();
  wireVisit(V, canLog, snap);
}
// Checks a past commitment against today's numbers. Store numbers use this week if the report has it,
// otherwise the month. 6 Elements and practice commitments check what was scored on this visit.
const numOf = t => { const n = parseFloat(String(t ?? '').replace(/[$,%\s]/g, '')); return isFinite(n) ? n : null; };
function metricKeyFor(a) {
  if (a?.key && STORE_METRICS.some(m => m.key === a.key)) return a.key;
  const w = String(a?.what || a?.behavior || '').toLowerCase();
  const hits = STORE_METRICS.filter(m => w.includes(m.label.toLowerCase().replace(' w/ cancellations', '')) || (PLAIN_LABELS[m.key] && w.includes(PLAIN_LABELS[m.key])));
  return hits.sort((x, y) => y.label.length - x.label.length)[0]?.key || null;
}
const PLAIN_LABELS = { protectionSph: 'protection per hour', beddingSph: 'bedding per hour', spg: 'spg', closeRate: 'close rate', financePct: 'finance', beddingPct: 'bedding', protectionAttach: 'protection attach', protectionPct: 'protection', deliveryPct: 'delivery', cancelPct: 'cancel', appsToTraffic: 'apps', sph: 'sales per hour', avgTicket: 'ticket' };
const COACH_WATCH = [
  'Opened with a win before the opportunity',
  'Used the number, then asked a question before telling',
  'Showed the behavior or ran it with them, did not just talk about it',
  'Got a commitment in the consultant\'s words, with a date'];
// Coaching through the store leader: help the leader coach the consultant. The numbers open a
// behavior conversation, then you role-play it with the leader before they do it for real.
function leaderCoachText(cc, first, leader) {
  const L = titleName(leader).split(' ')[0] || 'the leader';
  const main = (cc.items || []).find(f => !f.stretch) || (cc.items || [])[0];
  const out = [`Help ${L} coach ${first}. Use the numbers to start a conversation about behavior, not a lecture about the number.`];
  const facts = [];
  if (cc.strength) facts.push(`strong on ${PLAIN[cc.strength.key] || cc.strength.label.toLowerCase()} (${fmtMetric(cc.strength.key, cc.strength.value)} vs ${fmtMetric(cc.strength.key, cc.strength.goal)} goal)`);
  (cc.items || []).filter(f => !f.stretch).forEach(f => facts.push(`${PLAIN[f.key] || f.label.toLowerCase()} at ${fmtMetric(f.key, f.value)} vs ${fmtMetric(f.key, f.goal)} goal`));
  if (facts.length) out.push(`What the data says about ${first}: ${facts.join('; ')}.`);
  if (cc.strength) out.push(`${L} opens with the win: "${first}, your ${PLAIN[cc.strength.key] || 'numbers'} ${cc.strength.key === 'sph' ? 'are' : 'is'} ahead of goal. That's real."`);
  if (main) {
    const beh = String(main.coach?.doThis || '').split('. ')[0].replace(/\.$/, '');
    out.push(`${L} uses the number to ask, not tell: "Your ${PLAIN[main.key] || main.label.toLowerCase()} is ${fmtMetric(main.key, main.value)} and our goal is ${fmtMetric(main.key, main.goal)}. ${/^walk me through/i.test(main.coach?.ask?.[0] || '') ? main.coach.ask[0] : `Walk me through your last guest. ${main.coach?.ask?.[0] || 'What happened?'}`}"`);
    out.push(`The behavior ${L} coaches: ${beh}. ${L} shows it or runs it with ${first}, not just talks about it.`);
  }
  out.push(`Role-play it with ${L} first: you play ${first} and push back a little, ${L} coaches you. Give one tip, then run it again.`);
  if (cc.drill) out.push(`Then ${L} runs the "${cc.drill.title}" practice with ${first} on the floor.`);
  out.push(cc.commit ? `${L} gets ${first}'s commitment in their own words: ${cc.commit.what.replace(/^[^:]+:\s*/, '').replace(/^./, c => c.toLowerCase())}, ${PLAIN[cc.commit.key] || 'the number'} from ${cc.commit.from} to ${cc.commit.to}${main && cc.commit.to !== fmtMetric(main.key, main.goal) ? ` (a first step toward ${fmtMetric(main.key, main.goal)})` : ''}, and checks it before your next visit.` : `${L} gets ${first}'s commitment in their own words and checks it before your next visit.`);
  return out.join('\n');
}
function followBadge(v, au) {
  const lbl = { yes: ['done', 'Done'], partial: ['part', 'Moving'], no: ['not', 'Not yet'] }[v];
  return `<div class="review">${lbl ? `<span class="rv ${lbl[0]}">${lbl[1]}</span>` : ''}<span class="small">${esc(au?.text || 'Review it with the leader.')}</span></div>`;
}
function autoFollow(a, snap, V) {
  if (!a) return null;
  // A consultant's commitment: check their number in the latest RSA report.
  if (a.cid && a.rkey) {
    const p = (S.rsa?.people || []).find(q => q.cid === a.cid), now = p?.k?.[a.rkey];
    const from = numOf(a.from), to = numOf(a.to);
    if (now == null || from == null || to == null) return null;
    const lower = !!METRICS.find(m => m.key === a.rkey)?.lower;
    const v = (lower ? now <= to : now >= to) ? 'yes' : (lower ? now < from : now > from) ? 'partial' : 'no';
    return { v, now: fmtMetric(a.rkey, now), text: `${titleName(p.name)} is at ${fmtMetric(a.rkey, now)} month to date (was ${a.from}, goal ${a.to}).` };
  }
  const key = metricKeyFor(a);
  const from = numOf(a.from), to = numOf(a.to);
  if (key && from != null && to != null && snap) {
    const per = snap.wtd?.k?.[key] != null ? 'wtd' : 'mtd';
    const now = snap[per]?.k?.[key]; if (now == null) return null;
    const m = STORE_METRICS.find(x => x.key === key), lower = !!m?.lower;
    const better = (x, y) => lower ? x <= y : x >= y;
    const v = better(now, to) ? 'yes' : (lower ? now < from : now > from) ? 'partial' : 'no';
    return { v, now: fmtMetric(key, now), text: `Now ${fmtMetric(key, now)} ${per === 'wtd' ? 'this week' : 'this month'} (was ${a.from}, goal ${a.to}).` };
  }
  const w = String(a.what || '').trim();
  const el = ELEMENTS.find(e => e.t.toLowerCase() === w.toLowerCase());
  if (el) {
    const vals = [...Object.values(V.checks?.[el.key] || {}), ...(el.aor ? Object.values(V.aor || {}) : [])];
    if (!vals.length) return { v: null, text: `Score ${el.t} on today's walk and this updates.` };
    const bad = vals.filter(x => x === 'no' || x === 'needs').length, part = vals.filter(x => x === 'partial').length;
    const v = bad ? 'no' : part ? 'partial' : 'yes';
    return { v, text: `From today's walk: ${bad} not there, ${part} partial, ${vals.length - bad - part} good.` };
  }
  const c = (V.consultants || []).find(x => w.toLowerCase().startsWith(titleName(x.name).toLowerCase() + ':'));
  const sc = c ? Object.values(c.drill?.scored || {}) : [];
  if (c && sc.length) {
    const pts = sc.reduce((t, x) => t + (x === 'yes' ? 1 : x === 'partial' ? 0.5 : 0), 0), tot = drillFor(c.drill.key).watch.length;
    return { v: pts >= tot ? 'yes' : pts >= tot / 2 ? 'partial' : 'no', text: `Today's practice: ${pts} of ${tot}.` };
  }
  return null;
}
// Commitments to choose from, built from today's visit: last visit's commitments that aren't done yet,
// the store's opportunities (as the behavior that moves them), the people coached today, and the
// 6 Elements fixes. Every one is a behavior with a from X to Y, a how, an owner and a date.
function commitmentOptions(V, snap, priorRaw = []) {
  const nextDay = (S.vPlans || []).flatMap(p => p.days || []).filter(d => dayStores(d).includes(V.store) && d.date > V.date).map(d => d.date).sort()[0] || addDays(V.date, 7);
  const leader = V.leaderCommit?.name || V.leaderWin?.name || 'Store leader';
  const beh = k => String(COACHING[k]?.doThis || '').split('. ')[0].replace(/\.$/, '');
  const rest = k => String(COACHING[k]?.doThis || '').split('. ').slice(1).join('. ');
  const out = [];
  priorRaw.forEach(a => {
    const au = autoFollow(a, snap, V); if (au?.v === 'yes') return;
    out.push({ group: 'Carry forward (not there yet)', key: a.cid ? a.key : (a.key || metricKeyFor(a) || undefined), cid: a.cid, rkey: a.rkey, what: String(a.what || a.behavior || '').trim(), from: au?.now || a.from || 'Not there yet', to: a.to || 'Every guest', how: a.how || '', owner: a.owner || leader, due: nextDay, carried: true });
  });
  const L = V.lever ? leverStatus(snap?.mtd).find(l => l.key === V.lever) : null;
  if (L) [...L.inputs].sort((a, b) => ((V.leverInputs || []).includes(b.key) ? 1 : 0) - ((V.leverInputs || []).includes(a.key) ? 1 : 0)).forEach(inp => {
    const first = String(inp.behavior).split('. ')[0].replace(/\.$/, '');
    const how = `${String(inp.behavior).split('. ').slice(1).join('. ')} ${inp.fact || ''} Practice "${drillFor(inp.drill).title}" with the team, then the leader checks it every shift.`.replace(/\s+/g, ' ').trim();
    if (inp.metric && inp.value != null) out.push({ group: `Lever: ${L.label}`, key: inp.metric, what: `${first} (${inp.label})`, from: fmtMetric(inp.metric, inp.value), to: fmtMetric(inp.metric, inp.goal), how, owner: leader, due: nextDay, lever: (V.leverInputs || []).includes(inp.key) });
    else if (!inp.metric) out.push({ group: `Lever: ${L.label}`, what: `${first} (${inp.label})`, from: `Count today: ${inp.count.toLowerCase()}`, to: COUNT_TARGET[inp.key] || 'Every guest', how, owner: leader, due: nextDay, lever: (V.leverInputs || []).includes(inp.key) });
  });
  const playGap = Object.values(V.play || {}).some(x => x === 'no' || x === 'partial');
  const bundleMiss = ['no', 'partial'].includes(V.play?.[3]);
  out.push({ group: 'Run the play', what: 'Present every option with financing + Protection + Premium Delivery (the bundle)', from: bundleMiss ? (V.remote ? 'Missed in today\'s role-play' : 'Not presented on today\'s guest') : (V.remote ? 'Leader\'s count: sales with the full bundle' : 'Count today: sales with the full bundle'), to: 'Every sale',
    how: `Connect and start a cart first, build value from Best, then buying power and the bundle.${activeOffer() ? ` Show the savings in dollars: $100 off every $1,000 with financing AND the bundle (${S.offer.name}).` : ''} Practice "Running the play" at the huddle.`, owner: leader, due: nextDay, playFirst: playGap });
  if (['no', 'partial'].includes(V.play?.[0])) out.push({ group: 'Run the play', what: 'Connect first: greet like a referral and start a cart with every guest', from: V.remote ? 'Leader couldn\'t say how many carts' : 'Missed on today\'s guest', to: '8 of 10 guests with a cart', how: 'Practice "Building the cart" standing up. The leader counts carts started at every huddle.', owner: leader, due: nextDay, playFirst: true });
  if (V.remote && ['no', 'partial'].includes(V.floor?.knew)) out.push({ group: 'Run the play', what: 'Know the floor every shift: guests, carts, finance apps, bundles and coachings', from: 'Couldn\'t answer from memory', to: 'Knows every number at each huddle', how: 'Leader keeps a tally on the floor and opens each huddle with yesterday\'s guests, carts, apps, bundles and who they coached.', owner: leader, due: nextDay, playFirst: true });
  const vWho = S.users.find(u => u.email === V.email);
  if (hasFliq(vWho)) { const u = numOf(V.fliq?.using), fl = numOf(V.fliq?.floor);
    out.push({ group: 'FrontLine IQ', what: 'Every associate gets reps in FrontLine IQ before their first guest', from: u != null && fl ? `${u} of ${fl} associates` : 'Count today', to: fl ? `${fl} of ${fl} associates` : 'Every associate', how: 'Leader checks FrontLine IQ use at open and goes over what it flagged at the huddle.', owner: leader, due: nextDay, fliqFirst: true }); }
  const opps = storeFocus(snap, 6).filter(f => !f.stretch && COACHING[f.key]);
  const focusFirst = [...opps.filter(f => (V.focus || []).includes(f.key)), ...opps.filter(f => !(V.focus || []).includes(f.key))];
  focusFirst.forEach(f => out.push({ group: "Store opportunities", key: f.key, what: `${beh(f.key)} (${f.label})`, from: fmtMetric(f.key, f.value), to: fmtMetric(f.key, f.target ?? f.goal),
    how: `${rest(f.key)} ${leader === 'Store leader' ? 'The leader' : titleName(leader).split(' ')[0]} checks it at every huddle.`.trim(), owner: leader, due: nextDay, coached: (V.focus || []).includes(f.key) }));
  (V.consultants || []).forEach(c => {
    const p = (S.rsa?.people || []).find(q => q.cid === c.cid); if (!p) return;
    const cc = consultantCoaching({ p, store: p.store || V.store, why: c.why, goals: DEFAULT_GOALS, pace: paceFactor(S.rsa?.to), lever: c.lever });
    if (cc.commit) out.push({ group: 'People coached today', cid: c.cid, rkey: cc.commit.key, what: cc.commit.what.replace(/^[^:]+:/, `${titleName(c.name)}:`), from: cc.commit.from, to: cc.commit.to, how: c.drill?.adjust ? `${cc.commit.how} Focus on: ${c.drill.adjust}` : cc.commit.how, owner: titleName(c.name), due: nextDay });
  });
  visitSummary(V).fixes.forEach(fx => out.push({ group: '6 Elements fixes', what: `${fx}: fix what was flagged on today's walk`, from: 'Needs work today', to: 'Grand Opening Ready', how: 'Leader walks it at open every day and sends a photo when it\'s right.', owner: leader, due: nextDay }));
  // One commitment per number: the same store metric (or the same person's metric) only shows once.
  const seen = new Set();
  return out.filter(o => { const ks = [o.what.toLowerCase(), o.key && !o.cid ? 'm:' + o.key : null, o.cid ? `p:${o.cid}:${o.rkey}` : null].filter(Boolean); if (!o.what || ks.some(k => seen.has(k))) return false; ks.forEach(k => seen.add(k)); return true; });
}
// Fills only empty slots, best first: carry-forwards, what the leader is coaching today, people
// coached, then 6 Elements fixes and other store opportunities. Returns how many it filled.
function fillCommitments(V, focusAll, opts) {
  const all = opts || commitmentOptions(V, S.vSnap, S.vPriorRaw || []);
  const rank = o => o.carried ? 0 : o.lever ? 0.5 : o.playFirst ? 0.6 : o.fliqFirst ? 0.7 : o.coached ? 1 : o.group === 'People coached today' ? 2 : o.group === '6 Elements fixes' ? 3 : String(o.group).startsWith('Lever') ? 3.5 : 4;
  const sugg = [...all].sort((a, b) => rank(a) - rank(b));
  const have = new Set(V.actions.filter(hasCommitment).map(a => String(a.what || a.behavior).toLowerCase()));
  let n = 0;
  for (let i = 0; i < 3; i++) {
    if (hasCommitment(V.actions[i])) continue;
    const next = sugg.find(x => !have.has(x.what.toLowerCase()));
    if (!next) break;
    have.add(next.what.toLowerCase()); const { group, coached, lever, playFirst, fliqFirst, ...rest } = next; V.actions[i] = { ...(V.actions[i] || {}), ...rest, suggested: true }; n++;
  }
  return n;
}
function winsFor(m) {
  const out = [];
  if (m.vsBud.netSales >= 0) out.push(`Sales ${pct(m.vsBud.netSales)} to budget`);
  if (m.vsLy.spg >= 0) out.push(`SPG with cancellations ${pct(m.vsLy.spg)} vs LY`);
  if (m.vsBud.closeRate >= 0) out.push(`Close rate +${Math.round(m.vsBud.closeRate)} bps to budget`);
  if (m.k.financePct >= STORE_GOALS.financePct) out.push(`Finance ${m.k.financePct.toFixed(0)}% of sales`);
  if (m.k.protectionAttach >= STORE_GOALS.protectionAttach) out.push(`Protection attach ${m.k.protectionAttach.toFixed(0)}%`);
  return out;
}
function drawScore() {
  const el = $('#vscore'); if (!el || !S.V) return;
  const s = visitScore(S.V), sm = visitSummary(S.V);
  el.innerHTML = `<div><p class="eyebrow">Visit score</p><b class="big">${s ? s.pct + '%' : '--'}</b><span class="small"> ${s ? `${s.n} of ${s.t} scored items` : 'Builds as you score. Yes = 1, Partial = half.'}</span></div>
    ${sm.fixes.length ? `<p class="small" style="margin:6px 0 0">To fix: ${sm.fixes.map(esc).join(', ')}</p>` : ''}`;
}
function setPath(o, path, val) {
  const keys = path.split('.'); let cur = o;
  keys.slice(0, -1).forEach(k => { if (cur[k] == null || typeof cur[k] !== 'object') cur[k] = {}; cur = cur[k]; });
  const last = keys[keys.length - 1];
  if (val === null) delete cur[last]; else cur[last] = val;
}
function getPath(o, path) { return path.split('.').reduce((c, k) => (c == null ? undefined : c[k]), o); }
let vTimer = null;
// ---------------------------------------------------------------- edits after submit
// Once a visit is submitted, any later change is stamped on the visit (who, when, which parts) and
// sent to the VP's review list, so nothing gets quietly rewritten after the fact.
const EDIT_SKIP = new Set(['at', 'by', 'needScore', 'commitments', 'edits', 'lastEdit', 'submittedAt', 'status', 'follow', 'photosMissing', 'photosMissingList', 'id', 'email', 'name', 'role', 'store', 'date']);
const EDIT_LABEL = { leaderWin: 'Leader win', checks: '6 Elements walk', segs: 'Value segments', aor: 'AOR walk', elNotes: '6 Elements notes', segMeta: 'Value segments',
  consultants: 'Consultants', actions: 'Action plan', leaderCommit: 'Store leader commitment', working: 'What is working', notes: 'Notes', teamNotes: 'Team notes',
  reflection: 'Reflection', photoReason: 'Reason for missing photos', focus: 'Team focus', lever: 'Lever', play: 'Run the play', fliq: 'FrontLine IQ', floor: 'Leader knows the floor', intent: 'Why you are here', vtype: 'Visit type', kind: 'Visit type' };
const editParts = V => { const o = {}; for (const k of Object.keys(V).sort()) if (!EDIT_SKIP.has(k)) o[k] = JSON.stringify(V[k] ?? null); return o; };
const editLabel = k => EDIT_LABEL[k] || k.replace(/([A-Z])/g, ' $1').replace(/^./, c => c.toUpperCase());
// Called after every save of a submitted visit: compares with the last known version.
function noteEdit(V, extra) {
  if (V.status !== 'done' || !S.vBase || S.vBaseFor !== V.id) return;
  if (!V.submittedAt) V.submittedAt = V.date;
  const now = editParts(V), before = S.vBase;
  const changed = [...new Set([...Object.keys(now), ...Object.keys(before)])].filter(k => now[k] !== before[k]).map(editLabel);
  if (extra) changed.push(extra);
  if (!changed.length) return;
  S.vBase = now;
  V.edits = V.edits || [];
  const at = new Date().toISOString(), last = V.edits[V.edits.length - 1];
  // One entry per sitting: changes by the same person within 30 minutes merge.
  if (last && last.by === S.user.email && Date.parse(at) - Date.parse(last.at) < 30 * 60000) { last.at = at; last.parts = [...new Set([...last.parts, ...changed])]; }
  else V.edits.push({ at, by: S.user.email, name: S.user.name || S.user.email, parts: [...new Set(changed)] });
  V.lastEdit = V.edits[V.edits.length - 1];
  S.vEditPending = V.id;
}
// Tell the VP once per sitting, when the leader leaves the visit or taps Update.
async function sendEditAlert(V) {
  if (S.vEditPending !== V?.id || !V.lastEdit || isAdmin()) { S.vEditPending = null; return; }
  S.vEditPending = null;
  const e = V.lastEdit;
  const a = { id: `${V.email}_edit_${Date.now()}`, type: 'visitEdit', email: S.user.email, name: S.user.name || S.user.email, forEmail: V.email, visitId: V.id, store: V.store, visitDate: V.date,
    weekStart: weekStartOf(V.date), submittedAt: V.submittedAt, parts: e.parts, at: e.at, reason: 'Edited after submit', changes: [], seen: false };
  try { await S.be.saveAlert(a); } catch (err) { console.warn('edit alert', err); }
}
// Recount missing photos after the visit is submitted, so adding them later clears the flag.
function refreshMissing(V) {
  if (V.remote) return;
  const missing = ELEMENTS.flatMap(e => photoAreas(e).filter(a => !(S.vPhotos || []).some(p => p.el === e.key && p.item === a)).map(a => `${e.t}: ${a}`));
  V.photosMissing = missing.length; V.photosMissingList = missing; if (!missing.length) V.photoReason = '';
}
const editedText = V => V?.edits?.length ? V.edits.map(e => `${dayLabel(e.at.slice(0, 10))} ${new Date(e.at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })} by ${e.name}: ${e.parts.join(', ')}`) : [];
function saveDraft(now) {
  S.vEdits = (S.vEdits || 0) + 1;
  const V = S.V; if (!V) return;
  V.at = new Date().toISOString(); V.by = S.user.email;
  V.needScore = S.scores[V.store]?.score ?? null;
  V.commitments = visitSummary(V).commitments.join('\n');
  noteEdit(V);
  localSet(V);
  const s = $('#vsaved'); if (s && V.status !== 'done') s.textContent = 'Saving…';
  clearTimeout(vTimer);
  vTimer = setTimeout(async () => {
    try {
      await S.be.saveVisit(structuredClone(V));
      const i = S.visits.findIndex(y => y.id === V.id); if (i >= 0) S.visits[i] = structuredClone(V); else S.visits.push(structuredClone(V));
      const t = $('#vsaved'); if (t) t.textContent = V.status === 'done' ? `Submitted · saved ${new Date().toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}` : `Draft saved ${new Date().toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`;
    } catch (e) { const t = $('#vsaved'); if (t) t.textContent = 'Saved on this device. Will retry when you are back online.'; }
  }, now ? 0 : 900);
}
function wireVisit(V, canLog, snap) {
  const v = $('#view');
  if (S.vBaseFor !== V.id) { S.vBaseFor = V.id; S.vBase = editParts(V); }
  wireWho(v, () => viewVisit());
  $('#back').onclick = () => { stopMic(); if (canLog) { saveDraft(true); sendEditAlert(V); } S.vBaseFor = null; S.visit = null; S.V = null; renderShell(); };
  v.querySelectorAll('.vsec-hd').forEach(h => h.onclick = () => {
    const secEl = h.parentElement, k = secEl.dataset.sec;
    secEl.classList.toggle('open'); h.setAttribute('aria-expanded', secEl.classList.contains('open'));
    if (secEl.classList.contains('open')) V_OPEN.add(k); else V_OPEN.delete(k);
  });
  v.querySelectorAll('[data-period]').forEach(b => b.onclick = () => { S.vPeriod = b.dataset.period; viewVisit(); });
  if (!canLog) return;
  wireMics(v);
  const vt = $('#vtype'); vt.onchange = () => { V.vtype = vt.value; saveDraft(); };
  v.querySelectorAll('[data-path]').forEach(b => b.onclick = () => {
    const path = b.dataset.path, cur = getPath(V, path), val = cur === b.dataset.v ? null : b.dataset.v;
    setPath(V, path, val);
    if (/^consultants\.\d+\.mode$/.test(path)) { if (!val) setPath(V, path, 'direct'); saveDraft(); return viewVisit(); }
    if (path.startsWith('follow.')) { setPath(V, 'followAuto.' + path.slice(7), null); const au = b.closest('.item')?.querySelector('.auto .pill'); if (au) au.remove(); }
    b.parentElement.querySelectorAll('button').forEach(o => o.classList.toggle('on', o.dataset.v === val));
    const ip = b.closest('.item')?.querySelector('.iphoto'); if (ip) ip.hidden = !(val === 'no' || val === 'partial' || val === 'needs');
    const ap = b.closest('.item')?.querySelector('.aphoto:not(.done) > span'); if (ap) { const bad = val === 'no' || val === 'partial' || val === 'needs'; ap.textContent = bad ? 'Needs work: get a photo and say what has to change.' : 'Photo of this area'; ap.className = 'small ' + (bad ? 'warn' : 'muted'); }
    const dc = path.includes('.drill.scored.') && b.closest('.drill')?.querySelector('.dcount'); if (dc) dc.textContent = `(${Object.keys(getPath(V, path.split('.').slice(0, 4).join('.')) || {}).length} of ${dc.dataset.total})`;
    const prac = b.closest('details.prac'); if (prac && path.includes('.practice.')) prac.querySelector('summary .small').textContent = `${Object.keys(getPath(V, path.split('.').slice(0, 3).join('.')) || {}).length} of ${PRACTICE.length} scored`;
    drawScore(); saveDraft();
  });
  v.querySelectorAll('[data-field]').forEach(inp => inp.oninput = inp.onchange = () => { setPath(V, inp.dataset.field, inp.value); if (inp.dataset.field === 'leaderWin.text') { V.leaderWin.touched = true; V.leaderWin.suggested = false; } saveDraft(); });
  v.querySelectorAll('[data-lpick]').forEach(b => b.onclick = () => {
    const k = b.dataset.lpick; V.lever = V.lever === k ? null : k;
    const L = leverStatus(S.vSnap?.mtd).find(l => l.key === V.lever);
    V.leverInputs = L ? defaultInputs(L) : [];
    // Suggested commitments follow the lever; anything the leader typed or picked stays.
    V.actions = V.actions.map(a => a?.suggested && !a.carried ? {} : a); fillCommitments(V);
    saveDraft(); V_OPEN.add('lever'); Promise.resolve(viewVisit()).then(() => setTimeout(() => $('#lvinputs')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50));
  });
  v.querySelectorAll('[data-linput]').forEach(b => b.onclick = () => {
    const k = b.dataset.linput, cur = V.leverInputs || [];
    V.leverInputs = cur.includes(k) ? cur.filter(z => z !== k) : [...cur, k].slice(-2);
    V.actions = V.actions.map(a => a?.suggested && !a.carried ? {} : a); fillCommitments(V);
    saveDraft(); V_OPEN.add('lever'); viewVisit();
  });
  // The list is rebuilt when opened, so it reflects what was scored since the page drew.
  const apickHtml = () => { S.cOpts = commitmentOptions(V, S.vSnap, S.vPriorRaw || []); return `<option value="">Choose from today's opportunities and coaching…</option>${[...new Set(S.cOpts.map(o => o.group))].map(gr => `<optgroup label="${esc(gr)}">${S.cOpts.map((o, k) => o.group === gr ? `<option value="${k}">${esc(o.what)}: ${esc(o.from)} to ${esc(o.to)}</option>` : '').join('')}</optgroup>`).join('')}`; };
  v.querySelectorAll('[data-apick]').forEach(sel => { sel.onfocus = sel.onmousedown = () => { if (sel.dataset.fresh !== String(S.vEdits || 0)) { sel.innerHTML = apickHtml(); sel.dataset.fresh = String(S.vEdits || 0); } }; });
  v.querySelectorAll('[data-apick]').forEach(sel => sel.onchange = () => {
    const o = S.cOpts[+sel.value]; if (!o) return;
    const { group, coached, lever, playFirst, fliqFirst, ...rest } = o; V.actions[+sel.dataset.apick] = { ...rest, suggested: false };
    saveDraft(); V_OPEN.add('action'); viewVisit(); toast('Commitment filled in. Change anything.');
  });
  wireLeaderPicks(v);
  v.querySelectorAll('[data-focus]').forEach(b => b.onclick = () => {
    const k = b.dataset.focus;
    if (V.focus.includes(k)) V.focus = V.focus.filter(x => x !== k);
    else { if (V.focus.length >= 2) return toast('Two focus items at most. Tap one to drop it first.', true); V.focus.push(k); }
    v.querySelectorAll('[data-focus]').forEach(o => { const on = V.focus.includes(o.dataset.focus); o.classList.toggle('primary', on); o.textContent = on ? 'Coaching this ✓' : 'Coach this'; o.setAttribute('aria-pressed', on); o.closest('.fcard').classList.toggle('on', on); });
    saveDraft();
  });
  v.querySelectorAll('[data-segp]').forEach(sel => sel.onchange = () => {
    const gi = sel.dataset.segp, p = (S.rsa?.people || []).find(q => q.cid === sel.value) || V.consultants.find(c => c.cid === sel.value);
    setPath(V, `segMeta.${gi}.cid`, sel.value || null); setPath(V, `segMeta.${gi}.name`, p?.name || null); saveDraft();
  });
  v.querySelectorAll('[data-helpc]').forEach(b => b.onclick = () => {
    const p = (S.rsa?.people || []).find(q => q.cid === b.dataset.helpc); if (!p) return;
    V.consultants.push({ cid: p.cid, name: p.name, why: 'model', lever: b.dataset.lever, notes: '', practice: {} });
    V_OPEN.add('people'); saveDraft(); viewVisit(); toast(`${titleName(p.name)} added. Recognize them and have them show the team.`);
  });
  v.querySelectorAll('[data-dragc]').forEach(b => b.onclick = () => {
    const p = (S.rsa?.people || []).find(q => q.cid === b.dataset.dragc); if (!p) return;
    const lever = b.dataset.lever, rk = STORE_TO_RSA[lever];
    const drillKey = { creditApps: 'appsToTraffic', protectionSph: 'protectionPct', beddingSph: 'beddingPct' }[rk] || rk;
    V.consultants.push({ cid: p.cid, name: p.name, why: 'drag', lever, notes: '', practice: {}, drill: { key: drillKey, title: drillFor(drillKey).title } });
    V_OPEN.add('people'); saveDraft(); viewVisit(); toast(`${titleName(p.name)} added. Coaching starts with ${(STORE_METRICS.find(m => m.key === lever)?.label || lever).toLowerCase()}.`);
  });
  v.querySelectorAll('[data-drill]').forEach(sel => sel.onchange = () => {
    const c = V.consultants[+sel.dataset.drill];
    c.drill = { key: sel.value, title: drillFor(sel.value).title, well: c.drill?.well || '', adjust: c.drill?.adjust || '', manual: true };
    saveDraft(); viewVisit();
  });
  v.querySelectorAll('[data-rmc]').forEach(b => b.onclick = () => { V.consultants.splice(+b.dataset.rmc, 1); saveDraft(); viewVisit(); });
  const addGo = () => {
    const typed = ($('#addc').value || '').trim(); if (!typed) return toast('Type or say a name first.', true);
    const key = typed.toLowerCase().replace(/[^a-z]/g, '');
    const all = S.rsa?.people || [];
    const p = all.find(q => q.name.toLowerCase().replace(/[^a-z]/g, '') === key) || all.find(q => q.name.toLowerCase().replace(/[^a-z]/g, '').startsWith(key));
    const r = p ? null : (S.roster || []).find(q => q.name.toLowerCase().replace(/[^a-z]/g, '') === key);
    const cid = p ? p.cid : r ? r.cid : cidOf(typed);
    if (V.consultants.some(c => c.cid === cid)) return toast('Already on this visit.', true);
    V.consultants.push({ cid, name: p ? p.name : r ? r.name : typed, why: 'added', notes: '', practice: {} });
    toast(p ? `${titleName(p.name)} added with suggested coaching.` : `${typed} added. Not in the RSA report, so the suggestion is the stand-up practice.`);
    saveDraft(); viewVisit();
  };
  const ab = $('#addcgo'); if (ab) ab.onclick = addGo;
  const ai = $('#addc'); if (ai) ai.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); addGo(); } };
  v.querySelectorAll('[data-usec]').forEach(b => b.onclick = () => {
    const ci = +b.dataset.usec, box = $('#cn' + ci), text = b.closest('.suggest').querySelector('.stext').textContent;
    box.value = box.value.trim() ? box.value.trim() + '\n\n' + text : text;
    V.consultants[ci].notes = box.value; saveDraft(); toast('Added to your notes. Edit anything.');
  });
  // Photos: shrunk on the phone before saving so they stay small.
  const pin = $('#photoIn'), plib = $('#photoLib'); let pel = 'general', pitem = '';
  const room = () => PHOTO_MAX - (S.vPhotos || []).length;
  const pick = (el, input, item = '') => { if (room() <= 0) return toast(`${PHOTO_MAX} photos per visit. Remove one to add another.`, true); pel = el; pitem = item; input.value = ''; input.click(); };
  v.querySelectorAll('[data-photo]').forEach(b => b.onclick = () => pick(b.dataset.photo, b.dataset.src === 'lib' ? plib : pin, b.dataset.item || ''));
  const pt = $('#ptag'); if (pt) pt.onchange = () => { S.pTag = pt.value; };
  const pc = $('#pcam'); if (pc) pc.onclick = () => pick(S.pTag || 'general', pin);
  const pl = $('#plib'); if (pl) pl.onclick = () => pick(S.pTag || 'general', plib);
  const addFiles = async files => {
    const list = [...files].slice(0, room()); if (!list.length) return;
    try {
      for (const f of list) {
        const data = await shrinkImage(f, 900, 0.6);
        const ph = { id: `${V.id}_${Date.now()}_${Math.round(Math.random() * 1e4)}`, visitId: V.id, email: V.email, store: V.store, date: V.date, el: pel, item: pitem, caption: '', data };
        await S.be.savePhoto(ph); S.vPhotos = [...(S.vPhotos || []), ph]; if (V.status === 'done') { refreshMissing(V); noteEdit(V, 'Photos'); saveDraft(); }
      }
      if (pel === 'general') V_OPEN.add('photos'); viewVisit(); toast(list.length > 1 ? `${list.length} photos added. Add a caption to each.` : 'Photo added. Add a caption.');
    } catch (e) { toast('Could not add that photo. Try again.', true); }
  };
  pin.onchange = () => addFiles(pin.files || []);
  plib.onchange = () => addFiles(plib.files || []);
  const capTimers = {};
  v.querySelectorAll('[data-pcap]').forEach(inp => inp.oninput = () => {
    const ph = S.vPhotos.find(x => x.id === inp.dataset.pcap); if (!ph) return;
    ph.caption = inp.value; clearTimeout(capTimers[ph.id]);
    capTimers[ph.id] = setTimeout(() => S.be.savePhoto(ph).catch(() => {}), 800);
  });
  v.querySelectorAll('[data-delphoto]').forEach(b => b.onclick = async () => {
    await S.be.deletePhoto(b.dataset.delphoto); S.vPhotos = S.vPhotos.filter(p => p.id !== b.dataset.delphoto); if (V.status === 'done') { refreshMissing(V); noteEdit(V, 'Photos'); saveDraft(); } viewVisit();
  });
  const sg = $('#apsugg'); if (sg) sg.onclick = () => {
    const n = fillCommitments(V, storeFocus(snap, 4));
    if (!n) return toast('All 3 commitments are filled. Clear one to get a new suggestion.', true);
    saveDraft(); viewVisit(); toast(`${n} suggested. Change anything.`);
  };
  // Editing a suggested commitment makes it yours.
  v.querySelectorAll('[data-field^="actions."]').forEach(inp => inp.addEventListener('input', () => {
    const i = +inp.dataset.field.split('.')[1], a = V.actions[i]; if (!a) return;
    a.suggested = false;
    const pv = $('#apv' + i); if (pv) { pv.hidden = !hasCommitment(a); pv.textContent = hasCommitment(a) ? commitmentText(a) : ''; }
  }));
  $('#vprint').onclick = () => { V_OPEN.clear(); ['why', 'win', 'follow', 'lever', 'people', 'action', 'reflect', 'el1', 'el2', 'el3', 'el4', 'el5', 'el6'].forEach(k => V_OPEN.add(k)); document.querySelectorAll('.vsec').forEach(s => s.classList.add('open')); setTimeout(() => window.print(), 200); };
  // Take the leader straight to what needs fixing: open its section, scroll it to the middle of the
  // screen, put the cursor in it and outline it until they change it.
  const goFix = (secKey, target, msg) => {
    V_OPEN.add(secKey); const sec0 = document.querySelector(`[data-sec="${secKey}"]`); if (sec0) { sec0.classList.add('open'); sec0.querySelector('.vsec-hd')?.setAttribute('aria-expanded', 'true'); }
    const el = (typeof target === 'string' ? document.querySelector(target) : target) || sec0;
    toast(msg, true);
    requestAnimationFrame(() => setTimeout(() => {
      el?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      const box = el?.closest('.item, .acard, .aphoto, label, div') || el; box?.classList.add('needfix');
      const clear = () => box?.classList.remove('needfix'); el?.addEventListener?.('input', clear, { once: true }); el?.addEventListener?.('click', clear, { once: true });
      if (el?.matches?.('input, textarea, select')) setTimeout(() => el.focus({ preventScroll: true }), 400);
    }, 30));
  };
  const goPhoto = m => goFix('el' + m.e.n, [...document.querySelectorAll(`[data-sec="el${m.e.n}"] [data-photo="${m.e.key}"]`)].find(b => b.dataset.item === m.a)?.closest('.aphoto'), `Take the photo of ${m.a}, then tap Submit again.`);
  const photoReasonBox = missing => {
    document.querySelector('#photoreason')?.remove();
    const box = document.createElement('section'); box.id = 'photoreason'; box.className = 'panel editbox';
    const byEl = {}; missing.forEach(m => (byEl[m.e.t] ||= []).push(m));
    box.innerHTML = `<h3 style="margin:0 0 4px">${missing.length} photo${missing.length > 1 ? 's' : ''} missing</h3>
      <p class="small" style="margin:0 0 8px">You can submit without them. The visit will be flagged with what's missing and your reason, and your VP sees it.</p>
      ${Object.entries(byEl).map(([t, ms]) => `<p class="small" style="margin:4px 0"><b>${esc(t)}:</b> ${ms.map(m => `<button type="button" class="link" data-gophoto="${esc(m.e.key)}|${esc(m.a)}">${esc(m.a)}</button>`).join(', ')}</p>`).join('')}
      <label for="prsn" style="margin-top:10px">Why are they missing?<select id="prsn"><option value="">Pick a reason</option>${['Area was blocked or under repair', 'Ran out of time on the visit', 'Phone or camera problem', 'Photos taken but did not upload', 'Other'].map(r => `<option>${r}</option>`).join('')}</select></label>
      ${fieldBox('prnote', 'Details', '', 2, 'What happened, and when the photos will be taken.', '', '')}
      <div class="row" style="gap:8px;margin-top:8px"><button type="button" class="btn primary" id="prok">Submit without these photos</button><button type="button" class="btn" id="prgo">Take the photos first</button></div>`;
    const bar = document.querySelector('.vbar'); bar.parentNode.insertBefore(box, bar);
    box.scrollIntoView({ behavior: 'smooth', block: 'center' }); box.classList.add('needfix');
    box.querySelectorAll('[data-gophoto]').forEach(b => b.onclick = () => { const [k, a] = b.dataset.gophoto.split('|'); goPhoto(missing.find(m => m.e.key === k && m.a === a)); });
    $('#prgo').onclick = () => goPhoto(missing[0]);
    wireMics?.(box);
    $('#prok').onclick = () => {
      const r = $('#prsn').value, n = ($('#prnote').value || '').trim();
      if (!r) return toast('Pick a reason first.', true);
      if (r === 'Other' && n.length < 5) return toast('Add a few words on why.', true);
      V.photoReason = n ? `${r}: ${n}` : r; saveDraft(true); box.remove(); $('#vsubmit').click();
    };
  };
  $('#vsubmit').onclick = async () => {
    stopMic();
    const sm = visitSummary(V);
    const pi = V.actions.findIndex(a => hasCommitment(a) && (!String(a.from || '').trim() || !String(a.to || '').trim()));
    if (pi >= 0) { const a = V.actions[pi]; return goFix('action', String(a.from || '').trim() ? `#apt${pi}` : `#apf${pi}`, `Commitment ${pi + 1} needs a ${String(a.from || '').trim() ? 'To' : 'From'}. Fill it in, then tap Submit again.`); }
    if (!sm.commitments.length) return goFix('action', '#apw0', 'Add at least one commitment in the action plan, then tap Submit again.');
    if (!V.remote) {
      const missing = ELEMENTS.flatMap(e => photoAreas(e).filter(a => !(S.vPhotos || []).some(p => p.el === e.key && p.item === a)).map(a => ({ e, a })));
      V.photosMissing = missing.length;
      V.photosMissingList = missing.map(m => `${m.e.t}: ${m.a}`);
      if (!missing.length) V.photoReason = '';
      // Photos can be skipped, but the Market Leader says why, and the visit is flagged with what's missing.
      else if (!String(V.photoReason || '').trim()) return photoReasonBox(missing);
    }
    const wasDone = V.status === 'done';
    if (!wasDone) S.vBaseFor = null; // the first submit is not an edit
    V.status = 'done'; V.submittedAt = V.submittedAt || new Date().toISOString();
    saveDraft(true);
    if (wasDone) sendEditAlert(V); else { S.vBaseFor = V.id; S.vBase = editParts(V); }
    try {
      const plan = await S.be.plan(V.email, weekStartOf(V.date));
      if (V.remote) {
        if (plan) { plan.remoteDone = [...new Set([...(plan.remoteDone || []), V.store])]; await S.be.savePlan(plan); }
      } else {
        const day = plan?.days.find(d => d.date === V.date && d.store === V.store);
        const stopDay = !day && plan?.days.find(d => d.date === V.date && (d.stops || []).some(x => x.store === V.store));
        if (day && day.status !== 'done') { day.status = 'done'; await S.be.savePlan(plan); }
        else if (stopDay) { stopDay.stops.find(x => x.store === V.store).status = 'done'; await S.be.savePlan(plan); }
      }
    } catch (e) {}
    localDrop(V.id);
    S.lastVisit = latestVisitMap([...S.visits.filter(y => y.id !== V.id), V]); S.scores = scoresFor(S.daily);
    toast(V.remote ? 'Remote coaching logged. The recap message is ready in Team messages.' : 'Visit submitted. The recap message is ready in Team messages.');
    S.msgType = 'recap'; S.msgVisit = V.id;
    viewVisit();
  };
}
function shrinkImage(file, max, q) {
  return new Promise((res, rej) => {
    const img = new Image(), url = URL.createObjectURL(file);
    img.onload = () => {
      const r = Math.min(1, max / Math.max(img.width, img.height));
      const c = document.createElement('canvas'); c.width = Math.round(img.width * r); c.height = Math.round(img.height * r);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
      res(c.toDataURL('image/jpeg', q));
    };
    img.onerror = rej; img.src = url;
  });
}

const fmtMetric = (key, v) => {
  const m = STORE_METRICS.find(x => x.key === key) || { fmt: /pct|margin|rate/i.test(key) ? 'pct' : 'money' };
  if (key === 'creditApps') return String(Math.round(v));
  return fmt(m, v);
};

// ---------------------------------------------------------------- stores
function viewStores() {
  const v = $('#view');
  const ls = leaders();
  let list = seesAll() ? STORES.map(s => s.name) : (S.user.stores || []);
  if (seesAll() && S.storeFilter?.startsWith('mkt:')) list = (S.markets.find(m => 'mkt:' + m.id === S.storeFilter)?.stores || []);
  else if (seesAll() && S.storeFilter && S.storeFilter !== '*') list = S.storeFilter === '__none' ? list.filter(s => !leaderOf(s)) : (ls.find(l => l.email === S.storeFilter)?.stores || []);
  list = list.slice().sort((a, b) => (S.scores[b]?.score ?? -1) - (S.scores[a]?.score ?? -1));
  const cell = (val, cls = '') => `<td class="num ${cls}">${val}</td>`;
  const sign = v => v == null ? '' : v >= 0 ? 'good' : v <= -10 ? 'bad' : 'warn';
  const row = s => {
    const snap = S.daily?.stores?.[s], m = snap?.mtd, w = snap?.wtd, sc = S.scores[s];
    const lv = S.lastVisit[s], l = leaderOf(s);
    return `<tr class="click" data-store="${esc(s)}">
      <td class="nm">${esc(s)}${seesAll() ? `<br><small class="muted">${esc(marketOf(s)?.name ? marketOf(s).name + ' · ' : '')}${esc(l ? l.name || l.email : 'No Market Leader')}</small>` : ''}</td>
      <td>${needChip(sc?.score)}</td>
      ${cell(w?.vsBud.netSales != null ? pct(w.vsBud.netSales) : '--', sign(w?.vsBud.netSales))}
      ${cell(m?.vsBud.netSales != null ? pct(m.vsBud.netSales) : '--', sign(m?.vsBud.netSales))}
      ${cell(m?.k.spg != null ? '$' + m.k.spg.toFixed(0) : '--')}
      ${cell(m?.vsLy.spg != null ? pct(m.vsLy.spg) : '--', sign(m?.vsLy.spg))}
      ${cell(m?.vsBud.closeRate != null ? Math.round(m.vsBud.closeRate) : '--', m?.vsBud.closeRate == null ? '' : m.vsBud.closeRate >= 0 ? 'good' : m.vsBud.closeRate > -200 ? 'warn' : 'bad')}
      ${cell(m?.vsLy.traffic != null ? pct(m.vsLy.traffic) : '--')}
      ${cell(m?.k.protectionAttach != null ? m.k.protectionAttach.toFixed(0) + '%' : '--', m?.k.protectionAttach == null ? '' : m.k.protectionAttach >= STORE_GOALS.protectionAttach ? 'good' : 'warn')}
      ${cell(m?.k.cancelPct != null ? m.k.cancelPct.toFixed(1) + '%' : '--', m?.k.cancelPct == null ? '' : m.k.cancelPct <= STORE_GOALS.cancelPct ? 'good' : 'warn')}
      ${(() => { const cx = S.carts?.stores?.[s]; return cell(cx ? `${cx.n} · $${Math.round(cx.value / 1000)}k` : '--', cx?.due ? 'warn' : ''); })()}
      <td class="small">${(() => { const tm = S.teams?.[s]; if (!tm) return '<span class="muted">--</span>'; const bits = []; if (tm.below.length) bits.push(`<span class="bad">${tm.below.length} below</span>`); if (tm.slipping.length) bits.push(`<span class="warn">${tm.slipping.length} slipping</span>`); return bits.join(' · ') || '<span class="good">On track</span>'; })()}</td>
      <td>${lv ? `${daysApart(lv, today())}d ago` : '<span class="warn">None</span>'}</td>
      <td class="small" style="white-space:normal;min-width:200px">${esc(sc?.parts?.[0]?.text || '')}</td>
    </tr>`;
  };
  v.innerHTML = `
  <div class="spread">
    <div><h2 class="big" style="margin:0">${seesAll() ? 'All stores' : 'My stores'}, ranked by need</h2><p class="small" style="margin:4px 0 0">${dataLine()}</p></div>
    ${seesAll() ? `<label for="sf" style="margin:0">Show<select id="sf"><option value="*">All stores</option>${ls.map(l => `<option value="${esc(l.email)}" ${S.storeFilter === l.email ? 'selected' : ''}>${esc(l.name || l.email)}</option>`).join('')}<option value="__none" ${S.storeFilter === '__none' ? 'selected' : ''}>No Market Leader</option>${(S.markets || []).length ? `<optgroup label="Markets">${S.markets.map(m => `<option value="mkt:${esc(m.id)}" ${S.storeFilter === 'mkt:' + m.id ? 'selected' : ''}>${esc(m.name)}</option>`).join('')}</optgroup>` : ''}</select></label>` : ''}
  </div>
  <div class="panel" style="padding:0;overflow:hidden"><div class="scroller"><table class="grid">
    <thead><tr><th>Store</th><th>Need</th><th class="num">Sales WTD vs bud</th><th class="num">Sales MTD vs bud</th><th class="num">SPG w/ canc</th><th class="num">SPG vs LY</th><th class="num">Close rate bps</th><th class="num">Traffic vs LY</th><th class="num">Prot attach</th><th class="num">Cancel %</th><th class="num">Open carts</th><th>Consultants</th><th>Last visit</th><th>Top reason</th></tr></thead>
    <tbody>${list.map(row).join('') || '<tr><td colspan="14">No stores.</td></tr>'}</tbody>
  </table></div></div>
  <p class="small">Tap a store for its coaching plan. SPG uses SPG with cancellations. Close rate is basis points against budget.</p>`;
  v.querySelectorAll('[data-store]').forEach(r => r.onclick = () => openVisit({ store: r.dataset.store, date: today(), email: seesAll() ? (leaderOf(r.dataset.store)?.email || S.user.email) : S.user.email, kind: 'drop-in' }));
  const sf = $('#sf'); if (sf) sf.onchange = () => { S.storeFilter = sf.value; viewStores(); };
}

// ---------------------------------------------------------------- daily brief
// The first thing a Market Leader or director opens each morning: yesterday's wins and opportunities
// for their stores and people, where they're going today, and what's due. It's for them, not the team.
// The VP's review list sits on top of the daily brief: newest first, unreviewed at the top.
async function viewBrief() {
  await viewBriefInner();
  if (!seesAll()) return;
  const list = (S.alerts || []).filter(a => !a.seen || a.at >= addDays(today(), -7)).sort((a, b) => (a.seen - b.seen) || b.at.localeCompare(a.at)).slice(0, 12);
  if (!list.length) return;
  const el = document.createElement('section');
  el.className = 'panel'; el.id = 'alerts'; el.style.borderLeft = '6px solid #F68C2C';
  el.innerHTML = `<h3>Changes to review</h3>
    <p class="small">Schedule changes (with the reason) and visits edited after they were submitted land here.</p>
    ${list.map(a => a.type === 'visitEdit' ? `<div class="ap" style="${a.seen ? 'opacity:.7' : ''}">
      <div class="row" style="justify-content:space-between"><b>${esc(a.name)}</b><span class="small muted">${esc(dayLabel(a.at.slice(0, 10)))} ${esc(new Date(a.at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }))}</span></div>
      <p style="margin:4px 0"><span class="pill edited">Edited after submit</span> ${esc(a.store)} visit from ${esc(dayLabel(a.visitDate))}, submitted ${esc(dayLabel((a.submittedAt || a.visitDate).slice(0, 10)))}. Changed: ${esc((a.parts || []).join(', '))}.</p>
      <div class="row" style="gap:8px">${a.seen ? `<span class="small muted">Reviewed${a.seenByName ? ' by ' + esc(a.seenByName) : ''}</span>` : isAdmin() ? `<button class="btn tiny" data-seen="${esc(a.id)}" type="button">Mark reviewed</button>` : ''}<button class="btn tiny" type="button" data-aov='${esc(JSON.stringify({ store: a.store, date: a.visitDate, email: a.forEmail, remote: /_remote$/.test(a.visitId || '') }))}'>Open visit</button></div>
    </div>` : `<div class="ap" style="${a.seen ? 'opacity:.7' : ''}">
      <div class="row" style="justify-content:space-between"><b>${esc(a.name)}</b><span class="small muted">${esc(dayLabel(a.at.slice(0, 10)))} · week of ${esc(shortDate(a.weekStart))}</span></div>
      <p style="margin:4px 0"><span class="pill check">${esc(a.reason)}</span>${a.note ? ` ${esc(a.note)}` : ''}</p>
      ${(a.changes || []).length ? `<ul class="blist small" style="margin:4px 0">${changeLines(a).map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : ''}
      ${a.seen ? `<span class="small muted">Reviewed${a.seenByName ? ' by ' + esc(a.seenByName) : ''}</span>` : isAdmin() ? `<button class="btn tiny" data-seen="${esc(a.id)}" type="button">Mark reviewed</button>` : ''}
    </div>`).join('')}`;
  const v = $('#view'); v.insertBefore(el, v.firstChild);
  el.querySelectorAll('[data-aov]').forEach(b => b.onclick = () => openVisit(JSON.parse(b.dataset.aov)));
  el.querySelectorAll('[data-seen]').forEach(b => b.onclick = async () => {
    const a = S.alerts.find(x => x.id === b.dataset.seen); if (!a) return;
    Object.assign(a, { seen: true, seenBy: S.user.email, seenByName: S.user.name || S.user.email, seenAt: new Date().toISOString() });
    try { await S.be.saveAlert(a); renderShell(); } catch (e) { toast(friendly(e), true); }
  });
}
// Each store's budget: today's revenue and SPG, yesterday against its budget, and the month to date.
const $k = n => n == null ? '--' : '$' + Math.round(n).toLocaleString('en-US');
const vsPct = (a, b) => a != null && b ? Math.round((a / b - 1) * 1000) / 10 : null;
const vsTag = v => v == null ? '' : `<span class="${v >= 0 ? 'good' : v <= -10 ? 'bad' : 'warn'}">${v >= 0 ? '+' : ''}${v.toFixed(1)}%</span>`;
function budgetPanel(stores, t) {
  const y = S.meta.latestDaily, hasDay = (S.daily?.periods || []).includes('day');
  const rows = stores.map(s => ({ s, td: budgetFor(s, t), yb: y ? budgetFor(s, y) : null, mb: y ? budgetToDate(s, y) : null, snap: S.daily?.stores?.[s] })).filter(r => r.td || r.mb);
  if (!rows.length) return '';
  const tot = rows.reduce((a, r) => (a.sales += r.td?.sales || 0, a.traffic += r.td?.traffic || 0, a), { sales: 0, traffic: 0 });
  return `<section class="panel budget">
    <h3 style="margin:0 0 4px">Today's budget</h3>
    <p class="small" style="margin:0 0 8px">Every store leader should know these two numbers walking in: revenue and SPG for the day. Today across your stores: <b>${$k(tot.sales)}</b> on about ${Math.round(tot.traffic)} guests.</p>
    <div class="scroller"><table class="grid"><thead><tr><th>Store</th><th class="num">Revenue today</th><th class="num">SPG today</th><th class="num">Guests</th>${hasDay ? `<th class="num">${esc(dayLabel(y))} actual</th><th class="num">vs budget</th>` : ''}<th class="num">MTD sales</th><th class="num">vs budget to date</th></tr></thead><tbody>
      ${rows.map(r => { const d = hasDay ? r.snap?.day?.k : null, m = r.snap?.mtd?.k;
        return `<tr><td class="nm">${esc(r.s)}</td><td class="num"><b>${$k(r.td?.sales)}</b></td><td class="num"><b>${r.td?.spg ? '$' + Math.round(r.td.spg) : '--'}</b></td><td class="num">${r.td ? Math.round(r.td.traffic) : '--'}</td>
        ${hasDay ? `<td class="num">${$k(d?.netSales)}${d?.spg != null ? ` <span class="small muted">SPG $${Math.round(d.spg)}</span>` : ''}</td><td class="num">${vsTag(vsPct(d?.netSales, r.yb?.sales))}</td>` : ''}
        <td class="num">${$k(m?.netSales)}</td><td class="num">${vsTag(vsPct(m?.netSales, r.mb?.sales))}</td></tr>`; }).join('')}
    </tbody></table></div>
    <p class="small muted" style="margin:6px 0 0">SPG budget is the day's sales budget divided by its traffic budget. Month to date compares the daily report through ${esc(dayLabel(y || t))} with the budget through that day.</p>
  </section>`;
}
function cartsPanel(stores) {
  const c = S.carts; if (!c?.stores) return '';
  const rows = stores.map(s => ({ s, x: c.stores[s], top: Object.values(c.people || {}).filter(p => p.store === s).sort((a, b) => b.value - a.value)[0] })).filter(r => r.x);
  if (!rows.length) return '';
  const tot = rows.reduce((a, r) => (a.n += r.x.n, a.v += r.x.value, a.d += r.x.due, a), { n: 0, v: 0, d: 0 });
  return `<section class="panel carts">
    <h3 style="margin:0 0 4px">Open carts: the money already in the building</h3>
    <p class="small" style="margin:0 0 8px">${tot.n.toLocaleString('en-US')} open carts across your stores, about <b>${$k(tot.v)}</b> estimated. <b>${tot.d}</b> are due a follow-up today (day 1, 3 or 7). Make sure every leader has a follow-up plan and is working it with their consultants.</p>
    <div class="scroller"><table class="grid"><thead><tr><th>Store</th><th class="num">Open carts</th><th class="num">Est. value</th><th class="num">Started this week</th><th class="num">Due today</th><th class="num">Over 2 weeks</th><th>Most to follow up</th></tr></thead><tbody>
      ${rows.sort((a, b) => b.x.value - a.x.value).map(r => `<tr><td class="nm">${esc(r.s)}</td><td class="num">${r.x.n}</td><td class="num"><b>${$k(r.x.value)}</b></td><td class="num">${r.x.wk} · ${$k(r.x.wkValue)}</td><td class="num ${r.x.due ? 'warn' : ''}">${r.x.due}</td><td class="num">${r.x.old}</td><td>${r.top ? `${esc(titleName(r.top.name))} <span class="small muted">(${r.top.n}, ${$k(r.top.value)})</span>` : ''}</td></tr>`).join('')}
    </tbody></table></div>
    <p class="small muted" style="margin:6px 0 0">Open carts as of ${esc(shortDate(c.date))}. Cart value is estimated from the cart lines at today's price. It isn't booked and isn't used for pay.</p>
  </section>`;
}
async function viewBriefInner() {
  const v = $('#view');
  if (seesAll() && !S.viewEmail) S.viewEmail = leaders()[0]?.email || null;
  const email = seesAll() ? S.viewEmail : S.user.email;
  const who = S.users.find(u => u.email === email) || (email === S.user.email ? S.user : null);
  const picker = seesAll() ? `<label for="bp" style="margin:0">Field leader<select id="bp">${leaders().map(l => `<option value="${esc(l.email)}" ${l.email === email ? 'selected' : ''}>${esc(whoLabel(l))}</option>`).join('')}</select></label>` : '';
  const wirePick = () => { const b = $('#bp'); if (b) b.onchange = () => { S.viewEmail = b.value; viewBrief(); }; };
  const stores = who?.stores || [];
  if (!stores.length) { v.innerHTML = `<div class="spread">${picker}</div><div class="panel"><p>No stores assigned yet.</p></div>`; wirePick(); return; }
  if (!S.daily) { v.innerHTML = `<div class="panel"><h2>Waiting on the first daily report</h2><p>Your brief builds from the daily report and RSA report. Once they're uploaded, it's here every morning.</p></div>`; return; }
  v.innerHTML = '<p class="loading">Building your brief…</p>';
  const t = today(), week = weekStartOf(t);
  try { await Promise.race([schedWeek(week), new Promise(r => setTimeout(r, 2500))]); } catch (e) {}
  const plan = await S.be.plan(email, week).catch(() => null);
  const asOf = S.meta.latestDaily;
  const people = S.rsa?.people || [];
  const hasDay = (S.daily.periods || []).includes('day');
  const per = snap => (hasDay ? snap?.day : snap?.wtd) || snap?.mtd;
  const perLabel = hasDay ? `yesterday (${dayLabel(asOf)})` : `the week through ${dayLabel(asOf)}`;
  const rows = stores.map(s => ({ s, snap: S.daily.stores[s], sc: S.scores[s] })).filter(r => r.snap);
  const money = n => (n < 0 ? '-$' : '$') + Math.abs(Math.round(n)).toLocaleString('en-US');
  const pace = paceFactor(S.rsa?.to);

  // Stores: wins and opportunities
  const storeWins = [], storeOpps = [];
  rows.forEach(({ s, snap, sc }) => {
    const p = per(snap), opp = [];
    if (p?.vsBud.netSales != null) {
      if (p.vsBud.netSales >= 0) storeWins.push({ s, t: `${money(p.k.netSales)} in sales, ${pct(p.vsBud.netSales)} to budget` });
      else if (p.vsBud.netSales <= -10) opp.push(`${money(p.k.netSales)} in sales, ${pct(p.vsBud.netSales)} to budget`);
    }
    if (p?.vsLy.spg != null && p.vsLy.spg >= 5) storeWins.push({ s, t: `SPG with cancellations ${pct(p.vsLy.spg)} vs LY` });
    if (p?.vsBud.closeRate != null && p.vsBud.closeRate >= 100) storeWins.push({ s, t: `close rate +${Math.round(p.vsBud.closeRate)} bps to budget` });
    const f = storeFocus(snap, 1)[0];
    if (f && !f.stretch) opp.push(`${f.label} at ${fmtMetric(f.key, f.value)}, goal ${fmtMetric(f.key, f.goal)}`);
    if (opp.length) storeOpps.push({ s, t: opp.join('. '), sc: sc?.score || 0, key: f && !f.stretch ? f.key : null });
  });
  storeOpps.sort((a, b) => b.sc - a.sc);

  // People: wins and opportunities
  const days = S.days || {};
  const uniq = (arr, k) => { const seen = new Set(); return arr.filter(x => { const id = k(x); if (seen.has(id)) return false; seen.add(id); return true; }); };
  const inMine = uniq(people.filter(p => stores.includes(p.store)), p => p.cid);
  const dayTop = inMine.map(p => ({ p, d: days[p.cid] })).filter(x => x.d?.hours >= 4 && x.d.sph != null && x.d.sales > 0).sort((a, b) => b.d.sales - a.d.sales).slice(0, 3);
  const rising = uniq(stores.flatMap(s => S.teams?.[s]?.rising || []), r => r.cid).filter(r => !dayTop.some(x => x.p.cid === r.cid)).slice(0, 3);
  const helpersList = [], dragList = [];
  rows.forEach(({ s, snap }) => storeFocus(snap, 2).forEach(f => {
    helpers(people, s, f.key, DEFAULT_GOALS, pace, 1).forEach(h => helpersList.push({ s, h, f }));
    draggers(people, s, f.key, DEFAULT_GOALS, pace, 1).forEach(d => dragList.push({ s, d, f }));
  }));
  const below = uniq(stores.flatMap(s => (S.teams?.[s]?.below || []).map(r => ({ s, r }))), o => o.r.cid);
  const slipping = uniq(stores.flatMap(s => (S.teams?.[s]?.slipping || []).map(r => ({ s, r }))), o => o.r.cid).filter(o => !below.some(b => b.r.cid === o.r.cid));

  // Commitments due by tomorrow that the numbers don't show as done, and support promised
  const due = S.visits.filter(x => stores.includes(x.store) && x.status !== 'draft')
    .flatMap(x => (x.actions || []).filter(hasCommitment).map(a => ({ x, a })))
    .filter(({ a }) => a.due && a.due <= addDays(t, 1))
    .map(o => ({ ...o, au: autoFollow(o.a, S.daily.stores[o.x.store], {}) }))
    .filter(o => o.au?.v !== 'yes').sort((a, b) => a.a.due.localeCompare(b.a.due)).slice(0, 8);
  const supportDue = S.visits.filter(x => x.email === email && x.status !== 'draft' && String(x.leaderCommit?.support || '').trim())
    .filter(x => !x.leaderCommit.supportBy || x.leaderCommit.supportBy <= addDays(t, 2)).slice(0, 5);

  const one = await latestOne(email);
  const oneActs = (one?.actions || []).filter(hasCommitment).map(a => ({ a, au: oneFollow(a, stores, email, one.weekStart) }));
  // Today
  const todayDay = plan?.days.find(d => d.date === t);
  const sugg = plan && S.meta.latestDaily > plan.basisDate ? pivotSuggestion({ plan, scores: S.scores, today: t, dismissed: plan.dismissed || [] }) : null;
  const oppStores = [...new Set([...storeOpps.map(o => o.s), ...below.map(o => o.s), ...slipping.map(o => o.s), ...dragList.map(o => o.s)])];
  const remoteNext = oppStores.filter(s => !dayStores(todayDay).includes(s)).sort((a, b) => (S.scores[b]?.score ?? 0) - (S.scores[a]?.score ?? 0));
  const li = arr => arr.length ? `<ul class="blist">${arr.join('')}</ul>` : '<p class="small muted">Nothing here today.</p>';
  const nm = n => esc(titleName(n));
  // Opportunities at stores you're not in today get a remote coaching button.
  const remoteToday = new Set(S.visits.filter(x => x.remote && x.email === email && x.date === t).map(x => x.store));
  const rbtn = (store, extra = {}) => dayStores(todayDay).includes(store) ? '<span class="pill set">You\'re there today</span>'
    : remoteToday.has(store) ? '<span class="pill done">Coached remotely today</span>'
    : `<button type="button" class="btn tiny" data-rc='${esc(JSON.stringify({ store, ...extra }))}'>Coach remotely</button>`;

  v.innerHTML = `
  <div class="spread">
    <div><p class="eyebrow">Daily brief · ${esc(longDate(t))}</p><h2 class="big" style="margin:0">Good morning${who.name ? ', ' + esc(who.name.split(' ')[0]) : ''}.</h2>
      <p class="small" style="margin:4px 0 0">Here's ${esc(perLabel)} across your ${stores.length} stores. ${dataLine()}</p></div>
    ${picker}
  </div>
  ${activeOffer() ? `<section class="panel offerbar"><p class="eyebrow" style="margin:0">Connect, build value, run the play · ${esc(S.offer.name)} through ${esc(shortDate(S.offer.end))}</p>
    <p style="margin:4px 0 0">6 or 12 month financing AND Protection + Premium Delivery is <b>$100 off every $1,000</b>. Either one alone is $50. Connection first, every guest.</p></section>` : ''}
  ${hasFliq(who) ? `<section class="panel" style="border-left:6px solid #3F738D"><h3 style="margin:0 0 4px">FrontLine IQ: today's focus</h3>
    <p style="margin:0">${esc(FLIQ_DAILY[dow(t)])}</p>
    <p class="small" style="margin:6px 0 0">Check in with each store leader: who used it yesterday, what it flagged, and who needs a follow-up. It's on every visit too.</p></section>` : ''}
  <section class="panel today">
    <h3>Today</h3>
    ${todayDay?.store ? `<div class="tstops">${dayStores(todayDay).map((st, j) => { const isP = j === 0, stop = isP ? todayDay : todayDay.stops[j - 1], part = todayDay.stops?.length ? (isP ? (todayDay.part || 'AM') : (stop.part || 'Stop')) : 'Full day';
        const m = j ? driveMin(dayStores(todayDay)[j - 1], st) : null;
        return `<div class="tstop"><div class="row" style="justify-content:space-between;flex-wrap:nowrap"><span><span class="part">${esc(part)}</span> <b>${esc(st)}</b>${isP && todayDay.anchor ? ' <span class="pill anchor">Anchor</span>' : ''}</span>${needChip(S.scores[st]?.score)}</div>
          ${m != null ? `<p class="small muted" style="margin:0">~${driveText(m)} drive from ${esc(dayStores(todayDay)[j - 1])}</p>` : ''}
          ${whoLine(st, t)}
          ${stop.status === 'done' ? '<span class="pill done">Visited</span>' : intentHtml(visitIntent(st, { anchor: isP && todayDay.anchor, kind: stop.kind, date: t }), true)}
          <div class="row" style="margin-top:6px"><button class="btn ${j === 0 ? 'primary' : ''}" type="button" data-bopen="${esc(st)}">${stop.status === 'done' ? 'See visit' : `Open ${part === 'Full day' ? '' : esc(part) + ' '}visit`}</button></div></div>`; }).join('')}</div>`
      : ''}
    ${todayDay?.store ? '' : `<p style="margin:0">${plan ? 'No store visit on your plan today.' : 'No plan yet this week.'}</p>`}
    ${sugg ? `<div class="warnbox" style="margin-top:10px"><b>Suggested swap:</b> add ${esc(sugg.to)}, drop ${esc(sugg.from)}. ${esc(sugg.reasonTo || '')} <button class="link" id="bweek">Review it on My week</button></div>` : ''}
    ${remoteNext.length ? `<p class="small" style="margin:10px 0 0"><b>Remote coaching today</b> (opportunities at stores you're not in): ${remoteNext.map(s => remoteToday.has(s) ? `${esc(s)} <span class="good">✓</span>` : `<button class="link" data-bremote="${esc(s)}">${esc(s)}</button>`).join(' · ')}</p>` : ''}
  </section>
  ${budgetPanel(stores, t)}
  ${cartsPanel(stores)}
  <div class="bgrid">
    <section class="panel bwin">
      <h3>Wins to celebrate</h3>
      <p class="eyebrow">Stores</p>
      ${li(storeWins.slice(0, 6).map(w => `<li><b>${esc(w.s)}:</b> ${esc(w.t)}</li>`))}
      <p class="eyebrow">People</p>
      ${li([
        ...dayTop.map(x => `<li><b>${nm(x.p.name)}</b> (${esc(x.p.store)}) wrote ${money(x.d.sales)} yesterday, ${money(x.d.sph)} an hour</li>`),
        ...rising.map(r => `<li><b>${nm(r.name)}</b> is up to ${money(r.wk.sph)} an hour this week, from ${money(r.wk.priorSph)}</li>`),
        ...helpersList.slice(0, 3).map(({ s, h, f }) => `<li><b>${nm(h.name)}</b> (${esc(s)}) leads the store in ${esc(f.label.toLowerCase())} at ${esc(fmtMetric(h.key, h.value))}</li>`)
      ].slice(0, 7))}
      <p class="small muted" style="margin:6px 0 0">Call or text these out today. People repeat what gets recognized.</p>
    </section>
    <section class="panel bopp">
      <h3>Opportunities</h3>
      <p class="eyebrow">Stores</p>
      ${li(storeOpps.slice(0, 6).map(o => `<li class="bopp-row"><span>${needChip(o.sc)} <b>${esc(o.s)}:</b> ${esc(o.t)}</span>${rbtn(o.s, o.key ? { focusKey: o.key } : {})}</li>`))}
      <p class="eyebrow">People</p>
      ${li([
        ...below.slice(0, 4).map(({ s, r }) => `<li class="bopp-row"><span><b>${nm(r.name)}</b> (${esc(s)}) is at $${Math.round(r.sph)} an hour, under the minimum</span>${rbtn(s, { addCid: r.cid, why: 'below' })}</li>`),
        ...slipping.slice(0, 3).map(({ s, r }) => `<li class="bopp-row"><span><b>${nm(r.name)}</b> (${esc(s)}) dropped to $${Math.round(r.wk.sph)} an hour this week, from $${Math.round(r.wk.priorSph)}</span>${rbtn(s, { addCid: r.cid, why: 'slipping' })}</li>`),
        ...dragList.slice(0, 4).map(({ s, d, f }) => `<li class="bopp-row"><span><b>${nm(d.name)}</b> (${esc(s)}) is pulling down ${esc(f.label.toLowerCase())}: ${esc(fmtMetric(d.key, d.value))} vs ${esc(fmtMetric(d.key, d.goal))} goal</span>${rbtn(s, { addCid: d.cid, why: 'drag', lever: f.key, focusKey: f.key })}</li>`)
      ].slice(0, 8))}
    </section>
  </div>
  <section class="panel">
    <h3>Commitments due</h3>
    ${li(due.map(({ x, a, au }) => `<li><b>${esc(x.store)}:</b> ${esc(commitmentText(a))} ${au?.v ? `<span class="rv ${au.v === 'partial' ? 'part' : 'not'}">${au.v === 'partial' ? 'Moving' : 'Not yet'}</span> <span class="small">${esc(au.text)}</span>` : '<span class="small muted">Check in with the leader.</span>'}</li>`))}
    ${supportDue.length ? `<p class="eyebrow" style="margin-top:10px">Support you promised</p>${li(supportDue.map(x => `<li><b>${esc(x.store)}:</b> ${esc(x.leaderCommit.support)}${x.leaderCommit.supportBy ? ` (by ${esc(shortDate(x.leaderCommit.supportBy))})` : ''}</li>`))}` : ''}
  </section>
  ${oneActs.length ? `<section class="panel" style="border-left:6px solid #F68C2C">
    <h3>From your 1 on 1 <span class="small muted">(week of ${esc(weekRange(one.weekStart))})</span></h3>
    <ul class="blist">${oneActs.map(({ a, au }) => `<li>${esc(commitmentText(a))} ${au?.v ? `<span class="rv ${au.v === 'yes' ? 'done' : au.v === 'partial' ? 'part' : 'not'}">${au.v === 'yes' ? 'Done' : au.v === 'partial' ? 'Moving' : 'Not yet'}</span> <span class="small">${esc(au.text)}</span>` : ''}</li>`).join('')}</ul>
    ${String(one.support || '').trim() ? `<p class="small" style="margin:0"><b>Support promised:</b> ${esc(one.support)}${one.supportBy ? ` (by ${esc(shortDate(one.supportBy))})` : ''}</p>` : ''}
  </section>` : ''}
  <section class="panel">
    <h3>Send to your team</h3>
    <p class="small">Your daily store huddle and market recap are ready to copy and send.</p>
    <div class="row"><button class="btn primary" id="bmsg">Open team messages</button></div>
  </section>`;
  wirePick();
  v.querySelectorAll('[data-bopen]').forEach(b => b.onclick = () => openVisit({ store: b.dataset.bopen, date: t, email, kind: (todayDay.store === b.dataset.bopen ? todayDay.kind : todayDay.stops?.find(y => y.store === b.dataset.bopen)?.kind) || 'first', dayIndex: plan.days.indexOf(todayDay) }));
  v.querySelectorAll('[data-bstop]').forEach(b => b.onclick = () => openVisit({ store: b.dataset.bstop, date: t, email, kind: 'first', dayIndex: plan.days.indexOf(todayDay) }));
  const g = $('#bgo'); if (g) g.onclick = () => openVisit({ store: todayDay.store, date: t, email, kind: todayDay.kind, dayIndex: plan.days.indexOf(todayDay) });
  const w = $('#bweek'); if (w) w.onclick = () => { S.tab = 'week'; renderShell(); };
  v.querySelectorAll('[data-bremote]').forEach(x => x.onclick = () => openVisit({ store: x.dataset.bremote, date: t, email, kind: 'remote', remote: true }));
  v.querySelectorAll('[data-rc]').forEach(x => x.onclick = () => { const o = JSON.parse(x.dataset.rc); openVisit({ ...o, date: t, email, kind: 'remote', remote: true }); });
  $('#bmsg').onclick = () => { S.tab = 'messages'; S.msgType = 'dailyStore'; renderShell(); };
}

// ---------------------------------------------------------------- VP 1 on 1 with a Market Leader
// Weekly, VP to Market Leader. Recaps the week that just closed (Monday to Sunday): which stores and
// people performed and which didn't, how the leader ran their week, the one lever the market needs
// pulled, and where the focus goes this week. Ends with commitments from X to Y by a date, and how.
const ONE_LEVERS = ['financePct', 'appsToTraffic', 'beddingPct', 'protectionAttach', 'deliveryPct'];   // core behaviors under the two levers
const O_OPEN = new Set(['glance', 'wins', 'opps', 'ran', 'prev', 'lever', 'focus', 'coach', 'acts', 'notes']);
const oneId = (email, ws) => `${email}_${ws}`;
const weekRange = ws => `${shortDate(ws)} to ${shortDate(addDays(ws, 6))}`;
const moneyK = n => (n < 0 ? '-$' : '$') + Math.abs(Math.round(n)).toLocaleString('en-US');
const firstOf = n => titleName(String(n || '').trim().split(/\s+/)[0] || '');
const plainOf = key => PLAIN[key] || (STORE_METRICS.find(m => m.key === key)?.label || key).toLowerCase();
const canHoldOne = () => isAdmin();

// Market totals for a period: sales and budget add up; rates are weighted by each store's sales.
function rollup(stores, daily, per) {
  const rows = stores.map(s => daily?.stores?.[s]?.[per]).filter(p => p?.k?.netSales != null);
  if (!rows.length) return null;
  const w = p => Math.max(p.k.netSales || 0, 1);
  const wavg = get => { let ww = 0, t = 0; rows.forEach(p => { const v = get(p); if (v != null && isFinite(v)) { ww += w(p); t += v * w(p); } }); return ww ? t / ww : null; };
  const sales = rows.reduce((t, p) => t + (p.k.netSales || 0), 0);
  const bud = rows.reduce((t, p) => t + (p.budget?.netSales ?? p.k.netSales ?? 0), 0);
  const k = { netSales: sales, traffic: rows.reduce((t, p) => t + (p.k.traffic || 0), 0) }, budget = { netSales: bud };
  for (const m of STORE_METRICS) if (!['netSales', 'traffic'].includes(m.key)) {
    k[m.key] = wavg(p => p.k[m.key]);
    const b = wavg(p => p.budget?.[m.key]); if (b != null) budget[m.key] = b;
  }
  return { k, budget, n: rows.length, vsBud: { netSales: bud ? (sales / bud - 1) * 100 : null, closeRate: wavg(p => p.vsBud?.closeRate) }, vsLy: { spg: wavg(p => p.vsLy?.spg), netSales: wavg(p => p.vsLy?.netSales) } };
}
// The lever: the store number where the market leaves the most on the table, weighted by store volume.
function marketLevers(stores, daily, per) {
  return ONE_LEVERS.map(key => {
    const m = STORE_METRICS.find(x => x.key === key); if (!m) return null;
    let score = 0; const hits = [];
    stores.forEach(s => {
      const p = daily?.stores?.[s]?.[per]; const v = p?.k?.[key]; const goal = p?.budget?.[key] ?? STORE_GOALS[key];
      if (v == null || goal == null) return;
      const ratio = m.lower ? (v ? goal / v : 2) : v / goal;
      if (ratio < 1) { const short = 1 - ratio; score += short * Math.max(p.k.netSales || 0, 1); hits.push({ store: s, value: v, goal, short }); }
    });
    return { key, label: m.label, lower: !!m.lower, score, hits: hits.sort((a, b) => b.short - a.short), pillar: COACHING[key].pillar, of: stores.filter(s => daily?.stores?.[s]?.[per]?.k?.[key] != null).length };
  }).filter(x => x && x.hits.length).sort((a, b) => b.score - a.score);
}
// Where a 1 on 1 commitment stands today, from the latest numbers.
function oneFollow(a, stores, email, ws) {
  if (!a) return null;
  if (/full-day visits/i.test(a.what || '') && email && ws) {
    const w0 = addDays(ws, 7), w1 = addDays(ws, 13), to = numOf(a.to);
    const n = S.visits.filter(x => x.email === email && !x.remote && x.status !== 'draft' && x.date >= w0 && x.date <= w1).length;
    return { v: to != null && n >= to ? 'yes' : n ? 'partial' : 'no', text: `${n} full-day visit${n === 1 ? '' : 's'} logged that week.` };
  }
  if (!S.daily) return null;
  if (a.cid && a.rkey) {
    const p = (S.rsa?.people || []).find(x => x.cid === a.cid); const now = p?.k?.[a.rkey];
    const from = numOf(a.from), to = numOf(a.to); if (now == null || from == null || to == null) return null;
    const lower = !!METRICS.find(x => x.key === a.rkey)?.lower;
    const v = (lower ? now <= to : now >= to) ? 'yes' : (lower ? now < from : now > from) ? 'partial' : 'no';
    return { v, text: `${titleName(p.name)} is at ${fmtMetric(a.rkey, now)} month to date (was ${a.from}, goal ${a.to}).` };
  }
  if (a.store && S.daily.stores[a.store]) return autoFollow(a, S.daily.stores[a.store], {});
  const mk = { wtd: rollup(stores, S.daily, 'wtd'), mtd: rollup(stores, S.daily, 'mtd') };
  return autoFollow(a, mk, {});
}

async function viewOne() {
  const v = $('#view');
  const holder = canHoldOne() || S.user.role === 'exec';
  const mms = holder ? leaders() : [S.user];
  if (holder && (!S.oneEmail || !mms.some(l => l.email === S.oneEmail))) S.oneEmail = mms.find(l => l.role === 'leader')?.email || mms[0]?.email || null;
  const email = holder ? S.oneEmail : S.user.email;
  const who = S.users.find(u => u.email === email) || (email === S.user.email ? S.user : null);
  if (!who) { v.innerHTML = `<div class="panel"><h2>1 on 1s</h2><p>No Market Leaders set up yet. Add them in Setup and put them on a market.</p></div>`; return; }
  const weeks = Array.from({ length: 8 }, (_, i) => addDays(weekStartOf(today()), -7 * (i + 1)));
  if (!S.oneWeek || !weeks.includes(S.oneWeek)) {
    S.oneWeek = weeks[0];
    // A Market Leader opens on their most recent held 1 on 1.
    if (!holder) for (const w of weeks.slice(0, 4)) { const d = await S.be.oneOnOne(oneId(email, w)).catch(() => null); if (d?.status === 'done') { S.oneWeek = w; break; } }
  }
  const ws = S.oneWeek, we = addDays(ws, 6);
  v.innerHTML = '<p class="loading">Pulling last week together…</p>';

  // The week's numbers: the last daily report inside the week (its WTD is the whole week), and the
  // RSA report at the end of the week minus the one before it.
  const dd = (S.meta.dailyDates || []).filter(d => d > ws && d <= we).sort().reverse()[0] || null;
  const daily = dd ? (dd === S.meta.latestDaily ? S.daily : await S.be.daily(dd).catch(() => null)) : null;
  const per = daily?.periods?.includes('wtd') ? 'wtd' : 'mtd';
  const rEnd = (S.meta.rsaDates || []).filter(d => d >= ws && d <= we).sort().reverse()[0];
  const rBase = (S.meta.rsaDates || []).filter(d => d < ws).sort().reverse()[0];
  const [rsaEnd, rsaBase, plan, prev, saved, thisPlan] = await Promise.all([
    rEnd ? S.be.rsaAt(rEnd).catch(() => null) : null, rBase ? S.be.rsaAt(rBase).catch(() => null) : null,
    S.be.plan(email, ws).catch(() => null), S.be.oneOnOne(oneId(email, addDays(ws, -7))).catch(() => null),
    S.be.oneOnOne(oneId(email, ws)).catch(() => null), S.be.plan(email, addDays(ws, 7)).catch(() => null)
  ]);
  const wkRsa = consultantWeeks(rsaEnd, rsaBase);
  const sameMonth = rsaBase?.to && rsaEnd?.to && rsaBase.to.slice(0, 7) === rsaEnd.to.slice(0, 7);
  const stores = (who.stores || []).filter(s => daily?.stores?.[s]);
  const allStores = who.stores || [];
  const first = firstOf(who.name) || 'them';
  const vp = S.users.find(u => u.email === saved?.heldBy) || S.user;
  const mk = rollup(stores, daily, per);
  const mtd = rollup(stores, daily, 'mtd');
  const mine = new Set(allStores);

  // Stores: ranked on sales to budget for the week.
  const storeRows = stores.map(s => {
    const p = daily.stores[s][per];
    const f = pickStoreFocus({ k: p.k, budget: p.budget }, STORE_GOALS, 1)[0];
    return { s, p, f, bud: p.vsBud?.netSales, spg: p.vsLy?.spg, cr: p.vsBud?.closeRate, need: S.scores[s]?.score };
  }).sort((a, b) => (b.bud ?? -999) - (a.bud ?? -999));
  const storeWins = storeRows.filter(r => r.bud >= 0 || r.spg >= 5).slice(0, 4);
  const storeMiss = storeRows.filter(r => r.bud != null && r.bud < -5).reverse().slice(0, 4);

  // People: who performed and who didn't, on the week.
  const seenC = new Set();
  const ppl = (rsaEnd?.people || []).filter(p => mine.has(p.store) && !seenC.has(p.cid) && seenC.add(p.cid)).map(p => ({ ...p, wk: wkRsa[p.cid] || {} }));
  const minFor = s => isOutlet(s) ? DEFAULT_GOALS.outletMinSph : DEFAULT_GOALS.minSph;
  const worked = ppl.filter(p => p.wk.hours >= 12 && p.wk.sph != null);
  const topPeople = [...worked].sort((a, b) => b.wk.sph - a.wk.sph).slice(0, 5);
  const topSet = new Set(topPeople.map(p => p.cid));
  const lowPeople = [...worked].filter(p => !topSet.has(p.cid) && (p.wk.sph < minFor(p.store) || (p.wk.priorSph && p.wk.sph < p.wk.priorSph * 0.75)))
    .sort((a, b) => a.wk.sph - b.wk.sph).slice(0, 6);
  const peopleLine = p => `${titleName(p.name)} (${p.store}): ${moneyK(p.wk.sales)} in ${Math.round(p.wk.hours)} hours, $${Math.round(p.wk.sph)} an hour${p.wk.priorSph ? `, was $${Math.round(p.wk.priorSph)}` : ''}`;

  // How the leader ran the week.
  const vw = S.visits.filter(x => x.email === email && x.date >= ws && x.date <= we && x.status !== 'draft');
  const inPerson = vw.filter(x => !x.remote), remote = vw.filter(x => x.remote);
  const planned = (plan?.days || []).reduce((t, d) => t + dayStores(d).length, 0);
  const coachedN = vw.reduce((t, x) => t + (x.consultants || []).length, 0);
  const practiceN = vw.reduce((t, x) => t + (x.consultants || []).filter(c => Object.keys(c.drill?.scored || c.lead?.scored || {}).length).length, 0);
  const commitN = vw.reduce((t, x) => t + (x.actions || []).filter(hasCommitment).length, 0);
  const seen = new Set(inPerson.map(x => x.store));
  const notSeen = allStores.filter(s => !seen.has(s) && !remote.some(r => r.store === s));
  const dueWk = S.visits.filter(x => x.email === email && x.status !== 'draft').flatMap(x => (x.actions || []).filter(hasCommitment).map(a => ({ x, a })))
    .filter(({ a }) => a.due && a.due >= ws && a.due <= we).map(o => ({ ...o, au: autoFollow(o.a, S.daily?.stores?.[o.x.store], {}) }));
  const dueNot = dueWk.filter(o => o.au?.v === 'no').length, dueDone = dueWk.filter(o => o.au?.v === 'yes').length;
  const missed = Math.max(0, planned - inPerson.length);
  const schedChanges = (S.alerts || []).filter(a => a.type !== 'visitEdit' && (a.forEmail || a.email) === email && (a.weekStart === ws || a.weekStart === addDays(ws, -1)));

  // The lever and where the week goes.
  // The lever first (close rate, average ticket or effective margin), then the input to coach under it.
  const LS = leverStatus(mk);
  const leverKey = (S.oneLeverPick || {})[oneId(email, ws)] || saved?.leverKey || suggestLever(mk);
  const topL = LS.find(l => l.key === leverKey) || null;
  const inMetrics = topL ? topL.inputs.map(i => i.metric).filter(Boolean) : [];
  const levers0 = marketLevers(stores, daily, per);
  const levers = topL ? [...levers0.filter(l => inMetrics.includes(l.key)), ...levers0.filter(l => !inMetrics.includes(l.key))] : levers0;
  const lever = levers[0] || null;
  const lever2 = levers.find(l => l !== lever && l.pillar !== lever?.pillar) || null;
  const pace = paceFactor(S.rsa?.to);
  const leverStores = lever ? lever.hits.slice(0, 2).map(h => h.store) : [];
  const leverPeople = lever ? leverStores.flatMap(s => draggers(S.rsa?.people || [], s, lever.key, DEFAULT_GOALS, pace, 2).map(d => ({ ...d, store: s }))) : [];
  const leverHelp = lever ? leverStores.flatMap(s => helpers(S.rsa?.people || [], s, lever.key, DEFAULT_GOALS, pace, 1).map(d => ({ ...d, store: s }))) : [];
  const drill = lever ? drillFor(lever.key) : null;
  const focusStores = allStores.filter(s => S.scores[s]).sort((a, b) => (S.scores[b]?.score ?? 0) - (S.scores[a]?.score ?? 0)).slice(0, 3);
  const onPlan = new Set((thisPlan?.days || []).filter(d => d.store).map(d => d.store));
  const leverSet = new Set(leverStores);
  const whereNow = [...new Set([...leverStores, ...focusStores])].slice(0, 3);

  // Wins and opportunities, in plain words.
  const wins = [
    ...storeWins.map(r => `${r.s}: ${moneyK(r.p.k.netSales)}, ${pct(r.bud ?? 0)} to budget${r.spg != null ? `, SPG with cancellations ${pct(r.spg)} vs LY` : ''}`),
    ...topPeople.slice(0, 3).map(peopleLine),
    ...(planned && missed === 0 ? [`${first} made all ${planned} planned visits`] : []),
    ...(practiceN >= 3 ? [`${practiceN} stand-up practices run on visits`] : []),
    ...(dueDone ? [`${dueDone} commitment${dueDone > 1 ? 's' : ''} from visits hit the goal`] : [])
  ];
  const opps = [
    ...storeMiss.map(r => `${r.s}: ${moneyK(r.p.k.netSales)}, ${pct(r.bud)} to budget${r.f && !r.f.stretch ? `. ${r.f.label} at ${fmtMetric(r.f.key, r.f.value)}, goal ${fmtMetric(r.f.key, r.f.goal)}` : ''}`),
    ...lowPeople.slice(0, 4).map(p => `${peopleLine(p)}${p.wk.sph < minFor(p.store) ? ' (under the minimum)' : ' (slipping)'}`),
    ...(missed ? [`${missed} planned visit${missed > 1 ? 's' : ''} didn't happen`] : []),
    ...(notSeen.length ? [`Not seen last week: ${notSeen.join(', ')}`] : []),
    ...(schedChanges.length ? [`Schedule changed ${schedChanges.length} time${schedChanges.length > 1 ? 's' : ''} (${[...new Set(schedChanges.map(a => a.reason.toLowerCase()))].join(', ')})`] : []),
    ...(dueNot ? [`${dueNot} of ${dueWk.length} commitments due last week ${dueNot > 1 ? "aren't" : "isn't"} moving`] : [])
  ];

  // The coaching conversation, written the way you'd say it. Editable.
  const coachText = () => {
    const L = [];
    const bestS = storeWins[0], bestP = topPeople[0];
    L.push(`Open with a win: ${bestS ? `${bestS.s} finished ${pct(bestS.bud ?? 0)} to budget` : bestP ? `${titleName(bestP.name)} did $${Math.round(bestP.wk.sph)} an hour` : `${first} got into the stores`}${bestS && bestP ? `, and ${titleName(bestP.name)} did $${Math.round(bestP.wk.sph)} an hour` : ''}. Ask what drove it and how the rest of the market copies it.`);
    if (mk) L.push(`The number: ${moneyK(mk.k.netSales)} for the week, ${pct(mk.vsBud.netSales ?? 0)} to budget${mk.vsLy.spg != null ? `, SPG with cancellations ${pct(mk.vsLy.spg)} vs LY` : ''}.${mtd?.vsBud?.netSales != null ? ` Month to date ${pct(mtd.vsBud.netSales)}.` : ''}`);
    if (topL) L.push(`The lever this week is ${topL.label.toLowerCase()}: market at ${fmtMetric(topL.key, topL.value)}${topL.goal != null ? ` against ${fmtMetric(topL.key, topL.goal)}` : ''}. ${topL.why} We move it through its inputs: ${topL.inputs.map(i => i.label.split(':')[0].toLowerCase()).join(', ')}.`);
    if (lever) L.push(`Start with ${plainOf(lever.key)}. ${lever.hits.length} of ${lever.of} stores are under goal, worst at ${lever.hits.slice(0, 2).map(h => `${h.store} (${fmtMetric(lever.key, h.value)})`).join(' and ')}. Ask: "What are you seeing on the floor that explains ${plainOf(lever.key)} at ${leverStores[0]}?"`);
    if (storeMiss[0] && storeMiss[0].s !== leverStores[0]) L.push(`${storeMiss[0].s} was ${pct(storeMiss[0].bud)} to budget. Ask: "Walk me through that store last week. Who did you coach and on what?"`);
    if (planned) L.push(`${first} made ${inPerson.length} of ${planned} planned visits${remote.length ? ` and ${remote.length} remote coaching call${remote.length > 1 ? 's' : ''}` : ''}, coached ${coachedN} ${coachedN === 1 ? 'person' : 'people'} and ran ${practiceN} stand-up practice${practiceN === 1 ? '' : 's'}.${missed ? ` Ask: "What got in the way of the visit${missed > 1 ? 's' : ''} you missed? How do we protect those days?"` : ''}`);
    if (dueNot) L.push(`${dueNot} commitment${dueNot > 1 ? 's' : ''} from last week's visits ${dueNot > 1 ? "aren't" : "isn't"} moving. Ask: "What's the plan to get ${dueNot > 1 ? 'them' : 'it'} there, and who owns it?"`);
    if (leverPeople.length) L.push(`People: see ${leverPeople.slice(0, 3).map(d => `${titleName(d.name)} (${d.store})`).join(', ')} first on ${plainOf(lever.key)}.${leverHelp[0] ? ` Use ${titleName(leverHelp[0].name)} as the model.` : ''}`);
    if (whereNow.length) L.push(`This week: ${whereNow.join(', ')}.${whereNow.filter(s => !onPlan.has(s)).length && thisPlan ? ` ${whereNow.filter(s => !onPlan.has(s)).join(' and ')} ${whereNow.filter(s => !onPlan.has(s)).length > 1 ? "aren't" : "isn't"} on the plan yet. Swap or add a remote call.` : ''}`);
    L.push(`Close with ${first}'s commitments in their words: from X to Y by a date, and how they'll get there.`);
    return L.join('\n\n');
  };

  // Suggested commitments: the lever, the worst store, the people, and visits if they slipped.
  const due = addDays(ws, 13);
  const suggest = () => {
    const out = [];
    if (lever && mk?.k?.[lever.key] != null) out.push({ key: lever.key, what: `${lever.label} across the market`, from: fmtMetric(lever.key, mk.k[lever.key]), to: fmtMetric(lever.key, mk.budget?.[lever.key] ?? STORE_GOALS[lever.key] ?? lever.hits[0].goal),
      how: `Run the "${drill.title}" stand-up on every visit this week, starting with ${leverStores.join(' and ')}. ${COACHING[lever.key].doThis} Check it at every huddle call.`, owner: who.name || first, due, suggested: true });
    const ws1 = storeMiss.find(r => r.f && !r.f.stretch && !leverSet.has(r.s)) || storeMiss.find(r => r.f && !r.f.stretch);
    if (ws1) out.push({ store: ws1.s, key: ws1.f.key, what: `${ws1.s}: ${ws1.f.label}`, from: fmtMetric(ws1.f.key, ws1.f.value), to: fmtMetric(ws1.f.key, ws1.f.target ?? ws1.f.goal),
      how: `Full-day visit this week. ${COACHING[ws1.f.key].doThis} Coach the leader to check it every shift.`, owner: who.name || first, due, suggested: true });
    const d0 = leverPeople[0];
    if (d0) out.push({ cid: d0.cid, rkey: d0.key, what: `${titleName(d0.name)} (${d0.store}): ${plainOf(d0.key)}`, from: fmtMetric(d0.key, d0.value), to: fmtMetric(d0.key, d0.goal),
      how: `Coach ${firstOf(d0.name)} in person and practice "${drillFor(lever.key).title}" standing up. Remote check-in midweek through the store leader.`, owner: who.name || first, due, suggested: true });
    if (missed || (planned && inPerson.length < planned)) out.push({ what: 'Full-day visits', from: `${inPerson.length} of ${planned}`, to: `${planned} of ${planned}`, how: 'Days off set by Sunday. Visit days are protected. If a visit moves, it moves inside the week and the store gets a remote call.', owner: who.name || first, due, suggested: true });
    return out;
  };

  const doc = saved || { id: oneId(email, ws), email, name: who.name || email, weekStart: ws, status: 'draft', coaching: '', actions: [{}, {}, {}], support: '', supportBy: '', mmNotes: '', vpNotes: '' };
  if (!doc.actions) doc.actions = [{}, {}, {}];
  while (doc.actions.length < 3) doc.actions.push({});
  if (!saved && canHoldOne()) {
    doc.coaching = coachText();
    const sg = suggest(); for (let i = 0; i < 3 && sg[i]; i++) doc.actions[i] = sg[i];
  }
  if (holder) { const pv = await S.be.onePrivate(doc.id).catch(() => null); doc.vpNotes = pv?.vpNotes || ''; }
  S.O = doc;
  const canEdit = canHoldOne();
  const dis = canEdit ? '' : 'disabled';

  if (!holder && !saved) {
    v.innerHTML = `<div class="panel"><p class="eyebrow">1 on 1 · Week of ${esc(weekRange(ws))}</p><h2>Nothing yet for this week</h2><p>Your 1 on 1 shows up here once it's held.</p>
      <label for="owk" style="max-width:260px">Week<select id="owk">${weeks.map(w => `<option value="${w}" ${w === ws ? 'selected' : ''}>Week of ${esc(weekRange(w))}</option>`).join('')}</select></label></div>`;
    $('#owk').onchange = e => { S.oneWeek = e.target.value; viewOne(); };
    return;
  }

  const sec = (key, num, title, q, body, accent) => `<section class="vsec ${O_OPEN.has(key) ? 'open' : ''}" data-sec="${key}">
    <button type="button" class="vsec-hd" aria-expanded="${O_OPEN.has(key)}"><span class="vnum ${accent ? 'acc' : ''}">${num}</span><span class="vtl"><b>${title}</b><small>${q}</small></span><span class="chev" aria-hidden="true">›</span></button>
    <div class="vsec-bd">${body}</div></section>`;
  const li = arr => arr.length ? `<ul class="blist">${arr.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : '<p class="small muted">Nothing to call out.</p>';
  const tone = (v, good, bad) => v == null ? '' : v >= good ? 'green' : v <= bad ? 'red' : 'amber';
  const glance = mk ? `<div class="tiles">
      ${tileFor('Sales', moneyK(mk.k.netSales), mk.vsBud.netSales != null ? `${pct(mk.vsBud.netSales)} to budget` : '', tone(mk.vsBud.netSales, 0, -5))}
      ${tileFor('SPG w/ cancellations', mk.k.spg != null ? '$' + Math.round(mk.k.spg) : '--', mk.vsLy.spg != null ? `${pct(mk.vsLy.spg)} vs LY` : '', tone(mk.vsLy.spg, 0, -5))}
      ${tileFor('Close rate', mk.k.closeRate != null ? mk.k.closeRate.toFixed(1) + '%' : '--', mk.vsBud.closeRate != null ? `${mk.vsBud.closeRate >= 0 ? '+' : ''}${Math.round(mk.vsBud.closeRate)} bps to budget` : '', tone(mk.vsBud.closeRate, 0, -100))}
      ${tileFor('Finance', mk.k.financePct != null ? mk.k.financePct.toFixed(0) + '%' : '--', `goal ${STORE_GOALS.financePct}%`, tone(mk.k.financePct - STORE_GOALS.financePct, 0, -8))}
      ${tileFor('Protection attach', mk.k.protectionAttach != null ? mk.k.protectionAttach.toFixed(0) + '%' : '--', `goal ${STORE_GOALS.protectionAttach}%`, tone(mk.k.protectionAttach - STORE_GOALS.protectionAttach, 0, -8))}
      ${tileFor('Cancels', mk.k.cancelPct != null ? mk.k.cancelPct.toFixed(1) + '%' : '--', 'goal under 4%', mk.k.cancelPct == null ? '' : mk.k.cancelPct <= 4 ? 'green' : mk.k.cancelPct >= 6 ? 'red' : 'amber')}
    </div>
    <p class="small">${per === 'wtd' ? `Week to date through ${esc(dayLabel(dd))}` : `Month to date through ${esc(dayLabel(dd))} (the report had no week-to-date columns)`}.${mtd?.vsBud?.netSales != null ? ` Month to date: ${moneyK(mtd.k.netSales)}, ${pct(mtd.vsBud.netSales)} to budget.` : ''}</p>` : `<p>No daily report was uploaded for that week, so there are no store numbers to recap.</p>`;
  const storeTable = storeRows.length ? `<div class="scroller"><table class="grid"><thead><tr><th>Store</th><th class="num">Sales</th><th class="num">vs budget</th><th class="num">SPG vs LY</th><th class="num">Close vs bud</th><th>Biggest gap</th><th class="num">Need now</th></tr></thead><tbody>
      ${storeRows.map(r => `<tr><td class="nm">${esc(r.s)}</td><td class="num">${moneyK(r.p.k.netSales)}</td><td class="num ${r.bud >= 0 ? 'up' : r.bud <= -5 ? 'down' : ''}">${r.bud != null ? pct(r.bud) : '--'}</td><td class="num ${r.spg >= 0 ? 'up' : r.spg <= -5 ? 'down' : ''}">${r.spg != null ? pct(r.spg) : '--'}</td><td class="num">${r.cr != null ? (r.cr >= 0 ? '+' : '') + Math.round(r.cr) + ' bps' : '--'}</td><td>${r.f && !r.f.stretch ? `${esc(r.f.label)} ${esc(fmtMetric(r.f.key, r.f.value))} vs ${esc(fmtMetric(r.f.key, r.f.goal))}` : '<span class="muted">At goal</span>'}</td><td class="num">${needChip(r.need)}</td></tr>`).join('')}
    </tbody></table></div>` : '';
  const pTable = (arr, empty) => arr.length ? `<ul class="blist">${arr.map(p => `<li><b>${esc(titleName(p.name))}</b> (${esc(p.store)}): $${Math.round(p.wk.sph)} an hour on ${Math.round(p.wk.hours)} hours, ${moneyK(p.wk.sales)}${p.wk.priorSph ? ` <span class="small muted">(was $${Math.round(p.wk.priorSph)})</span>` : ''}</li>`).join('')}</ul>` : `<p class="small muted">${empty}</p>`;
  const prevActs = (prev?.actions || []).filter(hasCommitment);

  v.innerHTML = `
  <div class="spread">
    <div><p class="eyebrow">1 on 1 · Week of ${esc(weekRange(ws))}</p><h2 class="big" style="margin:0">${esc(who.name || who.email)}</h2>
      <p class="small" style="margin:4px 0 0">${esc(marketsOf(email).map(m => m.name).join(', ') || allStores.length + ' stores')} · ${doc.status === 'done' ? `<span class="pill done">Held ${esc(shortDate(doc.heldAt?.slice(0, 10) || today()))}${doc.heldByName ? ' with ' + esc(doc.heldByName) : ''}</span>` : '<span class="pill check">Not held yet</span>'}</p></div>
    <div class="row">
      ${holder ? `<label for="omm" style="margin:0">Market Leader<select id="omm">${mms.map(l => `<option value="${esc(l.email)}" ${l.email === email ? 'selected' : ''}>${esc(whoLabel(l))}</option>`).join('')}</select></label>` : ''}
      <label for="owk" style="margin:0">Week<select id="owk">${weeks.map(w => `<option value="${w}" ${w === ws ? 'selected' : ''}>Week of ${esc(weekRange(w))}</option>`).join('')}</select></label>
    </div>
  </div>
  ${sec('glance', '1', 'Last week at a glance', 'The market, Monday to Sunday.', `${glance}${storeTable}`)}
  ${sec('wins', '2', 'Wins to call out', 'Stores, people, and how the week was run. Start here.', `${li(wins.filter(w => !topPeople.slice(0, 3).some(p => w === peopleLine(p))))}
    ${pTable(topPeople, 'No consultant numbers for that week yet.').replace('<ul class="blist">', '<p class="eyebrow">Who performed (sales per hour for the week, 12+ hours)</p><ul class="blist">')}`)}
  ${sec('opps', '3', 'Opportunities', "Who didn't perform, and where the week slipped.", `${li(opps.filter(o => !lowPeople.some(p => o.startsWith(peopleLine(p)))))}
    ${pTable(lowPeople, 'Nobody under the minimum or slipping last week.').replace('<ul class="blist">', `<p class="eyebrow">Who didn't (under the minimum or down 25%+ from their month)</p><ul class="blist">`)}
    ${rsaEnd && !sameMonth ? `<p class="small muted">The month turned over during this week, so consultant numbers are month to date through ${esc(shortDate(rsaEnd.to))}.</p>` : ''}`)}
  ${sec('ran', '4', `How ${esc(first)} ran the week`, 'Visits, coaching and follow-through.', `
    <div class="tiles">
      ${tileFor('Visits', `${inPerson.length}${planned ? ' of ' + planned : ''}`, planned ? 'full-day, in person' : 'no plan saved that week', planned ? (missed ? 'red' : 'green') : '')}
      ${tileFor('Remote coaching', String(remote.length), 'calls logged', remote.length ? 'green' : '')}
      ${tileFor('People coached', String(coachedN), `${practiceN} stand-up practices`, coachedN >= 5 ? 'green' : coachedN ? 'amber' : 'red')}
      ${tileFor('Commitments', String(commitN), 'set on visits', '')}
    </div>
    ${notSeen.length ? `<div class="warnbox"><b>No visit or call:</b> ${notSeen.map(esc).join(', ')}</div>` : ''}
    ${hasFliq(who) ? (() => { const fv = vw.filter(x => numOf(x.fliq?.floor)); const u = fv.reduce((t2, x) => t2 + (numOf(x.fliq.using) || 0), 0), fl = fv.reduce((t2, x) => t2 + numOf(x.fliq.floor), 0);
      return `<p class="small" style="margin:8px 0 0"><b>FrontLine IQ:</b> ${fv.length ? `${u} of ${fl} associates using it across ${fv.length} visit${fv.length > 1 ? 's' : ''} (${Math.round(u / fl * 100)}%).` : 'not checked on any visit that week.'}</p>`; })() : ''}
    ${schedChanges.length ? `<p class="eyebrow" style="margin-top:10px">Schedule changes that week</p><ul class="blist">${schedChanges.map(a => `<li><b>${esc(a.reason)}</b>${a.note ? ': ' + esc(a.note) : ''} <span class="small muted">(${esc(changeLines(a).join('; '))})</span></li>`).join('')}</ul>` : ''}
    <p class="eyebrow" style="margin-top:10px">Commitments due last week</p>
    ${dueWk.length ? `<ul class="blist">${dueWk.map(({ x, a, au }) => `<li><b>${esc(x.store)}:</b> ${esc(commitmentText(a))} ${au?.v ? `<span class="rv ${au.v === 'yes' ? 'done' : au.v === 'partial' ? 'part' : 'not'}">${au.v === 'yes' ? 'Done' : au.v === 'partial' ? 'Moving' : 'Not yet'}</span>` : ''}</li>`).join('')}</ul>` : '<p class="small muted">None were due that week.</p>'}`)}
  ${prevActs.length ? sec('prev', '5', "Last 1 on 1's commitments", 'Where they stand today. Review them first.', `<ul class="blist">${prevActs.map(a => { const au = oneFollow(a, stores, email, prev.weekStart); return `<li>${esc(commitmentText(a))} ${au?.v ? `<span class="rv ${au.v === 'yes' ? 'done' : au.v === 'partial' ? 'part' : 'not'}">${au.v === 'yes' ? 'Done' : au.v === 'partial' ? 'Moving' : 'Not yet'}</span> <span class="small">${esc(au.text)}</span>` : '<span class="small muted">Talk it through.</span>'}</li>`; }).join('')}</ul>`) : ''}
  ${sec('lever', prevActs.length ? '6' : '5', 'The lever to pull', 'The outcome to move, then the input to coach under it.', `${LS.length ? `<div class="fcard" style="border-left:5px solid #003B4A;margin:0 0 12px">
      <div class="spread" style="margin:0"><div><p class="eyebrow">${topL && topL.key === suggestLever(mk) ? 'The lever (suggested from the numbers)' : 'The lever'}</p><h3 style="margin:0 0 4px">${topL ? esc(topL.label) : 'Every lever is at goal'}</h3>
        ${topL ? `<p style="margin:0">Market at <b>${esc(fmtMetric(topL.key, topL.value))}</b>${topL.goal != null ? `, goal ${esc(fmtMetric(topL.key, topL.goal))}` : ''}. ${esc(topL.why)}</p>` : ''}</div>
        ${canEdit ? `<label for="olever" style="margin:0">Lever<select id="olever">${LS.map(l => `<option value="${l.key}" ${topL?.key === l.key ? 'selected' : ''}>${esc(l.label)}${l.value != null ? ` (${esc(fmtMetric(l.key, l.value))})` : ''}</option>`).join('')}</select></label>` : ''}</div>
      ${topL ? `<ul class="blist small" style="margin:8px 0 0">${topL.inputs.map(i => `<li><b>${esc(i.label)}</b>: ${i.metric ? `${esc(fmtMetric(i.metric, i.value))} vs ${esc(fmtMetric(i.metric, i.goal))} goal` : `count it on visits (${esc(i.count.toLowerCase())})`}. ${esc(i.behavior)}${i.fact ? ` <i>${esc(i.fact)}</i>` : ''}</li>`).join('')}</ul>` : ''}
    </div>` : ''}${lever ? `
    <div class="fcard" style="border-left:5px solid #F68C2C">
      <p class="eyebrow">Input to coach first · ${esc(lever.pillar)}</p>
      <h3 style="margin:0 0 4px">${esc(lever.label)}</h3>
      <p style="margin:0 0 8px">${lever.hits.length} of ${lever.of} stores under goal. Market at ${esc(fmtMetric(lever.key, mk.k[lever.key]))}.</p>
      <p class="small" style="margin:0 0 4px"><b>Where:</b> ${lever.hits.slice(0, 4).map(h => `${esc(h.store)} ${esc(fmtMetric(lever.key, h.value))} vs ${esc(fmtMetric(lever.key, h.goal))}`).join(' · ')}</p>
      ${leverPeople.length ? `<p class="small" style="margin:0 0 4px"><b>Who's pulling it down:</b> ${leverPeople.map(d => `${esc(titleName(d.name))} (${esc(d.store)}, ${esc(fmtMetric(d.key, d.value))})`).join(', ')}</p>` : ''}
      ${leverHelp.length ? `<p class="small" style="margin:0 0 4px"><b>Who's carrying it:</b> ${leverHelp.map(d => `${esc(titleName(d.name))} (${esc(d.store)}, ${esc(fmtMetric(d.key, d.value))})`).join(', ')}. Have them show the team.</p>` : ''}
      <p class="small" style="margin:6px 0 0"><b>How:</b> ${esc(COACHING[lever.key].doThis)}</p>
      <p class="small" style="margin:4px 0 0"><b>Practice it standing up:</b> ${esc(drill.title)}. ${esc(drill.guest)}</p>
    </div>
    ${lever2 ? `<p class="small" style="margin-top:8px"><b>Next input:</b> ${esc(lever2.label)}, ${lever2.hits.length} of ${lever2.of} stores under goal (${lever2.hits.slice(0, 2).map(h => esc(h.store)).join(', ')}).</p>` : ''}` : '<p>Every store number is at goal for the week. Pick a stretch goal together.</p>'}`, true)}
  ${sec('focus', prevActs.length ? '7' : '6', 'Where the focus goes this week', 'Stores with the most need right now, and whether they are on the plan.', `
    ${focusStores.length ? `<ul class="blist">${[...new Set([...leverStores, ...focusStores])].map(s => `<li>${needChip(S.scores[s]?.score)} <b>${esc(s)}</b>: ${esc((S.scores[s]?.parts || []).slice(0, 2).map(p => p.text).join('. '))}${thisPlan ? (onPlan.has(s) ? ' <span class="pill done">On the plan</span>' : ' <span class="pill check">Not on the plan</span>') : ''}${leverSet.has(s) ? ' <span class="pill">Lever store</span>' : ''}</li>`).join('')}</ul>` : '<p class="small muted">No scores yet.</p>'}
    ${thisPlan ? '' : `<p class="small muted">${esc(first)} has no plan saved for this week yet.</p>`}`)}
  ${!holder ? '' : sec('coach', prevActs.length ? '8' : '7', `Coaching for ${esc(first)}`, 'Written from the numbers. Change anything.', `
    ${fieldBox('ocoach', 'Talk track', doc.coaching, 12, '', 'coaching', dis)}
    ${canEdit ? `<button type="button" class="link" id="oregen">Rewrite from the numbers</button>` : ''}`, true)}
  ${sec('acts', '✓', 'Commitments', `Up to 3, each from X to Y by a date, and how. ${esc(first)} owns them.`, `
    ${[0, 1, 2].map(i => { const a = doc.actions[i] || {}; return `<div class="ap"><div class="row" style="justify-content:space-between"><p class="eyebrow" style="margin:0">Commitment ${i + 1}</p>${a.suggested ? '<span class="pill check">Suggested</span>' : ''}</div>
      ${fieldInput(`ow${i}`, 'What', a.what, `actions.${i}.what`, dis, 'placeholder="The number or behavior" style="width:100%"')}
      <div class="two">
        <div>${fieldInput(`of${i}`, 'From (now)', a.from, `actions.${i}.from`, dis, 'placeholder="Where it is now" style="width:100%"')}</div>
        <div>${fieldInput(`ot${i}`, 'To', a.to, `actions.${i}.to`, dis, 'placeholder="Where it will be" style="width:100%"')}</div>
      </div>
      ${fieldBox(`oh${i}`, 'How (the most important part)', a.how, 3, 'What they will do, where, with who, and how often. Specific enough to check.', `actions.${i}.how`, dis)}
      <div class="row"><div style="flex:1;min-width:180px">${fieldInput(`oo${i}`, 'Owner', a.owner, `actions.${i}.owner`, dis, 'style="width:100%"')}</div>
      <label for="od${i}" style="margin:0">By<input id="od${i}" type="date" data-field="actions.${i}.due" value="${esc(a.due || '')}" ${dis}></label></div>
      <p class="small preview" id="opv${i}" ${hasCommitment(a) ? '' : 'hidden'}>${hasCommitment(a) ? esc(commitmentText(a)) : ''}</p></div>`; }).join('')}
    ${canEdit ? `<button type="button" class="btn" id="osugg">Suggest commitments</button> <span class="small muted">Fills any empty ones from the lever, the stores and the people.</span>` : ''}`, true)}
  ${sec('notes', '✓', 'Notes', `${esc(first)}'s words, and what they need from you.`, `
    ${fieldBox('omm_notes', `What ${first} said`, doc.mmNotes, 3, 'Their read on the week and what they will do differently.', 'mmNotes', dis)}
    ${fieldBox('osupport', `Support ${first} needs from ${esc(firstOf(vp.name) || 'you')}`, doc.support, 2, 'People, schedule, product, training, a call with someone. Be specific.', 'support', dis)}
    <label for="osby" style="max-width:220px">Support by<input id="osby" type="date" data-field="supportBy" value="${esc(doc.supportBy || '')}" ${dis}></label>
    ${holder ? fieldBox('ovp', 'Your notes (private: admins and the exec team only)', doc.vpNotes, 3, 'Anything else from the conversation.', 'vpNotes', dis) : ''}`, true)}
  <div class="vbar"><span class="small" id="osaved">${doc.status === 'done' ? 'Held. Changes still save.' : canEdit ? 'Draft saves as you go' : ''}</span>
    <button class="btn" id="ocopy" type="button">Copy recap for ${esc(first)}</button>
    <button class="btn" id="oprint" type="button">Print / PDF</button>
    ${canEdit ? `<button class="btn primary" id="odone" type="button">${doc.status === 'done' ? 'Update 1 on 1' : 'Mark 1 on 1 held'}</button>` : ''}</div>`;

  // ---- wiring
  const omm = $('#omm'); if (omm) omm.onchange = () => { S.oneEmail = omm.value; viewOne(); };
  const olv = $('#olever'); if (olv) olv.onchange = async () => {
    S.oneLeverPick = { ...(S.oneLeverPick || {}), [doc.id]: olv.value }; doc.leverKey = olv.value;
    if (canEdit) { const { vpNotes, ...pub } = doc; await S.be.saveOneOnOne({ ...pub, updatedAt: new Date().toISOString() }).catch(() => null); }
    viewOne();
  };
  $('#owk').onchange = e => { S.oneWeek = e.target.value; viewOne(); };
  v.querySelectorAll('.vsec-hd').forEach(h => h.onclick = () => {
    const el = h.parentElement, k = el.dataset.sec; el.classList.toggle('open'); h.setAttribute('aria-expanded', el.classList.contains('open'));
    if (el.classList.contains('open')) O_OPEN.add(k); else O_OPEN.delete(k);
  });
  const recap = () => {
    const acts = doc.actions.filter(hasCommitment);
    return [`${first}, here's our 1 on 1 for the week of ${weekRange(ws)}.`, '',
      'Wins', ...wins.slice(0, 5).map(w => `- ${w}`), '',
      'Opportunities', ...opps.slice(0, 5).map(w => `- ${w}`), '',
      ...(lever ? [`The lever: ${lever.label}. Focus on ${leverStores.join(' and ')}${leverPeople.length ? `. See ${leverPeople.slice(0, 3).map(d => titleName(d.name)).join(', ')} first` : ''}.`, ''] : []),
      ...(acts.length ? ['Your commitments', ...acts.map((a, i) => `${i + 1}. ${commitmentText(a)}`), ''] : []),
      ...(String(doc.support || '').trim() ? [`My part: ${doc.support.trim()}${doc.supportBy ? ` by ${shortDate(doc.supportBy)}` : ''}`, ''] : []),
      "Let's go.", firstOf(vp.name) || ''].join('\n');
  };
  $('#ocopy').onclick = async () => { try { await navigator.clipboard.writeText(recap()); toast('Recap copied. Paste it into a text or Teams.'); } catch (e) { toast('Could not copy. Select the text and copy it.', true); } };
  $('#oprint').onclick = () => { ['glance', 'wins', 'opps', 'ran', 'prev', 'lever', 'focus', 'coach', 'acts', 'notes'].forEach(k => O_OPEN.add(k)); v.querySelectorAll('.vsec').forEach(s => s.classList.add('open')); setTimeout(() => window.print(), 200); };
  if (!canEdit) return;
  wireMics(v);
  let tmr = null;
  const save = async (msg) => {
    clearTimeout(tmr);
    try { const { vpNotes, ...pub } = doc; await Promise.all([S.be.saveOneOnOne({ ...pub, updatedAt: new Date().toISOString() }), S.be.saveOnePrivate({ id: doc.id, email: doc.email, vpNotes: vpNotes || '' })]); $('#osaved').textContent = msg || (doc.status === 'done' ? 'Held. Changes saved.' : 'Draft saved'); }
    catch (e) { $('#osaved').textContent = 'Could not save. Check your connection.'; }
  };
  const later = () => { clearTimeout(tmr); $('#osaved').textContent = 'Saving…'; tmr = setTimeout(save, 900); };
  v.querySelectorAll('[data-field]').forEach(el => el.addEventListener('input', () => {
    setPath(doc, el.dataset.field, el.value);
    const m = /^actions\.(\d)\./.exec(el.dataset.field);
    if (m) { const a = doc.actions[+m[1]]; delete a.suggested; const pv = $('#opv' + m[1]); if (pv) { pv.hidden = !hasCommitment(a); pv.textContent = hasCommitment(a) ? commitmentText(a) : ''; } }
    later();
  }));
  $('#oregen').onclick = () => { doc.coaching = coachText(); $('#ocoach').value = doc.coaching; later(); };
  $('#osugg').onclick = () => {
    const have = new Set(doc.actions.filter(hasCommitment).map(a => String(a.what).toLowerCase()));
    let n = 0; const sg = suggest().filter(s => !have.has(s.what.toLowerCase()));
    for (let i = 0; i < 3 && sg.length; i++) if (!hasCommitment(doc.actions[i])) { doc.actions[i] = sg.shift(); n++; }
    if (!n) return toast('All 3 are filled. Clear one to get a suggestion.');
    save(); viewOne();
  };
  $('#odone').onclick = async () => {
    const acts = doc.actions.filter(hasCommitment);
    if (!acts.length) return toast('Add at least one commitment before you close it out.', true);
    const bad = acts.findIndex(a => !String(a.from || '').trim() || !String(a.to || '').trim() || !String(a.how || '').trim() || !a.due);
    if (bad >= 0) { O_OPEN.add('acts'); toast(`Commitment ${doc.actions.indexOf(acts[bad]) + 1} needs From, To, How and a By date.`, true); return; }
    doc.status = 'done'; doc.heldAt = doc.heldAt || new Date().toISOString(); doc.heldBy = S.user.email; doc.heldByName = S.user.name || S.user.email;
    await save('Held. Saved.'); toast(`1 on 1 with ${first} saved. Copy the recap to send it.`); viewOne();
  };
  if (!saved) save('Draft saved');
}
// The Market Leader's latest held 1 on 1, for their daily brief.
async function latestOne(email) {
  const ws = addDays(weekStartOf(today()), -7);
  for (const w of [ws, addDays(ws, -7)]) { const d = await S.be.oneOnOne(oneId(email, w)).catch(() => null); if (d?.status === 'done') return d; }
  return null;
}

// ---------------------------------------------------------------- leaders (Frank and exec team)
async function viewLeaders() {
  const v = $('#view');
  v.innerHTML = '<p class="loading">Loading plans…</p>';
  const week = weekStartOf(today());
  const next = addDays(week, 7);
  const [plans, nextOff] = await Promise.all([S.be.plansForWeek(week), S.be.timeOffForWeek(next)]);
  const t = today();
  const ls = leaders();
  const unassigned = STORES.filter(s => !leaderOf(s.name)).map(s => s.name);
  v.innerHTML = `
  <div class="spread">
    <div><h2 class="big" style="margin:0">Field leaders this week</h2><p class="small" style="margin:4px 0 0">Week of ${esc(dayLabel(week))} · ${dataLine()}</p></div>
  </div>
  ${unassigned.length ? `<div class="warnbox"><b>${unassigned.length} stores have no Market Leader:</b> ${unassigned.map(esc).join(', ')}.${isAdmin() ? ' Assign them in Setup.' : ''}</div>` : ''}
  <div class="focus">${ls.map(l => {
    const p = plans.find(x => x.email === l.email);
    const done = p ? p.days.filter(d => d.status === 'done').length : 0;
    const due = p ? p.days.filter(d => d.date < t && d.status !== 'done').length : 0;
    const top = (l.stores || []).slice().sort((a, b) => (S.scores[b]?.score ?? -1) - (S.scores[a]?.score ?? -1))[0];
    const sugg = p && S.meta.latestDaily > p.basisDate ? pivotSuggestion({ plan: p, scores: S.scores, today: t, dismissed: p.dismissed || [] }) : null;
    return `<div class="fcard">
      <div class="row" style="justify-content:space-between"><h3 style="margin:0">${esc(l.name || l.email)}</h3><span class="row" style="gap:6px"><span class="pill ${l.role === 'director' ? 'check' : 'set'}">${l.role === 'director' ? 'Director' : 'Market Leader'}</span>${marketsOf(l.email).map(m => `<span class="pill">${esc(m.name)}</span>`).join('')}<span class="pill">${(l.stores || []).length} stores</span></span></div>
      <p class="small" style="margin:6px 0">${p ? `<b>${done}</b> of ${p.days.length} visits logged${due ? ` · <span class="warn">${due} not logged</span>` : ''}${(p.pivots || []).length ? ` · ${p.pivots.length} changes` : ''} · ${S.visits.filter(v => v.remote && v.email === l.email && v.date >= week && v.date <= addDays(week, 6)).length} remote` : '<span class="warn">No plan yet this week</span>'}</p>
      ${p ? `<ul class="small" style="padding-left:18px;margin:4px 0">${p.days.map(d => `<li>${esc(dayLabel(d.date))}: ${esc(d.store || 'Open')} ${d.status === 'done' ? '<span class="good">✓</span>' : d.date < t ? '<span class="warn">not logged</span>' : ''}</li>`).join('')}</ul>` : ''}
      <p class="small" style="margin:6px 0">Days off this week: <b>${safeOff(p?.off || l.off, l.role).map(x => DAY_NAMES[x]).join(', ')}</b> · Next week: ${(() => { const n = nextOff.find(x => x.email === l.email); return n ? `<b>${n.off.map(x => DAY_NAMES[x]).join(', ')}</b>` : '<span class="warn">not set</span>'; })()}</p>
      ${top ? `<p class="small" style="margin:6px 0">Highest need: <b>${esc(top)}</b> ${needChip(S.scores[top]?.score)}</p>` : ''}
      ${sugg ? `<p class="small warn">Pivot waiting: add ${esc(sugg.to)}, drop ${esc(sugg.from)}</p>` : ''}
      <button class="btn tiny" data-lw="${esc(l.email)}">Open week</button>
    </div>`;
  }).join('') || '<div class="panel"><p>No Market Leaders or directors set up yet. Add them in Setup.</p></div>'}</div>`;
  v.querySelectorAll('[data-lw]').forEach(b => b.onclick = () => { S.viewEmail = b.dataset.lw; S.tab = 'week'; S.week = null; renderShell(); });
}

// ---------------------------------------------------------------- team messages
// Written for the Market Leader to copy into Teams, a group text or email. Nothing is sent from the app.
const MSG_TYPES = [
  ['dailyStore', 'Daily: store huddle', 'Every morning, to one store. Yesterday, the week, one thing for today, and the top performers this week.'],
  ['dailyMarket', 'Daily: market recap', 'Every morning, to all your store leaders. Yesterday ranked, and the top performers in the market this week.'],
  ['kickoff', 'Weekly: store kickoff', 'Monday, to one store. The month, what is working, 2 focus items, top performers and category leaders, your visit days.'],
  ['marketUpdate', 'Weekly: market update', 'Monday, to all your store leaders. The month ranked, the top 5 consultants and category leaders across the market, and your visit schedule.'],
  ['recap', 'Visit recap', 'After a visit, to that store. What is working and the commitments.']
];
async function viewMessages() {
  const v = $('#view');
  if (seesAll() && !S.viewEmail) S.viewEmail = leaders()[0]?.email || null;
  const email = seesAll() ? S.viewEmail : S.user.email;
  const who = S.users.find(u => u.email === email) || (email === S.user.email ? S.user : null);
  const stores = who?.stores || [];
  const picker = seesAll() ? `<label for="mlp" style="margin:0">Field leader<select id="mlp">${leaders().map(l => `<option value="${esc(l.email)}" ${l.email === email ? 'selected' : ''}>${esc(whoLabel(l))}</option>`).join('')}</select></label>` : '';
  if (!stores.length) { v.innerHTML = `<div class="spread">${picker}</div><div class="panel"><p>No stores assigned yet.</p></div>`; wireMlp(); return; }
  const type = S.msgType || 'dailyStore';
  if (!stores.includes(S.msgStore)) S.msgStore = stores.slice().sort((a, b) => (S.scores[b]?.score ?? 0) - (S.scores[a]?.score ?? 0))[0];
  const week = weekStartOf(today());
  const plan = await S.be.plan(email, week);
  const myVisits = S.visits.filter(x => x.email === email).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 20);
  if (!myVisits.some(x => x.id === S.msgVisit)) S.msgVisit = myVisits[0]?.id || null;
  const needsStore = ['dailyStore', 'kickoff'].includes(type);
  const asOf = S.meta.latestDaily || today();
  const snap = S.daily?.stores?.[S.msgStore];
  const visitDays = (plan?.days || []).filter(d => d.store === S.msgStore && d.date >= today()).map(d => d.date);
  const pastDays = (plan?.days || []).filter(d => d.store === S.msgStore && d.date < today() && d.status === 'done').map(d => d.date);
  const tp = { people: S.rsa?.people || [], weeks: S.weeks || {}, tops: S.msgTops !== false };
  let text = '';
  if (!S.daily && type !== 'recap') text = 'Waiting on the first daily report upload.';
  else if (type === 'dailyStore') text = dailyStore({ store: S.msgStore, snap, asOf, plan, today: today(), sender: who.name, ...tp, offer: activeOffer(), fliq: hasFliq(who) ? FLIQ_DAILY[dow(today())] : null });
  else if (type === 'dailyMarket') text = dailyMarket({ stores, daily: S.daily, asOf, plan, today: today(), sender: who.name, ...tp });
  else if (type === 'kickoff') text = kickoff({ store: S.msgStore, snap, asOf, people: S.rsa?.people || [], goals: DEFAULT_GOALS, pace: paceFactor(S.rsa?.to), visitDays, pastDays, sender: who.name, weekStart: week, ...tp });
  else if (type === 'marketUpdate') text = marketUpdate({ stores, daily: S.daily, scores: S.scores, plan, sender: who.name, asOf, ...tp });
  else if (type === 'recap') {
    const vis = myVisits.find(x => x.id === S.msgVisit);
    const next = vis ? (plan?.days || []).find(d => d.store === vis.store && d.date > vis.date && d.status !== 'done')?.date : null;
    text = vis ? visitRecap({ visit: vis, nextVisit: next, sender: who.name }) : 'No visits logged yet. Log a visit, then come back for the recap.';
  }
  const info = MSG_TYPES.find(t => t[0] === type);
  v.innerHTML = `
  <div class="spread">
    <div><h2 class="big" style="margin:0">Team messages</h2><p class="small" style="margin:4px 0 0">Written from the latest numbers. Edit anything, then copy and send in Teams, a text or email. ${dataLine()}</p></div>
    ${picker}
  </div>
  <div class="panel">
    <div class="msgtypes" role="tablist" aria-label="Message type">${MSG_TYPES.map(([k, l]) => `<button type="button" role="tab" aria-selected="${k === type}" class="${k === type ? 'on' : ''}" data-mt="${k}">${esc(l)}</button>`).join('')}</div>
    <p class="small" style="margin:10px 0">${esc(info[2])}</p>
    <div class="row" style="margin-bottom:10px">
      ${needsStore ? `<label for="mstore" style="margin:0">Store<select id="mstore">${stores.map(s => `<option ${s === S.msgStore ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></label>` : ''}
      ${type !== 'recap' ? `<label class="check" for="mtops" style="margin:0"><input type="checkbox" id="mtops" ${S.msgTops !== false ? 'checked' : ''}>Include top performers</label>` : ''}
      ${type === 'recap' && myVisits.length ? `<label for="mvisit" style="margin:0">Visit<select id="mvisit">${myVisits.map(x => `<option value="${esc(x.id)}" ${x.id === S.msgVisit ? 'selected' : ''}>${esc(dayLabel(x.date))}: ${esc(x.store)}${x.remote ? ' (remote)' : ''}</option>`).join('')}</select></label>` : ''}
    </div>
    <div class="fieldhead"><label for="msgout">Message</label>${micBtn('msgout')}</div>
    <textarea id="msgout" rows="18" spellcheck="true">${esc(text)}</textarea>
    <div class="row" style="margin-top:10px"><button class="btn primary" id="mcopy">Copy message</button><button class="btn" id="mreset">Start over</button><span class="small muted">Nothing is sent from the app.</span></div>
  </div>`;
  wireMlp(); wireMics(v);
  v.querySelectorAll('[data-mt]').forEach(b => b.onclick = () => { S.msgType = b.dataset.mt; viewMessages(); });
  const ms = $('#mstore'); if (ms) ms.onchange = () => { S.msgStore = ms.value; viewMessages(); };
  const mv = $('#mvisit'); if (mv) mv.onchange = () => { S.msgVisit = mv.value; viewMessages(); };
  $('#mreset').onclick = () => viewMessages();
  const mt = $('#mtops'); if (mt) mt.onchange = () => { S.msgTops = mt.checked; viewMessages(); };
  $('#mcopy').onclick = async () => {
    const ta = $('#msgout');
    try { await navigator.clipboard.writeText(ta.value); toast('Copied. Paste it into Teams, a text or email.'); }
    catch (e) { ta.focus(); ta.select(); try { document.execCommand('copy'); toast('Copied. Paste it into Teams, a text or email.'); } catch (x) { toast('Text is selected. Press Ctrl+C (or Cmd+C) to copy.'); } }
  };
}
function wireMlp() { const l = $('#mlp'); if (l) l.onchange = () => { S.viewEmail = l.value; viewMessages(); }; }

// ---------------------------------------------------------------- visit log
function viewVisits() {
  const v = $('#view');
  const mine = S.user.stores || [];
  let list = seesAll() ? S.visits : S.visits.filter(x => x.email === S.user.email || mine.includes(x.store));
  if (S.visitFilter && S.visitFilter !== '*') list = list.filter(x => x.email === S.visitFilter || x.store === S.visitFilter);
  list = list.slice().sort((a, b) => b.date.localeCompare(a.date) || a.store.localeCompare(b.store));
  v.innerHTML = `
  <div class="spread">
    <h2 class="big" style="margin:0">Visit log</h2>
    <label for="vf" style="margin:0">Show<select id="vf"><option value="*">All visits</option>
      ${seesAll() ? `<optgroup label="Market Leader">${leaders().map(l => `<option value="${esc(l.email)}" ${S.visitFilter === l.email ? 'selected' : ''}>${esc(l.name || l.email)}</option>`).join('')}</optgroup>` : ''}
      <optgroup label="Store">${(seesAll() ? STORES.map(s => s.name) : mine).map(s => `<option ${S.visitFilter === s ? 'selected' : ''}>${esc(s)}</option>`).join('')}</optgroup></select></label>
  </div>
  <div class="panel">${list.map(x => {
    const sm = visitSummary(x), fixes = sm.fixes;
    const vid = `${x.email}_${x.date}_${slug(x.store)}${x.remote ? '_remote' : ''}`;
    return `<details class="vis" data-vid="${esc(vid)}"><summary><b>${esc(x.store)}</b><span class="small">${esc(longDate(x.date))}</span><span class="small muted">${esc(x.name)}</span>${x.remote ? '<span class="pill">Remote</span>' : `<span class="pill" data-pcount="${esc(vid)}" hidden></span>`}${x.edits?.length ? `<span class="pill edited" title="${esc(editedText(x).join(' | '))}">Edited after submit</span>` : ''}${!x.remote && x.status === 'done' && x.photosMissing ? `<span class="pill edited">${x.photosMissing} photo${x.photosMissing > 1 ? 's' : ''} missing</span>` : ''}<span class="pill ${x.vtype === 'Follow-Up' ? 'check' : x.vtype === 'Priority' ? 'set' : ''}">${esc(x.vtype || kindToType(x.kind))}</span>${x.status === 'draft' ? '<span class="pill off">Draft</span>' : ''}${sm.score ? `<span class="pill">${sm.score.pct}%</span>` : ''}${fixes.length ? `<span class="pill off">${fixes.length} to fix</span>` : ''}</summary>
      <div class="vbody">${!x.remote && x.status === 'done' && x.photosMissing ? `<b class="warn">Photos missing (${x.photosMissing}):</b> ${esc((x.photosMissingList || []).join(', ') || 'not listed')}\n<b>Reason:</b> ${esc(x.photoReason || 'none given')}\n` : ''}${x.edits?.length ? `<b class="warn">Edited after submit:</b>\n${editedText(x).map(t => '  ' + esc(t)).join('\n')}\n` : ''}${sm.win ? `<b>Leader win:</b> ${esc(sm.winName ? titleName(sm.winName) + ': ' : '')}${esc(sm.win)}\n` : ''}<b>Working:</b> ${esc(x.working || '--')}
<b>Commitments:</b>
${sm.commitments.length ? sm.commitments.map((c, i) => `${i + 1}. ${esc(c)}`).join('\n') : '--'}
${sm.leaderCommit ? `<b>Leader commits to:</b> ${esc(sm.leaderCommit)}\n` : ''}${sm.support ? `<b>Support needed:</b> ${esc(sm.support)}\n` : ''}${x.leaderCommit?.notes ? `<b>Leader notes:</b> ${esc(x.leaderCommit.notes)}\n` : ''}${Object.entries(x.segMeta || {}).filter(([, m]) => m?.how || m?.name).map(([gi, m]) => { const vals = Object.values(x.segs?.[gi] || {}); const pts = vals.reduce((t, v) => t + (v === 'yes' ? 1 : v === 'partial' ? 0.5 : 0), 0); return `<b>${esc(SEGMENTS[gi]?.name || '')}:</b> ${m.how === 'practice' ? 'practiced with' : 'watched'} ${esc(titleName(m.name || 'a team member'))}${m.how === 'observed' ? ' on a live guest' : ''}${vals.length ? `, ${pts} of ${SEGMENTS[gi].items.length}` : ''}${m.notes ? `. ${esc(m.notes)}` : ''}\n`; }).join('')}${sm.coached.length ? `<b>Consultants coached:</b>\n${sm.coached.map(c => `- ${esc(titleName(c.name))}${c.via ? `: coached through the store leader${c.score ? `, ${c.score}` : ''}` : ''}${c.drill ? `: ${esc(c.drill)}${c.ran ? ` practice, ${c.score}` : ', practice not run'}${c.rerun === 'better' ? ', second rep better' : c.rerun === 'same' ? ', second rep same' : ''}` : ''}${c.adjust ? `. Adjustment: ${esc(c.adjust)}` : ''}${c.notes ? `\n  Notes: ${esc(c.notes)}` : ''}`).join('\n')}\n` : ''}${x.teamNotes ? `<b>Team notes:</b> ${esc(x.teamNotes)}\n` : ''}${x.reflection ? `<b>Coach next visit:</b> ${esc(x.reflection)}\n` : ''}<b>Notes:</b> ${esc(x.notes || '--')}${fixes.length ? `\n<b>6 Elements to fix:</b> ${esc(fixes.join(', '))}` : ''}</div>
      ${x.remote ? '' : `<div class="logph" data-lp="${esc(vid)}"></div>`}
      <button class="btn tiny" data-ov='${esc(JSON.stringify({ store: x.store, date: x.date, email: x.email, kind: x.kind, remote: !!x.remote }))}' style="margin-top:8px">Open</button></details>`;
  }).join('') || '<p class="muted">No visits logged yet.</p>'}</div>`;
  $('#vf').onchange = e => { S.visitFilter = e.target.value; viewVisits(); };
  v.querySelectorAll('[data-ov]').forEach(b => b.onclick = () => openVisit(JSON.parse(b.dataset.ov)));
  logPhotos(v);
}
// Photos in the visit log: a count on each visit, and thumbnails inside it. Tap one to see it full size.
const LOG_PH = {};
async function logPhotos(v) {
  const fill = (vid, list) => {
    const c = v.querySelector(`[data-pcount="${CSS.escape(vid)}"]`); if (c && list.length) { c.hidden = false; c.textContent = `${list.length} photo${list.length > 1 ? 's' : ''}`; }
    const box = v.querySelector(`[data-lp="${CSS.escape(vid)}"]`);
    if (box) box.innerHTML = list.length ? `<p class="small" style="margin:8px 0 4px"><b>Photos (${list.length})</b></p><div class="lpgrid">${list.map((ph, i) => `<button type="button" class="lpimg" data-lpv="${esc(vid)}" data-lpi="${i}"><img src="${ph.data}" alt="${esc(ph.caption || ph.item || 'Visit photo')}" loading="lazy"><span>${esc(ph.item || (ph.el === 'general' ? 'General' : ELEMENTS.find(e => e.key === ph.el)?.t || ''))}</span></button>`).join('')}</div>` : '<p class="small muted" style="margin:8px 0 0">No photos on this visit.</p>';
    box?.querySelectorAll('[data-lpi]').forEach(b => b.onclick = () => photoZoom(LOG_PH[b.dataset.lpv][+b.dataset.lpi]));
  };
  const els = [...v.querySelectorAll('details.vis[data-vid]')];
  // The first 25 load right away; the rest when they are opened.
  const load = async vid => { if (!LOG_PH[vid]) { try { LOG_PH[vid] = await S.be.photos(vid); } catch (e) { LOG_PH[vid] = []; } } fill(vid, LOG_PH[vid]); };
  els.forEach((d, i) => { if (!v.querySelector(`[data-lp="${CSS.escape(d.dataset.vid)}"]`)) return; if (i < 25) load(d.dataset.vid); else d.addEventListener('toggle', () => d.open && load(d.dataset.vid), { once: true }); });
}
function photoZoom(ph) {
  if (!ph) return;
  const o = document.createElement('div'); o.className = 'pzoom';
  o.innerHTML = `<figure><img src="${ph.data}" alt="${esc(ph.caption || 'Visit photo')}"><figcaption>${esc([ph.store, ph.item || (ph.el === 'general' ? 'General' : ELEMENTS.find(e => e.key === ph.el)?.t), ph.caption].filter(Boolean).join(' · '))}</figcaption></figure><button type="button" class="btn tiny">Close</button>`;
  const close = () => { o.remove(); document.removeEventListener('keydown', esc_); };
  const esc_ = e => { if (e.key === 'Escape') close(); };
  o.onclick = close; document.addEventListener('keydown', esc_); document.body.appendChild(o);
}

// ---------------------------------------------------------------- upload (Frank)
function viewUpload() {
  const v = $('#view'), m = S.meta;
  const when = x => x ? `${esc(new Date(x.at).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }))} · ${esc(x.file || '')}` : 'Never';
  v.innerHTML = `
  <section class="panel">
    <h2>Upload the reports</h2>
    <p>Drop in any of the files. The app figures out which is which.</p>
    <ul class="small" style="color:var(--body);padding-left:18px">
      <li><b>Daily report</b> (every morning, <code>daily-report-YYYY-MM-DD.csv</code>): drives the need scores, the weekly plan (Monday to Sunday) and the mid-week pivots. Keep the WTD and MTD columns in the export. Last upload: ${when(m.lastDaily)}</li>
      <li><b>RSA report</b> (every morning with the daily report, <code>rsa_report_…_to_….csv</code>): drives the consultant conversations and the consultant side of the need score. Each day's copy is kept, so the app compares this week against the month before it and flags who is slipping. Last upload: ${when(m.lastRsa)}</li>
      <li><b>RSA report, year to date</b> (once a month is enough, run it from January 1: <code>rsa_report_YYYY-01-01_to_….csv</code>): fills the YTD column on every consultant card so leaders can see the trend. The app adds this month on top of it. Last upload: ${when(m.lastRsaYtd)}</li>
      <li><b>Daily budget</b> (once a month, the "Month YYYY Daily Budgets" workbook with a sheet per store): shows each store's revenue and SPG budget for the day on the daily brief and every visit. Months on file: ${Object.keys(S.budgets || {}).map(esc).join(', ') || 'none'}.</li>
      <li><b>Open carts</b> (daily if you can, the Storis open cart detail <code>open_carts_detail_YYYY-MM-DD.csv</code>): number of open carts and estimated value per consultant and store, with who is due a follow-up. Guest phone numbers and emails are dropped and never saved. Last upload: ${S.carts?.date ? esc(shortDate(S.carts.date)) + ' · ' + (S.carts.n || 0).toLocaleString('en-US') + ' carts' : 'Never'}.</li>
      <li><b>Store leader list</b> (when leaders change, the store-leader-logins file): names the store leader on every visit. ${(S.storeLeaders || []).length} on file.</li>
      <li><b>Sales team roster</b> (when people change, the Paylocity "Sales Team" export): tells the app which store each consultant works in. ${S.roster.length} people on file.</li>
    </ul>
    <label class="drop" id="drop" for="file"><input type="file" id="file" accept=".csv,.xlsx,.xls" multiple><span><b>Choose files</b> or drag them here</span></label>
    <div id="pending"></div>
  </section>`;
  const drop = $('#drop'), input = $('#file');
  input.onchange = () => handleFiles([...input.files]);
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); handleFiles([...e.dataTransfer.files]); };
}
async function handleFiles(files) {
  const box = $('#pending');
  box.innerHTML = '<p class="loading">Reading…</p>';
  const out = [], read = [];
  for (const f of files) {
    try { const rows = await readSpreadsheet(f); read.push({ f, rows, heads: new Set(Object.keys(rows[0] || {}).map(h => h.trim().toLowerCase().replace(/[^a-z0-9]+/g, '_'))) }); }
    catch (e) { out.push({ file: f.name, error: e.message }); }
  }
  // A roster in the same batch is used to match the RSA report.
  // The store leadership file looks like the roster (location, role, name) but lists GMs and AGMs.
  const isLeaders = (rows, h) => h.has('role') && h.has('name') && (h.has('stores') || rows.some(r => /general manager/i.test(String(r.Role ?? r.role ?? ''))));
  read.forEach(x => { x.leaders = isLeaders(x.rows, x.heads); });
  const isRoster = (h, x) => h.has('location') && h.has('name') && !x.leaders;
  const batchRoster = read.filter(x => isRoster(x.heads, x)).map(x => prepRoster(x.f.name, x.rows)).find(x => x.people);
  for (const { f, rows, heads } of read) {
    try {
      if (rows.budget) out.push(prepBudget(f.name, rows.budget));
      else if (heads.has('associate') && [...heads].some(h => /cart_value/.test(h))) out.push(prepCarts(f.name, rows));
      else if (heads.has('segment') && heads.has('metric')) out.push(prepDaily(f.name, rows));
      else if (heads.has('sales_associate')) out.push(prepRsa(f.name, rows, batchRoster?.people));
      else if (read.find(x => x.f === f)?.leaders) out.push(prepStoreLeaders(f.name, rows));
      else if (heads.has('location') && heads.has('name')) out.push(prepRoster(f.name, rows));
      else out.push({ file: f.name, error: 'Not a daily report, RSA report, roster or store leader list. Check that the first row has the column names.' });
    } catch (e) { out.push({ file: f.name, error: e.message }); }
  }
  // Roster first so the RSA report matches against the new list.
  out.sort((a, b) => (a.kind === 'roster' ? -1 : 0) - (b.kind === 'roster' ? -1 : 0));
  const good = out.filter(x => !x.error);
  box.innerHTML = (good.length > 1 ? `<div class="row" style="margin:0 0 12px"><button class="btn accent" id="puball">Publish all ${good.length}</button><span class="small muted">Or publish them one at a time below.</span></div>` : '') + out.map((x, i) => `<div class="panel" style="background:var(--soft)" id="pf${i}">
    <h3>${esc(x.file)}</h3>
    ${x.error ? `<p class="err">${esc(x.error)}</p>` : `<p class="small">${x.summary}</p>${x.extra || ''}<button class="btn primary" data-pub="${i}">Publish ${esc(x.label)}</button>`}
  </div>`).join('');
  const pa = $('#puball');
  if (pa) pa.onclick = async () => {
    pa.disabled = true; pa.textContent = 'Publishing…';
    try { for (const x of good) await x.publish(); toast(`${good.length} files published.`); await loadShared(); S.tab = 'upload'; renderShell(); }
    catch (e) { toast(friendly(e), true); pa.disabled = false; pa.textContent = `Publish all ${good.length}`; }
  };
  box.querySelectorAll('[data-pub]').forEach(b => b.onclick = async () => {
    const x = out[+b.dataset.pub];
    b.disabled = true; b.textContent = 'Publishing…';
    try { await x.publish(); toast(`${x.label[0].toUpperCase() + x.label.slice(1)} published.`); b.textContent = 'Published'; await loadShared(); S.tab = 'upload'; renderShell(); }
    catch (e) { toast(friendly(e), true); b.disabled = false; b.textContent = `Publish ${x.label}`; }
  });
}
function prepDaily(file, rows) {
  const d = parseDaily(rows);
  if (d.missing.length) return { file, error: `Missing columns: ${d.missing.join(', ')}.` };
  const n = Object.keys(d.stores).length;
  const missingStores = STORES.filter(s => !d.stores[s.name]).map(s => s.name);
  return { file, kind: 'daily', label: 'daily report',
    summary: `Daily report through <b>${esc(longDate(d.date))}</b> · ${n} stores · periods found: ${d.periods.map(p => p.toUpperCase()).join(', ')}${!d.periods.includes('wtd') ? ' <span class="warn">(no WTD columns, the plan will use month to date only)</span>' : ''}`,
    extra: (missingStores.length ? `<p class="small warn">Not in this file: ${missingStores.map(esc).join(', ')}</p>` : '') + (d.unknown.length ? `<p class="small warn">Rows skipped (store name not recognized): ${d.unknown.map(esc).join(', ')}</p>` : ''),
    publish: () => S.be.publishDaily({ date: d.date, periods: d.periods, stores: d.stores, file }) };
}
function prepRsa(file, rows, rosterOverride) {
  const r = parseRsa(rows);
  if (r.missing.length) return { file, error: `Missing columns: ${r.missing.join(', ')}.` };
  const range = rangeFromFileName(file) || { from: null, to: S.meta.latestDaily || today() };
  const roster = rosterOverride || S.roster;
  const res = resolveReportNames(r.people, roster);
  const people = res.matched.map(({ p, d }) => ({ ...p, name: d.name || p.name, store: d.store }));
  const sugg = res.unmatched.filter(p => res.suggestions[p.cid] && res.suggestions[p.cid].how !== 'possible');
  sugg.forEach(p => { const d = roster.find(x => x.cid === res.suggestions[p.cid].cid); people.push({ ...p, store: d.store }); });
  const left = res.unmatched.filter(p => !sugg.includes(p));
  // A report that starts January 1 and runs past January is the year-to-date copy. It is kept on its own
  // so it never replaces this month's numbers.
  if (range.from && /-01-01$/.test(range.from) && range.to && range.to.slice(0, 7) !== range.from.slice(0, 7))
    return { file, kind: 'rsa', label: 'RSA report, year to date',
      summary: `Year-to-date RSA report ${esc(shortDate(range.from))} to ${esc(shortDate(range.to))} · <b>${people.length}</b> consultants matched to a store. Fills the YTD column on every consultant card.${left.length ? ` · <span class="warn">${left.length} not on the roster</span>` : ''}`,
      extra: left.length ? `<p class="small">Not matched: ${left.map(p => esc(titleName(p.name))).join(', ')}</p>` : '',
      publish: () => S.be.publishRsaYtd({ from: range.from, to: range.to, file, people }) };
  // Every daily copy has to be month to date (from the 1st). A one-day or one-week report would be read as the
  // whole month and throw off the month, the week and the trends.
  if (range.from && range.to && range.from !== range.to.slice(0, 8) + '01')
    return { file, error: `This report runs ${esc(shortDate(range.from))} to ${esc(shortDate(range.to))}. Run the RSA report from the 1st of the month (${esc(shortDate(range.to.slice(0, 8) + '01'))} to ${esc(shortDate(range.to))}) so the app reads it as month to date. For year to date, start it January 1.` };
  return { file, kind: 'rsa', label: 'RSA report',
    summary: `RSA report ${range.from ? esc(shortDate(range.from)) + ' to ' : 'through '}${esc(shortDate(range.to))} · <b>${people.length}</b> consultants matched to a store${left.length ? ` · <span class="warn">${left.length} not on the roster</span>` : ''}${range.guessed ? `<br><span class="warn">The file name has one date, so this is read as month to date, ${esc(shortDate(range.from))} through ${esc(shortDate(range.to))}. If it covers something else, rename it rsa_report_YYYY-MM-DD_to_YYYY-MM-DD before uploading.</span>` : ''}${!r.people.some(p => p.k.cancelPct != null) ? '<br><span class="small muted">No cancellation or discount columns in this export, so those show blank for this day.</span>' : ''}`,
    extra: left.length ? `<p class="small">Not matched (left out of coaching until the roster has them): ${left.map(p => esc(titleName(p.name))).join(', ')}</p>` : '',
    publish: () => S.be.publishRsa({ from: range.from, to: range.to, file, people }) };
}
// Store leader list (store-leader-logins export): one leader per store. Names the leader on visits.
// Store leadership: either the Store Leadership Contacts export (Location, Role, Name, Email; one row
// per person, OPEN for empty seats) or the older logins file (email, name, role, stores).
// It never touches the sales team roster. ASMs and Sales Leads can be on both: they lead and sell.
const LEAD_RANK = r => /^general manager/i.test(r) ? 0 : /assistant general/i.test(r) ? 1 : /assistant selling|asm/i.test(r) ? 2 : /sales lead/i.test(r) ? 3 : 4;
const isGM = r => /^general manager|^gm$/i.test(String(r || '').trim());
function prepStoreLeaders(file, rows) {
  const leaders = [], open = [], bad = new Set();
  rows.forEach(r => {
    const o = {}; for (const [k, v] of Object.entries(r)) o[k.trim().toLowerCase()] = String(v ?? '').trim();
    const list = o.location ? [o.location] : String(o.stores || '').split(/[;|,]/).map(x => x.trim()).filter(Boolean);
    if (!o.name || !list.length) return;
    list.forEach(st => {
      const store = canonicalStore(st);
      if (!isKnownStore(store)) { bad.add(st); return; }
      const role = o.location ? o.role : (/leader/i.test(o.role) ? 'Store leader' : o.role);
      if (/^open$/i.test(o.name)) { open.push({ store, role }); return; }
      leaders.push({ store, name: o.name, email: (o.email || '').toLowerCase(), role });
    });
  });
  leaders.sort((a, b) => a.store.localeCompare(b.store) || LEAD_RANK(a.role) - LEAD_RANK(b.role));
  const covered = new Set(leaders.map(l => l.store));
  const missing = STORES.filter(s => !covered.has(s.name)).map(s => s.name);
  const hasRoles = leaders.some(l => isGM(l.role));
  const noGM = hasRoles ? STORES.filter(s => !leaders.some(l => l.store === s.name && isGM(l.role))).map(s => s.name) : [];
  const byRole = {}; leaders.forEach(l => { byRole[l.role] = (byRole[l.role] || 0) + 1; });
  return { file, kind: 'leaders', label: 'store leader list',
    summary: `Store leaders · <b>${leaders.length}</b> across ${covered.size} stores${hasRoles ? ` · ${Object.entries(byRole).map(([r, n]) => `${n} ${esc(r)}${n > 1 ? 's' : ''}`).join(', ')}` : ''}${open.length ? ` · ${open.length} open seats` : ''}. The sales team roster is not changed.`,
    extra: (noGM.length ? `<p class="small warn"><b>No GM (${noGM.length}):</b> ${noGM.map(esc).join(', ')}. The app weighs these when it picks anchor stores.</p>` : '') + (missing.length ? `<p class="small warn">No leader listed for: ${missing.map(esc).join(', ')}</p>` : '') + (bad.size ? `<p class="small warn">Rows skipped (store not recognized): ${[...bad].map(esc).join(', ')}</p>` : ''),
    publish: async () => { await S.be.saveStoreLeaders(leaders); S.storeLeaders = leaders; } };
}
function prepRoster(file, rows) {
  const r = parseTeamRoster(rows);
  if (r.missing.length) return { file, error: `Missing columns: ${r.missing.join(', ')}.` };
  const old = new Map(S.roster.map(p => [p.cid, p]));
  const people = r.people.map(p => ({ ...p, aliases: old.get(p.cid)?.aliases || [] }));
  return { file, kind: 'roster', label: 'roster', people,
    summary: `Roster · <b>${people.length}</b> people in ${new Set(people.map(p => p.store)).size} stores${r.open.length ? ` · ${r.open.length} open seats` : ''}`,
    extra: r.badStores.length ? `<p class="small warn">Rows skipped (store not recognized): ${r.badStores.map(esc).join(', ')}</p>` : '',
    publish: async () => { await S.be.saveRoster(people); S.roster = people; } };
}

// ---------------------------------------------------------------- setup (Frank): logins and store assignment
// ---------------------------------------------------------------- markets
// Name a market, give it its stores, a Market Leader and a director. Their store lists follow the markets.
function marketsSection(groups) {
  const ms = S.markets || [];
  const em = S.editMarket;
  const m = em ? (ms.find(x => x.id === em) || { id: '', name: '', leader: '', director: '', stores: [] }) : null;
  const byRole = r => S.users.filter(u => u.role === r).sort((a, b) => (a.name || a.email).localeCompare(b.name || b.email));
  const nameOf = e => { const u = S.users.find(x => x.email === e); return u ? u.name || u.email : ''; };
  const unassigned = STORES.filter(st => !ms.some(x => x.stores.includes(st.name))).map(st => st.name);
  return `<section class="panel">
    <div class="spread" style="margin:0 0 8px"><h2 style="margin:0">Markets</h2><button class="btn primary" id="addmkt" type="button">Add a market</button></div>
    <p class="small">Name each market, pick its stores, and name its Market Leader and director. Their store lists and weekly plans follow the market. A store belongs to one market.</p>
    ${unassigned.length ? `<div class="warnbox"><b>${unassigned.length} stores are not in a market:</b> ${unassigned.map(esc).join(', ')}</div>` : ''}
    ${m ? `<form class="panel" id="mform" style="background:var(--soft)">
      <h3>${em === '__new' ? 'Add a market' : 'Edit ' + esc(m.name)}</h3>
      <div class="focus" style="gap:0 16px">
        <div style="margin:0 0 12px">${fieldInput('mname', 'Market name', m.name, '', '', 'required placeholder="For example: Jacksonville" style="width:100%"')}</div>
        <label for="mleader">Market Leader<select id="mleader"><option value="">None yet</option>${byRole('leader').map(u => `<option value="${esc(u.email)}" ${m.leader === u.email ? 'selected' : ''}>${esc(u.name || u.email)}</option>`).join('')}</select></label>
        <label for="mdirector">Director<select id="mdirector"><option value="">None yet</option>${byRole('director').map(u => `<option value="${esc(u.email)}" ${m.director === u.email ? 'selected' : ''}>${esc(u.name || u.email)}</option>`).join('')}</select></label>
        <label for="manchor">Anchor store<select id="manchor"><option value="">Let the app decide (default)</option><option value="__none" ${m.anchorMode === 'none' ? 'selected' : ''}>Never anchor this market</option>${STORES.map(st => `<option ${m.anchor === st.name ? 'selected' : ''}>${esc(st.name)}</option>`).join('')}</select></label>
        <label for="manchordays">Anchor mornings per week<select id="manchordays">${[5, 4, 3, 2, 1].map(n => `<option value="${n}" ${(+m.anchorDays || 5) === n ? 'selected' : ''}>${n === 5 ? 'Every work day (5)' : n}</option>`).join('')}</select></label>
      </div>
      <p class="small" style="margin:0 0 8px">Anchor store: the Market Leader spends mornings there to set the tone, then goes to a second store for the afternoon. By default the app decides each week from the numbers (the store with the most at stake, weighting need by revenue, that is well above the rest and short on sales or SPG with cancellations, or has no GM). Pick a store here to lock one in. The Market Leader can still change it for a week with a reason.</p>
      <p class="small" style="margin:0 0 8px">Not on the list? Add them under Logins below with the Market Leader or Director role, then come back.</p>
      <p class="small" style="margin:4px 0"><b>Stores</b> (greyed out = in another market; checking it moves it here)</p>
      <div class="storepick">${Object.entries(groups).map(([k, arr]) => `<div class="grp">${esc(DISTRICTS[k])}</div>${arr.map(st => {
        const other = ms.find(x => x.id !== m.id && x.stores.includes(st.name));
        return `<label class="check" for="ms_${st.id}" ${other ? `title="Now in ${esc(other.name)}" style="opacity:.6"` : ''}><input type="checkbox" id="ms_${st.id}" value="${esc(st.name)}" ${m.stores.includes(st.name) ? 'checked' : ''}>${esc(st.name)}${other ? ` <small class="muted">(${esc(other.name)})</small>` : ''}</label>`;
      }).join('')}`).join('')}</div>
      <div class="row" style="margin-top:14px"><button class="btn primary" type="submit">Save market</button><button class="link" type="button" id="mcancel">Cancel</button>
        ${em !== '__new' ? `<button class="link" type="button" id="mdel" style="margin-left:auto;color:var(--red)">Delete market</button>` : ''}</div>
      <div id="mdelconfirm"></div>
    </form>` : ''}
    <div class="scroller"><table class="grid"><thead><tr><th>Market</th><th>Market Leader</th><th>Director</th><th>Stores</th><th></th></tr></thead><tbody>
      ${ms.map(x => `<tr><td class="nm">${esc(x.name)}</td><td>${esc(nameOf(x.leader)) || '<span class="warn">None</span>'}</td><td>${esc(nameOf(x.director)) || '<span class="muted">None</span>'}</td>
        <td style="white-space:normal;min-width:220px">${x.stores.map(esc).join(', ') || '<span class="muted">--</span>'}<br><span class="pill ${x.anchor ? 'anchor' : ''}">Anchor: ${x.anchor ? `${esc(x.anchor)}, ${+x.anchorDays === 5 || !x.anchorDays ? 'every work day' : x.anchorDays + ' mornings'}` : x.anchorMode === 'none' ? 'never' : 'app decides'}</span></td><td><button class="btn tiny" type="button" data-em="${esc(x.id)}">Edit</button></td></tr>`).join('') || '<tr><td colspan="5">No markets yet. Add your first one.</td></tr>'}
    </tbody></table></div>
  </section>`;
}
function wireMarkets() {
  const v = $('#view');
  wireMics(v);
  $('#addmkt').onclick = () => { S.editMarket = '__new'; viewSetup(); };
  v.querySelectorAll('[data-em]').forEach(b => b.onclick = () => { S.editMarket = b.dataset.em; viewSetup(); });
  const f = $('#mform'); if (!f) return;
  $('#mcancel').onclick = () => { S.editMarket = null; viewSetup(); };
  const del = $('#mdel');
  if (del) del.onclick = () => {
    const m = S.markets.find(x => x.id === S.editMarket);
    $('#mdelconfirm').innerHTML = `<div class="warnbox">Delete ${esc(m.name)}? Its stores become unassigned and come off its leaders' lists. Visits stay in the log. <button class="btn tiny" type="button" id="mdelyes">Delete</button></div>`;
    $('#mdelyes').onclick = () => saveMarketList(S.markets.filter(x => x.id !== m.id), 'Market deleted.');
  };
  f.onsubmit = e => {
    e.preventDefault();
    const name = $('#mname').value.trim(); if (!name) return toast('Give the market a name.', true);
    const stores = [...f.querySelectorAll('.storepick input:checked')].map(x => x.value);
    const id = S.editMarket === '__new' ? slug(name) + '-' + Date.now().toString(36) : S.editMarket;
    const av = $('#manchor').value, anchor = av && av !== '__none' ? av : '';
    if (anchor && !stores.includes(anchor)) return toast(`${anchor} isn't checked as one of this market's stores.`, true);
    const doc = { id, name, leader: $('#mleader').value, director: $('#mdirector').value, stores, anchor, anchorMode: av === '__none' ? 'none' : anchor ? 'store' : 'auto', anchorDays: anchor ? +$('#manchordays').value : null };
    const list = (S.markets || []).filter(x => x.id !== id).map(x => ({ ...x, stores: x.stores.filter(st => !stores.includes(st)) }));
    list.push(doc); list.sort((a, b) => a.name.localeCompare(b.name));
    saveMarketList(list, `${name} saved.`);
  };
}
// Saves the markets, then updates each Market Leader's and director's store list to match.
async function saveMarketList(list, msg) {
  try {
    const before = new Set((S.markets || []).flatMap(m => [m.leader, m.director]).filter(Boolean));
    await S.be.saveMarkets(list);
    S.markets = list;
    const members = new Set(list.flatMap(m => [m.leader, m.director]).filter(Boolean));
    for (const u of S.users.filter(x => members.has(x.email) || before.has(x.email))) {
      const stores = [...new Set(list.filter(m => m.leader === u.email || m.director === u.email).flatMap(m => m.stores))];
      if (stores.join('|') !== (u.stores || []).join('|')) { u.stores = stores; await S.be.saveUser(u); if (u.email === S.user.email) S.user = u; }
    }
    S.editMarket = null; toast(msg); viewSetup();
  } catch (e) { toast(friendly(e), true); }
}
function viewSetup() {
  const v = $('#view');
  if (FIELD_TEAM.some(t => !S.users.some(u => u.email === t.email))) { seedFieldTeam().then(() => { if (S.tab === 'setup') viewSetup(); }); }
  const rank = r => ({ leader: 0, director: 1, admin: 2, exec: 3 }[r] ?? 4);
  const users = S.users.slice().sort((a, b) => rank(a.role) - rank(b.role) || (a.name || a.email).localeCompare(b.name || b.email));
  const edit = S.editUser || null;
  const u = edit ? (S.users.find(x => x.email === edit) || { email: '', name: '', role: 'leader', stores: [], off: DEFAULT_OFF }) : null;
  const taken = s => S.users.find(l => l.role === 'leader' && l.email !== u?.email && (l.stores || []).includes(s));
  const groups = {};
  STORES.forEach(st => (groups[st.district] ||= []).push(st));
  const of = S.offer || OFFER_DEFAULT;
  v.innerHTML = `
  <section class="panel">
    <h2 style="margin:0 0 4px">In-store offer</h2>
    <p class="small">Shows on every visit under "Run the play", on the daily brief, and in the daily huddle message while it's running. Update it when the next event starts.</p>
    <div class="focus" style="gap:0 16px">
      <div>${fieldInput('ofname', 'Event name', of.name, '', '', 'style="width:100%"')}</div>
      <label for="ofstart">Starts<input type="date" id="ofstart" value="${esc(of.start || '')}"></label>
      <label for="ofend">Ends<input type="date" id="ofend" value="${esc(of.end || '')}"></label>
    </div>
    ${fieldBox('oflines', 'The offer, one line each', (of.lines || []).join('\n'), 4)}
    ${fieldBox('offine', 'Fine print', of.fine, 2)}
    <div class="row" style="margin-top:10px"><button class="btn primary" type="button" id="ofsave">Save offer</button><span class="small ${offerActive(of, today()) ? 'good' : 'muted'}">${offerActive(of, today()) ? 'Running now' : 'Not running today'}</span></div>
  </section>
  ${marketsSection(groups)}
  <section class="panel">
    <div class="spread" style="margin:0 0 8px"><h2 style="margin:0">Logins and store assignment</h2><button class="btn primary" id="adduser">Add a person</button></div>
    <p class="small">Market Leaders and directors see their own stores, weekly plan and visits. Each store has one Market Leader; directors can share stores with them. Executives see everything, read only. People create their own password with their @${esc(EMAIL_DOMAIN)} email; they get in once they are listed here.</p>
    ${u ? `<form class="panel" id="uform" style="background:var(--soft)">
      <h3>${edit === '__new' ? 'Add a person' : 'Edit ' + esc(u.name || u.email)}</h3>
      <div class="focus" style="gap:0 16px">
        <label for="uemail">Work email<input id="uemail" type="email" value="${esc(u.email)}" ${edit !== '__new' ? 'readonly' : ''} required></label>
        <div style="margin:0 0 12px">${fieldInput('uname', 'Name', u.name, '', '', 'required style="width:100%"')}</div>
        <label for="urole">Role<select id="urole">${ROLES.map(([k, l]) => `<option value="${k}" ${u.role === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
      </div>
      ${marketsOf(u.email).length ? `<p class="small" style="margin:4px 0"><b>Stores</b> come from their market${marketsOf(u.email).length > 1 ? 's' : ''}: ${marketsOf(u.email).map(m => `${esc(m.name)} (${m.stores.length})`).join(', ')}. Change them in Markets above.</p><div hidden>` : '<div>'}
      <p class="small" style="margin:4px 0"><b>Stores</b> (greyed out = already has a Market Leader. For a Market Leader, checking it moves it here. Directors can share stores.)</p>
      <div class="storepick">${Object.entries(groups).map(([k, arr]) => `<div class="grp">${esc(DISTRICTS[k])}</div>${arr.map(st => {
        const t = taken(st.name);
        return `<label class="check" for="st_${st.id}" ${t ? `title="Now with ${esc(t.name || t.email)}" style="opacity:.6"` : ''}><input type="checkbox" id="st_${st.id}" value="${esc(st.name)}" ${(u.stores || []).includes(st.name) ? 'checked' : ''}>${esc(st.name)}</label>`;
      }).join('')}`).join('')}</div></div>
      <p class="small" style="margin:12px 0 4px"><b>Pilots</b></p>
      <label class="check" for="upfliq"><input type="checkbox" id="upfliq" ${hasFliq(u) ? 'checked' : ''}>FrontLine IQ (AI sales coach). Adds it to their visits and daily brief.</label>
      <p class="small" style="margin:12px 0 4px"><b>Default days off</b> (pick 2; Tue, Wed or Thu works best. They can change any week from their own page.)</p>
      <div class="dayspick">${DAY_LONG.map((n, i) => `<label class="check" for="uo${i}"><input type="checkbox" id="uo${i}" value="${i}" ${safeOff(u.off).includes(i) ? 'checked' : ''}>${n}</label>`).join('')}</div>
      <div class="row" style="margin-top:14px"><button class="btn primary" type="submit">Save</button><button class="link" type="button" id="ucancel">Cancel</button>
        ${edit !== '__new' && u.email !== OWNER_EMAIL ? `<button class="link" type="button" id="udel" style="margin-left:auto;color:var(--red)">Remove access</button>` : ''}</div>
      <div id="delconfirm"></div>
    </form>` : ''}
    <div class="scroller"><table class="grid"><thead><tr><th>Name</th><th>Role</th><th>Stores</th><th>Default off</th><th></th></tr></thead><tbody>
      ${users.map(x => `<tr><td class="nm">${esc(x.name || x.email)}<br><small class="muted">${esc(x.email)}</small></td><td>${esc(roleLabel(x.role))}</td>
        <td style="white-space:normal;min-width:220px">${(x.stores || []).map(esc).join(', ') || '<span class="muted">--</span>'}</td>
        <td>${FIELD.includes(x.role) ? safeOff(x.off, x.role).map(i => DAY_NAMES[i]).join(', ') : ''}</td>
        <td><button class="btn tiny" data-eu="${esc(x.email)}">Edit</button></td></tr>`).join('')}
    </tbody></table></div>
  </section>`;
  $('#adduser').onclick = () => { S.editUser = '__new'; viewSetup(); };
  $('#ofsave').onclick = async () => {
    const o = { name: $('#ofname').value.trim() || 'In-store offer', start: $('#ofstart').value, end: $('#ofend').value, lines: $('#oflines').value.split('\n').map(x => x.trim()).filter(Boolean), fine: $('#offine').value.trim() };
    try { await S.be.saveOffer(o); S.offer = o; toast('Offer saved.'); viewSetup(); } catch (e) { toast(friendly(e), true); }
  };
  wireMarkets();
  wireMics(v);
  v.querySelectorAll('[data-eu]').forEach(b => b.onclick = () => { S.editUser = b.dataset.eu; viewSetup(); });
  if (!u) return;
  $('#ucancel').onclick = () => { S.editUser = null; viewSetup(); };
  const del = $('#udel');
  if (del) del.onclick = () => {
    $('#delconfirm').innerHTML = `<div class="warnbox">Remove ${esc(u.name || u.email)}? They lose access and their stores become unassigned. Their visits stay in the log. <button class="btn tiny" type="button" id="delyes">Remove</button></div>`;
    $('#delyes').onclick = async () => { await S.be.deleteUser(u.email); S.users = S.users.filter(x => x.email !== u.email); S.editUser = null; toast('Access removed.'); viewSetup(); };
  };
  $('#uform').onsubmit = async e => {
    e.preventDefault();
    const email = $('#uemail').value.trim().toLowerCase();
    if (!email.endsWith('@' + EMAIL_DOMAIN) && !DEMO) return toast(`Use a @${EMAIL_DOMAIN} email.`, true);
    const stores = marketsOf(email).length ? (S.users.find(x => x.email === email)?.stores || []) : [...v.querySelectorAll('#uform .storepick input:checked')].map(x => x.value);
    const off = [...v.querySelectorAll('.dayspick input:checked')].map(x => +x.value);
    const role = $('#urole').value;
    if (FIELD.includes(role) && !validOff(off)) return toast('Pick exactly 2 default days off.', true);
    const prevPilots = (S.users.find(x => x.email === email)?.pilots || []).filter(p => p !== 'frontlineiq');
    const doc = { ...(S.users.find(x => x.email === email) || {}), email, name: $('#uname').value.trim(), role, stores, off, pilots: [...prevPilots, ...($('#upfliq').checked ? ['frontlineiq'] : [])] };
    try {
      // A store has one Market Leader: take moved stores off any other Market Leader. Directors can overlap.
      for (const other of S.users.filter(x => role === 'leader' && x.role === 'leader' && x.email !== email && (x.stores || []).some(s => stores.includes(s)))) {
        other.stores = other.stores.filter(s => !stores.includes(s)); await S.be.saveUser(other);
      }
      await S.be.saveUser(doc);
      S.users = [...S.users.filter(x => x.email !== email), doc];
      if (email === S.user.email) S.user = doc;
      S.editUser = null; toast('Saved.'); viewSetup();
    } catch (x) { toast(friendly(x), true); }
  };
}

// ---------------------------------------------------------------- how it works
function viewGuide() {
  $('#view').innerHTML = `<section class="panel guide">
    <p class="eyebrow">Field Leader Guide</p>
    <h2 class="big">Where to go, and what to coach when you get there</h2>
    <p>For Market Leaders and directors. The app builds your week, points you to the stores and people that need you, and walks you through the visit. You should never need to build a spreadsheet.</p>
    <h3 style="margin-top:18px">Every morning: your daily brief</h3>
    <ul>
      <li>Open the app and start on <b>Daily brief</b>. It covers yesterday for your stores and your people: wins to celebrate, opportunities, commitments due, and where you're going today.</li>
      <li>Then open <b>Team messages</b>, copy the daily huddle for each store and the market recap, and send them to your team.</li>
    </ul>
    <h3 style="margin-top:18px">Sunday: next week is built for you</h3>
    <ul>
      <li>Pick your 2 days off for each week on <b>My week</b> (tap the › arrow for next week). Any 2 days work; Tuesday, Wednesday or Thursday works best. Your schedule builds as soon as you save.</li>
      <li>If you don't pick, the plan uses your default days off (Wednesday and Thursday unless Frank set others).</li>
      <li>Weeks run Monday to Sunday, the same as WTD in the daily report. Open <b>My week</b> on Sunday and plan next week from it. It's built from Saturday's numbers: 5 visit days around your 2 days off. Until Monday it keeps updating with each new daily report, unless you've set days by hand.</li>
      <li>Every store gets a visit. Extra days go to the stores that need you most, as a second visit late in the week. Visit 1 sets the plan, visit 2 checks it.</li>
      <li>Visits are full days in one store. More stores than visit days? The lowest-need stores are marked Call.</li>
      <li>While you're on a full-day visit, coach your other stores remotely from <b>Remote coaching</b> on My week: phone, video or Teams. Log it the same way (numbers, focus items, consultants, from-to commitments), minus the 6 Elements walk.</li>
      <li>You can change any day by hand.</li>
    </ul>
    <h3 style="margin-top:18px">Monday: your 1 on 1</h3>
    <ul>
      <li>Each week your VP holds a 1 on 1 with you on the week that just closed, Monday to Sunday. It covers wins and opportunities for your stores and people, who performed and who didn't, how your visits went, and the one lever your market needs to pull.</li>
      <li>You leave with up to 3 commitments, each from X to Y by a date, with how you'll get there. They show on your <b>Daily brief</b> all week with where each one stands, and on <b>My 1 on 1</b>.</li>
      <li>Next week's 1 on 1 starts by reviewing them.</li>
    </ul>
    <h3 style="margin-top:18px">During the week: pivot when the numbers move</h3>
    <ul>
      <li>Frank uploads the daily report and RSA report each morning. If a store got worse since your week was built and now needs you more than a store still ahead on your plan, you get a suggested swap.</li>
      <li>The swap always drops the visit with the least opportunity left this week, and the remaining days re-rank so the highest-need store is next.</li>
      <li>You decide: <b>Make the swap</b> or <b>Keep my plan</b>. Either way it is recorded.</li>
    </ul>
    <h3 style="margin-top:18px">In the store: open the visit and work top to bottom</h3>
    <ol>
      <li><b>Why you are here.</b> Why this store is on your plan, and the scorecard for the prior day, week or month. Read the traffic line first. If traffic is way down, look at the schedule before you coach effort.</li>
      <li><b>Leader win.</b> Start by celebrating the leader.</li>
      <li><b>Last visit's commitments.</b> Mark each one done, partial or not done before anything new.</li>
      <li><b>Two focus items.</b> The app picks the two furthest from goal. Tap to swap one. Keep it to two.</li>
      <li><b>Consultants coached.</b> Who to see first is already there: below minimum, slipping this week, biggest gap, and the top seller to recognize. Run the stand-up practice with them.</li>
      <li><b>The 6 Elements walk.</b> Score each item Yes, Partial or No. Element 1 includes the 3 value segments. Element 3 is the AOR walk. Add photos where they help.</li>
      <li><b>Photos.</b> Take or add photos, tag each one to an element and add a caption.</li>
      <li><b>Action plan.</b> Up to 3 commitments, each from X to Y (where it is today, where it will be by the next visit), with how, an owner and a date. They are suggested for you; change anything. The next visit starts with these.</li>
      <li><b>Your reflection</b>, then <b>Submit visit</b>. The recap message is ready in Team messages.</li>
    </ol>
    <p class="small">Everything saves as you go, even if the store wifi drops. Tap Talk on any box to say it instead of typing.</p>
    <h3 style="margin-top:18px">How the need score works</h3>
    <p>0 to 100, higher needs you more. Store points come from sales against budget (month and week), SPG with cancellations against LY, close rate against budget, cancellations over 4%, protection attach under 60%, finance under 55% of sales, and days since the last visit. Consultant points come from the daily RSA report: each consultant below the SPH minimum, and each consultant slipping this week (at least 12 hours on the floor and this week's SPH under 75% of their month before it).</p>
    <h3 style="margin-top:18px">Where to take things</h3>
    <ul><li>Numbers that look wrong (traffic counters, missing stores, a consultant in the wrong store): Frank.</li><li>Schedule problems you cannot solve inside your stores: Frank.</li><li>Talent and retention conversations: Frank and Leah.</li></ul>
    <p class="small" style="margin-top:14px">This page never emails anyone. You deliver the coaching.</p>
  </section>`;
}

boot();
