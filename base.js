// Shared scorecard logic (copied from the Consultant Scorecard so both apps read the reports the same way).
// Pure logic, no Firebase, no DOM.
//   RSA report  (rsa_report_YYYY-MM-DD_to_YYYY-MM-DD.csv): one row per consultant, month to date,
//                ratios already computed by the report (sum first, divide once).
//   Daily report (daily-report-YYYY-MM-DD.csv): long format, one row per store x metric with
//                daily / WTD / MTD / YTD values and LY / budget comparisons.

// ---------------------------------------------------------------- basics
export function normalizeHeader(h) {
  return String(h ?? '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}
export function toNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).trim(), neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  s = s.replace(/[$,%\s]/g, '');
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return neg ? -n : n;
}
const numOrNull = v => (v === '' || v === null || v === undefined ? null : toNumber(v));
function pad(n) { return String(n).padStart(2, '0'); }
export function parseDate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  if (typeof v === 'number') { const d = new Date(Math.round((v - 25569) * 86400000)); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${pad(m[1])}-${pad(m[2])}`; }
  return null;
}
export function slug(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'x'; }
// Consultant id: their name as it appears in the RSA report.
export const cidOf = name => slug(String(name).trim());
export function daysBetween(a, b) {
  const t = s => { const [y, m, d] = s.split('-').map(Number); return Date.UTC(y, m - 1, d); };
  return Math.round((t(b) - t(a)) / 86400000) + 1;
}
export const monthsBack = (month, n) => {
  const [y, m] = month.split('-').map(Number);
  const out = [];
  for (let i = 0; i <= n; i++) out.push(new Date(Date.UTC(y, m - 1 - i, 1)).toISOString().slice(0, 7));
  return out;
};

// All 41 selling stores, from STORIS locations (Sept 2026). district = STORIS district code.
export const STORES = [
  ['1001', 'Tallahassee', 'CN'], ['1002', 'Thomasville', 'CN'], ['1003', 'Albany', 'CN'], ['1004', 'Macon', 'CN'],
  ['1005', 'Warner Robins', 'CN'], ['1006', 'Dothan', 'CN'], ['1007', 'Enterprise', 'CN'], ['1008', 'Panama City', 'CN'],
  ['1009', 'Valdosta', 'CN'], ['1010', 'Opelika', 'CN'], ['1011', 'Columbus', 'CN'],
  ['1012', 'Town Center', 'JX'], ['1013', 'North', 'JX'], ['1014', 'Orange Park', 'JX'], ['1015', 'Brunswick', 'JX'],
  ['1016', 'Yulee', 'JX'], ['1017', 'St. Augustine', 'JX'], ['1018', 'Outlet Regency', 'JX'],
  ['1101', 'Mobile', 'GC'], ['1102', "D'Iberville", 'GC'], ['1103', 'Spanish Fort', 'GC'], ['1104', 'Pensacola', 'GC'],
  ['1105', 'Crestview', 'GC'], ['1106', 'Ft. Walton', 'GC'], ['1107', 'Outlet Pensacola', 'GC'],
  ['1201', 'Greensboro', 'NC'], ['1202', 'Winston Salem', 'NC'], ['1203', 'Burlington', 'NC'], ['1204', 'Danville', 'NC'],
  ['1205', 'Outlet Greensboro', 'NC'],
  ['1301', 'Baton Rouge', 'LA'], ['1302', 'Lafayette', 'LA'], ['1303', 'Gonzales', 'LA'], ['1304', 'Harahan', 'LA'],
  ['1305', 'Houma', 'LA'], ['1306', 'Lake Charles', 'LA'], ['1307', 'Opelousas', 'LA'], ['1308', 'Ponchatoula', 'LA'],
  ['1309', 'Hattiesburg', 'LA'], ['1310', 'Flowood', 'LA'], ['1311', 'Harvey', 'LA']
].map(([id, name, district]) => ({ id, name, district }));
export const DISTRICTS = { CN: 'Capital / Central', JX: 'Jacksonville', GC: 'Gulf Coast', NC: 'Carolinas / VA', LA: 'Louisiana / Mississippi' };
export const OUTLETS = ['Outlet Regency', 'Outlet Pensacola', 'Outlet Greensboro'];

const storeKey = s => String(s).toLowerCase().replace(/ashley|furniture|homestore|store/g, '').replace(/d\W?i?iberville/, 'iberville').replace(/[^a-z0-9]/g, '');
const STORE_INDEX = new Map();
STORES.forEach(st => {
  STORE_INDEX.set(st.id, st.name);
  STORE_INDEX.set(storeKey(st.name), st.name);
});
STORE_INDEX.set(storeKey('Pensacola Outlet'), 'Outlet Pensacola');
STORE_INDEX.set(storeKey('Greensboro Outlet'), 'Outlet Greensboro');
STORE_INDEX.set(storeKey('Regency Outlet'), 'Outlet Regency');
STORE_INDEX.set(storeKey('Regency'), 'Outlet Regency');
STORE_INDEX.set(storeKey('Fort Walton'), 'Ft. Walton');
STORE_INDEX.set(storeKey('Fort Walton Beach'), 'Ft. Walton');
STORE_INDEX.set(storeKey('Saint Augustine'), 'St. Augustine');
STORE_INDEX.set(storeKey('Jacksonville North'), 'North');
STORE_INDEX.set(storeKey('Jax North'), 'North');
STORE_INDEX.set(storeKey('Jax Orange Park'), 'Orange Park');
STORE_INDEX.set(storeKey('Jax Town Center'), 'Town Center');
STORE_INDEX.set(storeKey('FT Walton Beach'), 'Ft. Walton');
// Maps an export's store name (or STORIS ID) to the store list name. Unknown names come back as-is.
export function canonicalStore(raw) {
  const t = String(raw ?? '').trim();
  const id = t.match(/^(\d{4})\b/);
  if (id && STORE_INDEX.has(id[1])) return STORE_INDEX.get(id[1]);
  return STORE_INDEX.get(storeKey(t)) || t;
}
export const isKnownStore = name => STORES.some(s => s.name === name);

// Region rows in the daily report (not stores). Online and Total are skipped too.
export const REGIONS = ['Big Bend', 'Capital & Acadiana', 'Central', 'Crescent', 'East', 'Fall Line', 'Golden Isles',
  'Gulf Coast', 'Magnolia', 'St. Johns', 'The Piedmont', 'West', 'Wiregrass'];

// ---------------------------------------------------------------- metrics
// Consultant card. Order follows the Core 4. lower: true means lower is better.
export const METRICS = [
  { key: 'netSales',      label: 'Net Sales',          fmt: 'money', monthly: true, col: 'net_sales' },
  { key: 'sph',           label: 'Sales / Hour',       fmt: 'money', col: 'sph' },
  { key: 'avgTicket',     label: 'Avg Ticket w/ Del',  fmt: 'money', col: 'avg_ticket_w_del' },
  { key: 'effMargin',     label: 'Eff. Margin',        fmt: 'pct',   col: 'eff_margin' },
  { key: 'financePct',    label: 'Finance %',          fmt: 'pct',   col: 'fin_of_sales' },
  { key: 'creditApps',    label: 'Credit Apps',        fmt: 'int', monthly: true, col: 'credit_apps' },
  { key: 'beddingPct',    label: 'Bedding %',          fmt: 'pct',   col: 'bed_of_sales' },
  { key: 'beddingSph',    label: 'Bedding / Hour',     fmt: 'money', derived: 'SPH x Bedding %' },
  { key: 'protectionPct', label: 'Protection %',       fmt: 'pct',   col: 'prot_of_sales' },
  { key: 'protectionSph', label: 'Protection / Hour',  fmt: 'money', derived: 'SPH x Protection %' },
  { key: 'deliveryPct',   label: 'Delivery %',         fmt: 'pct',   col: 'del_of_sales' },
  { key: 'cancelPct',     label: 'Cancellation %',     fmt: 'pct', lower: true, col: 'cancellation' },
  { key: 'discountPct',   label: 'Discount %',         fmt: 'pct', lower: true, col: 'discount' }
];
export const CONSULTANT_METRICS = METRICS;

// Store total, straight from the daily report (MTD column). budget: 'pct' = % vs budget,
// 'bps' = basis points vs budget. Where a budget exists it is the store's goal.
export const STORE_METRICS = [
  { key: 'netSales',      label: 'Net Sales',             fmt: 'money', src: 'Net Sales (Stores)', budget: 'pct', ly: true },
  { key: 'spg',           label: 'SPG w/ Cancellations',  fmt: 'money2', src: 'Sales per Guest w. Cancellations', budget: 'pct', ly: true },
  { key: 'closeRate',     label: 'Close Rate',            fmt: 'pct',   src: 'Close Rate', budget: 'bps' },
  { key: 'traffic',       label: 'Traffic',               fmt: 'int',   src: 'Traffic', budget: 'pct', ly: true },
  { key: 'sph',           label: 'Sales / Hour',          fmt: 'money', src: 'Sales per Hour' },
  { key: 'avgTicket',     label: 'Avg Ticket w/ Del',     fmt: 'money', src: 'Avg Ticket w. Del.' },
  { key: 'effMargin',     label: 'Eff. Margin',           fmt: 'pct',   src: 'Eff. Margin' },
  { key: 'financePct',    label: 'Finance %',             fmt: 'pct',   src: 'Finance % of Sales' },
  { key: 'appsToTraffic', label: 'Apps to Traffic',       fmt: 'pct',   src: 'Finance Apps to Traffic' },
  { key: 'beddingPct',    label: 'Bedding %',             fmt: 'pct',   src: 'Bedding % of Sales' },
  { key: 'beddingSph',    label: 'Bedding / Hour',        fmt: 'money', src: 'Bedding SPH' },
  { key: 'protectionPct', label: 'Protection %',          fmt: 'pct',   src: 'Protection % of Sales' },
  { key: 'protectionSph', label: 'Protection / Hour',     fmt: 'money', src: 'Protection SPH' },
  { key: 'protectionAttach', label: 'Protection Attach',  fmt: 'pct',   src: 'Protection Attachment' },
  { key: 'deliveryPct',   label: 'Delivery %',            fmt: 'pct',   src: 'Delivery % of sales' },
  { key: 'cancelPct',     label: 'Cancellation %',        fmt: 'pct', lower: true, calc: 'Cancellations / Gross Sales' }
];

export const DEFAULT_GOALS = {
  standard: { netSales: 46648, sph: 400, avgTicket: 2200, effMargin: 55.5, financePct: 65, creditApps: 18,
    beddingPct: 20, beddingSph: 60, protectionPct: 8, protectionSph: 32, protectionAttach: 60, deliveryPct: 8,
    cancelPct: 4, discountPct: 12, appsToTraffic: 10 },
  outlet: null,
  outletStores: OUTLETS,
  minSph: 250, outletMinSph: 150
};
export const isOutlet = (goals, store) => ((goals || DEFAULT_GOALS).outletStores || OUTLETS).includes(store);
export function goalsFor(goals, store) {
  const g = goals || DEFAULT_GOALS;
  const std = { ...DEFAULT_GOALS.standard, ...(g.standard || {}) };
  if (g.outlet && isOutlet(g, store)) return { ...std, ...g.outlet };
  return std;
}
export function minSphFor(goals, store) {
  const g = { ...DEFAULT_GOALS, ...(goals || {}) };
  return isOutlet(g, store) ? g.outletMinSph : g.minSph;
}

// ---------------------------------------------------------------- RSA report (consultants)
const SKIP_NAMES = /^(rsa goal|house sales.*|zzz|conv|total.*|employee discount.*)$/i;
// Reads the date range out of names like rsa_report_2026-09-01_to_2026-09-28.csv
export function rangeFromFileName(name) {
  const m = String(name || '').match(/(\d{4}-\d{2}-\d{2})\D+(\d{4}-\d{2}-\d{2})/);
  return m ? { from: m[1], to: m[2] } : null;
}
export function parseRsa(rawRows) {
  const rows = rawRows.map(r => { const o = {}; for (const [k, v] of Object.entries(r)) o[normalizeHeader(k)] = v; return o; });
  const present = new Set(rows.flatMap(r => Object.keys(r)));
  const need = ['sales_associate', ...METRICS.filter(m => m.col).map(m => m.col)];
  const missing = need.filter(c => !present.has(c));
  if (missing.length) return { missing, people: [], skipped: [], nonSellers: [] };
  const people = [], skipped = [], nonSellers = [];
  let goalRow = null;
  for (const r of rows) {
    const name = String(r.sales_associate ?? '').trim();
    if (!name) continue;
    if (/^rsa goal$/i.test(name)) { goalRow = r; continue; }
    if (SKIP_NAMES.test(name)) { skipped.push(name); continue; }
    const k = {};
    METRICS.forEach(m => { if (m.col) k[m.key] = numOrNull(r[m.col]); });
    const hours = k.sph ? k.netSales / k.sph : 0;
    // No hours on the floor = not a selling consultant this month (leaders, returns only, etc.).
    if (!hours || hours < 0) { nonSellers.push({ name, netSales: k.netSales }); continue; }
    k.beddingSph = k.sph * (k.beddingPct ?? 0) / 100;
    k.protectionSph = k.sph * (k.protectionPct ?? 0) / 100;
    Object.keys(k).forEach(x => { if (k[x] !== null) k[x] = Math.round(k[x] * 100) / 100; });
    people.push({ name, cid: cidOf(name), k, hours: Math.round(hours * 100) / 100 });
  }
  return { missing: [], people, skipped, nonSellers, goalRow };
}

// Adds rank = { store: {metric: [place, of]}, company: {...} } to each person that has a store.
export function addRanks(people) {
  const list = people.filter(p => p.store);
  const place = (group, m) => {
    const vals = group.map(a => a.k[m.key]).filter(v => v !== null && v !== undefined);
    return a => {
      const v = a.k[m.key];
      if (v === null || v === undefined) return null;
      return [1 + vals.filter(x => (m.lower ? x < v : x > v)).length, vals.length];
    };
  };
  const byStore = {};
  list.forEach(a => (byStore[a.store] ||= []).push(a));
  list.forEach(a => a.rank = { store: {}, company: {} });
  for (const m of METRICS) {
    const co = place(list, m);
    list.forEach(a => { a.rank.company[m.key] = co(a); });
    for (const group of Object.values(byStore)) { const st = place(group, m); group.forEach(a => { a.rank.store[m.key] = st(a); }); }
  }
  return people;
}

// ---------------------------------------------------------------- daily report (stores)
export function parseDailyReport(rawRows) {
  const rows = rawRows.map(r => { const o = {}; for (const [k, v] of Object.entries(r)) o[normalizeHeader(k)] = v; return o; });
  const present = new Set(rows.flatMap(r => Object.keys(r)));
  const missing = ['report_date', 'segment', 'metric', 'mtd_ty'].filter(c => !present.has(c));
  if (missing.length) return { missing, stores: [] };
  const bySeg = {}, ignored = new Set(), unknown = new Set();
  let date = null;
  for (const r of rows) {
    const seg = String(r.segment ?? '').trim();
    const d = parseDate(r.report_date); if (d && (!date || d > date)) date = d;
    if (!seg || /^(total|online)$/i.test(seg) || REGIONS.includes(seg)) { ignored.add(seg); continue; }
    const store = canonicalStore(seg);
    if (!isKnownStore(store)) { unknown.add(seg); continue; }
    (bySeg[store] ||= {})[String(r.metric).trim()] = { mtd: numOrNull(r.mtd_ty), ly: numOrNull(r.mtd_ly), bud: numOrNull(r.mtd_budget) };
  }
  const stores = Object.entries(bySeg).map(([store, m]) => {
    const k = {}, budget = {}, vsBudget = {}, vsLy = {};
    for (const sm of STORE_METRICS) {
      if (!sm.src) continue;
      const x = m[sm.src]; if (!x) { k[sm.key] = null; continue; }
      k[sm.key] = x.mtd;
      if (sm.budget && x.bud !== null && x.mtd !== null) {
        if (sm.budget === 'pct') { vsBudget[sm.key] = x.bud; budget[sm.key] = x.bud > -100 ? Math.round(x.mtd / (1 + x.bud / 100) * 100) / 100 : null; }
        else { vsBudget[sm.key] = x.bud; budget[sm.key] = Math.round((x.mtd - x.bud / 100) * 100) / 100; }
      }
      if (sm.ly && x.ly !== null) vsLy[sm.key] = x.ly;
    }
    const canc = m['Cancellations']?.mtd, gross = m['Gross Sales']?.mtd;
    k.cancelPct = gross ? Math.round(Math.abs(canc || 0) / gross * 10000) / 100 : null;
    return { store, k, budget, vsBudget, vsLy };
  });
  return { missing: [], date, month: date ? date.slice(0, 7) : null, stores, ignored: [...ignored], unknown: [...unknown] };
}

// ---------------------------------------------------------------- goals and status
export function paceFactor(asOf) {
  if (!asOf) return 1;
  const [y, m, d] = asOf.split('-').map(Number);
  return Math.min(1, d / new Date(y, m, 0).getDate());
}
// green at or better than goal, amber within 20 percent, red beyond that.
export function status(metric, value, goal, pace = 1) {
  if (value === null || value === undefined || goal === null || goal === undefined || goal === '') return 'none';
  const target = metric.monthly ? goal * pace : goal;
  if (metric.lower) {
    if (value <= target) return 'green';
    return value <= target * 1.25 ? 'amber' : 'red';
  }
  if (!target) return 'none';
  const r = value / target;
  return r >= 1 ? 'green' : r >= 0.8 ? 'amber' : 'red';
}
export function fmt(metric, v) {
  if (v === null || v === undefined || Number.isNaN(v)) return '--';
  if (metric.fmt === 'money') return (v < 0 ? '-$' : '$') + Math.abs(Math.round(v)).toLocaleString('en-US');
  if (metric.fmt === 'money2') return '$' + v.toFixed(2);
  if (metric.fmt === 'pct') return v.toFixed(1) + '%';
  return Math.round(v).toLocaleString('en-US');
}
export function fmtGoal(metric, g) {
  if (g === null || g === undefined || g === '') return '--';
  if (metric.fmt === 'money' || metric.fmt === 'money2') return '$' + Math.round(Number(g)).toLocaleString('en-US');
  if (metric.fmt === 'pct') return Math.round(g * 10) / 10 + '%';
  return String(Math.round(g));
}

// ---------------------------------------------------------------- minimum standard
// Rolling SPH = last month's final sales and hours plus this month to date.
// cur / prev are consultant cards ({ k, hours, from, asOf }). prev may be missing.
export function rollingSph(cur, prev) {
  const parts = [cur, prev].filter(Boolean);
  const sales = parts.reduce((a, c) => a + (c.k.netSales || 0), 0);
  const hours = parts.reduce((a, c) => a + (c.hours || 0), 0);
  const from = (prev || cur).from;
  return { sph: hours ? sales / hours : null, sales, hours, from, to: cur.asOf, days: daysBetween(from, cur.asOf) };
}
export function minStatus(sph, min) {
  if (sph === null || sph === undefined || !min) return 'none';
  if (sph < min) return 'below';
  if (sph < min * 1.1) return 'watch';
  return 'ok';
}

// ---------------------------------------------------------------- weekly 1:1 coaching
// Talk tracks tied to the Core 4: Connection, Finance, Bedding, Presenting Every Option as
// Protected and Delivered. Net Sales is the result, so it is never picked as a focus.
export const COACHING = {
  sph: {
    pillar: 'Connection',
    why: 'Sales per hour tells us how well we use our time with guests.',
    ask: ['Tell me about your last three guests. How did each one end?', 'What do you do between ups to set up your next sale?'],
    doThis: 'Take your ups in order on the Up System. Call 3 be-backs every shift.'
  },
  avgTicket: {
    pillar: 'Presenting Every Option',
    why: 'When the ticket is low, we usually sold a piece instead of the room.',
    ask: ['On your last sale, what else did you show before you wrote it up?', 'What keeps you from showing the whole room?'],
    doThis: 'Start at Best and show the whole room: tables, rug, lighting, accents. Let the guest take things out.'
  },
  effMargin: {
    pillar: 'Presenting Every Option',
    why: 'Margin slips when we go to price before we build value.',
    ask: ['When a guest pushes on price, what do you say first?', 'On your last discount, did the guest ask or did you offer?'],
    doThis: 'Build value first, then show the monthly payment. Get a leader before you go below price.'
  },
  financePct: {
    pillar: 'Finance',
    why: 'Finance gives the guest buying power. If we wait until the end, it usually never comes up.',
    ask: ['When do you bring up financing?', 'How do you talk about the monthly payment?'],
    doThis: 'Get every guest their buying power. Show the monthly payment next to the price.'
  },
  beddingPct: {
    pillar: 'Bedding',
    why: 'Every guest sleeps on something. The healthy-sleep conversation is the easiest add we have.',
    ask: ['How many guests did you ask about their sleep this week?', 'What keeps you from walking guests to bedding?'],
    doThis: 'Have the healthy-sleep conversation with every guest. Walk at least 2 guests a day to the bedding gallery.'
  },
  protectionPct: {
    pillar: 'Protected and Delivered',
    why: 'Protection belongs in the price we present, from the first quote.',
    ask: ['When and how do you present protection today?', 'What do guests say when they pass on it?'],
    doThis: 'Present every option protected and delivered. Quote protection inside the price on every item.'
  },
  deliveryPct: {
    pillar: 'Protected and Delivered',
    why: 'Delivered is our standard. Carry-outs lead to more damage and more returns.',
    ask: ['Do you quote delivered or carry-out first?', 'Why did your last carry-out guest pass on delivery?'],
    doThis: 'Quote the delivered price first on every sale. Carry-out is the exception.'
  },
  beddingSph: {
    pillar: 'Bedding',
    why: 'Bedding per hour shows if the healthy-sleep conversation is happening with every guest, not just mattress shoppers.',
    ask: ['How many guests this week did you ask how they are sleeping?', 'On a bedroom guest, when do you bring up the mattress?'],
    doThis: 'Ask every guest how they are sleeping. Walk every bedroom guest to bedding before you write it up.'
  },
  protectionSph: {
    pillar: 'Protected and Delivered',
    why: 'Protection per hour shows how often protection is part of the presentation.',
    ask: ['Walk me through how you presented protection on your last sale.', 'Who did you not show protection to this week, and why?'],
    doThis: 'Present every item protected and delivered. Quote it inside the price, not at the end.'
  },
  protectionAttach: {
    pillar: 'Protected and Delivered',
    why: 'Attach rate is how many guests leave protected. Our standard is 6 out of 10.',
    ask: ['Out of your last 10 sales, how many left protected?', 'What do you say when a guest says they don\'t need it?'],
    doThis: 'Quote every item protected. Keep a tally this week and hit 6 out of 10.'
  },
  cancelPct: {
    pillar: 'Connection',
    why: 'A cancel is a sale we already had. Most come from a delivery date, a finance approval or the product not being clear.',
    ask: ['Walk me through your last cancel. When did you first know it was shaky?', 'What do you go over with the guest before they leave?'],
    doThis: 'Before the guest leaves, go over the delivery date, the payment and what they bought. Call them the next day.'
  },
  discountPct: {
    pillar: 'Presenting Every Option',
    why: 'Every point of discount comes out of margin. It should be the last tool we use.',
    ask: ['On your last discounted sale, who brought up price first?', 'What did you offer before the discount?'],
    doThis: 'Build value and show the monthly payment first. No discount without a leader, and only after finance.'
  },
  closeRate: {
    pillar: 'Connection',
    why: 'Close rate is how many of our guests leave as customers. Every point is money with no extra traffic.',
    ask: ['Which guests walked this week, and why?', 'Who got a TO to a leader before they left?'],
    doThis: 'No guest leaves without a TO to a leader or a second consultant.'
  },
  appsToTraffic: {
    pillar: 'Finance',
    why: 'Apps to traffic shows how many guests we offer finance to. Our standard is 10%.',
    ask: ['When in the visit is the team offering the app?', 'Who on the team runs the most apps, and what do they do differently?'],
    doThis: 'Get every guest their buying power. Leaders check apps at every huddle.'
  },
  creditApps: {
    pillar: 'Finance',
    why: 'No app, no approval. Apps open the door to finance.',
    ask: ['Who did you offer an app to this week, and who did you skip?', 'How do you ask for the app?'],
    doThis: 'Offer the app to every guest who likes the room but pauses on price. Get 4 or more this week.'
  }
};

// Picks at most 2 things to coach: the metrics furthest from goal, from different Core 4
// pillars when possible. If everything is at goal, returns one stretch item.
const gap = (m, v, goal) => (m.lower ? (v ? goal / v : 2) : v / goal);
export function pickFocus(k, goals, pace = 1, max = 2) {
  const cands = METRICS.filter(m => COACHING[m.key] && k[m.key] !== null && k[m.key] !== undefined && goals[m.key])
    .map(m => {
      const goal = m.monthly ? goals[m.key] * pace : goals[m.key];
      return { key: m.key, label: m.label, value: k[m.key], goal, lower: !!m.lower, ratio: gap(m, k[m.key], goal),
        perWeek: m.monthly ? Math.ceil(goals[m.key] / 4.33) : null };
    })
    .sort((a, b) => a.ratio - b.ratio);
  const below = cands.filter(x => x.ratio < 1);
  const out = [], used = new Set();
  for (const x of below) { if (out.length >= max) break; const p = COACHING[x.key].pillar; if (used.has(p)) continue; out.push(x); used.add(p); }
  for (const x of below) { if (out.length >= max) break; if (!out.includes(x)) out.push(x); }
  if (!out.length && cands.length) out.push({ ...cands[0], stretch: true });
  return out.map(x => ({ ...x, target: weeklyTarget(x) }));
}
// Next-week target: close half the gap to goal. Monthly counts get a this-week number.
export function weeklyTarget(x) {
  if (x.perWeek) return x.perWeek;
  let t;
  if (x.lower) t = x.value <= x.goal ? x.value * 0.95 : x.value - (x.value - x.goal) / 2;
  else t = x.ratio >= 1 ? x.value * 1.05 : x.value + (x.goal - x.value) / 2;
  return Math.round(t * 10) / 10;
}

// ---------------------------------------------------------------- store (team) coaching
// Store focus candidates. Results (Net Sales, SPG, Traffic) are left out; the team works the drivers.
export const TEAM_FOCUS = ['closeRate', 'sph', 'avgTicket', 'effMargin', 'financePct', 'appsToTraffic', 'beddingPct', 'beddingSph',
  'protectionPct', 'protectionSph', 'protectionAttach', 'deliveryPct', 'cancelPct'];
export function pickStoreFocus(t, goals, max = 2) {
  const cands = TEAM_FOCUS.map(key => {
    const m = STORE_METRICS.find(x => x.key === key);
    const v = t.k?.[key], goal = t.budget?.[key] ?? goals[key];
    if (v === null || v === undefined || goal === null || goal === undefined || !COACHING[key]) return null;
    return { key, label: m.label, value: v, goal, lower: !!m.lower, ratio: m.lower ? (v ? goal / v : 2) : v / goal, perWeek: null };
  }).filter(Boolean).sort((a, b) => a.ratio - b.ratio);
  const below = cands.filter(x => x.ratio < 1);
  const out = [], used = new Set();
  for (const x of below) { if (out.length >= max) break; const p = COACHING[x.key].pillar; if (used.has(p)) continue; out.push(x); used.add(p); }
  for (const x of below) { if (out.length >= max) break; if (!out.includes(x)) out.push(x); }
  if (!out.length && cands.length) out.push({ ...cands[0], stretch: true });
  return out.map(x => ({ ...x, target: weeklyTarget(x) }));
}

// ---------------------------------------------------------------- team roster (Paylocity export)
// Sheet "Sales Team": Region, Quartile, Location, Role, Name, Email. OPEN rows are empty seats.
export const TITLES = { 'assistant selling manager': 'ASM', 'sales lead': 'Sales Lead', 'rsa': 'RSA' };
export function parseTeamRoster(rawRows) {
  const rows = rawRows.map(r => { const o = {}; for (const [k, v] of Object.entries(r)) o[normalizeHeader(k)] = typeof v === 'string' ? v.trim() : v; return o; });
  const missing = ['location', 'role', 'name'].filter(c => !rows.some(r => c in r));
  if (missing.length) return { missing, people: [] };
  const people = [], open = [], badStores = new Set();
  for (const r of rows) {
    const name = String(r.name || '').trim(); if (!name) continue;
    const store = canonicalStore(r.location);
    if (!isKnownStore(store)) { badStores.add(r.location); continue; }
    const title = TITLES[String(r.role || '').toLowerCase()] || 'RSA';
    if (/^open$/i.test(name)) { open.push({ store, title }); continue; }
    people.push({ cid: cidOf(name), name, store, title, email: String(r.email || '').trim().toLowerCase() || null });
  }
  return { missing: [], people, open, badStores: [...badStores] };
}

// Links names in the RSA report to people on the roster. Exact name first, then saved aliases,
// then a suggestion when the last name matches and the first names share a start
// (Danny / Daniel, Nathan / Nathaniel, "Michael Mike Harding", "Daniel Martinez Jr").
const nameParts = s => String(s).toLowerCase().replace(/\b(jr|sr|ii|iii)\b/g, '').replace(/[^a-z\s-]/g, '').trim().split(/\s+/);
// Returns { cid, how } or null. how: 'name' (nickname / suffix), 'email' (first initial + last name
// matches their work email, catches name changes), 'possible' (same last name only: confirm by hand).
export function suggestMatch(reportName, candidates) {
  const [f, ...rest] = nameParts(reportName); const l = rest[rest.length - 1];
  if (!f || !l) return null;
  const one = arr => (arr.length === 1 ? arr[0] : null);
  const byName = one(candidates.filter(c => {
    const [cf, ...cr] = nameParts(c.name); const cl = cr[cr.length - 1];
    if (cl !== l) return false;
    return cf.slice(0, 3) === f.slice(0, 3) || (cf.startsWith(f.slice(0, 2)) && f.startsWith(cf.slice(0, 2))) || rest.includes(cf) || cr.includes(f);
  }));
  if (byName) return { cid: byName.cid, how: 'name' };
  const handle = (f[0] + l).replace(/[^a-z]/g, '');
  const byEmail = one(candidates.filter(c => c.email && c.email.split('@')[0].replace(/\d+$/, '') === handle));
  if (byEmail) return { cid: byEmail.cid, how: 'email' };
  const byLast = one(candidates.filter(c => { const cr = nameParts(c.name); return cr[cr.length - 1] === l; }));
  if (byLast) return { cid: byLast.cid, how: 'possible' };
  return null;
}
export function resolveReportNames(reportPeople, directory) {
  const byCid = new Map(directory.map(d => [d.cid, d]));
  const byAlias = new Map();
  directory.forEach(d => (d.aliases || []).forEach(a => byAlias.set(a, d)));
  const matched = [], unmatched = [];
  for (const p of reportPeople) {
    const d = byCid.get(p.cid) || byAlias.get(p.cid);
    if (d) matched.push({ p, d }); else unmatched.push(p);
  }
  const used = new Set(matched.map(x => x.d.cid));
  const open = directory.filter(d => !used.has(d.cid) && d.store !== '__skip');
  const suggestions = {};
  for (const p of unmatched) { const s = suggestMatch(p.name, open); if (s) suggestions[p.cid] = s; }
  return { matched, unmatched, suggestions, openRoster: open };
}
