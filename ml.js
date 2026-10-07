// Market Leader App: pure logic (no Firebase, no DOM). Safe to test in Node.
// Reads the same daily report as the Consultant Scorecard, but keeps the daily, week-to-date
// and month-to-date columns so the week plan can react to this week, not just the month.
import {
  normalizeHeader, toNumber, parseDate, slug, canonicalStore, isKnownStore, REGIONS, STORE_METRICS,
  COACHING, METRICS, pickFocus, pickStoreFocus, goalsFor, DEFAULT_GOALS, isOutlet, minSphFor, fmt, weeklyTarget
} from './base.js?v=202610071024';

const numOrNull = v => (v === '' || v === null || v === undefined ? null : toNumber(v));

// ---------------------------------------------------------------- dates (local, YYYY-MM-DD)
const pad = n => String(n).padStart(2, '0');
export const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
export const fromIso = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
export const addDays = (s, n) => { const d = fromIso(s); d.setDate(d.getDate() + n); return iso(d); };
export const daysApart = (a, b) => Math.round((fromIso(b) - fromIso(a)) / 86400000);
// Weeks run Sunday to Saturday. The plan for a week is built on its Sunday.
// Weeks run Monday to Sunday (the same week as WTD in the daily report).
export const weekStartOf = s => { const d = fromIso(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return iso(d); };
export const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
export const DAY_LONG = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
export const dow = s => fromIso(s).getDay();
// Market Leaders and directors pick any 2 days off each week. The default is Wednesday and Thursday.
export const DEFAULT_OFF = [3, 4];
export const DEFAULT_OFF_CHOICES = [2, 3, 4];   // Tue, Wed, Thu: what we suggest
export const blackoutFor = () => [];
export const offChoicesFor = () => [0, 1, 2, 3, 4, 5, 6];
export const validOff = off => Array.isArray(off) && off.length === 2 && new Set(off).size === 2 && off.every(d => d >= 0 && d <= 6);
export const safeOff = off => (validOff(off) ? [...off].sort() : DEFAULT_OFF);
export const VISIT_DAYS = 5;

// Store-level targets (company scorecard): $400 SPH, 55% finance of sales, 10% apps to traffic,
// 60% protection attachment. Everything else comes from the store's budget where it has one.
export const STORE_GOALS = { ...DEFAULT_GOALS.standard, financePct: 55 };

// ---------------------------------------------------------------- daily report
// Long format: report_date, segment, metric, then value columns per period.
// Period columns are matched loosely because exports name them differently.
const PERIODS = {
  day: ['day', 'daily', 'dtd', 'today'],
  wtd: ['wtd', 'week', 'weekly'],
  mtd: ['mtd', 'month']
};
const SIDES = { ty: ['ty', 'actual', 'act'], ly: ['ly', 'ly_var', 'vs_ly'], bud: ['budget', 'bud', 'vs_budget', 'plan'] };
function findCols(present) {
  const cols = {};
  for (const [p, pn] of Object.entries(PERIODS)) {
    cols[p] = {};
    for (const [side, sn] of Object.entries(SIDES)) {
      const hit = [...present].find(h => pn.some(a => sn.some(b => h === `${a}_${b}`)));
      if (hit) cols[p][side] = hit;
    }
    if (!cols[p].ty) delete cols[p];
  }
  return cols;
}
export function parseDaily(rawRows) {
  const rows = rawRows.map(r => { const o = {}; for (const [k, v] of Object.entries(r)) o[normalizeHeader(k)] = v; return o; });
  const present = new Set(rows.flatMap(r => Object.keys(r)));
  const missing = ['report_date', 'segment', 'metric'].filter(c => !present.has(c));
  const cols = findCols(present);
  if (!cols.mtd) missing.push('mtd_ty');
  if (missing.length) return { missing, stores: {} };
  const raw = {}, unknown = new Set();
  let date = null;
  for (const r of rows) {
    const seg = String(r.segment ?? '').trim();
    const d = parseDate(r.report_date); if (d && (!date || d > date)) date = d;
    if (!seg || /^(total|online)$/i.test(seg) || REGIONS.includes(seg)) continue;
    const store = canonicalStore(seg);
    if (!isKnownStore(store)) { unknown.add(seg); continue; }
    const m = (raw[store] ||= {});
    const key = String(r.metric).trim();
    m[key] = {};
    for (const [p, c] of Object.entries(cols)) m[key][p] = { ty: numOrNull(r[c.ty]), ly: c.ly ? numOrNull(r[c.ly]) : null, bud: c.bud ? numOrNull(r[c.bud]) : null };
  }
  const stores = {};
  for (const [store, m] of Object.entries(raw)) {
    const snap = {};
    for (const p of Object.keys(cols)) snap[p] = periodSnap(m, p);
    stores[store] = snap;
  }
  return { missing: [], date, periods: Object.keys(cols), stores, unknown: [...unknown] };
}
// k = values, vsBud = % (or bps for close rate) vs budget, vsLy = % vs LY, budget = the goal value.
function periodSnap(m, p) {
  const k = {}, vsBud = {}, vsLy = {}, budget = {};
  for (const sm of STORE_METRICS) {
    if (!sm.src) continue;
    const x = m[sm.src]?.[p]; if (!x) { k[sm.key] = null; continue; }
    k[sm.key] = x.ty;
    if (sm.budget && x.bud !== null && x.ty !== null) {
      vsBud[sm.key] = x.bud;
      budget[sm.key] = sm.budget === 'pct' ? (x.bud > -100 ? r2(x.ty / (1 + x.bud / 100)) : null) : r2(x.ty - x.bud / 100);
    }
    if (sm.ly && x.ly !== null) vsLy[sm.key] = x.ly;
  }
  const canc = m['Cancellations']?.[p]?.ty, gross = m['Gross Sales']?.[p]?.ty;
  k.cancelPct = gross ? r2(Math.abs(canc || 0) / gross * 100) : null;
  return { k, vsBud, vsLy, budget };
}
const r2 = v => Math.round(v * 100) / 100;

// ---------------------------------------------------------------- store need score
// Higher = needs the Market Leader more. Built from what the store controls and what it owes
// the budget. SPG with cancellations is the north star, so it carries the most weight after sales.
// Returns { score, parts: [{key, pts, text}] } with the biggest reasons first.
const cap = (v, max) => Math.max(0, Math.min(max, v));
export function needScore(snap, { lastVisit = null, today = null, team = null } = {}) {
  const parts = [];
  const add = (key, pts, text) => { if (pts > 0.5) parts.push({ key, pts: Math.round(pts), text }); };
  const mtd = snap?.mtd, wtd = snap?.wtd;
  if (mtd) {
    const s = mtd.vsBud.netSales;
    if (s != null && s < 0) add('sales', cap(-s * 1.2, 25), `Sales ${pct(s)} to budget this month`);
    const spg = mtd.vsLy.spg ?? mtd.vsBud.spg;
    if (spg != null && spg < 0) add('spg', cap(-spg * 1.2, 20), `SPG w/ cancellations ${pct(spg)} vs ${mtd.vsLy.spg != null ? 'LY' : 'budget'}`);
    const cr = mtd.vsBud.closeRate;
    if (cr != null && cr < 0) add('closeRate', cap(-cr / 50, 15), `Close rate ${Math.round(cr)} bps to budget`);
    const c = mtd.k.cancelPct;
    if (c != null && c > STORE_GOALS.cancelPct + 2) add('cancelPct', cap((c - STORE_GOALS.cancelPct) * 0.75, 4), `Cancellations at ${c.toFixed(1)}% of gross`);
    const pa = mtd.k.protectionAttach;
    if (pa != null && pa < STORE_GOALS.protectionAttach) add('protectionAttach', cap((STORE_GOALS.protectionAttach - pa) / 2, 8), `Protection attach ${pa.toFixed(0)}% (goal ${STORE_GOALS.protectionAttach}%)`);
    const fp = mtd.k.financePct;
    if (fp != null && fp < STORE_GOALS.financePct) add('financePct', cap((STORE_GOALS.financePct - fp) / 2, 8), `Finance ${fp.toFixed(0)}% of sales (goal ${STORE_GOALS.financePct}%)`);
  }
  if (wtd) {
    const s = wtd.vsBud.netSales;
    if (s != null && s < 0) add('wtd', cap(-s * 0.6, 15), `This week ${pct(s)} to budget`);
  }
  // Consultant side (from the daily RSA report): who is below the minimum, who is slipping this week.
  if (team) {
    if (team.below.length) add('below', cap(team.below.length * 4, 12), `${team.below.length} consultant${team.below.length > 1 ? 's' : ''} below the $${team.min} SPH minimum`);
    if (team.slipping.length) add('slipping', cap(team.slipping.length * 3, 9), `${team.slipping.length} consultant${team.slipping.length > 1 ? 's' : ''} slipping this week (${team.slipping.slice(0, 2).map(p => shortName(p.name)).join(', ')})`);
  }
  if (today) {
    const gap = lastVisit ? daysApart(lastVisit, today) : null;
    if (gap === null) add('visit', 8, 'No visit logged yet');
    else if (gap > 14) add('visit', 10, `Last visit ${gap} days ago`);
    else if (gap > 7) add('visit', 5, `Last visit ${gap} days ago`);
  }
  parts.sort((a, b) => b.pts - a.pts);
  return { score: Math.min(100, parts.reduce((a, p) => a + p.pts, 0)), parts };
}
export const pct = v => (v > 0 ? '+' : '') + (Math.round(v * 10) / 10) + '%';
export const band = score => (score >= 45 ? 'high' : score >= 20 ? 'mid' : 'low');

// Traffic vs LY: what the team does not control. Environment before effort.
export function environment(snap) {
  const p = snap?.wtd?.vsLy.traffic != null ? snap.wtd : snap?.mtd;
  const t = p?.vsLy.traffic;
  if (t == null) return null;
  const n = Math.abs(Math.round(t));
  if (t <= -15) return { kind: 'headwind', traffic: t, text: `Traffic is down ${n}% from last year. Check the schedule against traffic before you coach effort.` };
  if (t >= 10) return { kind: 'tailwind', traffic: t, text: `Traffic is up ${n}% from last year. The guests are coming in. It's on us to close them.` };
  return { kind: 'normal', traffic: t, text: `Traffic is ${t >= 0 ? 'up' : 'down'} ${n}% from last year, about normal. Results come down to how we sell.` };
}

// ---------------------------------------------------------------- the Sunday plan
// 5 visit days. Every store gets at least one visit when there are 5 or fewer; extra days go to
// the stores that need it most, as a second visit late in the week (visit 1 sets the plan,
// visit 2 checks it). With more than 5 stores, the lowest-need stores become a call.
export function workDays(weekStart, off = DEFAULT_OFF, role = 'leader') {
  off = safeOff(off, role);
  const out = [];
  for (let i = 0; i < 7 && out.length < VISIT_DAYS; i++) { const d = addDays(weekStart, i); if (!off.includes(dow(d))) out.push(d); }
  return out;
}
export function buildPlan({ weekStart, stores, scores, off = DEFAULT_OFF, role = 'leader' }) {
  const days = workDays(weekStart, off, role);
  const ranked = [...stores].sort((a, b) => (scores[b]?.score ?? 0) - (scores[a]?.score ?? 0) || a.localeCompare(b));
  const visited = ranked.slice(0, days.length);
  const calls = ranked.slice(days.length);
  const extra = Math.max(0, days.length - visited.length);
  const seconds = visited.slice(0, extra);
  const seq = [...visited.map(s => ({ store: s, kind: 'first' })), ...[...seconds].reverse().map(s => ({ store: s, kind: 'second' }))];
  // Only one store: every day is that store.
  while (seq.length < days.length && visited.length) seq.push({ store: visited[0], kind: 'second' });
  return {
    weekStart, off: safeOff(off, role),
    days: days.map((date, i) => ({ date, store: seq[i]?.store || null, kind: seq[i]?.kind || 'open', status: 'planned', score: scores[seq[i]?.store]?.score ?? null })),
    calls,
    basis: Object.fromEntries(stores.map(s => [s, scores[s]?.score ?? null]))
  };
}



// ---------------------------------------------------------------- levers and their inputs
// The Market Leader picks the lever (the outcome to move), then coaches the inputs that drive it.
// metric: the store number for that input, when the daily report has one. Inputs without a number
// are counted on the floor during the visit.
export const LEVERS = [
  { key: 'closeRate', label: 'Close Rate', why: 'More guests say yes today. Our job is to help the guest decide, and make it easy to say yes.', inputs: [
    { key: 'cart', label: 'Connection: cart creation', metric: null, count: 'Guests with a cart started', drill: 'cart',
      behavior: 'Start a cart with every guest. Connection shows up as a cart, and a guest with a cart is a guest who buys.',
      ask: 'How many of your last 10 guests left with a cart started?' },
    { key: 'value', label: 'Build value: product, experience, brand', metric: null, count: 'Guests who heard value in all three before price', drill: 'value',
      behavior: 'Build value before price comes up, in three places: the product, the experience and the brand.',
      fact: 'A guest decides to buy when the value is bigger than the price. If price comes up before value, price wins.',
      examples: [
        { t: 'Value in the product', say: 'Sit here and feel the cushion. Look at how the frame and the fabric are made. This is built for how your family actually lives, so it still looks this good years from now.', why: 'A guest cannot see quality on a price tag. When they understand what they are paying for, the price makes sense.' },
        { t: 'Value in the experience', say: 'We bring it in, set it up in your room and take the boxes with us. With protection, a spill or a tear gets taken care of. You are never on your own after the sale.', why: 'The guest is buying how it feels to own it, not just the piece. Delivery, setup and protection are worth more than a lower price somewhere else.' },
        { t: 'Value in the brand', say: 'You know the Ashley name, and we are your local Ashley store. We are here after the sale for delivery, service and the next room.', why: 'Trust takes the risk out of the decision. A guest who trusts us says yes today instead of shopping around.' }
      ],
      ask: 'Before your last guest saw a price, what did you tell them about the product, the experience and our brand?' },
    { key: 'finance', label: 'Finance: buying power early', metric: 'financePct', alt: 'appsToTraffic', drill: 'finance',
      behavior: 'Get every guest their buying power so the yes is easy.',
      fact: '93% of guests who get approved buy today, and 97% buy within 7 days.',
      ask: 'When in the conversation are you bringing up buying power?' },
    { key: 'options', label: 'Options: every option protected and delivered', metric: null, count: 'Guests shown their options on the options calculator', drill: 'options',
      behavior: 'Use the options calculator to show every option protected and delivered, with the monthly payment next to each one.',
      fact: 'Options turn "should I buy?" into "which one do I want?". That is an easier yes. We are there to help the guest decide, not to push one price.',
      ask: 'How many of your last 10 guests saw their options on the calculator?' }
  ] },
  { key: 'avgTicket', label: 'Average Ticket', why: 'More on every sale.', inputs: [
    { key: 'value', label: 'Build value: worth more than the price', metric: null, count: 'Guests who heard value in all three before price', drill: 'value',
      behavior: 'When the value is clear, guests choose the better piece and finish the room. Build value in the product, the experience and the brand before you show a price.',
      fact: 'A guest decides to buy when the value is bigger than the price. The more value they see, the more they are comfortable investing.',
      examples: [
        { t: 'Value in the product', say: 'Sit here and feel the cushion. Look at how the frame and the fabric are made. This is built for how your family actually lives, so it still looks this good years from now.', why: 'A guest cannot see quality on a price tag. When they understand what they are paying for, the price makes sense.' },
        { t: 'Value in the experience', say: 'We bring it in, set it up in your room and take the boxes with us. With protection, a spill or a tear gets taken care of. You are never on your own after the sale.', why: 'The guest is buying how it feels to own it, not just the piece. Delivery, setup and protection are worth more than a lower price somewhere else.' },
        { t: 'Value in the brand', say: 'You know the Ashley name, and we are your local Ashley store. We are here after the sale for delivery, service and the next room.', why: 'Trust takes the risk out of the decision. A guest who trusts us says yes today instead of shopping around.' }
      ],
      ask: 'What did you show your last guest that made the better piece worth it?' },
    { key: 'finance', label: 'Finance: buying power makes it affordable', metric: 'financePct', alt: 'appsToTraffic', drill: 'finance',
      behavior: 'Get buying power early and show the monthly payment, so the whole room is affordable.',
      fact: '93% of guests who get approved buy today, and 97% buy within 7 days.',
      ask: 'Are you showing the monthly payment next to the price?' },
    { key: 'pieces', label: 'More pieces: sell the room', metric: null, count: 'Pieces per ticket', drill: 'room',
      behavior: 'Show the whole room: tables, rug, lighting, accents. Let the guest take pieces out instead of never offering them.',
      ask: 'What did your last guest leave without that would have finished the room?' },
    { key: 'quality', label: 'Quality of the pieces: start at Best', metric: null, count: 'Guests shown Best first', drill: 'quality',
      behavior: 'Start at Best and walk down only if the guest asks.',
      ask: 'Which option do you show first?' },
    { key: 'bundle', label: 'Run the play: the bundle', metric: null, count: 'Sales with financing + Protection + Premium Delivery', drill: 'bundle',
      behavior: 'Present every option with 6 or 12 month financing plus Protection and Premium Delivery. The bundle is how the guest gets the biggest savings.',
      ask: 'How many of your last 10 sales had the full bundle?' },
    { key: 'bedding', label: 'Bedding: start the conversation', metric: 'beddingPct', drill: 'bedding',
      behavior: 'Ask every guest how they are sleeping. Start the conversation even when they came in for something else.',
      ask: 'How many guests today did you ask about their sleep?' },
    { key: 'protection', label: 'Protection', metric: 'protectionAttach', drill: 'protection', also: 'Also raises effective margin.',
      behavior: 'Present every option protected and delivered. Quote protection inside the price.',
      ask: 'Are you quoting protection inside the price, or adding it at the end?' },
    { key: 'delivery', label: 'Delivery', metric: 'deliveryPct', drill: 'delivery', also: 'Also raises effective margin.',
      behavior: 'Quote the delivered price first on every sale. Carry-out is the exception.',
      ask: 'Which price do you quote first?' }
  ] },
  { key: 'effMargin', label: 'Effective Margin', why: 'Profit is oxygen. We have to stay healthy.', inputs: [
    { key: 'protection', label: 'Protection', metric: 'protectionAttach', drill: 'protection',
      behavior: 'Present every option protected and delivered. Quote protection inside the price.',
      ask: 'Are you quoting protection inside the price, or adding it at the end?' },
    { key: 'delivery', label: 'Delivery', metric: 'deliveryPct', drill: 'delivery',
      behavior: 'Quote the delivered price first on every sale.',
      ask: 'Which price do you quote first?' },
    { key: 'bundle', label: 'Run the play: the bundle', metric: null, count: 'Sales with financing + Protection + Premium Delivery', drill: 'bundle',
      behavior: 'Present every option with 6 or 12 month financing plus Protection and Premium Delivery. The bundle is how the guest gets the biggest savings, and it protects our margin.',
      ask: 'How many of your last 10 sales had the full bundle?' },
    { key: 'price', label: 'Hold price', metric: null, count: 'Sales with no discount', drill: 'effMargin',
      behavior: 'Build value and show the monthly payment first. No discount without a leader, and only after finance.',
      ask: 'What happened the last time a guest asked for a better price?' }
  ] }
];
const LEVER_DEFAULT = { closeRate: null, avgTicket: 2200, effMargin: 55.5 };
// Where each lever and input stands for a store (or a market rollup): value, goal, and how far off.
export function leverStatus(p) {
  if (!p?.k) return [];
  return LEVERS.map(L => {
    const value = p.k[L.key], goal = p.budget?.[L.key] ?? (L.key === 'avgTicket' && p.store && isOutlet(DEFAULT_GOALS, p.store) ? DEFAULT_GOALS.outlet.avgTicket : LEVER_DEFAULT[L.key]);
    const ratio = value != null && goal ? value / goal : null;
    const inputs = L.inputs.map(inp => {
      if (!inp.metric) return { ...inp, value: null, goal: null, ratio: null };
      const v = p.k[inp.metric], g = p.budget?.[inp.metric] ?? STORE_GOALS[inp.metric];
      return { ...inp, value: v, goal: g, ratio: v != null && g ? v / g : null };
    });
    return { ...L, value, goal, ratio, inputs };
  });
}
// The lever to pull: the one furthest below goal. Null if every lever is at goal or there are no numbers.
// The app suggests one of the two levers. Effective margin is there to pick, not suggested.
export function suggestLever(p) {
  const st = leverStatus(p).filter(l => l.ratio != null && l.key !== 'effMargin');
  const worst = st.sort((a, b) => a.ratio - b.ratio)[0];
  return worst && worst.ratio < 1 ? worst.key : st.length ? null : null;
}


// ---------------------------------------------------------------- run the play (the in-store offer)
// Connect, build value, run the play. The offer is the current in-store event; Frank can change it in Setup.
export const OFFER_DEFAULT = {
  name: 'Fall Savings Event', start: '2026-09-15', end: '2026-10-26',
  lines: [
    '$50 off every $1,000 with 6 or 12 month financing, OR with the Protection + Premium Delivery bundle (cash, credit or Acima)',
    'Double your savings: $100 off every $1,000 with 6 or 12 month financing AND the Protection + Premium Delivery bundle',
    'Plus $300 off any premium mattress $1,999 or more, or 5% military discount (active and retired)'
  ],
  fine: 'Not combined with other offers or financing over 12 months. Excludes clearance, floor models, protection plans, and select premium mattress brands.'
};
export const PLAY = [
  { key: 'connect', t: 'Connect', d: 'Greet like a referral, learn how they live, and start a cart with every guest.' },
  { key: 'value', t: 'Build value', d: 'Start at Best, show the whole room, tell the Ashley story. Value before price, every time.' },
  { key: 'play', t: 'Run the play', d: 'Buying power early with 6 or 12 month financing, then present every option with Protection + Premium Delivery. That is the double savings.' }
];
export const PLAY_CHECKS = [
  'Connected first: greeted like a referral and started a cart',
  'Built value before price: started at Best and showed the whole room',
  'Buying power: got the guest their buying power with 6 or 12 month financing',
  'Bundle presented: every option with Protection + Premium Delivery',
  'Showed the guest their savings in dollars and asked which option feels right'
];
// Remote coaching can't watch the floor: ask the leader and role-play it on video. Same order as PLAY_CHECKS.
export const PLAY_CHECKS_REMOTE = [
  'Connection: the leader can tell you how many guests got a cart yesterday, and who',
  'Value: the leader walked you through a recent sale that started at Best and showed the whole room',
  'Buying power: the leader knows which guests got their buying power yesterday, and who didn\'t',
  'Bundle: role-played every option with Protection + Premium Delivery with the leader or an associate on video',
  'Savings in dollars: the double savings was said out loud in the role-play'
];
export const FLIQ_CHECKS_REMOTE = [
  'Per the leader: associates on the floor today have used FrontLine IQ',
  'Per the leader: reps in before the first guest at open',
  'Leader walked you through what FrontLine IQ flagged this week',
  'You and the leader picked one associate to coach on what it flagged'
];
export const offerActive = (o, day) => !!o && (!o.start || day >= o.start) && (!o.end || day <= o.end);
// What the bundle saves on a ticket, so the consultant can say it in dollars.
export const offerMath = amount => { const k = Math.floor(amount / 1000); return { one: k * 50, both: k * 100 }; };

// ---------------------------------------------------------------- FrontLine IQ pilot
// An AI sales coach in the store for sales associates. Only for Market Leaders flagged as pilots.
export const FLIQ_CHECKS = [
  'Associates on the floor today have used FrontLine IQ',
  'Reps in before the first guest: practiced with it at open',
  'Leader went over what FrontLine IQ flagged at the huddle',
  'One associate coached on something FrontLine IQ flagged'
];
export const FLIQ_DAILY = [
  'Sunday: go over last week\'s FrontLine IQ use with each store leader. Who used it, who didn\'t.',
  'Monday: reps before the first guest. Every associate runs one FrontLine IQ practice at open.',
  'Tuesday: connection. Practice the greeting and starting a cart in FrontLine IQ.',
  'Wednesday: buying power. Practice getting every guest their buying power.',
  'Thursday: run the play. Practice presenting the bundle and the double savings.',
  'Friday: weekend prep. Every associate role-plays the weekend guest in FrontLine IQ.',
  'Saturday: recognize the associate who used FrontLine IQ the most this week.'
];

// ---------------------------------------------------------------- drive time between stores
// City-level coordinates for each store. Drive time is an estimate: straight-line miles x 1.15 for
// roads, at 60 mph. Good enough to keep a Market Leader from a 3-hour afternoon drive; not routing.
export const STORE_GEO = {
  'Tallahassee': [30.44, -84.28], 'Thomasville': [30.84, -83.98], 'Albany': [31.58, -84.16], 'Macon': [32.84, -83.63],
  'Warner Robins': [32.61, -83.62], 'Dothan': [31.22, -85.39], 'Enterprise': [31.32, -85.86], 'Panama City': [30.18, -85.66],
  'Valdosta': [30.83, -83.28], 'Opelika': [32.65, -85.38], 'Columbus': [32.46, -84.99],
  'Town Center': [30.2597, -81.5246], 'North': [30.4788, -81.6376], 'Orange Park': [30.1992, -81.7381], 'Brunswick': [31.1994, -81.4809],
  'Yulee': [30.6317, -81.5489], 'St. Augustine': [29.8684, -81.3313], 'Outlet Regency': [30.3180, -81.5564],
  'Mobile': [30.68, -88.15], "D'Iberville": [30.43, -88.89], 'Spanish Fort': [30.67, -87.92], 'Pensacola': [30.47, -87.21],
  'Crestview': [30.76, -86.57], 'Ft. Walton': [30.42, -86.62], 'Outlet Pensacola': [30.47, -87.21],
  'Greensboro': [36.07, -79.79], 'Winston Salem': [36.10, -80.24], 'Burlington': [36.10, -79.44], 'Danville': [36.59, -79.40],
  'Outlet Greensboro': [36.07, -79.79],
  'Baton Rouge': [30.45, -91.15], 'Lafayette': [30.22, -92.02], 'Gonzales': [30.24, -90.92], 'Harahan': [29.94, -90.20],
  'Houma': [29.60, -90.72], 'Lake Charles': [30.23, -93.22], 'Opelousas': [30.53, -92.08], 'Ponchatoula': [30.44, -90.44],
  'Hattiesburg': [31.33, -89.29], 'Flowood': [32.31, -90.14], 'Harvey': [29.90, -90.08]
};
// Road drive times in minutes (typical, no traffic), looked up store to store. Pairs not here fall back to
// the straight-line estimate below.
export const ROAD_MIN = { 'North|Town Center': 28, 'Orange Park|Town Center': 29, 'Outlet Regency|Town Center': 13, 'St. Augustine|Town Center': 45, 'Town Center|Yulee': 49, 'Brunswick|Town Center': 96, 'North|Orange Park': 34, 'North|Outlet Regency': 20, 'North|St. Augustine': 64, 'North|Yulee': 24, 'Brunswick|North': 70, 'Orange Park|Outlet Regency': 32, 'Orange Park|St. Augustine': 48, 'Orange Park|Yulee': 56, 'Brunswick|Orange Park': 101, 'Outlet Regency|St. Augustine': 49, 'Outlet Regency|Yulee': 42, 'Brunswick|Outlet Regency': 88, 'St. Augustine|Yulee': 86, 'Brunswick|St. Augustine': 132, 'Brunswick|Yulee': 68 };
// A Market Leader's day starts and ends at home. Over this, one way, the day is flagged as a long drive.
export const LONG_DRIVE_MIN = 90;
export function driveMin(a, b) {
  if (a && b && a !== b) { const k = [a, b].sort().join('|'); if (ROAD_MIN[k] != null) return ROAD_MIN[k]; }
  const p = STORE_GEO[a], q = STORE_GEO[b];
  if (!p || !q || a === b) return a === b ? 0 : null;
  const R = 3959, rad = x => x * Math.PI / 180;
  const dLat = rad(q[0] - p[0]), dLon = rad(q[1] - p[1]);
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(rad(p[0])) * Math.cos(rad(q[0])) * Math.sin(dLon / 2) ** 2;
  const miles = 2 * R * Math.asin(Math.sqrt(h)) * 1.15;
  return Math.round(miles / 60 * 60 / 5) * 5;
}
export const driveText = m => m == null ? '' : m < 60 ? `${Math.max(m, 5)} min` : `${Math.floor(m / 60)} hr${m % 60 >= 15 ? ` ${m % 60} min` : ''}`;
// A half-day split only works if the afternoon store is close. Over this, it's a full day instead.
export const MAX_SPLIT_MIN = 75;

// ---------------------------------------------------------------- mid-week pivot
// After a new upload: is there a store that now needs the leader more than one still on the plan?
// Suggest one swap at a time. The leader accepts or keeps the plan.
export const PIVOT_GAP = 15;
export function pivotSuggestion({ plan, scores, today, dismissed = [] }) {
  // Anchor days stay put: the leader is there on purpose, every one of those mornings.
  const ahead = plan.days.map((d, i) => ({ ...d, i })).filter(d => d.date > today && d.status === 'planned' && d.store && !d.anchor);
  if (!ahead.length) return null;
  const count = s => plan.days.filter(d => d.store === s && (d.date > today || d.status === 'done')).length;
  const aheadSet = new Set([...ahead.map(d => d.store), ...plan.days.filter(d => d.date > today).flatMap(d => [d.anchor ? d.store : null, ...(d.stops || []).map(x => x.store)]).filter(Boolean)]);
  const candidates = Object.keys(plan.basis)
    .filter(s => !aheadSet.has(s))
    .map(s => ({ store: s, score: scores[s]?.score ?? 0, was: plan.basis[s] ?? 0 }))
    // Only stores that got worse since Sunday. Sunday's ranking already weighed everything else.
    .filter(c => c.score >= 25 && c.score - c.was >= 10)
    .sort((a, b) => b.score - a.score);
  // Rank the rest of the week by priority. The visit with the least opportunity (lowest need
  // right now) is the one that comes off. On a tie, a store's second visit goes first.
  const sc = st => scores[st]?.score ?? 0;
  const target = [...ahead].sort((a, b) => sc(a.store) - sc(b.store) || (a.kind === 'second' ? -1 : 0) - (b.kind === 'second' ? -1 : 0) || b.date.localeCompare(a.date))[0];
  const tScore = sc(target.store);
  for (const c of candidates) {
    if (c.score - tScore < PIVOT_GAP) continue;
    const key = `${plan.weekStart}|${target.date}|${c.store}`;
    if (dismissed.includes(key)) continue;
    const rose = c.score - c.was;
    // After the swap, the remaining days run in priority order: highest need on the next open day.
    const stores = ahead.filter(d => d !== target).map(d => d.store).concat(c.store).sort((a, b) => sc(b) - sc(a));
    const seenBefore = new Set(plan.days.filter(d => d.date <= today && d.store).map(d => d.store));
    const reorder = ahead.map(d => d.i).sort((x, y) => plan.days[x].date.localeCompare(plan.days[y].date)).map((i, n) => {
      const st = stores[n], kind = seenBefore.has(st) ? 'second' : 'first';
      seenBefore.add(st);
      return { i, date: plan.days[i].date, store: st, kind, score: sc(st) };
    });
    return {
      key, dayIndex: target.i, date: target.date, from: target.store, to: c.store, fromScore: tScore, toScore: c.score,
      loses: count(target.store) <= 1, reorder,
      reasonTo: `${c.store} is now at ${c.score} need${rose >= 5 ? ` (up ${rose} since Sunday)` : ''}.`,
      reasonFrom: `${target.store} is the lowest priority left this week at ${tScore}.`,
      reason: `${c.store} is now at ${c.score} need${rose >= 5 ? ` (up ${rose} since Sunday)` : ''}. ${target.store} is the lowest priority left this week at ${tScore}.`,
      why: scores[c.store]?.parts?.slice(0, 2).map(p => p.text) || []
    };
  }
  return null;
}

// ---------------------------------------------------------------- visit agenda
// What the leader does in the store: 2 team focuses, up to 3 RSA conversations, a 6 Elements walk.
// The 6 Elements visit walk (from the director Store Visit app). Each item is scored Yes / Partial / No.
export const ELEMENTS = [
  { key: 'culture', n: 1, t: 'High Performing Sales Culture', q: 'Is the leader building a team that runs the play?', items: [
    'Team understands our selling structure: can explain the Core 4 and why it wins',
    'Consultant coaching and scorecards in use: 1:1 coaching happening, scorecards reviewed',
    'Role clarity and training discipline: a daily and weekly coaching rhythm is in place',
    'Performance coaching is happening on the floor'] },
  { key: 'assortment', n: 2, t: 'Assortment', q: 'Do we have what customers actually want?', items: [
    'Floor reflects what customers are actually buying',
    'Floor space maximized: heroes placed, no dead space',
    'Markdown and clearance strategy: right product, right time, right place',
    'Best sellers in stock and on the floor',
    'Accessories and attachments available to complete the sale'] },
  { key: 'visual', n: 3, t: 'Visual Presentation', q: 'Is your store Grand Opening Ready, every single day?', aor: true, items: [] },
  { key: 'facilities', n: 4, t: 'Facilities', q: 'Is the store clean, safe, and working?', items: [
    'All tech and TV equipment working', 'Bathrooms spotless', 'Backroom clean and safe',
    'Parking lot and signage clean and visible', 'Lighting fully functional, storefront and windows clean'] },
  { key: 'backoffice', n: 5, t: 'Back Office Controls', q: 'Are we managing the business behind the sale?', items: [
    'Order management: delivery dates managed and accurate', 'Customer service: owning the guest experience',
    'Margin and discounting: protecting profitability', 'VAMOO worked for open orders'] },
  { key: 'inventory', n: 6, t: 'Inventory Control', q: 'Are the controls tight and accurate?', items: [
    'Tight, accurate controls in place', 'Shrink minimized through discipline',
    'Visual and operational alignment with what is on hand', 'Cycle counts current; damages and RTV processed'] }
];
// Element 1 broken into the 3 value segments, watched on live guests.
export const SEGMENTS = [
  { name: 'Value in Product', must: 'What the guest must KNOW before they say yes', items: [
    "Consultant connected features to the guest's specific words", 'Guest understood WHY, not just what',
    'Good / Better / Best presented: full range, never narrowed by assumed budget', 'Started at Best and walked down gracefully: no pre-qualifying the guest'] },
  { name: 'Value in Experience', must: 'What the guest must FEEL throughout', items: [
    "Greeting felt like a referral, not a transaction: guest's name used", 'Guest was heard before any product was shown',
    'Guided, not sold: zero pressure, genuine care', 'Guest felt like the only person in the store'] },
  { name: 'Value in Ashley Brand', must: 'What the guest must BELIEVE when they leave', items: [
    'Ashley story told with conviction: believed, not memorized', "Guest understood why Ashley's size and scale means quality and value",
    'Brand tied naturally into the protection conversation', 'Guest would send a friend here'] }
];
// Element 3 AOR walk.
export const AORS = ['Front Entrance and Windows', 'Customer Service Desk and Greeter', 'Living Room / Upholstery', 'Bedroom', 'Mattress / Bedding Gallery',
  'Dining', 'Occasional / Accents / Accessories', 'Clearance / Outlet', 'Restrooms', 'Backroom / Warehouse'];
// Stand-up practice: run the play with a consultant and score what you see.
export const PRACTICE = [
  'Core 4 delivered in order: Connection, Finance, Bedding, Protected and Delivered',
  'Connection felt real, starting at the door', 'Finance introduced early: buying power first',
  'Healthy-sleep (bedding) conversation included',
  'All 4 options presented as protected and delivered; closed with "which feels right?"', 'Ashley story told with conviction'];
export const VISIT_TYPES = ['Priority', 'Follow-Up', 'Coaching Stop'];
export const kindToType = k => (k === 'second' ? 'Follow-Up' : k === 'first' ? 'Priority' : 'Coaching Stop');

// Visit score: Yes = 1, Partial = half, No = 0 across the element items, value segments and AOR walk.
export function visitScore(v) {
  let n = 0, t = 0;
  const add = x => { if (!x) return; t++; n += x === 'yes' || x === 'pass' ? 1 : x === 'partial' ? 0.5 : 0; };
  Object.values(v.checks || {}).forEach(o => Object.values(o || {}).forEach(add));
  Object.values(v.segs || {}).forEach(o => Object.values(o || {}).forEach(add));
  Object.values(v.aor || {}).forEach(add);
  return t ? { pct: Math.round(n / t * 100), n, t } : null;
}
// Commitments are always "from X to Y": where it is today, and where it will be by the next visit.
export const hasCommitment = a => !!(a && (String(a.what || '').trim() || String(a.behavior || '').trim()));
export function commitmentText(a) {
  const what = String(a.what || a.behavior || '').trim();
  const ft = a.from || a.to ? `: from ${String(a.from || '?').trim()} to ${String(a.to || '?').trim()}` : '';
  const how = String(a.how || '').trim();
  const md = d => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d || ''); return m ? `${+m[2]}/${+m[3]}` : d; };
  const who = a.owner ? ` (${String(a.owner).trim()}${a.due ? ', by ' + md(a.due) : ''})` : a.due ? ` (by ${md(a.due)})` : '';
  return `${what}${ft}${how ? '. ' + how.replace(/\.$/, '') : ''}${who}`;
}
// What a visit left behind, in plain terms, for the log, the next visit and the recap message.
export function visitSummary(v) {
  const fixes = ELEMENTS.filter(e => {
    if (e.aor) return Object.values(v.aor || {}).includes('needs') || v.elements?.[e.key] === 'fix';
    return Object.values(v.checks?.[e.key] || {}).includes('no') || (e.key === 'culture' && Object.values(v.segs || {}).some(o => Object.values(o || {}).includes('no'))) || v.elements?.[e.key] === 'fix';
  }).map(e => e.t);
  const actions = (v.actions || []).filter(hasCommitment);
  const commitments = actions.length ? actions.map(commitmentText)
    : String(v.commitments || '').split(/\n+/).map(x => x.trim()).filter(Boolean);
  const coached = (v.consultants || []).map(c => {
    const d = c.drill || {}, sc = Object.values(d.scored || {});
    const pts = sc.reduce((a, x) => a + (x === 'yes' ? 1 : x === 'partial' ? 0.5 : 0), 0);
    const dd = d.key ? drillFor(d.key) : null;
    if (c.mode === 'leader') { const ls = Object.values(c.lead?.scored || {}); const lp = ls.reduce((a, x) => a + (x === 'yes' ? 1 : x === 'partial' ? 0.5 : 0), 0);
      return { name: c.name, notes: String(c.notes || '').trim(), via: 'leader', drill: null, ran: ls.length > 0, score: ls.length ? `${lp} of 4 on the leader's coaching` : null, rerun: null, well: c.lead?.well || '', adjust: c.lead?.adjust || '' }; }
    return { name: c.name, notes: String(c.notes || '').trim(), drill: dd?.title || d.title || null, ran: sc.length > 0, score: sc.length ? `${pts} of ${dd ? dd.watch.length : sc.length}` : null, rerun: d.rerun || null, well: d.well || '', adjust: d.adjust || '' };
  });
  const lc = v.leaderCommit || {};
  const mdd = d => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d || ''); return m ? `${+m[2]}/${+m[3]}` : ''; };
  const lcBase = String(lc.what || '').trim() ? `${String(lc.what).trim()}${lc.from || lc.to ? `: from ${String(lc.from || '?').trim()} to ${String(lc.to || '?').trim()}` : ''}` : String(lc.commit || '').trim();
  const leaderCommit = lcBase ? lcBase + (lc.by ? ` by ${mdd(lc.by)}` : '') : '';
  const support = String(lc.support || '').trim() ? String(lc.support).trim() + (lc.supportBy ? ` (by ${mdd(lc.supportBy)})` : '') : '';
  return { fixes, commitments, coached, leaderCommit, support, leaderName: lc.name || v.leaderWin?.name || '', win: v.leaderWin?.text || '', winName: v.leaderWin?.name || '', working: v.working || '', score: visitScore(v) };
}
// Coaching comes from the two levers (close rate and average ticket) through our core behaviors:
// connection, buying power, bedding, protection and delivery. Cancellations and the like are watched,
// not coached as the focus.
export const CORE_FOCUS = ['financePct', 'appsToTraffic', 'beddingPct', 'protectionAttach', 'deliveryPct'];
export function storeFocus(snap, max = 2) {
  const p = snap?.mtd; if (!p) return [];
  const all = pickStoreFocus({ k: p.k, budget: p.budget }, STORE_GOALS, 20).filter(f => CORE_FOCUS.includes(f.key));
  const below = all.filter(f => !f.stretch);
  const pick = (below.length ? below : all).slice(0, max);
  if (!pick.length) { const st = pickStoreFocus({ k: p.k, budget: p.budget }, STORE_GOALS, 1)[0]; if (st && CORE_FOCUS.includes(st.key)) pick.push(st); }
  return pick.map(f => ({ ...f, coach: COACHING[f.key] }));
}
// RSA picks for one store: below minimum standard first, then biggest gap, then one person to
// recognize and use as the model. Each gets one focus item, never a checklist.
export function rsaPicks(people, store, goals = DEFAULT_GOALS, pace = 1, max = 3, weeks = {}) {
  const list = people.filter(p => p.store === store && p.k?.sph != null);
  if (!list.length) return [];
  const min = minSphFor(goals, store);
  const g = goalsFor(goals, store);
  const scored = list.map(p => {
    const focus = pickFocus(p.k, g, pace, 6).filter(f => !['cancelPct', 'discountPct'].includes(f.key))[0] || null;
    const w = weeks[p.cid] || {};
    const slipping = w.hours >= 12 && w.priorSph && w.sph != null && w.sph < w.priorSph * SLIP;
    return { ...p, focus, below: p.k.sph < min, slipping, wk: w, gap: focus && !focus.stretch ? focus.ratio : 1.5 };
  });
  const out = [];
  scored.filter(p => p.below).sort((a, b) => a.k.sph - b.k.sph).forEach(p => out.length < max - 1 && out.push({ ...p, why: 'below' }));
  scored.filter(p => p.slipping && !p.below).sort((a, b) => a.wk.sph / a.wk.priorSph - b.wk.sph / b.wk.priorSph).forEach(p => out.length < max - 1 && out.push({ ...p, why: 'slipping' }));
  scored.filter(p => !out.some(o => o.cid === p.cid) && !p.below).sort((a, b) => a.gap - b.gap)
    .forEach(p => out.length < max - 1 && p.gap < 1 && !out.some(o => o.cid === p.cid) && out.push({ ...p, why: 'gap' }));
  const star = [...scored].filter(p => !out.some(o => o.cid === p.cid)).sort((a, b) => b.k.sph - a.k.sph)[0];
  if (star) out.push({ ...star, why: 'model' });
  return out.map(p => ({ cid: p.cid, name: p.name, sph: p.k.sph, hours: p.hours, why: p.why, wk: p.wk, min, focus: p.focus, coach: p.focus ? COACHING[p.focus.key] : null }));
}

// ---------------------------------------------------------------- consultants, day to day
// The RSA report is month to date. Uploading it daily and keeping the copy from Saturday lets the
// app take this week out of it: this week = today's month-to-date minus Saturday's month-to-date.
// If the month turned over this week, this week is simply the month to date.
export function consultantWeeks(latest, base) {
  const out = {};
  if (!latest?.people) return out;
  const sameMonth = base?.to && latest.to && base.to.slice(0, 7) === latest.to.slice(0, 7);
  const prior = new Map((sameMonth ? base.people : []).map(p => [p.cid, p]));
  for (const p of latest.people) {
    const b = prior.get(p.cid);
    const sales = (p.k?.netSales || 0) - (b?.k?.netSales || 0);
    const hours = (p.hours || 0) - (b?.hours || 0);
    const priorSph = b?.hours > 20 ? b.k.netSales / b.hours : null;
    out[p.cid] = { sales, hours, sph: hours >= 1 ? sales / hours : null, priorSph };
  }
  return out;
}
const shortName = n => { const t = String(n).trim().split(/\s+/); const f = t[0][0] + t[0].slice(1).toLowerCase(); return t.length > 1 ? `${f} ${t[t.length - 1][0]}.` : f; };
// Slipping = at least 12 hours on the floor this week and this week's SPH under 75% of their month before it.
export const SLIP = 0.75;
export function teamSignals(people, weeks, store, goals = DEFAULT_GOALS) {
  const list = people.filter(p => p.store === store && p.k?.sph != null);
  const min = minSphFor(goals, store);
  const rows = list.map(p => {
    const w = weeks[p.cid] || {};
    const slipping = w.hours >= 12 && w.priorSph && w.sph != null && w.sph < w.priorSph * SLIP;
    const rising = w.hours >= 12 && w.priorSph && w.sph != null && w.sph > w.priorSph * 1.25;
    return { cid: p.cid, name: p.name, sph: p.k.sph, hours: p.hours, wk: w, below: p.k.sph < min, slipping, rising };
  }).sort((a, b) => b.sph - a.sph);
  const hot = rows.filter(r => r.wk.hours >= 12 && r.wk.sph != null).sort((a, b) => b.wk.sph - a.wk.sph)[0] || null;
  return { min, rows, below: rows.filter(r => r.below), slipping: rows.filter(r => r.slipping && !r.below), rising: rows.filter(r => r.rising), hot };
}

// ---------------------------------------------------------------- stand-up practice drills
// One drill per coaching lever. The leader plays the guest, the consultant runs the rep,
// the leader scores what they see, gives one adjustment, and they run it again.
const D = {
  value: { title: 'Building value before price', guest: 'Ask "How much is this one?" in the first minute.',
    watch: ['Acknowledged the question and came back to price after value', 'Built value in the product: what it is made of and why it lasts', 'Built value in the experience: delivery, setup, protection', 'Built value in the brand: Ashley, and us as the local store', 'Showed the price with the monthly payment, after value'] },
  options: { title: 'Presenting options on the calculator', guest: 'Say "I need to think about it" when you see the first price.',
    watch: ['Opened the options calculator instead of dropping the price', 'Showed every option protected and delivered', 'Put the monthly payment next to each option', 'Asked which option fits best, not whether they want to buy', 'Let the guest choose and confirmed the yes'] },
  bundle: { title: 'Running the play', guest: 'Ask "Is there any deal going on right now?" before the salesperson brings it up.',
    watch: ['Connected and started a cart before talking about the offer', 'Built value first: started at Best and showed the whole room', 'Got buying power with 6 or 12 month financing', 'Presented every option with Protection + Premium Delivery and showed the double savings'] },
  cart: { title: 'Building the cart', guest: 'Say "I\'m just looking at sofas today." Like two pieces, but don\'t ask for anything.',
    watch: ['Started a cart early in the conversation', 'Added every piece the guest liked as they went', 'Asked about the rest of the room and added to the cart', 'Walked the guest through the cart before any talk of price'] },
  quality: { title: 'Starting at Best', guest: 'Ask "What\'s the difference between these three?"',
    watch: ['Showed the Best option first', 'Explained why Best is better in the guest\'s own words', 'Walked down only when the guest asked', 'Quoted Best as a monthly payment'] },
  connection: { title: 'Greeting to discovery', guest: 'Walk in and say "Just looking." Keep your answers short until they get you talking.',
    watch: ['Greeted within 10 seconds, used your name once they had it', 'Asked about the room and how you live in it before showing product', 'Listened more than they talked in the first 3 minutes', 'Set up the next step before walking to the floor'] },
  room: { title: 'Selling the room', guest: 'Pick one sofa and say "I\'ll figure out the rest later."',
    watch: ['Showed the full room: tables, rug, lighting, accents', 'Tied each add-on back to something you said', 'Let you take pieces out instead of never offering them', 'Gave one total for the room, delivered and protected'] },
  price: { title: 'Holding price', guest: 'Say "That\'s more than I wanted to spend. Can you do better?"',
    watch: ['Did not jump to a discount', 'Went back to value in your own words', 'Offered the monthly payment before any price move', 'Brought in a leader before going below price'] },
  finance: { title: 'Getting the guest their buying power', guest: 'Shop a bedroom set. Don\'t bring up a budget.',
    watch: ['Got the guest their buying power before price came up', 'Framed it as buying power, not credit', 'Quoted a monthly payment next to the price', 'Asked for the app with a clear, easy ask'] },
  bedding: { title: 'The sleep question', guest: 'Buy a bed frame and say your mattress is "fine."',
    watch: ['Asked how you are sleeping, not whether you need a mattress', 'Asked a follow-up about age of mattress, pain or partner', 'Walked you to the bedding gallery', 'Had you lie down and compared two options'] },
  protection: { title: 'Protection inside the price', guest: 'Say "I don\'t need protection. I\'m careful."',
    watch: ['Quoted the price with protection already included', 'Tied it to your life: kids, pets, food, the fabric', 'Handled the objection without dropping it right away', 'Presented it on every item, not just the big one'] },
  delivery: { title: 'Delivered first', guest: 'Say "I brought my truck. I\'ll take it today."',
    watch: ['Quoted the delivered price first', 'Explained what delivery covers: set up, placement, haul-away', 'Connected delivery to protecting the purchase', 'Treated carry-out as the exception'] },
  cancel: { title: 'Locking in the sale', guest: 'Say yes, then start heading for the door.',
    watch: ['Confirmed the delivery date and window out loud', 'Confirmed the payment and finance approval', 'Walked you through exactly what you bought', 'Set up the next-day thank-you call'] },
  close: { title: 'Turnover before a guest walks', guest: 'Say "Let me think about it and talk to my husband."',
    watch: ['Asked what they would need to feel good today', 'Offered a hold, a cart or a photo of the room', 'Brought in a leader or second consultant for a turnover', 'Got a name and number with a reason to follow up'] },
  apps: { title: 'Asking for the app', guest: 'Love the room, then go quiet when you see the total.',
    watch: ['Recognized the pause as the moment to offer the app', 'Explained it takes a few minutes and has no cost to check', 'Made a clear ask instead of "if you want"', 'Walked you through it or handed it off to a leader'] }
};
export const DRILLS = {
  value: D.value, options: D.options, bundle: D.bundle, cart: D.cart, quality: D.quality, finance: D.finance, room: D.room, bedding: D.bedding, protection: D.protection, delivery: D.delivery,
  sph: D.connection, closeRate: D.close, cancelPct: D.cancel, avgTicket: D.room, effMargin: D.price, discountPct: D.price,
  financePct: D.finance, appsToTraffic: D.apps, creditApps: D.apps, beddingPct: D.bedding, beddingSph: D.bedding,
  protectionPct: D.protection, protectionSph: D.protection, protectionAttach: D.protection, deliveryPct: D.delivery
};
export const drillFor = key => DRILLS[key] || D.connection;

// Suggested 1:1 coaching for any consultant, from their month-to-date numbers.
// Opens with a real strength, then at most 2 levers from different parts of the Core 4, then one commitment.
const fmtK = (key, v) => {
  if (v == null) return '--';
  const m = METRICS.find(x => x.key === key);
  if (m?.fmt === 'pct') return `${(Math.round(v * 10) / 10)}%`;
  if (m?.fmt === 'int') return String(Math.round(v));
  return '$' + Math.round(v).toLocaleString('en-US');
};
// ---------------------------------------------------------------- store opportunity -> people
// When the store is short on a number, name the consultants pulling it down. The store number
// maps to the consultant number that drives it; close rate has no consultant number in the RSA
// report, so it uses sales per hour.
export const STORE_TO_RSA = { sph: 'sph', closeRate: 'sph', avgTicket: 'avgTicket', effMargin: 'effMargin', financePct: 'financePct',
  appsToTraffic: 'creditApps', beddingPct: 'beddingPct', beddingSph: 'beddingSph', protectionPct: 'protectionPct',
  protectionSph: 'protectionSph', protectionAttach: 'protectionPct', deliveryPct: 'deliveryPct', cancelPct: 'cancelPct' };
const fmtD = v => '$' + Math.round(v).toLocaleString('en-US');
// Ranked by how much each person costs the store on that number: the gap times their volume.
// helpers() is the other side: who is above goal and carrying the store on that number.
export function draggers(people, store, storeKey, goals = DEFAULT_GOALS, pace = 1, n = 3) { return rankOn(people, store, storeKey, goals, pace, n, false); }
export function helpers(people, store, storeKey, goals = DEFAULT_GOALS, pace = 1, n = 3) { return rankOn(people, store, storeKey, goals, pace, n, true); }
function rankOn(people, store, storeKey, goals, pace, n, up) {
  const key = STORE_TO_RSA[storeKey]; if (!key) return [];
  const m = METRICS.find(x => x.key === key); if (!m) return [];
  const g = goalsFor(goals, store);
  const goal = m.monthly ? g[key] * pace : g[key];
  if (goal == null) return [];
  const seen = new Set();
  return people.filter(p => p.store === store && p.k?.[key] != null && (p.hours || 0) >= 20 && !seen.has(p.cid) && seen.add(p.cid)).map(p => {
    const v = p.k[key], sales = p.k.netSales || 0;
    // gap > 0 means below goal (for lower-is-better numbers, above goal).
    const gap = m.lower ? v - goal : goal - v;
    const size = Math.abs(gap);
    let impact, impactText;
    if (m.lower) impact = size / 100 * sales;
    else if (m.fmt === 'pct') impact = size / 100 * sales;
    else if (key === 'creditApps') impact = size * 1000;
    else if (key === 'avgTicket') impact = size * (sales / Math.max(v, 1));
    else impact = size * (p.hours || 0);
    if (gap > 0) impactText = m.lower ? `about ${fmtD(impact)} in ${key === 'cancelPct' ? 'cancels' : 'discounts'} over goal` : key === 'creditApps' ? `${Math.round(size)} apps short` : `about ${fmtD(impact)} short this month`;
    else impactText = m.lower ? `about ${fmtD(impact)} better than goal` : key === 'creditApps' ? `${Math.round(size)} apps over goal` : `about ${fmtD(impact)} over goal this month`;
    return { cid: p.cid, name: p.name, key, value: v, goal, gap, impact, impactText, hours: p.hours };
  }).filter(x => (up ? x.gap < 0 : x.gap > 0)).sort((a, b) => b.impact - a.impact).slice(0, n);
}

// Plain names for the numbers, the way we say them in the store.
export const PLAIN = { sph: 'sales per hour', avgTicket: 'average ticket', effMargin: 'margin', financePct: 'finance', creditApps: 'credit apps',
  beddingPct: 'bedding', beddingSph: 'bedding per hour', protectionPct: 'protection', protectionSph: 'protection per hour', protectionAttach: 'protection attach',
  deliveryPct: 'delivery', cancelPct: 'cancels', discountPct: 'discount', closeRate: 'close rate', appsToTraffic: 'apps to traffic', netSales: 'sales' };
export function consultantCoaching({ p, store, why = 'added', wk = null, goals = DEFAULT_GOALS, pace = 1, teamFocus = null, lever = null }) {
  const name = String(p?.name || '').trim().split(/\s+/)[0];
  const first = name ? name[0].toUpperCase() + name.slice(1).toLowerCase() : 'them';
  if (!p?.k || p.k.sph == null) {
    return { items: [], strength: null, drillKey: 'sph', drill: drillFor('sph'), text: `${first} isn't in the RSA report yet, so there are no numbers to go on.\nWatch ${first} with a live guest and see how far they get through the Core 4. Coach the first step they skip.\nPractice it: run the full Core 4 standing up, start to finish.\n${first}'s commitment: one full Core 4 presentation on the next guest.` };
  }
  const g = goalsFor(goals, store);
  // A consultant's coaching points to the core behaviors, never cancels or discount on their own.
  let items = pickFocus(p.k, g, pace, 6).map(f => ({ ...f, coach: COACHING[f.key] })).filter(f => f.coach && !['cancelPct', 'discountPct'].includes(f.key)).slice(0, 2);
  // Coaching this person on a store opportunity: that number goes first.
  const lk = lever && STORE_TO_RSA[lever];
  const lm = lk && METRICS.find(x => x.key === lk);
  if (lm && p.k[lk] != null && g[lk] != null) {
    const goal = lm.monthly ? g[lk] * pace : g[lk], value = p.k[lk];
    const ratio = lm.lower ? (value ? goal / value : 2) : value / goal;
    if (ratio < 1) {
      const f = { key: lk, label: lm.label, value, goal, lower: !!lm.lower, ratio, perWeek: lm.monthly ? Math.ceil(g[lk] / 4.33) : null, coach: COACHING[lk] };
      f.target = weeklyTarget(f);
      items = [f, ...items.filter(x => x.key !== lk && COACHING[x.key]?.pillar !== f.coach.pillar)].slice(0, 2);
    }
  }
  const strengths = METRICS.filter(m => !m.lower && COACHING[m.key] && p.k[m.key] != null && g[m.key])
    .map(m => ({ key: m.key, label: m.label, value: p.k[m.key], goal: m.monthly ? g[m.key] * pace : g[m.key] }))
    .map(x => ({ ...x, ratio: x.value / x.goal })).filter(x => x.ratio >= 1).sort((a, b) => b.ratio - a.ratio);
  const strength = strengths[0] || null;
  const min = minSphFor(goals, store);
  const nm = k => PLAIN[k] || k;
  const lines = [];
  lines.push(strength ? `Open with a win. ${first}'s ${nm(strength.key)} is ${fmtK(strength.key, strength.value)}, above our ${fmtK(strength.key, strength.goal)} goal. Tell them.` : `Open with a win. Find one thing ${first} did well on the floor today and say it out loud.`);
  if (why === 'below' || p.k.sph < min) lines.push(`${first}'s sales per hour is $${Math.round(p.k.sph)}. Our minimum is $${min}. Sit down with the store leader and write a plan today.`);
  if (wk?.priorSph && wk.sph != null && wk.hours >= 12 && wk.sph < wk.priorSph * 0.75) lines.push(`${first} is at $${Math.round(wk.sph)} an hour this week, down from $${Math.round(wk.priorSph)}. Ask what's going on before you talk numbers.`);
  if (why === 'model' && teamFocus) lines.push(`Call ${first} out in the huddle. Have ${first} show the team how they handle ${teamFocus.toLowerCase()}.`);
  items.filter(f => !f.stretch).forEach((f, i) => lines.push(`${i === 0 ? 'Coach' : 'Then'} ${nm(f.key)}: ${fmtK(f.key, f.value)} now, goal ${fmtK(f.key, f.goal)} (${f.coach.pillar}). Ask: "${f.coach.ask[0]}" Then show them: ${f.coach.doThis}`));
  const stretch = items.find(f => f.stretch);
  if (stretch) lines.push(`${first} is at goal across the board. Push ${nm(stretch.key)} from ${fmtK(stretch.key, stretch.value)} to ${fmtK(stretch.key, stretch.target)}.`);
  const main = items.find(f => !f.stretch);
  const drillKey = (main || items[0])?.key || 'sph';
  const drill = drillFor(drillKey);
  lines.push(`Practice it standing up: ${drill.title}. You're the guest. ${drill.guest} Let ${first} run it, give one tip, then run it again.`);
  // The commitment is a behavior with a number on it: what they'll do, and where the number goes.
  const behavior = main ? main.coach.doThis.split('. ')[0].replace(/\.$/, '') : null;
  const commit = main ? { key: main.key, what: `${first}: ${behavior}`, from: fmtK(main.key, main.value), to: fmtK(main.key, main.target ?? main.goal),
    how: `Practice "${drill.title}" standing up today, then do it with every guest. Store leader watches for it on the floor.` } : null;
  lines.push(commit ? `${first}'s commitment: ${behavior.replace(/^./, c => c.toLowerCase())}. ${nm(main.key).replace(/^./, c => c.toUpperCase())} from ${commit.from} to ${commit.to} by the next visit. We inspect it then.` : `${first}'s commitment this week: teach one teammate what they do best.`);
  return { items, strength, drillKey, drill, commit, text: lines.join('\n') };
}
export { COACHING, METRICS, STORE_METRICS, fmt, isOutlet, slug };

// ---------------------------------------------------------------- consultant trends (YTD, last month, MTD, this week)
// The RSA report gives ratios. Turn them back into running totals so periods can be added and subtracted
// (this week = today's month-to-date minus last Saturday's), then back into ratios.
const PCT_SUMS = [['fin', 'financePct'], ['bed', 'beddingPct'], ['prot', 'protectionPct'], ['del', 'deliveryPct'], ['mar', 'effMargin'], ['can', 'cancelPct'], ['dis', 'discountPct']];
export function toSums(p) {
  if (!p?.k) return null;
  const k = p.k, sales = k.netSales ?? 0, hours = p.hours ?? (k.sph ? sales / k.sph : 0);
  const s = { sales, hours, apps: k.creditApps ?? null, tickets: k.avgTicket ? sales / k.avgTicket : null };
  PCT_SUMS.forEach(([a, m]) => { s[a] = k[m] != null ? k[m] / 100 * sales : null; });
  return s;
}
const comb = (a, b, sign) => { if (!a) return b && sign > 0 ? { ...b } : null; if (!b) return { ...a }; const o = {}; for (const x of Object.keys(a)) o[x] = a[x] != null && b[x] != null ? a[x] + sign * b[x] : null; return o; };
export const addSums = (a, b) => comb(a, b, 1);
export const subSums = (a, b) => comb(a, b, -1);
export function fromSums(s) {
  if (!s || !(s.hours >= 1)) return null;
  const sales = s.sales, ok = sales > 0, pc = v => ok && v != null ? v / sales * 100 : null;
  const k = { netSales: sales, sph: sales / s.hours, avgTicket: s.tickets > 0.5 ? sales / s.tickets : null, appsPer40: s.apps != null ? s.apps / s.hours * 40 : null };
  PCT_SUMS.forEach(([a, m]) => { k[m] = pc(s[a]); });
  k.beddingSph = k.beddingPct != null ? k.sph * k.beddingPct / 100 : null;
  k.protectionSph = k.protectionPct != null ? k.sph * k.protectionPct / 100 : null;
  return { k, hours: s.hours };
}
const monthOf = d => d ? d.slice(0, 7) : '';
// src: { mtd, lastMonth, ytd, ytdSnap, monthEnds, wk: [{ plus, minus }], day: [{ plus, minus }] }
// Each piece is an RSA copy (people list); a period made of pieces is the sum of (plus minus minus).
export function consultantTrends(src) {
  const by = r => new Map((r?.people || []).map(p => [p.cid, p]));
  const M = by(src.mtd), L = by(src.lastMonth), Y = by(src.ytd), YS = by(src.ytdSnap);
  const ends = (src.monthEnds || []).map(by);
  const pieces = list => (list || []).map(x => ({ plus: by(x.plus), minus: x.minus ? by(x.minus) : null }));
  const WK = pieces(src.wk), DY = pieces(src.day);
  const sumPieces = (ps, cid) => {
    if (!ps.length) return null;
    let tot = null;
    for (const x of ps) {
      const a = toSums(x.plus.get(cid)); if (!a) continue;
      const part = x.minus ? (x.minus.get(cid) ? subSums(a, toSums(x.minus.get(cid))) : a) : a;
      tot = tot ? addSums(tot, part) : part;
    }
    return tot;
  };
  const mtdMonth = monthOf(src.mtd?.to);
  const out = {};
  const cids = new Set([...M.keys(), ...Y.keys(), ...L.keys(), ...WK.flatMap(x => [...x.plus.keys()])]);
  for (const cid of cids) {
    const m = toSums(M.get(cid));
    let y = null;
    if (src.ytd) {
      const ys = toSums(Y.get(cid));
      if (!m || !src.mtd || src.ytd.to >= src.mtd.to) y = ys;
      else if (monthOf(src.ytd.to) === mtdMonth) { const sn = toSums(YS.get(cid)); y = src.ytdSnap && sn ? addSums(ys, subSums(m, sn)) : ys; }
      else y = ys ? addSums(ys, m) : m;
    } else if (ends.length) {
      y = m; ends.forEach(e => { const v = toSums(e.get(cid)); y = y ? addSums(y, v) : v; });
    }
    out[cid] = { ytd: fromSums(y), lastMonth: fromSums(toSums(L.get(cid))), mtd: fromSums(m), wtd: fromSums(sumPieces(WK, cid)), day: fromSums(sumPieces(DY, cid)) };
  }
  return out;
}
// Rows shown on a consultant card. better: which way is good.
export const TREND_ROWS = [
  { key: 'netSales', label: 'Revenue', say: 'revenue', fmt: 'money', total: true },
  { key: 'sph', label: 'SPH', say: 'SPH', fmt: 'money' }, { key: 'avgTicket', label: 'Avg ticket', say: 'average ticket', fmt: 'money' },
  { key: 'financePct', label: 'Finance', say: 'finance', fmt: 'pct' }, { key: 'appsPer40', label: 'Apps / 40 hrs', say: 'credit apps', fmt: 'num' },
  { key: 'beddingPct', label: 'Bedding', say: 'bedding', fmt: 'pct' }, { key: 'protectionPct', label: 'Protection', say: 'protection', fmt: 'pct' },
  { key: 'deliveryPct', label: 'Delivery', say: 'delivery', fmt: 'pct' }, { key: 'effMargin', label: 'Margin', say: 'margin', fmt: 'pct' },
  { key: 'cancelPct', label: 'Cancel', say: 'cancellations', fmt: 'pct', lower: true }
];
export const trendFmt = (r, v) => v == null ? '--' : r.fmt === 'money' ? '$' + Math.round(v).toLocaleString('en-US') : r.fmt === 'pct' ? v.toFixed(1) + '%' : v.toFixed(1);
// Recent (this week if 8+ hours, else this month) against the longer baseline (YTD, else last month).
// Returns a direction per row and the biggest win and slip, so the leader can open with a question.
export function trendRead(t) {
  if (!t) return { rows: {}, up: null, down: null };
  const useWk = t.wtd?.hours >= 8;
  const baseP = t.ytd ? 'ytd' : t.lastMonth ? 'lastMonth' : null;
  // Most recent period with enough to go on: this week, then this month, then last month (early in a month).
  const fallback = t.mtd?.hours >= 8 ? 'mtd' : baseP === 'ytd' && t.lastMonth ? 'lastMonth' : t.mtd ? 'mtd' : null;
  const recentP = useWk ? 'wtd' : fallback;
  const rows = {}; let up = null, down = null;
  if (!baseP || !recentP || recentP === baseP) return { rows, up, down, recentP, baseP };
  for (const r of TREND_ROWS) {
    if (r.total) continue; // a total depends on how long the period is, so it gets no arrow
    const per = useWk && t.wtd.k[r.key] != null ? 'wtd' : fallback;
    const a = t[baseP].k[r.key], b = t[per]?.k?.[r.key];
    if (a == null || b == null || !a) continue;
    let ch = (b - a) / Math.abs(a);
    if (r.fmt === 'pct') ch = (b - a) / Math.max(Math.abs(a), 5);
    const good = r.lower ? ch < 0 : ch > 0, big = Math.abs(ch) >= 0.1;
    rows[r.key] = big ? (good ? 'up' : 'down') : 'flat';
    const score = Math.abs(ch);
    if (big && good && r.key !== 'appsPer40' && (!up || score > up.score)) up = { ...r, score, from: a, to: b, per };
    if (big && !good && r.key !== 'appsPer40' && (!down || score > down.score)) down = { ...r, score, from: a, to: b, per };
  }
  return { rows, up, down, recentP, baseP };
}
export const TREND_LABEL = { ytd: 'this year', lastMonth: 'last month', mtd: 'this month', wtd: 'this week', day: 'yesterday' };
