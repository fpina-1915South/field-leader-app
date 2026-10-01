// Team messages: plain text the Market Leader copies into Teams, text or email.
// Pure functions. Written plain and direct, no em dashes, SPG always with cancellations, close with "Let's go."
import { DAY_LONG, DAY_NAMES, dow, fromIso, weekStartOf, STORE_GOALS, storeFocus, rsaPicks, visitSummary, draggers, helpers } from './ml.js?v=202610010723';

const money = v => (v < 0 ? '-$' : '$') + Math.abs(Math.round(v)).toLocaleString('en-US');
const p1 = v => (Math.round(v * 10) / 10).toString();
const md = s => { const d = fromIso(s); return `${d.getMonth() + 1}/${d.getDate()}`; };
const day = s => `${DAY_LONG[dow(s)]} ${md(s)}`;
const first = name => String(name || '').split(/[\s(]/)[0];
const titleName = n => /^[A-Z\s.'-]+$/.test(n) ? n.toLowerCase().replace(/(^|[\s'-])([a-z])/g, (a, b, c) => b + c.toUpperCase()) : n;
const vsBud = v => v == null ? '' : v >= 0 ? `${p1(v)}% ahead of budget` : `${p1(-v)}% behind budget`;
const vsLy = v => v == null ? '' : v >= 0 ? `up ${p1(v)}% vs LY` : `down ${p1(-v)}% vs LY`;
const bps = v => v == null ? '' : v >= 0 ? `${Math.round(v)} bps ahead of budget` : `${Math.round(-v)} bps behind budget`;
const bullets = lines => lines.filter(Boolean).map(l => `- ${l}`).join('\n');
const lines = text => String(text || '').split(/\n+/).map(s => s.trim().replace(/^[-*•]\s*/, '')).filter(Boolean);

// What the store is beating: goals it hits, shown as wins.
function wins(m) {
  const out = [];
  if (m.vsBud.netSales >= 0) out.push(`Sales are ${vsBud(m.vsBud.netSales)}`);
  if (m.vsLy.spg >= 0) out.push(`SPG with cancellations is ${vsLy(m.vsLy.spg)}`);
  if (m.vsBud.closeRate >= 0) out.push(`Close rate is ${bps(m.vsBud.closeRate)}`);
  if (m.k.financePct >= STORE_GOALS.financePct) out.push(`Finance is at ${p1(m.k.financePct)}% of sales (goal ${STORE_GOALS.financePct}%)`);
  if (m.k.protectionAttach >= STORE_GOALS.protectionAttach) out.push(`Protection attach is at ${p1(m.k.protectionAttach)}% (goal ${STORE_GOALS.protectionAttach}%)`);
  if (m.k.cancelPct != null && m.k.cancelPct <= STORE_GOALS.cancelPct) out.push(`Cancellations are at ${p1(m.k.cancelPct)}%, inside the ${STORE_GOALS.cancelPct}% goal`);
  return out.slice(0, 3);
}
const fmtVal = (f, v) => /Pct|pct|Rate|Attach|Margin/.test(f.key) || f.key === 'appsToTraffic' ? `${p1(v)}%` : money(v);

// ---------------------------------------------------------------- top performers
// Consultant call-outs from the RSA report. Week = this week's sales and hours (today's upload
// minus Saturday's). Month leaders need 40+ hours so a small sample doesn't win a category.
const CATS = [
  ['netSales', 'Sales', v => money(v)],
  ['financePct', 'Finance', v => `${p1(v)}% of sales`],
  ['protectionPct', 'Protection', v => `${p1(v)}% of sales`],
  ['beddingPct', 'Bedding', v => `${p1(v)}% of sales`],
  ['creditApps', 'Credit apps', v => `${Math.round(v)}`]
];
export function topPerformers({ people = [], weeks = {}, stores, period = 'week', n = 3, withStore = false }) {
  const list = people.filter(p => stores.includes(p.store) && p.k?.sph != null);
  const tag = p => withStore ? ` (${p.store})` : '';
  const out = [];
  if (period === 'week') {
    const wk = list.map(p => ({ p, w: weeks[p.cid] })).filter(x => x.w?.hours >= 12 && x.w.sph != null).sort((a, b) => b.w.sph - a.w.sph).slice(0, n);
    if (wk.length) {
      out.push('Top performers this week (sales per hour)');
      wk.forEach((x, i) => out.push(`${i + 1}. ${titleName(x.p.name)}${tag(x.p)}: $${Math.round(x.w.sph)} an hour, ${money(x.w.sales)} in ${Math.round(x.w.hours)} hours`));
      const jump = list.map(p => ({ p, w: weeks[p.cid] })).filter(x => x.w?.hours >= 12 && x.w.priorSph && x.w.sph > x.w.priorSph * 1.25)
        .sort((a, b) => b.w.sph / b.w.priorSph - a.w.sph / a.w.priorSph)[0];
      if (jump) out.push(`Biggest jump: ${titleName(jump.p.name)}${tag(jump.p)}, from $${Math.round(jump.w.priorSph)} to $${Math.round(jump.w.sph)} an hour.`);
    }
    return out;
  }
  const qual = list.filter(p => (p.hours || 0) >= 40);
  const top = qual.slice().sort((a, b) => b.k.sph - a.k.sph).slice(0, n);
  if (top.length) {
    out.push('Top performers this month (sales per hour)');
    top.forEach((p, i) => out.push(`${i + 1}. ${titleName(p.name)}${tag(p)}: $${Math.round(p.k.sph)} an hour`));
    const leaders = CATS.map(([k, label, f]) => {
      const best = qual.filter(p => p.k[k] != null).sort((a, b) => b.k[k] - a.k[k])[0];
      return best && best.k[k] > 0 ? `- ${label}: ${titleName(best.name)}${tag(best)}, ${f(best.k[k])}` : null;
    }).filter(Boolean);
    if (leaders.length) out.push('', 'Category leaders', ...leaders);
  }
  return out;
}

// 1. Weekly kickoff to one store's team.
export function kickoff({ store, snap, asOf, people = [], weeks = {}, goals, pace = 1, visitDays = [], pastDays = [], sender, weekStart, tops = true }) {
  if (!snap?.mtd) return `${store} team, no numbers for this store in the latest daily report yet.`;
  const m = snap.mtd, w = snap.wtd;
  const lastWeek = asOf < weekStart;
  const out = [`${store} team, here's our week.`, ''];
  out.push(`Where we are this month (through ${md(asOf)})`);
  out.push(bullets([
    m.k.netSales != null && `Sales: ${money(m.k.netSales)}${m.vsBud.netSales != null ? ', ' + vsBud(m.vsBud.netSales) : ''}`,
    m.k.spg != null && `SPG with cancellations: $${Math.round(m.k.spg)}${m.vsLy.spg != null ? ', ' + vsLy(m.vsLy.spg) : ''}`,
    m.k.closeRate != null && `Close rate: ${p1(m.k.closeRate)}%${m.vsBud.closeRate != null ? ', ' + bps(m.vsBud.closeRate) : ''}`,
    w?.k.netSales != null && `${lastWeek ? 'Last week' : 'This week so far'}: ${money(w.k.netSales)}${w.vsBud.netSales != null ? ', ' + vsBud(w.vsBud.netSales) : ''}`
  ]));
  const good = wins(m);
  if (good.length) out.push('', "What's working", bullets(good));
  const focus = storeFocus(snap);
  if (focus.length) {
    out.push('', `Our ${focus.length === 1 ? 'focus' : focus.length + ' focus items'} this week`);
    focus.forEach((f, i) => { const dr = draggers(people, store, f.key).map(d => titleName(d.name)); const hp = helpers(people, store, f.key, undefined, 1, 1).map(d => titleName(d.name)); out.push(`${i + 1}. ${f.label}: ${fmtVal(f, f.value)} (goal ${fmtVal(f, f.goal)}). ${f.coach.doThis}${hp.length ? ` ${hp[0]} is doing it well, so learn from them.` : ''}${dr.length ? ` Leaders: start with ${dr.join(', ')}.` : ''}`); });
  }
  const best = tops ? topPerformers({ people, weeks, stores: [store], period: 'month' }) : [];
  if (best.length) out.push('', ...best);
  out.push('');
  if (visitDays.length) out.push(`I'll be in the store ${visitDays.map(day).join(' and ')}. Let's go.`);
  else if (pastDays.length) out.push(`Thanks for having me ${pastDays.map(d => DAY_LONG[dow(d)]).join(' and ')}. Call me if you need me. Let's go.`);
  else out.push("I'll check in by phone this week. Let's go.");
  if (sender) out.push(first(sender));
  return out.join('\n');
}

// 2. Recap after a visit, to the store leader and team.
export function visitRecap({ visit, nextVisit, sender }) {
  const sm = visitSummary(visit);
  const out = [visit.remote ? `${visit.store}, thanks for the time on the call ${visit.date ? day(visit.date) : 'today'}.` : `${visit.store}, thanks for ${visit.date ? day(visit.date) : 'today'}.`];
  if (sm.win) out.push('', `Shout-out to ${sm.winName ? titleName(sm.winName) : 'the leadership team'}: ${sm.win}`);
  const work = lines(visit.working);
  if (work.length) out.push('', "What's working", bullets(work));
  if (sm.commitments.length) out.push('', 'Our commitments before my next visit', sm.commitments.map((c, i) => `${i + 1}. ${c}`).join('\n'));
  if (sm.fixes.length) out.push('', '6 Elements to tighten up', bullets(sm.fixes));
  if (sm.leaderCommit) out.push('', `${sm.leaderName ? titleName(sm.leaderName) : 'Store leader'} commits to: ${sm.leaderCommit}`);
  if (sm.support) out.push(`My part: ${sm.support}`);
  out.push('', nextVisit ? `I'll be back ${day(nextVisit)} and we'll start with these. Let's go.` : "I'll follow up on these at my next visit. Let's go.");
  if (sender) out.push(first(sender));
  return out.join('\n');
}

// 3. Market update to all the store leaders in the market.
export function marketUpdate({ stores, daily, scores, plan, sender, asOf, people = [], weeks = {}, tops = true }) {
  const rows = stores.map(s => ({ s, m: daily?.stores?.[s]?.mtd })).filter(r => r.m);
  if (!rows.length) return 'No numbers for these stores in the latest daily report yet.';
  rows.sort((a, b) => (b.m.vsBud.netSales ?? -999) - (a.m.vsBud.netSales ?? -999));
  const out = [`Team, here's where our market stands through ${day(asOf)}.`, '', 'Month to date, ranked by sales to budget'];
  rows.forEach((r, i) => out.push(`${i + 1}. ${r.s}: sales ${r.m.vsBud.netSales != null ? (r.m.vsBud.netSales >= 0 ? '+' : '') + p1(r.m.vsBud.netSales) + '%' : 'n/a'}, SPG w/ cancellations ${r.m.vsLy.spg != null ? (r.m.vsLy.spg >= 0 ? '+' : '') + p1(r.m.vsLy.spg) + '% vs LY' : 'n/a'}, close rate ${r.m.vsBud.closeRate != null ? (r.m.vsBud.closeRate >= 0 ? '+' : '') + Math.round(r.m.vsBud.closeRate) + ' bps' : 'n/a'}`));
  const lead = rows[0];
  out.push('', `Leading the way: ${lead.s}. ${wins(lead.m)[0] ? wins(lead.m)[0] + '.' : ''} Call them and ask what they're doing.`.replace(/\s+\./g, '.').trim());
  const help = stores.slice().sort((a, b) => (scores[b]?.score ?? 0) - (scores[a]?.score ?? 0))[0];
  if (help && help !== lead.s && scores[help]?.parts?.length) out.push(`Needs the most help: ${help}. ${scores[help].parts[0].text}.`);
  const best = tops ? topPerformers({ people, weeks, stores, period: 'month', n: 5, withStore: true }) : [];
  if (best.length) out.push('', ...best);
  if (plan?.days?.length) {
    const ws = plan.weekStart || weekStartOf(asOf);
    out.push('', `My visits this week (${DAY_NAMES[0]} ${md(ws)})`);
    plan.days.filter(d => d.store).forEach(d => out.push(`- ${DAY_NAMES[dow(d.date)]} ${md(d.date)}: ${d.store}`));
    if (plan.calls?.length) out.push(`- Phone check-ins: ${plan.calls.join(', ')}`);
  }
  out.push('', "Let's go.");
  if (sender) out.push(first(sender));
  return out.join('\n');
}

// ---------------------------------------------------------------- daily
// Uses the report's daily columns when the export has them, otherwise week to date.
const dayOrWeek = snap => snap?.day ? { p: snap.day, label: 'day' } : snap?.wtd ? { p: snap.wtd, label: 'week' } : null;
function whereToday(plan, todayIso) {
  const d = plan?.days?.find(x => x.date === todayIso);
  return d?.store || null;
}
// 4. Daily huddle note to one store.
export function dailyStore({ store, snap, asOf, plan, today, sender, people = [], weeks = {}, tops = true }) {
  if (!snap?.mtd) return `${store} team, no numbers for this store in the latest daily report yet.`;
  const dw = dayOrWeek(snap), m = snap.mtd, w = snap.wtd;
  const out = [];
  if (dw?.label === 'day') {
    const p = dw.p;
    out.push(`${store}, here's ${day(asOf)}.`);
    out.push(bullets([
      p.k.netSales != null && `Sales: ${money(p.k.netSales)}${p.vsBud.netSales != null ? ', ' + vsBud(p.vsBud.netSales) : ''}`,
      p.k.spg != null && `SPG with cancellations: $${Math.round(p.k.spg)}${p.vsLy.spg != null ? ', ' + vsLy(p.vsLy.spg) : ''}`,
      p.k.closeRate != null && `Close rate: ${p1(p.k.closeRate)}%${p.vsBud.closeRate != null ? ', ' + bps(p.vsBud.closeRate) : ''}`,
      p.k.traffic != null && `Traffic: ${Math.round(p.k.traffic)}${p.vsLy.traffic != null ? ', ' + vsLy(p.vsLy.traffic) : ''}`
    ]));
  } else out.push(`${store}, here's where we stand through ${day(asOf)}.`);
  out.push('', bullets([
    w?.k.netSales != null && `Week so far: ${money(w.k.netSales)}${w.vsBud.netSales != null ? ', ' + vsBud(w.vsBud.netSales) : ''}`,
    m.k.netSales != null && `Month: ${money(m.k.netSales)}${m.vsBud.netSales != null ? ', ' + vsBud(m.vsBud.netSales) : ''}`,
    m.k.spg != null && dw?.label !== 'day' && `SPG with cancellations this month: $${Math.round(m.k.spg)}${m.vsLy.spg != null ? ', ' + vsLy(m.vsLy.spg) : ''}`
  ]));
  const good = wins(dw?.label === 'day' ? dw.p : m);
  if (good.length) out.push('', `Win: ${good[0]}.`);
  // One thing today. Alternates between the store's two focus items through the week.
  const focus = storeFocus(snap);
  if (focus.length) {
    const f = focus[dow(today) % focus.length];
    out.push('', `Today's one thing: ${f.label} (${fmtVal(f, f.value)} this month, goal ${fmtVal(f, f.goal)}). ${f.coach.doThis}`);
    const dr = draggers(people, store, f.key).map(d => titleName(d.name));
    const hp = helpers(people, store, f.key, undefined, 1, 1).map(d => titleName(d.name));
    if (hp.length) out.push(`${hp[0]} is doing this well. Have them show the team at the huddle.`);
    if (dr.length) out.push(`Leaders, work with ${dr.join(' and ').replace(/ and (?=.* and )/, ', ')} on this today.`);
  }
  const best = tops ? topPerformers({ people, weeks, stores: [store], period: 'week' }) : [];
  if (best.length) out.push('', ...best);
  const at = whereToday(plan, today);
  out.push('', at === store ? "I'll see you today. Let's go." : at ? `I'm at ${at} today. Call me if you need me. Let's go.` : "Call me if you need me today. Let's go.");
  if (sender) out.push(first(sender));
  return out.join('\n');
}
// 5. Daily market recap to all the store leaders.
export function dailyMarket({ stores, daily, asOf, plan, today, sender, people = [], weeks = {}, tops = true }) {
  const rows = stores.map(s => ({ s, snap: daily?.stores?.[s] })).filter(r => r.snap?.mtd);
  if (!rows.length) return 'No numbers for these stores in the latest daily report yet.';
  const useDay = rows.some(r => r.snap.day);
  const per = r => (useDay ? r.snap.day : r.snap.wtd || r.snap.mtd);
  rows.sort((a, b) => (per(b).vsBud.netSales ?? -999) - (per(a).vsBud.netSales ?? -999));
  const out = [`Team, ${useDay ? `here's ${day(asOf)}` : `here's the week so far through ${day(asOf)}`}, ranked by sales to budget.`, ''];
  rows.forEach((r, i) => {
    const p = per(r);
    out.push(`${i + 1}. ${r.s}: ${p.k.netSales != null ? money(p.k.netSales) : 'n/a'}${p.vsBud.netSales != null ? ` (${p.vsBud.netSales >= 0 ? '+' : ''}${p1(p.vsBud.netSales)}%)` : ''}${p.k.spg != null ? `, SPG w/ cancellations $${Math.round(p.k.spg)}` : ''}${p.k.closeRate != null ? `, close ${p1(p.k.closeRate)}%` : ''}`);
  });
  const top = rows[0];
  out.push('', `Top of the board: ${top.s}. Nice work.`);
  const best = tops ? topPerformers({ people, weeks, stores, period: 'week', withStore: true }) : [];
  if (best.length) out.push('', ...best, '');
  const at = whereToday(plan, today);
  out.push(at ? `I'm at ${at} today.` : '');
  out.push("Let's go.");
  if (sender) out.push(first(sender));
  return out.filter((l, i, a) => !(l === '' && a[i - 1] === '')).join('\n');
}
