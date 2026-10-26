/* Luck Swan Sales — Dashboard (admins only): revenue and pipeline summaries with charts */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  // validated categorical set (CVD-safe, direct-labelled because the amber is low contrast)
  const C = { teal: '#0E9F8E', red: '#B8333A', blue: '#4A62B8', amber: '#D99A00', gray: '#8D928F', ink: '#17211F', grid: '#EFEDE6', axis: '#D9D6CC', muted: '#6B716E' };
  const SEC = [['meet', 'ห้องประชุม', C.teal], ['food', 'อาหารและเครื่องดื่ม', C.red], ['room', 'ห้องพัก', C.blue], ['other', 'บริการอื่น ๆ', C.amber]];
  const ST = [['draft', 'ร่าง', C.gray], ['sent', 'ส่งแล้ว', C.blue], ['tentative', 'จองชั่วคราว', C.amber], ['confirmed', 'ยืนยันแล้ว', C.teal]];
  const PENDING = ['sent', 'tentative'];

  const short = v => { const x = Math.abs(v); return x >= 1e6 ? (v / 1e6).toFixed(x >= 1e7 ? 0 : 1) + 'M' : x >= 1e3 ? Math.round(v / 1e3) + 'K' : String(Math.round(v)); };
  const baht = v => '฿' + U.n(Math.round(v));

  /* ---- periods: weeks start Monday; revenue counts on the event's start date ---- */
  function weekStart(iso) { const d = U.d(iso), dow = (d.getDay() + 6) % 7; return U.addDays(iso, -dow); }
  // selectable years: from 2569 (2026, when the system started) to this year, or later once a confirmed event falls in a later year
  const FIRST_YEAR = 2026;
  function yearSpan() {
    const now = U.d(U.today()).getFullYear();
    const last = Store.db.quotes.filter(q => q.status === 'confirmed' && q.start).reduce((m, q) => Math.max(m, Number(q.start.slice(0, 4)) || 0), now);
    return { min: Math.min(FIRST_YEAR, now), max: Math.max(now, last) };
  }
  const UNIT = { day: 'วัน', week: 'สัปดาห์', month: 'เดือน' };
  // selected range: a calendar year, or one month of it
  function range(ui) {
    if (ui.month == null) return { a: ui.year + '-01-01', b: ui.year + '-12-31', label: 'ปี ' + U.be(ui.year) };
    const m = new Date(ui.year, ui.month, 1), e = new Date(ui.year, ui.month + 1, 0);
    return { a: U.iso(m), b: U.iso(e), label: U.ML[ui.month] + ' ' + U.be(ui.year) };
  }
  function prevRange(ui) { return range(ui.month == null ? { year: ui.year - 1, month: null } : { year: ui.month ? ui.year : ui.year - 1, month: ui.month ? ui.month - 1 : 11 }); }
  // year → by month or week; month → by day or week
  const modesFor = ui => ui.month == null ? ['week', 'month'] : ['day', 'week'];
  function periods(ui) {
    const t = U.today(), r = range(ui), out = [];
    const push = (a, b, label, tick) => out.push({ a, b, label, tick, cur: t >= a && t <= b });
    if (ui.mode === 'month') {
      for (let m = 0; m < 12; m++) { const x = new Date(ui.year, m, 1); push(U.iso(x), U.iso(new Date(ui.year, m + 1, 0)), U.MS[m] + ' ' + U.be(ui.year), U.MS[m]); }
    } else if (ui.mode === 'day') {
      for (const d of U.daysOf(r.a, r.b)) { const x = U.d(d); push(d, d, U.fmtD(d, { dow: true }), String(x.getDate())); }
    } else {
      // Monday-start weeks, clipped to the selected range
      for (let a = weekStart(r.a); a <= r.b; a = U.addDays(a, 7)) {
        const s0 = a < r.a ? r.a : a, e0 = U.addDays(a, 6) > r.b ? r.b : U.addDays(a, 6), ds = U.d(s0), de = U.d(e0);
        push(s0, e0, ds.getDate() + (ds.getMonth() !== de.getMonth() ? ' ' + U.MS[ds.getMonth()] : '') + '–' + de.getDate() + ' ' + U.MS[de.getMonth()], ds.getDate() + ' ' + U.MS[ds.getMonth()]);
      }
    }
    return out;
  }

  /* ---- chart pieces (inline SVG, one y-axis, recessive grid, hover tips via data-tip) ---- */
  function niceMax(v) { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))), n = v / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p; }
  function lineChart(ps, series) {
    const W = 1100, Hh = 300, L = 52, R = 20, T = 22, B = 34, iw = W - L - R, ih = Hh - T - B;
    const max = niceMax(Math.max(1, ...series.flatMap(s => s.vals)));
    const x = i => L + (ps.length === 1 ? iw / 2 : i * iw / (ps.length - 1)), y = v => T + ih - v / max * ih;
    const grid = [0, 0.25, 0.5, 0.75, 1].map(f => `<line x1="${L}" x2="${W - R}" y1="${y(max * f)}" y2="${y(max * f)}" stroke="${f ? C.grid : C.axis}"/><text x="${L - 8}" y="${y(max * f) + 4}" text-anchor="end">${short(max * f)}</text>`).join('');
    const every = Math.ceil(ps.length / 8);
    const ticks = ps.map((p, i) => i % every && i !== ps.length - 1 ? '' : `<text x="${x(i)}" y="${Hh - 10}" text-anchor="middle" ${p.cur ? `fill="${C.ink}" font-weight="600"` : ''}>${H(p.tick)}</text>`).join('');
    const cur = ps.findIndex(p => p.cur);
    const now = cur >= 0 ? `<line x1="${x(cur)}" x2="${x(cur)}" y1="${T}" y2="${T + ih}" stroke="${C.axis}" stroke-dasharray="3 4"/><text x="${x(cur)}" y="${T - 6}" text-anchor="middle" fill="${C.muted}">ตอนนี้</text>` : '';
    const lines = series.map(s => `<polyline fill="none" stroke="${s.color}" stroke-width="2" ${s.dash ? 'stroke-dasharray="6 5"' : ''} stroke-linejoin="round" stroke-linecap="round" points="${s.vals.map((v, i) => x(i) + ',' + y(v)).join(' ')}"/>` +
      s.vals.map((v, i) => `<circle cx="${x(i)}" cy="${y(v)}" r="4" fill="${s.dash ? '#fff' : s.color}" stroke="${s.color}" stroke-width="2"/>`).join('')).join('');
    // direct label at the last point of each series
    const ends = series.map(s => { const i = s.vals.length - 1; return !s.vals[i] ? '' : `<text x="${x(i) - 6}" y="${y(s.vals[i]) - 10}" text-anchor="end" fill="${C.ink}" font-weight="600">${short(s.vals[i])}</text>`; }).join('');
    const step = iw / Math.max(1, ps.length - 1);
    const hits = ps.map((p, i) => `<rect x="${x(i) - step / 2}" y="${T}" width="${step}" height="${ih}" fill="transparent" data-tip="${H(p.label + '\n' + series.map(s => s.name + ': ' + baht(s.vals[i])).join('\n'))}" data-cx="${x(i)}"/>`).join('');
    return `<svg class="dash-svg" viewBox="0 0 ${W} ${Hh}" role="img" aria-label="${H(series.map(s => s.name + ' ' + ps.map((p, i) => p.label + ' ' + baht(s.vals[i])).join(', ')).join(' · '))}"><g font-size="11" fill="${C.muted}">${grid}${ticks}${now}${lines}${ends}</g><line class="dash-cross" x1="0" x2="0" y1="${T}" y2="${T + ih}" stroke="${C.ink}" stroke-width="1" opacity="0"/>${hits}</svg>`;
  }
  function barChart(ps, vals, color) {
    const W = 760, Hh = 240, L = 48, R = 16, T = 18, B = 34, iw = W - L - R, ih = Hh - T - B;
    const max = niceMax(Math.max(1, ...vals)), slot = iw / ps.length, bw = Math.max(6, Math.min(40, slot - 8));
    const y = v => T + ih - v / max * ih;
    const grid = [0, 0.5, 1].map(f => `<line x1="${L}" x2="${W - R}" y1="${y(max * f)}" y2="${y(max * f)}" stroke="${f ? C.grid : C.axis}"/><text x="${L - 8}" y="${y(max * f) + 4}" text-anchor="end">${short(max * f)}</text>`).join('');
    const every = Math.ceil(ps.length / 8);
    const bars = ps.map((p, i) => {
      const cx = L + slot * i + slot / 2, v = vals[i], h = Math.max(v ? 2 : 0, ih - (y(v) - T));
      return `<path d="M${cx - bw / 2},${T + ih} v${-Math.max(0, h - 4)} q0,-4 4,-4 h${bw - 8} q4,0 4,4 v${Math.max(0, h - 4)} z" fill="${p.cur ? C.amber : color}" ${v ? '' : 'opacity="0"'}/>` +
        (v && (ps.length <= 12 || i % every === 0 || p.cur) ? `<text x="${cx}" y="${y(v) - 6}" text-anchor="middle" fill="${C.ink}" font-weight="600">${short(v)}</text>` : '') +
        (i % every && i !== ps.length - 1 ? '' : `<text x="${cx}" y="${Hh - 10}" text-anchor="middle" ${p.cur ? `fill="${C.ink}" font-weight="600"` : ''}>${H(p.tick)}</text>`) +
        `<rect x="${L + slot * i}" y="${T}" width="${slot}" height="${ih}" fill="transparent" data-tip="${H(p.label + '\n' + baht(v))}"/>`;
    }).join('');
    return `<svg class="dash-svg" viewBox="0 0 ${W} ${Hh}" role="img" aria-label="${H(ps.map((p, i) => p.label + ' ' + baht(vals[i])).join(', '))}"><g font-size="11" fill="${C.muted}">${grid}${bars}</g></svg>`;
  }
  function donut(parts, center, sub) {
    const tot = parts.reduce((a, p) => a + p.v, 0), r = 70, sw = 26, cx = 90, cy = 90, circ = 2 * Math.PI * r;
    let off = 0;
    const arcs = tot ? parts.filter(p => p.v).map(p => {
      const len = p.v / tot * circ, gap = parts.filter(q => q.v).length > 1 ? 2 : 0;
      const s = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${p.color}" stroke-width="${sw}" stroke-dasharray="${Math.max(0, len - gap)} ${circ}" stroke-dashoffset="${-off}" transform="rotate(-90 ${cx} ${cy})" data-tip="${H(p.name + '\n' + p.tip)}"/>`;
      off += len; return s;
    }).join('') : `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${C.grid}" stroke-width="${sw}"/>`;
    return `<svg class="dash-donut" viewBox="0 0 180 180" role="img" aria-label="${H(parts.map(p => p.name + ' ' + p.tip).join(', '))}">${arcs}<text x="${cx}" y="${cy}" text-anchor="middle" font-size="26" font-weight="700" fill="${C.ink}" font-family="Bai Jamjuree, sans-serif">${H(center)}</text><text x="${cx}" y="${cy + 20}" text-anchor="middle" font-size="11" fill="${C.muted}">${H(sub)}</text></svg>`;
  }
  const legend = parts => `<div class="dash-legend">${parts.map(p => `<div><i style="background:${p.color}"></i><span>${H(p.name)}</span><b>${H(p.tip)}</b></div>`).join('')}</div>`;
  const hbars = rows => { const max = Math.max(1, ...rows.map(r => r.v)); return rows.map(r => `<div class="dash-hbar" data-tip="${H(r.name + '\n' + baht(r.v))}"><span class="n">${H(r.name)}</span><span class="t"><i style="width:${r.v / max * 100}%;background:${r.color}"></i></span><b class="num">${baht(r.v)}</b></div>`).join(''); };

  App.V.dash = {
    title: 'Dashboard',
    render() {
      const me = App.me();
      if (!me || me.role !== 'admin') return '<div class="card pad" style="max-width:520px;margin:40px auto;text-align:center"><h2>เฉพาะผู้ดูแลระบบ</h2><p class="muted">หน้า Dashboard เปิดดูได้เฉพาะบทบาทผู้ดูแลระบบ</p><a class="btn ghost" href="#/home">กลับหน้าหลัก</a></div>';
      const ui = App.state.ui.dash || (App.state.ui.dash = { year: U.d(U.today()).getFullYear(), month: null, mode: 'month', table: false });
      const ys = yearSpan(); ui.year = Math.min(ys.max, Math.max(ys.min, ui.year));
      if (!modesFor(ui).includes(ui.mode)) ui.mode = ui.month == null ? 'month' : 'week';
      const db = Store.db, ps = periods(ui), rg = range(ui), pr = prevRange(ui), from = rg.a, to = rg.b;
      const live = db.quotes.filter(q => q.status !== 'cancelled');
      const inRange = q => q.start >= from && q.start <= to;
      const idx = q => ps.findIndex(p => q.start >= p.a && q.start <= p.b);
      const conf = ps.map(() => 0), pend = ps.map(() => 0);
      for (const q of live) { const i = idx(q); if (i < 0) continue; const n = App.calc(q).net; if (q.status === 'confirmed') conf[i] += n; else if (PENDING.includes(q.status)) pend[i] += n; }
      const cq = live.filter(q => q.status === 'confirmed' && inRange(q));
      const total = conf.reduce((a, v) => a + v, 0), pipe = pend.reduce((a, v) => a + v, 0);
      const pax = cq.reduce((a, q) => a + (Number(q.pax) || 0), 0);
      const prevV = live.filter(q => q.status === 'confirmed' && q.start >= pr.a && q.start <= pr.b).reduce((a, q) => a + App.calc(q).net, 0);
      const delta = prevV ? Math.round((total - prevV) / prevV * 100) : null;
      const unit = UNIT[ui.mode], thisY = U.d(U.today()).getFullYear(), thisM = U.d(U.today()).getMonth();
      // revenue by section (before discount, service charge and VAT)
      const bySec = SEC.map(([k, name, color]) => ({ name, color, v: cq.reduce((a, q) => a + (App.calc(q).by[k] || 0), 0) }));
      // quote status mix (count + value) for events in range
      const stParts = ST.map(([k, name, color]) => { const qs = db.quotes.filter(q => q.status === k && inRange(q)); return { name, color, v: qs.length, tip: qs.length + ' ใบ · ' + baht(qs.reduce((a, q) => a + App.calc(q).net, 0)) }; });
      const cancelled = db.quotes.filter(q => q.status === 'cancelled' && inRange(q)).length;
      const sent = db.quotes.filter(q => q.status !== 'draft' && inRange(q)), won = sent.filter(q => q.status === 'confirmed');
      // by sales person
      const people = {};
      for (const q of cq) people[q.owner] = (people[q.owner] || 0) + App.calc(q).net;
      const team = Object.entries(people).map(([id, v]) => ({ name: (db.users.find(u => u.id === id) || {}).name || 'ไม่ระบุ', v, color: C.teal })).sort((a, b) => b.v - a.v);
      const rangeTxt = rg.label + ' (' + U.fmtD(from, { year: false }) + ' – ' + U.fmtD(to) + ')';
      return `
<div class="page-hd"><div><h1>Dashboard</h1><div class="small muted">ช่วง ${H(rangeTxt)} · นับรายรับตามวันจัดงาน · เฉพาะผู้ดูแลระบบ</div></div>
  <div class="row wrap"><div class="seg white" role="group" aria-label="แสดงเป็น">${modesFor(ui).map(m => `<button class="${ui.mode === m ? 'on' : ''}" data-act="dashMode" data-v="${m}" aria-pressed="${ui.mode === m}">ราย${UNIT[m]}</button>`).join('')}</div>
  <button class="btn ghost" data-act="dashCsv">ดาวน์โหลด CSV</button></div></div>
<section class="card dash-pick" aria-label="เลือกช่วงเวลา">
  <div class="dash-year"><button class="icon-btn" data-act="dashYear" data-v="-1" aria-label="ปีก่อน" ${ui.year <= ys.min ? 'disabled' : ''}>${App.ICON.left}</button><b>ปี ${U.be(ui.year)}</b><button class="icon-btn" data-act="dashYear" data-v="1" aria-label="ปีถัดไป" ${ui.year >= ys.max ? 'disabled' : ''}>${App.ICON.right}</button></div>
  <div class="dash-months" role="group" aria-label="เลือกเดือน">
    <button class="chip ${ui.month == null ? 'on' : ''}" data-act="dashMonth" data-v="" aria-pressed="${ui.month == null}">ทั้งปี</button>
    ${U.MS.map((m, i) => `<button class="chip ${ui.month === i ? 'on' : ''} ${ui.year === thisY && i === thisM ? 'now' : ''}" data-act="dashMonth" data-v="${i}" aria-pressed="${ui.month === i}">${m}</button>`).join('')}
  </div>
  <button class="btn ghost sm dash-reset" data-act="dashReset" title="กลับไปเดือนปัจจุบัน"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/></svg>รีเซ็ต</button>
</section>
<div class="grid g4" style="margin-bottom:16px">
  <div class="card kpi" style="background:var(--ink);color:#fff;border-color:var(--ink)"><div class="small" style="color:#C9CFCC">รายรับยืนยันทั้งช่วง</div><div class="v" style="color:var(--yellow)">${baht(total)}</div><div class="small" style="color:#C9CFCC">${cq.length} งาน · ${U.n(pax)} ท่าน</div></div>
  <div class="card kpi"><div class="small muted">เทียบกับ${H(pr.label)}</div><div class="v ${delta == null ? '' : delta >= 0 ? 'teal' : 'red'}">${delta == null ? '—' : (delta >= 0 ? '▲ ' : '▼ ') + Math.abs(delta) + '<small>%</small>'}</div><div class="small muted">${H(pr.label)} ${baht(prevV)}${delta == null ? ' · ไม่มีงานยืนยันให้เทียบ' : ''}</div></div>
  <div class="card kpi"><div class="small muted">รอยืนยัน (ส่งแล้ว + จองชั่วคราว)</div><div class="v" style="color:var(--yellow-d)">${baht(pipe)}</div><div class="small muted">ถ้าปิดได้ทั้งหมด รวม ${baht(total + pipe)}</div></div>
  <div class="card kpi"><div class="small muted">ปิดการขายได้</div><div class="v">${sent.length ? Math.round(won.length / sent.length * 100) : 0}<small>%</small></div><div class="small muted">ยืนยัน ${won.length} จาก ${sent.length} ใบที่ส่ง · เฉลี่ย ${baht(cq.length ? total / cq.length : 0)}/งาน</div></div>
</div>
<section class="card" style="margin-bottom:16px;overflow:hidden">
  <div class="card-hd"><div><h2>รายรับราย${unit} · ${H(rg.label)}</h2><div class="small muted">ชี้หรือแตะที่กราฟเพื่อดูตัวเลข</div></div>
    <div class="row wrap" style="gap:14px"><span class="dash-key"><i style="background:${C.teal}"></i>ยืนยันแล้ว</span><span class="dash-key"><i class="dash" style="border-color:${C.amber}"></i>รอยืนยัน</span><button class="link" data-act="dashTable">${ui.table ? 'ดูเป็นกราฟ' : 'ดูเป็นตาราง'}</button></div></div>
  <div class="pad dash-chart">${ui.table ? `<div style="overflow-x:auto"><table class="tbl"><thead><tr><th>${unit}</th><th class="r">ยืนยันแล้ว (฿)</th><th class="r">รอยืนยัน (฿)</th></tr></thead><tbody>${ps.map((p, i) => `<tr${p.cur ? ' style="font-weight:600"' : ''}><td>${H(p.label)}${p.cur ? ' (วันนี้อยู่ในช่วงนี้)' : ''}</td><td class="r num">${U.n(Math.round(conf[i]))}</td><td class="r num">${U.n(Math.round(pend[i]))}</td></tr>`).join('')}</tbody></table></div>`
      : lineChart(ps, [{ name: 'ยืนยันแล้ว', vals: conf, color: C.teal }, { name: 'รอยืนยัน', vals: pend, color: C.amber, dash: true }])}</div>
</section>
<div class="split" style="margin-bottom:16px">
  <section class="card" style="overflow:hidden;min-width:0">
    <div class="card-hd"><div><h2>รายรับยืนยัน แยกตาม${unit}</h2><div class="small muted">แท่งสีเหลือง = ${unit}ปัจจุบัน</div></div></div>
    <div class="pad dash-chart">${barChart(ps, conf, C.teal)}</div>
  </section>
  <section class="card pad" style="min-width:0">
    <h2>สถานะใบเสนอราคา</h2><div class="small muted" style="margin-bottom:10px">งานในช่วงนี้${cancelled ? ' · ยกเลิก ' + cancelled + ' ใบ (ไม่นับ)' : ''}</div>
    <div class="dash-pie">${donut(stParts, String(stParts.reduce((a, p) => a + p.v, 0)), 'ใบ')}${legend(stParts)}</div>
  </section>
</div>
<div class="grid g2">
  <section class="card pad" style="min-width:0">
    <h2>รายรับแยกหมวด</h2><div class="small muted" style="margin-bottom:12px">งานที่ยืนยันในช่วงนี้ · ก่อนส่วนลด ค่าบริการ และ VAT</div>
    <div class="dash-pie">${donut(bySec.map(s => ({ ...s, tip: baht(s.v) })), short(bySec.reduce((a, s) => a + s.v, 0)), 'บาท')}${legend(bySec.map(s => ({ ...s, tip: baht(s.v) })))}</div>
  </section>
  <section class="card pad" style="min-width:0">
    <h2>ยอดตามพนักงานขาย</h2><div class="small muted" style="margin-bottom:12px">รายรับยืนยันในช่วงนี้</div>
    ${team.length ? hbars(team) : '<div class="empty small">ยังไม่มีงานที่ยืนยันในช่วงนี้</div>'}
  </section>
</div>`;
    },
    after() {
      // one floating tooltip for every [data-tip] mark: hover on desktop, tap on touch
      const root = document.querySelector('main');
      if (!root) return;
      let tip = document.getElementById('dashTip');
      if (!tip) { tip = document.createElement('div'); tip.id = 'dashTip'; tip.className = 'dash-tip'; tip.setAttribute('role', 'status'); document.body.appendChild(tip); }
      const hide = () => { tip.style.opacity = '0'; document.querySelectorAll('.dash-cross').forEach(l => l.setAttribute('opacity', '0')); };
      const show = (el, ev) => {
        tip.textContent = el.dataset.tip; tip.style.opacity = '1';
        const x = Math.min(window.innerWidth - tip.offsetWidth - 8, ev.clientX + 14), y = Math.max(8, ev.clientY - tip.offsetHeight - 12);
        tip.style.left = x + 'px'; tip.style.top = y + 'px';
        const svg = el.closest('svg'), cross = svg && svg.querySelector('.dash-cross');
        if (cross && el.dataset.cx) { cross.setAttribute('x1', el.dataset.cx); cross.setAttribute('x2', el.dataset.cx); cross.setAttribute('opacity', '.35'); }
      };
      root.onpointermove = ev => { const el = ev.target.closest && ev.target.closest('[data-tip]'); if (el) show(el, ev); else hide(); };
      root.onpointerleave = hide;
      root.onpointerdown = ev => { const el = ev.target.closest && ev.target.closest('[data-tip]'); if (el) show(el, ev); };
    }
  };

  Object.assign(App.ACT, {
    dashMode(el) { App.state.ui.dash.mode = el.dataset.v; App.render(); },
    dashYear(el) { const ui = App.state.ui.dash, ys = yearSpan(); ui.year = Math.min(ys.max, Math.max(ys.min, ui.year + Number(el.dataset.v))); App.render(); },
    // back to the current year and month
    dashReset() { const ui = App.state.ui.dash, t = U.d(U.today()); ui.year = t.getFullYear(); ui.month = t.getMonth(); App.render(); },
    dashMonth(el) { const ui = App.state.ui.dash; ui.month = el.dataset.v === '' ? null : Number(el.dataset.v); App.render(); },
    dashTable() { App.state.ui.dash.table = !App.state.ui.dash.table; App.render(); },
    dashCsv() {
      const ui = App.state.ui.dash, ps = periods(ui), db = Store.db;
      const rows = ps.map(p => {
        const qs = db.quotes.filter(q => q.start >= p.a && q.start <= p.b);
        const sum = st => Math.round(qs.filter(q => st.includes(q.status)).reduce((a, q) => a + App.calc(q).net, 0));
        return [p.label, p.a, p.b, sum(['confirmed']), sum(PENDING), qs.filter(q => q.status === 'confirmed').length].join(',');
      });
      U.download('dashboard-' + range(ui).a.slice(0, ui.month == null ? 4 : 7) + '-' + ui.mode + '.csv', '﻿' + ['ช่วง,ตั้งแต่,ถึง,รายรับยืนยัน (บาท),รอยืนยัน (บาท),จำนวนงานยืนยัน'].concat(rows).join('\n'), 'text/csv');
    }
  });
})(window);
