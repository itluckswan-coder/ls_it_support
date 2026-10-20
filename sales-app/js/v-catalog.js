/* Luck Swan Sales — catalog: food (1f), meeting rooms (1g), group rooms (1h) */
(function (g) {
  'use strict';
  const U = g.U, D = g.DATA, Store = g.Store, App = g.App, ICON = App.ICON;
  const H = U.esc;
  const TAGS = [['halal', 'ฮาลาล'], ['veg', 'มังสวิรัติ'], ['gf', 'ไม่มีกลูเตน']];

  /* =====================================================================
     FOOD / PACKAGES (1f)
     ===================================================================== */
  // price slider bounds follow the current (editable) menu list
  const priceBounds = () => {
    const prices = D.FOOD.filter(f => !/โต๊ะ|งาน/.test(f.unit)).map(f => Number(f.price) || 0);
    if (!prices.length) return [0, 100];
    const lo = Math.floor(Math.min(...prices) / 10) * 10, hi = Math.ceil(Math.max(...prices) / 10) * 10;
    return [lo, Math.max(hi, lo + 10)];
  };

  App.V.food = {
    title: 'จัดเลี้ยง',
    after() { bindGalleries(); },
    render() {
      const [PMIN, PMAX] = priceBounds();
      const ui = App.state.ui.food || (App.state.ui.food = { cat: 'all', lo: PMIN, hi: PMAX, tags: [], sort: 'rec' });
      // keep the filter fully open at the ends when prices change
      if (ui.lo <= ui._min || ui.lo < PMIN) ui.lo = PMIN;
      if (ui.hi >= ui._max || ui.hi > PMAX) ui.hi = PMAX;
      ui._min = PMIN; ui._max = PMAX;
      const editor = App.can('edit_prices');
      const q = App.ctxQuote();
      const inRange = f => /โต๊ะ|งาน/.test(f.unit) || (f.price >= ui.lo && f.price <= ui.hi);
      let list = D.FOOD.filter(f => (ui.cat === 'all' ? f.cat === 'package' || true : f.cat === ui.cat) && inRange(f) && ui.tags.every(t => (f.tags || []).includes(t)));
      if (ui.cat === 'all') list = list.sort((a, b) => (a.cat === 'package' ? 0 : 1) - (b.cat === 'package' ? 0 : 1));
      if (ui.sort === 'price') list = list.slice().sort((a, b) => a.price - b.price);
      else list = list.slice().sort((a, b) => (a.cat === 'package' ? 0 : 1) - (b.cat === 'package' ? 0 : 1) || (a.rec || 9) - (b.rec || 9));
      const count = c => D.FOOD.filter(f => c === 'all' || f.cat === c).length;
      const pct = v => (v - PMIN) / (PMAX - PMIN) * 100;
      return `
<div class="split left">
  <aside>
    <h1 style="margin-bottom:18px">จัดเลี้ยง</h1>
    <div class="small muted" style="margin-bottom:8px">ประเภท</div>
    <div class="cat-list">${D.FOOD_CATS.map(c => `<button class="${ui.cat === c.id ? 'on' : ''}" data-act="foodCat" data-v="${c.id}"><span>${c.name}</span><span class="n">${count(c.id)}</span></button>`).join('')}</div>
    <div class="small muted" style="margin:22px 0 8px">ราคา / ท่าน</div>
    <div class="range2"><div class="track"><i style="left:${pct(ui.lo)}%;right:${100 - pct(ui.hi)}%"></i></div>
      <input type="range" min="${PMIN}" max="${PMAX}" step="10" value="${ui.lo}" data-change="foodLo" aria-label="ราคาต่ำสุด"><input type="range" min="${PMIN}" max="${PMAX}" step="10" value="${ui.hi}" data-change="foodHi" aria-label="ราคาสูงสุด"></div>
    <div class="between small"><span>${U.n(ui.lo)} ฿</span><span>${U.n(ui.hi)} ฿</span></div>
    <div class="small muted" style="margin:22px 0 8px">เหมาะกับ</div>
    <div class="row wrap" style="gap:6px">${TAGS.map(([k, l]) => `<button class="chip ${ui.tags.includes(k) ? 'on-teal' : ''}" data-act="foodTag" data-v="${k}">${l}</button>`).join('')}</div>
  </aside>
  <div style="min-width:0">
    <div class="between" style="margin-bottom:16px;flex-wrap:wrap">${App.ctxBar(q)}
      <div class="row wrap">${editor ? `<button class="btn red sm" data-act="foodNew">+ เพิ่มเมนู / แพ็กเกจ</button>` : ''}<div class="seg white"><button class="${ui.sort === 'rec' ? 'on' : ''}" data-act="foodSort" data-v="rec">แนะนำ</button><button class="${ui.sort === 'price' ? 'on' : ''}" data-act="foodSort" data-v="price">ราคาต่ำ → สูง</button></div></div></div>
    <div class="pcards">${list.map(f => card(f, q)).join('') || '<div class="empty">ไม่พบรายการตามตัวกรอง</div>'}</div>
    <div style="height:90px"></div>
  </div>
</div>${App.dock(q)}`;
    }
  };
  // photo carousel: swipe the image, or tap the small dots under it
  function gallery(f) {
    const slides = D.foodSlides(f);
    const tag = f.best ? '<span class="tag">ขายดี</span>' : '';
    return `<div class="gal" data-gal>
      <div class="gal-track" tabindex="0" aria-label="ภาพเมนู ${H(f.name)}">${slides.map((sl, i) => `<figure class="gal-slide" aria-label="ภาพที่ ${i + 1} จาก ${slides.length}">
        ${sl.img ? `<img src="${H(sl.img)}" alt="${H(sl.cap)}" loading="lazy" draggable="false">` : `<div class="gal-ph"><span aria-hidden="true">🍽</span><b>${H(sl.cap)}</b><small>รอรูปเมนู</small></div>`}
        ${sl.img ? `<figcaption class="gal-cap">${H(sl.cap)}</figcaption>` : ''}</figure>`).join('')}</div>${tag}</div>
      ${slides.length > 1 ? `<div class="gal-dots" role="tablist">${slides.map((sl, i) => `<button role="tab" class="${i ? '' : 'on'}" data-act="galGo" data-i="${i}" aria-label="ดูภาพ ${H(sl.cap)}" aria-selected="${!i}"></button>`).join('')}</div>` : ''}`;
  }
  App.ACT.galGo = el => {
    const track = el.closest('.pcard').querySelector('.gal-track');
    track.scrollTo({ left: Number(el.dataset.i) * track.clientWidth, behavior: 'smooth' });
  };
  function bindGalleries() {
    document.querySelectorAll('.pcard .gal-track').forEach(track => {
      const dots = track.closest('.pcard').querySelectorAll('.gal-dots button');
      if (!dots.length) return;
      track.addEventListener('scroll', () => {
        const i = Math.round(track.scrollLeft / Math.max(1, track.clientWidth));
        dots.forEach((d, k) => { d.classList.toggle('on', k === i); d.setAttribute('aria-selected', String(k === i)); });
      }, { passive: true });
    });
  }

  function card(f, q) {
    const inQ = q && q.items.find(i => i.ref === f.id);
    const pkg = f.cat === 'package';
    return `<article class="card pcard">
      ${gallery(f)}
      <div class="bd">
        <div class="small b red">${H(f.sub || (D.FOOD_CATS.find(c => c.id === f.cat) || {}).name)}</div>
        <div class="between" style="align-items:flex-start;gap:8px"><h3 style="font-size:19px;margin:4px 0 0">${H(f.name)}</h3>${App.can('edit_prices') ? `<button class="btn ghost sm" style="height:30px;padding:0 12px;flex:none" data-act="foodEdit" data-v="${H(f.id)}">แก้ไข</button>` : ''}</div>
        <ul>${D.foodFeat(f).map(x => `<li>${H(x)}</li>`).join('')}</ul>
        <div class="price">${U.n(f.price)}.-${pkg ? ' Net' : ''} <small>/ ${H(f.unit)}${f.min ? ' · ขั้นต่ำ ' + f.min + ' ท่าน/วัน' : ''}</small></div>
        ${inQ ? `<button class="btn soft block" style="margin-top:12px;height:46px" data-act="foodRemove" data-v="${f.id}">${ICON.check} อยู่ในใบเสนอแล้ว · ${U.n(inQ.q1)} ${H(/โต๊ะ|งาน/.test(f.unit) ? f.unit : 'ท่าน')}</button>`
          : `<button class="btn ghost block" style="margin-top:12px;height:46px;border-color:var(--ink)" data-act="foodAdd" data-v="${f.id}">+ เพิ่มในใบเสนอ</button>`}
      </div></article>`;
  }
  function addFood(q, fid) {
    const f = D.FOOD.find(x => x.id === fid), days = U.daysOf(q.start, q.end).length;
    const times = f.cat === 'coffee' ? days * 2 : f.cat === 'lunch' ? days : (f.cat === 'package' || f.cat === 'wedding' || /โต๊ะ|งาน/.test(f.unit)) ? null : 1;
    const sched = [];
    if (f.cat === 'coffee') for (let d = 1; d <= days; d++) sched.push({ day: d, start: '10:00', end: '10:30' }, { day: d, start: '14:30', end: '15:00' });
    if (f.cat === 'lunch') for (let d = 1; d <= days; d++) sched.push({ day: d, start: '12:00', end: '13:00' });
    if (f.cat === 'gala') sched.push({ day: 1, start: '18:30', end: '21:30', place: f.place || '' });
    if (f.cat === 'cocktail') sched.push({ day: 1, start: '17:00', end: '19:00' });
    const qty = /โต๊ะ/.test(f.unit) ? Math.ceil(q.pax / 10) : /งาน/.test(f.unit) ? 1 : q.pax;
    const it = D.foodItem(fid, qty, times, f.cat === 'coffee' ? 'เช้า–บ่าย วันละ 2 รอบ' : f.place || f.sub || '', sched);
    if (f.cat === 'package') { it.name = f.name + ' (' + f.sub + ')'; it.desc = D.foodFeat(f).slice(0, 2).join(' · '); }
    q.items.push(it); App.touch(q, 'เพิ่ม ' + f.name);
  }
  Object.assign(App.ACT, {
    foodCat(el) { App.state.ui.food.cat = el.dataset.v; App.render(); },
    foodSort(el) { App.state.ui.food.sort = el.dataset.v; App.render(); },
    foodTag(el) { const t = App.state.ui.food.tags, v = el.dataset.v; const i = t.indexOf(v); if (i >= 0) t.splice(i, 1); else t.push(v); App.render(); },
    foodAdd(el) {
      const q = App.ctxQuote();
      const run = id => { const qq = Store.quote(id); App.setCtx(id); addFood(qq, el.dataset.v); App.render(); App.toast('เพิ่มใน ' + id + ' แล้ว'); };
      if (!q) App.pickQuoteModal('เพิ่มในใบเสนอราคาใด?', run); else run(q.id);
    },
    foodRemove(el) {
      const q = App.ctxQuote(); if (!q) return;
      const it = q.items.find(i => i.ref === el.dataset.v);
      App.confirm('นำ “' + it.name + '” ออกจากใบเสนอ?', () => { q.items = q.items.filter(i => i !== it); App.touch(q, 'นำ ' + it.name + ' ออก'); App.render(); }, 'นำออก');
    }
  });
  Object.assign(App.CHG, {
    foodLo(el) { const u = App.state.ui.food; u.lo = Math.min(Number(el.value), u.hi); App.render(); },
    foodHi(el) { const u = App.state.ui.food; u.hi = Math.max(Number(el.value), u.lo); App.render(); }
  });

  /* =====================================================================
     MEETING ROOMS (1g)
     ===================================================================== */
  const LAYOUT_ICON = {
    theatre: '<svg viewBox="0 0 72 48"><rect x="4" y="2" width="64" height="3" fill="#A8262D"/>' + Array.from({ length: 24 }, (_, i) => `<circle cx="${8 + (i % 8) * 8}" cy="${14 + Math.floor(i / 8) * 9}" r="2.2" fill="currentColor"/>`).join('') + '</svg>',
    classroom: '<svg viewBox="0 0 72 48"><rect x="4" y="2" width="64" height="3" fill="#A8262D"/>' + [0, 1, 2, 3].map(i => `<rect x="8" y="${12 + i * 9}" width="56" height="3" fill="#F5C200"/>`).join('') + '</svg>',
    ushape: '<svg viewBox="0 0 72 48"><rect x="4" y="2" width="64" height="3" fill="#A8262D"/><path d="M14 12v30h44V12" fill="none" stroke="currentColor" stroke-width="3"/></svg>',
    banquet: '<svg viewBox="0 0 72 48">' + [[14, 14], [36, 14], [58, 14], [14, 36], [36, 36], [58, 36]].map(([x, y]) => `<circle cx="${x}" cy="${y}" r="7" fill="currentColor"/>`).join('') + '</svg>',
    cocktail: '<svg viewBox="0 0 72 48">' + Array.from({ length: 22 }, (_, i) => `<circle cx="${6 + (i * 13) % 62}" cy="${6 + Math.floor(i * 37 / 62) % 5 * 9}" r="2" fill="currentColor"/>`).join('') + '</svg>'
  };
  function seatChart(layout, pax) {
    const W = 360; let body = '', h = 260;
    const stage = `<rect x="70" y="8" width="220" height="14" rx="3" fill="#A8262D"/><text x="180" y="18.5" text-anchor="middle" font-size="9" fill="#fff">เวที / จอ</text>`;
    if (layout === 'classroom' || layout === 'theatre') {
      const perRow = layout === 'classroom' ? 9 : 15, rows = Math.ceil(pax / perRow);
      const gap = layout === 'classroom' ? 20 : 14; h = 40 + rows * gap + 30;
      let n = 0;
      for (let r = 0; r < rows; r++) {
        const y = 40 + r * gap;
        if (layout === 'classroom') for (let b = 0; b < 3; b++) {
          const x0 = 12 + b * 118;
          body += `<rect x="${x0}" y="${y}" width="104" height="4" rx="2" fill="#0E9F8E"/>`;
          for (let s = 0; s < 3 && n < pax; s++, n++) body += `<circle cx="${x0 + 18 + s * 34}" cy="${y + 11}" r="3" fill="#17211F"/>`;
        } else for (let s = 0; s < perRow && n < pax; s++, n++) body += `<circle cx="${22 + s * 22 + (s >= perRow / 2 ? 14 : 0)}" cy="${y + 4}" r="3.2" fill="#17211F"/>`;
      }
    } else if (layout === 'ushape') {
      const side = Math.ceil((pax - Math.min(pax, 10)) / 2), base = Math.min(pax, 10);
      h = 60 + side * 16 + 40;
      body += `<path d="M70 50 V${50 + side * 16} H290 V50" fill="none" stroke="#0E9F8E" stroke-width="5"/>`;
      for (let i = 0; i < side; i++) body += `<circle cx="58" cy="${56 + i * 16}" r="3.2" fill="#17211F"/><circle cx="302" cy="${56 + i * 16}" r="3.2" fill="#17211F"/>`;
      for (let i = 0; i < base; i++) body += `<circle cx="${84 + i * (192 / Math.max(1, base - 1))}" cy="${62 + side * 16}" r="3.2" fill="#17211F"/>`;
    } else {
      const per = layout === 'banquet' ? 10 : 6, tables = Math.ceil(pax / per), cols = 5, rows = Math.ceil(tables / cols);
      h = 50 + rows * 62 + 20;
      for (let t = 0; t < tables; t++) {
        const cx = 40 + (t % cols) * 70, cy = 66 + Math.floor(t / cols) * 62, r = layout === 'banquet' ? 14 : 8;
        body += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#0E9F8E"/>`;
        const seats = Math.min(per, pax - t * per);
        for (let s = 0; s < seats; s++) { const a = s / per * Math.PI * 2; body += `<circle cx="${cx + Math.cos(a) * (r + 8)}" cy="${cy + Math.sin(a) * (r + 8)}" r="2.6" fill="#17211F"/>`; }
      }
    }
    return `<svg viewBox="0 0 ${W} ${h}" role="img" aria-label="ผังที่นั่ง ${pax} ที่">${stage}${body}<text x="180" y="${h - 8}" text-anchor="middle" font-size="9" fill="#8D928F">ทางเข้า ↓</text></svg>`;
  }

  App.V.meet = {
    title: 'ห้องประชุม',
    render() {
      const q = App.ctxQuote();
      const ui = App.state.ui.meet || (App.state.ui.meet = { layout: (q && (q.items.find(i => i.sec === 'meet' && i.layout) || {}).layout) || 'classroom', sel: null, av: true, half: false, holdDays: 10 });
      const pax = q ? q.pax : 50, days = q ? U.daysOf(q.start, q.end) : [U.today()];
      const inQ = q && q.items.find(i => i.sec === 'meet' && D.MEETING_ROOMS.some(r => r.id === i.ref));
      if (!ui.sel) ui.sel = (inQ && inQ.ref) || D.MEETING_ROOMS[1].id;
      if (q && ui._q !== q.id) { ui._q = q.id; ui.av = !q.items.length || q.items.some(i => i.ref === 'svc:av'); if (inQ) { ui.sel = inQ.ref; ui.layout = inQ.layout || ui.layout; } }
      const r = D.MEETING_ROOMS.find(x => x.id === ui.sel);
      const fit = room => { const [mn, mx] = room.cap[ui.layout]; return !mx ? ['ไม่รองรับ', 'red'] : pax > mx ? ['ไม่พอ', 'red'] : pax < mn ? ['กว้างขวาง', 'teal'] : ['พอดี', 'teal']; };
      const avail = room => {
        const b = days.map(d => [d, App.meetBooking(room.id, d, q && q.id)]).filter(x => x[1]);
        if (!b.length) return `<span class="badge st-confirmed">ว่างทั้ง ${days.length} วัน</span>`;
        return `<span class="badge st-tentative">${U.fmtD(b[0][0], { year: false })} มี${b[0][1].status === 'confirmed' ? 'การจองยืนยัน' : 'จองชั่วคราว'}</span>`;
      };
      const price = ui.half ? r.half : r.full, nd = days.length;
      const total = price * nd + (ui.av ? 8000 * nd : 0);
      const chosen = inQ && inQ.ref === r.id;
      return `
<div class="split wide-side">
  <div style="min-width:0">
    <div class="between" style="margin-bottom:16px;flex-wrap:wrap"><h1>ห้องประชุม</h1>
      <div class="row wrap">${q ? `<span class="pill">📅 ${U.fmtRange(q.start, q.end)}</span><span class="pill">${U.n(pax)} ท่าน</span>` : ''}<div class="seg white"><button class="${!ui.half ? 'on' : ''}" data-act="meetHalf" data-v="0">เต็มวัน</button><button class="${ui.half ? 'on' : ''}" data-act="meetHalf" data-v="1">ครึ่งวัน</button></div></div></div>
    <div style="margin-bottom:12px">${App.ctxBar(q, 'สำหรับ')}</div>
    <div class="layouts">${D.LAYOUTS.map(l => `<button class="${ui.layout === l.id ? 'on' : ''}" data-act="meetLayout" data-v="${l.id}">${LAYOUT_ICON[l.id]}${l.name}</button>`).join('')}</div>
    ${D.MEETING_ROOMS.map(room => {
      const [mn, mx] = room.cap[ui.layout]; const [fl, fc] = fit(room);
      return `<article class="card mroom ${room.id === ui.sel ? 'on' : ''}" data-act="meetSel" data-v="${room.id}">
        <div class="ph" style="${room.img ? `background-image:url(${room.img})` : 'background:repeating-linear-gradient(135deg,#dff3ef 0 10px,#cfeee8 10px 20px)'}"></div>
        <div><div class="row" style="align-items:baseline;flex-wrap:wrap"><h3 style="font-size:19px">${H(room.name)}</h3><span class="small muted">${room.area} ตร.ม. · ${H(room.desc)}</span></div>
          <div class="small muted" style="margin:8px 0 10px">${D.LAYOUTS.filter(l => l.id !== ui.layout && room.cap[l.id][1]).slice(0, 3).map(l => l.name + ' ' + room.cap[l.id][0] + '–' + room.cap[l.id][1]).join(' · ')}</div>${avail(room)}</div>
        <div class="cap"><span class="xs muted">${D.LAYOUTS.find(l => l.id === ui.layout).name}</span><b class="${fc === 'red' ? 'red' : ''}">${mx ? mn + '–' + mx : '—'}</b><span class="small ${fc === 'red' ? 'red' : 'muted'}">${fl}</span></div>
        <div class="pr" style="text-align:right"><b style="font:600 22px var(--fh)">${U.n(ui.half ? room.half : room.full)} ฿</b><div class="xs muted">เต็มวัน · ครึ่งวัน ${U.n(room.half)}</div></div>
      </article>`;
    }).join('')}
  </div>
  <aside class="card sticky" style="overflow:hidden">
    <div style="height:190px;background:${r.img ? `url(${r.img}) center/cover` : 'repeating-linear-gradient(135deg,#dff3ef 0 10px,#cfeee8 10px 20px)'};position:relative">${chosen ? `<span class="badge" style="position:absolute;left:14px;bottom:14px;background:#fff;color:var(--teal-d)">${ICON.check} เลือกแล้ว</span>` : ''}</div>
    <div style="padding:20px 22px 22px">
      <div class="small muted">${H(r.desc)}</div>
      <h2 style="font-size:22px;margin:4px 0">${H(r.name)} · ${D.LAYOUTS.find(l => l.id === ui.layout).name}</h2>
      <div class="small muted">${r.area} ตร.ม. · รองรับ ${r.cap[ui.layout][0]}–${r.cap[ui.layout][1]} ท่าน · ผังตัวอย่าง ${pax} ที่</div>
      <div class="seat">${seatChart(ui.layout, Math.min(pax, 400))}</div>
      <h3 style="font-size:14px;margin-bottom:10px">รวมในห้อง</h3>
      <div class="amen">${r.amen.map(a => `<span>${H(a)}</span>`).join('')}</div>
      <div class="card" style="padding:14px 16px;margin:16px 0;display:flex;align-items:center;gap:12px"><div style="flex:1"><b>AV Package</b><div class="small muted">โปรเจกเตอร์คู่ ไมค์ไร้สาย 4 · ช่าง</div></div>${App.sw('AV Package', ui.av, 'data-change="meetAv"')}</div>
      <div class="between"><span class="small muted">${nd} วัน × ${U.n(price)}${ui.av ? ' + AV ' + U.n(8000 * nd) : ''}</span><b style="font:700 22px var(--fh)" class="num">${U.n(total)} ฿</b></div>
      <div class="row" style="margin-top:14px;gap:8px"><span class="small muted">จองชั่วคราว</span><select class="input sm" style="width:auto" data-change="meetHold">${[0, 3, 7, 10, 14].map(d => `<option value="${d}" ${ui.holdDays === d ? 'selected' : ''}>${d ? d + ' วัน' : 'ไม่จอง'}</option>`).join('')}</select></div>
      <button class="btn red block" style="margin-top:14px" data-act="meetAdd">${chosen ? 'อัปเดตในใบเสนอ' : 'เพิ่มในใบเสนอ'}${ui.holdDays ? ' + จองชั่วคราว ' + ui.holdDays + ' วัน' : ''}</button>
    </div>
  </aside>
</div>`;
    }
  };
  Object.assign(App.ACT, {
    meetLayout(el) { App.state.ui.meet.layout = el.dataset.v; App.render(); },
    meetSel(el) { App.state.ui.meet.sel = el.dataset.v; App.render(); if (window.innerWidth <= 1100) document.querySelector('aside.card')?.scrollIntoView({ behavior: 'smooth' }); },
    meetHalf(el) { App.state.ui.meet.half = el.dataset.v === '1'; App.render(); },
    meetAdd() {
      const ui = App.state.ui.meet;
      const run = id => {
        const q = Store.quote(id); App.setCtx(id);
        const r = D.MEETING_ROOMS.find(x => x.id === ui.sel), days = U.daysOf(q.start, q.end);
        const busy = days.map(d => App.meetBooking(r.id, d, q.id)).find(Boolean);
        const apply = () => {
          const old = q.items.find(i => i.sec === 'meet' && D.MEETING_ROOMS.some(m => m.id === i.ref));
          const sched = days.map((d, i) => ({ day: i + 1, start: ui.half ? '13:00' : '08:30', end: '17:00', place: r.name }));
          const it = D.meetItem(r.id, days.length, ui.layout, { pax: q.pax, half: ui.half, desc: ui.half ? 'ครึ่งวัน 13:00–17:00' : 'เต็มวัน 08:00–17:00', sched });
          if (old) q.items.splice(q.items.indexOf(old), 1, it); else q.items.unshift(it);
          const av = q.items.find(i => i.ref === 'svc:av');
          if (ui.av && !av) q.items.splice(q.items.indexOf(it) + 1, 0, D.svcItem('av', days.length));
          if (ui.av && av) av.q1 = days.length;
          if (!ui.av && av) q.items = q.items.filter(i => i !== av);
          q.room = r.id;
          if (ui.holdDays) { q.holdUntil = U.addDays(U.today(), ui.holdDays); if (q.status === 'draft' || q.status === 'sent') q.status = 'tentative'; }
          App.touch(q, `เลือกห้อง ${r.name}${ui.holdDays ? ' + จองชั่วคราวถึง ' + U.fmtD(q.holdUntil, { year: false }) : ''}`);
          App.render(); App.toast('อัปเดต ' + q.id + ' แล้ว');
        };
        if (busy) App.confirm(`${r.name} มีการจองของ ${busy.client} ในช่วงนี้ ต้องการเพิ่มต่อหรือไม่?`, apply, 'เพิ่มต่อ'); else apply();
      };
      const q = App.ctxQuote(); if (!q) App.pickQuoteModal('เพิ่มในใบเสนอราคาใด?', run); else run(q.id);
    }
  });
  Object.assign(App.CHG, {
    meetAv(el) { App.state.ui.meet.av = el.checked; App.render(); },
    meetHold(el) { App.state.ui.meet.holdDays = Number(el.value); App.render(); }
  });

  /* =====================================================================
     GROUP ROOMS (1h)
     ===================================================================== */
  function roomState(q) {
    const st = { counts: {}, ci: q ? q.start : U.today(), nights: q ? Math.max(1, U.daysOf(q.start, q.end).length - 1) : 1, late: false, abf: true };
    if (q) {
      const rooms = q.items.filter(i => i.sec === 'room' && !i.free && i.ref !== 'late');
      // a stay across a change of season is several lines of one room type: count the rooms once, add up the nights
      const span = {};
      for (const it of rooms) {
        if (it.block) { for (const [k, v] of Object.entries(it.block)) st.counts[k] = (st.counts[k] || 0) + v; }
        else if (D.ROOM_TYPES.some(t => t.id === it.ref)) st.counts[it.ref] = Math.max(st.counts[it.ref] || 0, Number(it.q1) || 0);
        const ci = it.ci || q.start, n = Number(it.q2) || 1, k = it.ref || 'block';
        const sp = span[k] || (span[k] = { a: ci, b: U.addDays(ci, n) });
        if (ci < sp.a) sp.a = ci; if (U.addDays(ci, n) > sp.b) sp.b = U.addDays(ci, n);
      }
      const sp = Object.values(span)[0]; if (sp) { st.ci = sp.a; st.nights = Math.max(1, U.diffDays(sp.a, sp.b)); }
      st.late = q.items.some(i => i.ref === 'late');
      st.abf = !rooms.some(i => /Room only/i.test(i.name));
    }
    return st;
  }
  const stayTotal = (t, ui) => D.roomStay(t, ui.ci, ui.nights, ui.abf).reduce((a, x) => a + x.nights * x.price, 0);
  App.V.rooms = {
    nav: 'room', title: 'ห้องพัก',
    render() {
      const q = App.ctxQuote();
      let ui = App.state.ui.rooms;
      if (!ui || ui.q !== (q && q.id)) ui = App.state.ui.rooms = Object.assign(roomState(q), { q: q && q.id });
      const co = U.addDays(ui.ci, ui.nights);
      const sel = D.ROOM_TYPES.filter(t => ui.counts[t.id] > 0);
      const tot = sel.reduce((a, t) => a + ui.counts[t.id], 0), guests = sel.reduce((a, t) => a + ui.counts[t.id] * t.pax, 0);
      const amount = sel.reduce((a, t) => a + ui.counts[t.id] * stayTotal(t, ui), 0);
      const free = Math.floor(tot / 20);
      const cutoff = U.addDays(ui.ci, -14);
      const segs = D.roomStay(D.ROOM_TYPES[0], ui.ci, ui.nights, true), season = D.seasonOf(ui.ci);
      const card = t => {
        const c = ui.counts[t.id] || 0, pr = D.roomPrice(t, ui.ci, ui.abf), left = t.total;
        return `<article class="card rcard ${c ? 'on' : ''}">
      <div class="ph" style="${t.img ? `background-image:url(${t.img})` : ''}"><span class="zone">${H(t.zone)}</span>${t.img ? '' : `<span class="xs mono muted" style="position:absolute;left:12px;bottom:10px">รูปห้อง: ${H(t.fam)}</span>`}</div>
      <div class="bd"><div class="between" style="align-items:flex-start"><h3 style="font-size:18px">${H(t.fam)}</h3><span class="badge">${H(t.bed)}</span></div>
        <div class="small muted" style="margin:6px 0 12px">มี ${left} ห้อง · ราคาสำหรับ ${t.pax} ท่าน · ${ui.abf ? 'รวมอาหารเช้า' : 'ไม่รวมอาหารเช้า'}</div>
        <div class="between"><div><span class="strike num">${U.n(t.rack)}</span><span class="pr num">${U.n(pr)} ฿</span><span class="small muted"> / คืน</span></div>
        <span class="stepper"><button class="minus" data-act="rmStep" data-k="${t.id}" data-v="-1" aria-label="ลด">−</button><input inputmode="numeric" value="${c}" data-change="rmSet" data-k="${t.id}" aria-label="จำนวนห้อง ${H(t.full)}"><button class="plus" data-act="rmStep" data-k="${t.id}" data-v="1" aria-label="เพิ่ม">+</button></span></div>
        ${c > left ? `<div class="xs red" style="margin-top:6px">เกินจำนวนห้องที่มี (${left})</div>` : ''}
      </div></article>`;
      };
      const zones = ['Resident', 'Resort'];
      return `
<div class="page-hd"><div><h1>ห้องพัก</h1><div class="small muted">ราคาต่อห้องต่อคืน สำหรับ 2 ท่าน (Triple 3 ท่าน) · ราคาก่อน Service และ VAT · ช่วงราคา ${H(season.name)}${segs.length > 1 ? ' และ ' + segs.slice(1).map(x => H(x.season.name)).join(', ') + ' (คิดแยกตามคืน)' : ''}</div><div style="margin-top:8px">${App.ctxBar(q, 'สำหรับ')}</div></div>
  <div class="row wrap" style="gap:10px">${App.can('edit_prices') ? '<a class="btn ghost sm" href="#/roomrates">ตั้งราคาห้องพัก</a>' : ''}
  <div class="datebox"><label>เช็กอิน<input type="date" value="${ui.ci}" data-change="rmCi"></label><label style="border-left:1px solid var(--line)">เช็กเอาต์<input type="date" value="${co}" data-change="rmCo"></label><span class="n">${ui.nights} คืน</span></div></div></div>
<div class="seg white" style="margin-bottom:14px" role="group" aria-label="อาหารเช้า"><button class="${ui.abf ? 'on' : ''}" data-act="rmAbf" data-v="1">With ABF · รวมอาหารเช้า</button><button class="${ui.abf ? '' : 'on'}" data-act="rmAbf" data-v="0">Room only · ไม่รวมอาหารเช้า</button></div>
<div class="split">
  <div>${zones.map(z => `<h2 class="rm-zone">Zone : ${z}</h2><div class="rcards" style="margin-bottom:18px">${D.ROOM_TYPES.filter(t => t.zoneId === z).map(card).join('')}</div>`).join('')}</div>
  <aside class="card pad sticky">
    <h2 style="margin-bottom:12px">Room block</h2>
    ${sel.length ? sel.map(t => `<div class="sum-row"><span>${H(t.name)} × ${ui.counts[t.id]}</span><span class="num">${U.n(ui.counts[t.id] * stayTotal(t, ui))}</span></div>`).join('') : '<div class="small muted">ยังไม่ได้เลือกห้อง</div>'}
    <div class="sum-row" style="border-top:1px solid var(--line);margin-top:6px;padding-top:10px"><b>รวม ${tot} ห้อง × ${ui.nights} คืน</b><b class="num">${U.n(amount)}</b></div>
    <div class="small muted">รองรับ ${guests} ท่าน · ${ui.abf ? 'รวมอาหารเช้า' : 'ไม่รวมอาหารเช้า'}${q ? ' · ผู้เข้าร่วม ' + q.pax + ' ท่าน' : ''}</div>
    <div class="hr"></div>
    <div class="between"><div><div>ห้องฟรีสำหรับผู้จัด</div><div class="small muted">ทุก 20 ห้อง ฟรี 1 ห้อง</div></div><span class="badge" style="background:var(--yellow)">+${free} ห้อง</span></div>
    <div class="between" style="margin-top:14px"><span>Late check-out 14:00</span>${App.sw('Late check-out', ui.late, 'data-change="rmLate"')}</div>
    <div class="note" style="margin:16px 0">วันปล่อยห้อง (cut-off) ${U.fmtD(cutoff, { long: true })} — ห้องที่ยังไม่มีรายชื่อจะคืนเข้าระบบอัตโนมัติ</div>
    <button class="btn red block" data-act="rmApply" ${tot ? '' : 'disabled'}>อัปเดตใบเสนอ</button>
  </aside>
</div>`;
    }
  };
  // quote lines for a room type: one line per season the stay touches
  App.roomLines = (t, count, ci, nights, abf, desc) => D.roomStay(t, ci, nights, abf).map((x, k, all) =>
    D.roomItem(t.full + (abf ? ' รวมอาหารเช้า' : ' Room only (ไม่รวมอาหารเช้า)'), count, x.nights, x.price,
      (desc || '') + (all.length > 1 ? (desc ? ' · ' : '') + 'คืนที่ ' + (k ? all.slice(0, k).reduce((a, y) => a + y.nights, 0) + 1 : 1) + '–' + all.slice(0, k + 1).reduce((a, y) => a + y.nights, 0) + ' ราคาช่วง ' + x.season.name : ''), t.id, { ci: x.from }));
  Object.assign(App.ACT, {
    rmStep(el) { const ui = App.state.ui.rooms, k = el.dataset.k; ui.counts[k] = U.clamp((ui.counts[k] || 0) + Number(el.dataset.v), 0, 999); App.render(); },
    rmAbf(el) { App.state.ui.rooms.abf = el.dataset.v === '1'; App.render(); },
    rmApply() {
      const ui = App.state.ui.rooms;
      const run = id => {
        const q = Store.quote(id); App.setCtx(id);
        q.items = q.items.filter(i => i.sec !== 'room' && i.ref !== 'late');
        const sel = D.ROOM_TYPES.filter(t => ui.counts[t.id] > 0), tot = sel.reduce((a, t) => a + ui.counts[t.id], 0);
        const desc = `เช็กอิน ${U.fmtD(ui.ci, { year: false })} 14:00 → เช็กเอาต์ ${U.fmtD(U.addDays(ui.ci, ui.nights), { year: false })} ${ui.late ? '14:00' : '12:00'}`;
        for (const t of sel) q.items.push(...App.roomLines(t, ui.counts[t.id], ui.ci, ui.nights, ui.abf, desc));
        if (tot >= 20) q.items.push(D.roomItem('ห้องเจ้าหน้าที่ พักห้องละ 2 ท่าน' + (ui.abf ? ' รวมอาหารเช้า' : ''), Math.floor(tot / 20), ui.nights, D.roomPrice(sel[0], ui.ci, ui.abf), 'ห้องฟรีสำหรับผู้จัด ทุก 20 ห้อง ฟรี 1 ห้อง', 'comp', { free: true }));
        if (ui.late) q.items.push({ id: U.uid('i'), sec: 'room', ref: 'late', name: 'Late check-out 14:00', desc: 'สำหรับ room block', price: 0, q1: tot, u1: 'ห้อง', q2: null, u2: '', sched: [], free: true });
        if (ui.ci !== q.start) App.toast('หมายเหตุ: วันเช็กอินต่างจากวันงาน');
        App.touch(q, `อัปเดตห้องพัก ${tot} ห้อง × ${ui.nights} คืน`); App.state.ui.rooms = null;
        App.go('#/quote/' + encodeURIComponent(q.id));
      };
      const q = App.ctxQuote(); if (!q) App.pickQuoteModal('อัปเดตใบเสนอราคาใด?', run); else run(q.id);
    }
  });
  /* ---------- room rates (admins / edit_prices): counts, rack and season prices, season dates ---------- */
  const rrDraft = () => {
    const ui = App.state.ui;
    if (!ui.rr) {
      ui.rr = { seasons: D.seasons().map(x => Object.assign({}, x)), types: {} };
      for (const t of D.ROOM_TYPES) ui.rr.types[t.id] = { total: t.total, rack: t.rack, rates: JSON.parse(JSON.stringify(t.rates)) };
      ui.rrDirty = false;
    }
    return ui.rr;
  };
  const MON = U.MS;
  const mdPick = (si, k, v) => { const [m, d] = v.split('-').map(Number); return `<span class="row" style="gap:4px;flex-wrap:nowrap"><select class="input sm" data-change="rrSeason" data-s="${si}" data-k="${k}d" aria-label="วัน">${Array.from({ length: 31 }, (_, i) => `<option ${i + 1 === d ? 'selected' : ''}>${i + 1}</option>`).join('')}</select><select class="input sm" data-change="rrSeason" data-s="${si}" data-k="${k}m" aria-label="เดือน">${MON.map((x, i) => `<option value="${i + 1}" ${i + 1 === m ? 'selected' : ''}>${x}</option>`).join('')}</select></span>`; };
  App.V.roomrates = {
    nav: 'room', title: 'ตั้งราคาห้องพัก',
    render() {
      if (!App.can('edit_prices')) return '<div class="empty">เฉพาะผู้มีสิทธิ์แก้ไขราคาและแพ็กเกจ — <a href="#/rooms">กลับ</a></div>';
      const d = rrDraft(), ui = App.state.ui, ss = d.seasons;
      const num = (id, path, v, w = 86) => `<input class="input sm num" style="width:${w}px;text-align:right" inputmode="numeric" value="${v}" data-change="rrNum" data-t="${id}" data-p="${path}">`;
      const row = t => { const x = d.types[t.id]; return `<tr><td><b style="font-weight:500">${H(t.fam)}</b></td><td>${H(t.bed)}</td><td>${num(t.id, 'total', x.total, 64)}</td><td>${num(t.id, 'rack', x.rack)}</td>
        ${ss.map(se => { const r = x.rates[se.id] || [0, 0]; return `<td class="rr-s">${num(t.id, se.id + '.0', r[0])}</td><td>${num(t.id, se.id + '.1', r[1])}</td>`; }).join('')}</tr>`; };
      return `<div class="page-hd"><div><div class="small"><a class="link" href="#/rooms">← ห้องพัก</a></div><h1>ตั้งราคาห้องพัก</h1>
        <div class="small muted">ราคาต่อห้องต่อคืน สำหรับ 2 ท่าน (Triple 3 ท่าน) · ราคาก่อน Service และ VAT · ใช้กับทุกการจอง · ช่วงราคาวนซ้ำทุกปี</div></div>
        <div class="row wrap"><button class="btn ghost" data-act="rrReset">คืนค่าตามตารางราคา</button><button class="btn red" data-act="rrSave" ${ui.rrDirty ? '' : 'disabled'}>บันทึกราคา</button></div></div>
      <section class="card pad" style="margin-bottom:16px"><h3 style="margin-bottom:10px">ช่วงราคา</h3>
        ${ss.map((se, i) => `<div class="row wrap rr-season"><input class="input sm" style="max-width:200px" value="${H(se.name)}" data-change="rrSeason" data-s="${i}" data-k="name" aria-label="ชื่อช่วงราคา"><span class="small muted">ตั้งแต่</span>${mdPick(i, 'from', se.from)}<span class="small muted">ถึง</span>${mdPick(i, 'to', se.to)}</div>`).join('')}
        <div class="xs muted" style="margin-top:6px">ช่วงที่ข้ามปีได้ เช่น 1 ส.ค. – 28/29 ก.พ. · เข้าพักข้ามช่วงราคา ระบบคิดราคาแยกตามคืน</div></section>
      <section class="card" style="overflow:auto"><table class="tbl rr-tbl"><thead>
        <tr><th rowspan="2">ประเภทห้อง</th><th rowspan="2">เตียง</th><th rowspan="2">จำนวน (ห้อง)</th><th rowspan="2">ราคาเต็ม</th>${ss.map(se => `<th colspan="2" class="rr-s c">${H(se.name)}</th>`).join('')}</tr>
        <tr>${ss.map(() => '<th class="rr-s r">With ABF</th><th class="r">Room only</th>').join('')}</tr></thead><tbody>
        ${['Resident', 'Resort'].map(z => `<tr class="rr-z"><td colspan="${4 + ss.length * 2}">Zone : ${z}</td></tr>` + D.ROOM_TYPES.filter(t => t.zoneId === z).map(row).join('')).join('')}
      </tbody></table></section>`;
    }
  };
  Object.assign(App.ACT, {
    rrSave() {
      const d = rrDraft(), st = Store.db.settings || (Store.db.settings = {});
      st.roomRates = JSON.parse(JSON.stringify(d)); D.applyRoomRates(st.roomRates); Store.save();
      App.state.ui.rr = null; App.state.ui.rooms = null; App.render(); App.toast('บันทึกราคาห้องพักแล้ว');
    },
    rrReset() {
      App.confirm('คืนค่าราคา จำนวนห้อง และช่วงราคาตามตารางราคาเดิม?', () => {
        const def = D.ROOM_RATES_DEFAULT; App.state.ui.rr = JSON.parse(JSON.stringify(def)); App.state.ui.rrDirty = true; App.render();
      }, 'คืนค่า');
    }
  });
  Object.assign(App.CHG, {
    rrNum(el) {
      const d = rrDraft(), x = d.types[el.dataset.t], v = Math.max(0, Math.round(Number(String(el.value).replace(/[, ]/g, '')) || 0));
      const [a, b] = el.dataset.p.split('.');
      if (b == null) x[a] = v; else { x.rates[a] = x.rates[a] || [0, 0]; x.rates[a][Number(b)] = v; }
      App.state.ui.rrDirty = true; App.render();
    },
    rrSeason(el) {
      const se = rrDraft().seasons[Number(el.dataset.s)], k = el.dataset.k, pad = n => String(n).padStart(2, '0');
      if (k === 'name') se.name = el.value.trim().slice(0, 40) || se.name;
      else { const f = k.slice(0, -1), [m, dd] = se[f].split('-'); se[f] = k.endsWith('m') ? pad(el.value) + '-' + dd : m + '-' + pad(el.value); }
      App.state.ui.rrDirty = true; App.render();
    }
  });
  Object.assign(App.CHG, {
    rmSet(el) { const ui = App.state.ui.rooms; ui.counts[el.dataset.k] = U.clamp(parseInt(el.value, 10) || 0, 0, 999); App.render(); },
    rmCi(el) { const ui = App.state.ui.rooms; if (el.value) ui.ci = el.value; App.render(); },
    rmCo(el) { const ui = App.state.ui.rooms; if (el.value) ui.nights = Math.max(1, U.diffDays(ui.ci, el.value)); App.render(); },
    rmLate(el) { App.state.ui.rooms.late = el.checked; App.render(); }
  });
})(window);
