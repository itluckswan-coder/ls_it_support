/* Luck Swan Sales — quotes: list, editor (table + timeline), wizard, document */
(function (g) {
  'use strict';
  const U = g.U, D = g.DATA, Store = g.Store, App = g.App, ICON = App.ICON;
  const H = U.esc;
  const SEC = [
    { id: 'meet', name: 'ห้องประชุม', color: 'var(--teal)' },
    { id: 'food', name: 'อาหารและเครื่องดื่ม', color: 'var(--red)' },
    { id: 'room', name: 'ห้องพัก', color: 'var(--yellow)' },
    { id: 'other', name: 'บริการอื่น ๆ', color: 'var(--ink2)' }
  ];
  const qid = id => encodeURIComponent(id);
  const timeNow = iso => iso ? iso.slice(11, 16) : '';

  /* =====================================================================
     QUOTES LIST
     ===================================================================== */
  App.V.quotes = {
    nav: 'quote', title: 'ใบเสนอราคา',
    render() {
      const ui = App.state.ui.quotes || (App.state.ui.quotes = { st: 'all', q: '' });
      const all = Store.db.quotes;
      const qs = all.filter(q => (ui.st === 'all' || q.status === ui.st) && (!ui.q || [q.id, q.client, q.title, q.contact].join(' ').toLowerCase().includes(ui.q.toLowerCase())))
        .sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
      const cnt = s => all.filter(q => s === 'all' || q.status === s).length;
      return `
<div class="page-hd"><div><h1>ใบเสนอราคา</h1><div class="muted small">${all.length} ใบ · ยืนยันแล้ว ${cnt('confirmed')} · รอลูกค้า ${cnt('sent')}</div></div><a class="btn red" href="#/new">+ ใบเสนอราคาใหม่</a></div>
<div class="row wrap" style="margin-bottom:16px">
  <input class="input" style="max-width:280px;border-radius:999px" placeholder="ค้นหาลูกค้า เลขที่…" value="${H(ui.q)}" data-change="qlSearch" aria-label="ค้นหาใบเสนอราคา">
  ${['all', 'draft', 'sent', 'tentative', 'confirmed', 'cancelled'].map(s => `<button class="chip ${ui.st === s ? 'on' : ''}" data-act="qlSt" data-v="${s}">${s === 'all' ? 'ทั้งหมด' : D.STATUS[s]} <span class="n">${cnt(s)}</span></button>`).join('')}
</div>
<section class="card" style="overflow:hidden">
<table class="tbl resp"><thead><tr><th>เลขที่</th><th>ลูกค้า / งาน</th><th>วันงาน</th><th>ท่าน</th><th class="r">ยอดสุทธิ (฿)</th><th class="r">สถานะ</th></tr></thead><tbody>
${qs.map(q => `<tr class="click" data-act="openQuote" data-id="${H(q.id)}"><td class="mono small muted hide-m">${H(q.id)}</td>
  <td><b style="font-weight:500">${H(q.client)}</b><div class="sub">${H(q.title)} · <span class="mono">${H(q.id)}</span></div></td>
  <td class="hide-m">${U.fmtRange(q.start, q.end)}</td><td class="hide-m num">${U.n(q.pax)}</td>
  <td class="r num b">${U.n(App.calc(q).net, 2)}<div class="sub show-m" style="font-weight:400">${U.fmtRange(q.start, q.end, { year: false })} · ${U.n(q.pax)} ท่าน</div></td><td class="r">${App.statusBadge(q.status)}</td></tr>`).join('') || '<tr><td colspan="6"><div class="empty">ไม่พบใบเสนอราคา</div></td></tr>'}
</tbody></table></section>`;
    }
  };
  App.CHG.qlSearch = el => { App.state.ui.quotes.q = el.value; App.render(); };
  App.ACT.qlSt = el => { App.state.ui.quotes.st = el.dataset.v; App.render(); };

  /* =====================================================================
     QUOTE EDITOR (1b table / 1c timeline)
     ===================================================================== */
  const cur = () => Store.quote(App.state.params[0]);

  App.V.quote = {
    nav: 'quote', title: p => p[0],
    render(p) {
      const q = Store.quote(p[0]);
      if (!q) return `<div class="empty"><h2>ไม่พบใบเสนอราคา ${H(p[0])}</h2><p><a class="link" href="#/quotes">กลับไปรายการ</a></p></div>`;
      App.setCtx(q.id);
      const mode = p[1] === 'timeline' ? 'timeline' : 'table';
      const c = App.calc(q);
      const head = `
<div class="page-hd" style="align-items:flex-start">
  <div style="min-width:0">
    <div class="eyebrow"><a href="#/quotes" style="color:inherit">ใบเสนอราคา</a> / <b style="color:var(--ink)">${H(q.id)}</b></div>
    <h1>${H(q.title)} · ${H(q.client)}</h1>
    <div class="row wrap" style="margin-top:12px;gap:8px">
      <button class="pill" data-act="qDetails">${U.fmtRange(q.start, q.end)} (${U.DS[U.d(q.start).getDay()]}${q.end !== q.start ? '–' + U.DS[U.d(q.end).getDay()] : ''})</button>
      <button class="pill" data-act="qDetails">${U.n(q.pax)} ท่าน</button>
      <button class="pill" data-act="qDetails">ผู้ติดต่อ: ${H(q.contact || '—')}</button>
      <span class="pill gray">${H(D.STATUS[q.status])} · บันทึกอัตโนมัติ ${timeNow(q.updatedAt)}</span>
    </div>
  </div>
  <div class="row wrap">
    <div class="seg white"><button class="${mode === 'table' ? 'on' : ''}" data-act="qMode" data-v="table">ตาราง</button><button class="${mode === 'timeline' ? 'on' : ''}" data-act="qMode" data-v="timeline">ไทม์ไลน์</button></div>
    <button class="btn ghost" data-act="qDup">ทำสำเนา</button>
    <a class="btn red" href="#/doc/${qid(q.id)}">ดูตัวอย่างและส่ง ${ICON.arrow}</a>
  </div>
</div>`;
      return head + (mode === 'table' ? this.table(q, c) : this.timeline(q, c));
    },

    table(q, c) {
      const secs = SEC.filter(s => s.id !== 'other' || q.items.some(i => i.sec === 'other'));
      const disc = q.discount || { type: 'pct', value: 0 };
      const dv = disc.type === 'pct' && [0, 5, 10].includes(Number(disc.value)) ? String(disc.value) : 'custom';
      const over = c.pct > 10.0001;
      return `
<div class="split">
  <div>
    <div class="col-hd"><span>รายการ</span><span>จำนวน</span><span style="text-align:right">ราคา/หน่วย</span><span style="text-align:right">รวม (฿)</span><span></span></div>
    ${secs.map(s => {
      const items = q.items.filter(i => i.sec === s.id);
      return `<section class="card sec"><div class="sec-hd"><span class="sq" style="background:${s.color}"></span><h3>${s.name}</h3><b class="num" style="font:600 16px var(--fh)">${U.n(c.by[s.id] || 0)}</b></div>
      ${items.map(it => `<div class="li">
        <div style="min-width:0;cursor:pointer" data-act="itemEdit" data-id="${it.id}"><div class="name">${H(it.name)}${it.free ? ' <span class="badge st-confirmed">ฟรี</span>' : ''}</div><div class="desc">${H(it.desc || '')}</div></div>
        <div class="qcol"><span class="qty-chip"><input inputmode="numeric" data-change="itemF" data-id="${it.id}" data-f="q1" value="${H(it.q1)}" style="width:${String(it.q1).length + 1.5}ch" aria-label="จำนวน">${it.u1 && (it.q2 == null || it.sec !== 'food') ? ' ' + H(it.u1) : ''}${it.q2 != null ? ` × <input inputmode="numeric" data-change="itemF" data-id="${it.id}" data-f="q2" value="${H(it.q2)}" style="width:${String(it.q2).length + 1.5}ch" aria-label="จำนวนครั้ง"> ${H(it.u2 || '')}` : ''}</span></div>
        <div class="pcol"><input class="price-in num" inputmode="decimal" data-change="itemF" data-id="${it.id}" data-f="price" value="${U.n(it.price)}" aria-label="ราคาต่อหน่วย"></div>
        <div class="tcol num b" style="text-align:right">${it.free ? 'แถม' : U.n(U.lineTotal(it))}</div>
        <button class="del" data-act="itemDel" data-id="${it.id}" aria-label="ลบรายการ">${ICON.trash}</button>
      </div>`).join('')}
      <button class="add-row" data-act="itemAdd" data-sec="${s.id}">+ เพิ่มรายการใน${s.name}</button></section>`;
    }).join('')}
    ${q.items.some(i => i.sec === 'other') ? '' : `<button class="link" data-act="itemAdd" data-sec="other" style="margin:4px 0 0 4px">+ เพิ่มบริการอื่น ๆ</button>`}
  </div>
  <aside class="sticky stack">
    <section class="card pad">
      <h2 style="margin-bottom:10px">สรุปยอด</h2>
      <div class="sum-row"><span class="l">ยอดรวมก่อนส่วนลด</span><span class="num">${U.n2(c.sub)}</span></div>
      <div class="sum-row"><span class="l">ส่วนลด</span><div class="seg" style="padding:3px">${['0', '5', '10'].map(v => `<button class="${dv === v ? 'on' : ''}" data-act="qDisc" data-v="${v}" style="height:28px">${v}%</button>`).join('')}<button class="${dv === 'custom' ? 'on' : ''}" data-act="qDisc" data-v="custom" style="height:28px">${dv === 'custom' ? (disc.type === 'amt' ? U.n(disc.value) + ' ฿' : disc.value + '%') : 'กำหนดเอง'}</button></div></div>
      <div class="sum-row" style="justify-content:flex-end"><span class="num red">−${U.n2(c.disc)}</span></div>
      <div class="sum-row"><span class="l">ค่าบริการ ${Math.round(Store.db.settings.service * 100)}%</span><span class="num">${U.n2(c.svc)}</span></div>
      <div class="sum-row"><span class="l">ภาษีมูลค่าเพิ่ม ${Math.round(Store.db.settings.vat * 100)}%</span><span class="num">${U.n2(c.vat)}</span></div>
      <div class="net"><div class="small" style="opacity:.75">ยอดสุทธิ</div><div class="v num">฿${U.n2(c.net)}</div><div class="small" style="opacity:.75">เฉลี่ย ${U.n(Math.round(c.perPax))} บาท / ท่าน</div></div>
      ${over ? (q.approval === 'approved' ? `<div class="note" style="margin-top:12px;background:var(--teal-soft);color:var(--teal-d)">ส่วนลด ${U.n(c.pct, 1)}% อนุมัติแล้ว</div>` :
        `<div class="note red" style="margin-top:12px">ส่วนลดเกิน 10% ต้องขออนุมัติผู้จัดการฝ่ายขาย${App.can('discount_over') ? ` · <button class="link red" data-act="qApprove" style="font-weight:700">อนุมัติ</button>` : ''}</div>`) :
        `<div class="small muted" style="margin-top:12px">ส่วนลดเกิน 10% ต้องขออนุมัติผู้จัดการฝ่ายขาย</div>`}
    </section>
    <section class="card pad">
      <div class="sum-row"><span class="l">ยืนราคาถึง</span><b>${U.fmtD(q.validUntil)}</b></div>
      <div class="sum-row"><span class="l">มัดจำ</span><b style="text-align:right">${H(q.deposit || '—')}</b></div>
      <div class="sum-row"><span class="l">กันห้องประชุมไว้</span><b class="${q.holdUntil ? 'red' : ''}">${q.holdUntil ? 'จองชั่วคราวถึง ' + U.fmtD(q.holdUntil, { year: false }) : '—'}</b></div>
      <button class="link" data-act="qDetails" style="margin-top:8px">แก้ไขข้อมูลลูกค้าและเงื่อนไข</button>
    </section>
    <section class="card pad">
      <div class="between" style="margin-bottom:12px"><h3>สถานะ</h3>${App.statusBadge(q.status)}</div>
      <div class="row wrap">
        ${q.status !== 'confirmed' ? `<button class="btn teal sm" data-act="qStatus" data-v="confirmed">ยืนยันการจอง</button>` : ''}
        ${q.status !== 'tentative' && q.status !== 'confirmed' ? `<button class="btn ghost sm" data-act="qStatus" data-v="tentative">จองชั่วคราว</button>` : ''}
        ${App.holding(q) ? `<button class="btn ghost sm" data-act="qToBeo">ใบฟังก์ชัน</button><a class="btn ghost sm" href="#/staff/${qid(q.id)}">จัดรอบพนักงาน</a>` : ''}
        ${q.status !== 'cancelled' ? `<button class="btn ghost sm" data-act="qStatus" data-v="cancelled" style="color:var(--red)">ยกเลิก</button>` : `<button class="btn ghost sm" data-act="qStatus" data-v="draft">กู้คืนเป็นร่าง</button>`}
      </div>
      ${q.history && q.history.length ? `<div class="hr"></div><div class="small muted" style="margin-bottom:6px">ประวัติ</div>${q.history.slice(0, 5).map((h, i) => `<div class="row small" style="padding:4px 0;align-items:flex-start"><i class="dot" style="margin-top:6px;background:${i ? '#cfccc2' : 'var(--teal)'}"></i><span class="muted">${H(h.t)} — ${U.fmtD(h.at.slice(0, 10), { year: false })} ${h.at.slice(11, 16)}</span></div>`).join('')}` : ''}
    </section>
  </aside>
</div>`;
    },

    /* ---------- timeline ---------- */
    timeline(q, c) {
      const days = U.daysOf(q.start, q.end);
      const ui = App.state.ui.tl || (App.state.ui.tl = { f: 'all', s: '' });
      let h0 = 8, h1 = 22;
      for (const it of q.items) for (const s of it.sched || []) { h0 = Math.min(h0, Math.floor(U.hm(s.start))); h1 = Math.max(h1, Math.ceil(U.hm(s.end))); }
      const PX = 46, H0 = h0, height = (h1 - h0) * PX;
      const hours = []; for (let h = h0; h <= h1; h++) hours.push(h);
      const rooms = q.items.filter(i => i.sec === 'room');
      const roomTotal = rooms.reduce((a, i) => a + U.lineTotal(i), 0);
      const dayStats = days.map((d, i) => {
        const its = q.items.filter(it => (it.sched || []).some(s => s.day === i + 1));
        let tot = 0;
        for (const it of q.items) {
          const n = (it.sched || []).length, k = (it.sched || []).filter(s => s.day === i + 1).length;
          if (n) tot += U.lineTotal(it) * k / n;
        }
        return { n: its.length, tot };
      });
      const evClass = it => {
        if (it.sec === 'meet') return it.ref && it.ref.startsWith('svc:') ? 'other' : 'meet';
        const f = D.FOOD.find(x => x.id === it.ref);
        if (f && (f.cat === 'gala' || f.cat === 'wedding')) return 'gala';
        if (f && f.cat === 'lunch') return 'food big';
        return 'food';
      };
      const svcNoSched = q.items.filter(i => i.sec === 'meet' && i.ref && i.ref.startsWith('svc:') && !(i.sched || []).length);
      const cols = days.map((d, di) => {
        const evs = [];
        for (const it of q.items) (it.sched || []).forEach((s, si) => {
          if (s.day !== di + 1) return;
          const top = (U.hm(s.start) - H0) * PX, ht = Math.max(22, (U.hm(s.end) - U.hm(s.start)) * PX - 2);
          const cls = evClass(it);
          let inner;
          if (cls === 'meet') inner = `<b>${H(it.name.split(' — ')[0])}</b><br>${H((D.LAYOUTS.find(l => l.id === it.layout) || {}).name || '')} ${U.n(q.pax)}<br>${s.start}–${s.end}${svcNoSched.map(x => '<br>+ ' + H(x.name)).join('')}`;
          else if (cls === 'gala') inner = `<b>Gala dinner ${s.start}–${s.end}</b><br>${H(it.name)}${s.place ? ' · ' + H(s.place) : ''}<br>${U.n(it.q1)} × ${U.n(it.price)} = ${U.n(it.q1 * it.price)} ฿`;
          else if (cls === 'food big') inner = `<b>${s.start}–${s.end}</b><br>${H(it.name)}`;
          else inner = `<b>${s.start}</b> ${H(it.name)}`;
          evs.push(`<div class="ev ${cls}" style="top:${top + 2}px;height:${ht}px" data-act="schedEdit" data-id="${it.id}" data-i="${si}">${inner}</div>`);
        });
        return `<div><div style="margin-bottom:10px"><h3>วันที่ ${di + 1} · ${U.DS[U.d(d).getDay()]} ${U.fmtD(d, { year: false })}</h3><div class="small muted">${dayStats[di].n} รายการ · ${U.n(Math.round(dayStats[di].tot))} ฿</div></div>
          <div class="tl-col" data-day="${di + 1}" style="height:${height}px">${hours.map(h => `<div class="hl" style="top:${(h - H0) * PX}px"></div>`).join('')}${evs.join('')}</div></div>`;
      });
      const lib = App.V.quote.libItems().filter(x => (ui.f === 'all' || x.kind === ui.f) && (!ui.s || x.name.toLowerCase().includes(ui.s.toLowerCase())));
      return `
<div class="tl-wrap">
  <div style="min-width:0">
    ${rooms.length ? `<div style="background:var(--yellow);border-radius:12px;padding:12px 16px;margin:0 0 14px 66px;display:flex;gap:12px;align-items:center;flex-wrap:wrap"><b>ห้องพัก</b><span style="flex:1;min-width:200px" class="small">${rooms.map(r => H(r.name) + ' ' + r.q1 + ' ห้อง').join(' · ')} · เช็กอิน ${U.fmtD(q.start, { year: false })}</span><b class="num">${U.n(roomTotal)} ฿</b></div>` : ''}
    <div class="tl-days" style="--days:${days.length}">
      <div class="tl-hours" style="height:${height}px;margin-top:58px">${hours.map(h => `<div style="top:${(h - H0) * PX}px">${String(h).padStart(2, '0')}:00</div>`).join('')}</div>
      ${cols.join('')}
    </div>
  </div>
  <aside class="card pad sticky">
    <h3 style="margin-bottom:12px">คลังรายการ</h3>
    <input class="input fill sm" placeholder="ค้นหาเมนู แพ็กเกจ…" value="${H(ui.s)}" data-change="tlSearch" aria-label="ค้นหาคลังรายการ">
    <div class="row wrap" style="margin:12px 0;gap:6px">${[['all', 'ทั้งหมด'], ['food', 'อาหาร'], ['meet', 'ห้อง'], ['other', 'บริการเสริม']].map(([k, l]) => `<button class="chip ${ui.f === k ? 'on' : ''}" style="height:30px" data-act="tlF" data-v="${k}">${l}</button>`).join('')}</div>
    <div class="small muted" style="margin-bottom:10px">ลากรายการไปวางบนช่วงเวลา หรือแตะเพื่อเลือกวัน/เวลา</div>
    <div style="max-height:56vh;overflow:auto;padding-right:2px">${lib.map(x => `<div class="lib-item" draggable="true" data-lib="${H(x.key)}" data-act="libTap" data-key="${H(x.key)}"><span class="bar-l ${x.kind}"></span><div style="flex:1;min-width:0"><div style="font-weight:500;font-size:13.5px">${H(x.name)}</div><div class="xs muted">${H(x.sub)}</div></div><b class="num small">${U.n(x.price)}</b></div>`).join('')}</div>
  </aside>
</div>
<div class="bottom-dock">
  <span class="small" style="opacity:.8">จำนวนผู้เข้าร่วม</span>
  <span class="stepper"><button class="minus" data-act="paxStep" data-v="-1" aria-label="ลด">−</button><input inputmode="numeric" value="${q.pax}" data-change="paxSet" aria-label="จำนวนผู้เข้าร่วม"><button class="plus" data-act="paxStep" data-v="1" aria-label="เพิ่ม">+</button></span>
  <span class="small" style="opacity:.7;flex:1;min-width:200px">ก่อนส่วนลด ${U.n(c.sub)} · ส่วนลด ${U.n(c.pct, c.pct % 1 ? 1 : 0)}% · รวม Service ${Math.round(Store.db.settings.service * 100)}% + VAT ${Math.round(Store.db.settings.vat * 100)}%</span>
  <div style="text-align:right"><div class="xs" style="opacity:.6">ยอดสุทธิ</div><div style="font:700 26px var(--fh);color:var(--yellow)" class="num">฿${U.n2(c.net)}</div></div>
  <a class="btn red" href="#/doc/${qid(q.id)}">สร้างใบเสนอราคา ${ICON.arrow}</a>
</div>`;
    },
    libItems() {
      const out = [];
      D.FOOD.filter(f => f.cat !== 'package').forEach(f => out.push({ key: 'food:' + f.id, kind: 'food', name: f.name, sub: (D.FOOD_CATS.find(c => c.id === f.cat) || {}).name + (f.sub ? ' · ' + f.sub : ''), price: f.price }));
      D.MEETING_ROOMS.forEach(r => out.push({ key: 'meet:' + r.id, kind: 'meet', name: r.name, sub: 'ห้องประชุม ' + (r.cap.classroom[1] ? r.cap.classroom[0] + '–' + r.cap.classroom[1] + ' ท่าน' : 'กลางแจ้ง'), price: r.full }));
      D.SERVICES.forEach(s => out.push({ key: 'svc:' + s.id, kind: 'other', name: s.name, sub: 'บริการเสริม', price: s.price }));
      return out;
    },
    after(p) {
      if (p[1] !== 'timeline') return;
      document.querySelectorAll('.lib-item').forEach(el => el.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', el.dataset.lib); e.dataTransfer.effectAllowed = 'copy'; }));
      document.querySelectorAll('.tl-col').forEach(col => {
        col.addEventListener('dragover', e => { e.preventDefault(); col.classList.add('drop'); });
        col.addEventListener('dragleave', () => col.classList.remove('drop'));
        col.addEventListener('drop', e => {
          e.preventDefault(); col.classList.remove('drop');
          const key = e.dataTransfer.getData('text/plain'); if (!key) return;
          const q = cur(); let h0 = 8;
          for (const it of q.items) for (const s of it.sched || []) h0 = Math.min(h0, Math.floor(U.hm(s.start)));
          const y = e.clientY - col.getBoundingClientRect().top;
          const t = Math.round((h0 + y / 46) * 2) / 2;
          addFromLib(q, key, Number(col.dataset.day), U.fromHm(U.clamp(t, 6, 23)));
        });
      });
    }
  };

  function defaultDur(key) {
    const [k, id] = key.split(':');
    if (k === 'meet') return 8.5;
    if (k === 'svc') return 8;
    const f = D.FOOD.find(x => x.id === id) || {};
    return { coffee: 0.5, lunch: 1, gala: 3, cocktail: 2, wedding: 4 }[f.cat] || 1;
  }
  function addFromLib(q, key, day, start) {
    const [k, id] = key.split(':');
    const end = U.fromHm(Math.min(24, U.hm(start) + defaultDur(key)));
    const ref = k === 'svc' ? 'svc:' + id : id;
    let it = q.items.find(i => i.ref === ref);
    const s = { day, start, end, place: '' };
    if (it) {
      it.sched = it.sched || []; it.sched.push(s);
      if (it.sec === 'food' && it.q2 != null) it.q2 = it.sched.length;
      if (it.sec === 'meet') it.q1 = new Set(it.sched.map(x => x.day)).size;
    } else {
      if (k === 'food') { const f = D.FOOD.find(x => x.id === id); it = D.foodItem(id, /โต๊ะ/.test(f.unit) ? Math.ceil(q.pax / 10) : q.pax, /โต๊ะ|งาน/.test(f.unit) ? null : 1, f.place || '', [s]); }
      else if (k === 'meet') { it = D.meetItem(id, 1, 'classroom', { pax: q.pax, desc: start + '–' + end, sched: [s] }); }
      else { it = D.svcItem(id, 1, [s]); }
      q.items.push(it);
    }
    App.touch(q); App.render(); App.toast('เพิ่ม ' + it.name + ' วันที่ ' + day + ' เวลา ' + start);
  }

  Object.assign(App.ACT, {
    qMode(el) { const q = cur(); App.go('#/quote/' + qid(q.id) + (el.dataset.v === 'timeline' ? '/timeline' : '')); },
    tlF(el) { App.state.ui.tl.f = el.dataset.v; App.render(); },
    libTap(el) {
      const q = cur(), days = U.daysOf(q.start, q.end);
      App.modal(`<h2>เพิ่มลงไทม์ไลน์</h2><div class="g2 grid">
        <div><label class="f">วัน</label><select class="input" name="day">${days.map((d, i) => `<option value="${i + 1}">วันที่ ${i + 1} · ${U.fmtD(d, { year: false })}</option>`).join('')}</select></div>
        <div><label class="f">เวลาเริ่ม</label><input class="input" type="time" name="start" value="10:00" step="900"></div></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="okAdd">เพิ่ม</button></div>`,
      bg => bg.querySelector('#okAdd').onclick = () => { const f = App.formData(bg); App.closeModal(); addFromLib(q, el.dataset.key, Number(f.day), f.start || '10:00'); });
    },
    schedEdit(el) {
      const q = cur(), it = q.items.find(i => i.id === el.dataset.id), idx = Number(el.dataset.i), s = it.sched[idx], days = U.daysOf(q.start, q.end);
      App.modal(`<h2>${H(it.name)}</h2>
        <div class="grid g3"><div><label class="f">วัน</label><select class="input" name="day">${days.map((d, i) => `<option value="${i + 1}" ${s.day === i + 1 ? 'selected' : ''}>วันที่ ${i + 1}</option>`).join('')}</select></div>
        <div><label class="f">เริ่ม</label><input class="input" type="time" name="start" value="${H(s.start)}"></div><div><label class="f">สิ้นสุด</label><input class="input" type="time" name="end" value="${H(s.end)}"></div></div>
        <div class="field"><label class="f">สถานที่</label><input class="input" name="place" value="${H(s.place || '')}"></div>
        <div class="acts"><button class="btn ghost" id="del" style="color:var(--red);margin-right:auto">ลบช่วงเวลานี้</button><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`, bg => {
        bg.querySelector('#ok').onclick = () => { const f = App.formData(bg); Object.assign(s, { day: Number(f.day), start: f.start, end: f.end, place: f.place }); App.closeModal(); App.touch(q); App.render(); };
        bg.querySelector('#del').onclick = () => {
          it.sched.splice(idx, 1);
          if (it.sec === 'food' && it.q2 != null) it.q2 = Math.max(it.sched.length, 0);
          if (it.sec === 'meet' && it.sched.length) it.q1 = new Set(it.sched.map(x => x.day)).size;
          if (!it.sched.length && it.sec === 'food') q.items = q.items.filter(i => i !== it);
          App.closeModal(); App.touch(q); App.render();
        };
      });
    },
    paxStep(el) { setPax(cur(), cur().pax + Number(el.dataset.v)); },
    qDisc(el) {
      const q = cur(), v = el.dataset.v;
      if (v !== 'custom') { setDisc(q, { type: 'pct', value: Number(v) }); return; }
      App.modal(`<h2>ส่วนลดกำหนดเอง</h2><div class="grid g2"><div><label class="f">ประเภท</label><select class="input" name="type"><option value="pct">เปอร์เซ็นต์ (%)</option><option value="amt" ${q.discount.type === 'amt' ? 'selected' : ''}>จำนวนเงิน (฿)</option></select></div>
        <div><label class="f">มูลค่า</label><input class="input" name="value" inputmode="decimal" value="${H(q.discount.value)}"></div></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">ใช้ส่วนลด</button></div>`,
      bg => bg.querySelector('#ok').onclick = () => { const f = App.formData(bg); App.closeModal(); setDisc(q, { type: f.type, value: Math.max(0, parseFloat(String(f.value).replace(/,/g, '')) || 0) }); });
    },
    qApprove() { const q = cur(); q.approval = 'approved'; App.touch(q, 'อนุมัติส่วนลด ' + U.n(App.calc(q).pct, 1) + '%'); App.render(); },
    qStatus(el) {
      const q = cur(), v = el.dataset.v;
      const run = () => {
        q.status = v;
        if (v === 'tentative' && !q.holdUntil) q.holdUntil = U.addDays(U.today(), 10);
        App.touch(q, 'เปลี่ยนสถานะเป็น ' + D.STATUS[v]); App.render(); App.toast('อัปเดตสถานะแล้ว');
      };
      if (v === 'cancelled') App.confirm('ยกเลิกใบเสนอราคา ' + q.id + '?', run, 'ยกเลิกใบเสนอ'); else run();
    },
    qDup() {
      const q = cur(), n = JSON.parse(JSON.stringify(q));
      n.id = Store.nextId('QT'); n.status = 'draft'; n.history = []; n.views = 0; n.approval = null; n.holdUntil = '';
      n.createdAt = n.updatedAt = new Date().toISOString().slice(0, 19);
      n.items.forEach(i => i.id = U.uid('i'));
      n.history.unshift({ at: n.createdAt, t: 'สร้างจากสำเนา ' + q.id });
      Store.db.quotes.push(n); Store.save(); App.toast('ทำสำเนาเป็น ' + n.id); App.go('#/quote/' + qid(n.id));
    },
    qToBeo() {
      const q = cur(); const b = Store.db.beos.find(x => x.quoteId === q.id);
      if (b) App.go('#/beo/' + qid(b.id)); else { const nb = App.makeBEO(q); App.go('#/beo/' + qid(nb.id)); }
    },
    qDetails() { detailsModal(cur()); },
    itemDel(el) { const q = cur(); const it = q.items.find(i => i.id === el.dataset.id); q.items = q.items.filter(i => i !== it); App.touch(q); App.render(); App.toast('ลบ ' + it.name + ' แล้ว'); },
    itemEdit(el) { itemModal(cur(), cur().items.find(i => i.id === el.dataset.id)); },
    itemAdd(el) { addItemModal(cur(), el.dataset.sec); }
  });
  Object.assign(App.CHG, {
    itemF(el) {
      const q = cur(), it = q.items.find(i => i.id === el.dataset.id), f = el.dataset.f;
      const v = parseFloat(String(el.value).replace(/,/g, ''));
      it[f] = isNaN(v) ? 0 : Math.max(0, v);
      App.touch(q); App.render();
    },
    tlSearch(el) { App.state.ui.tl.s = el.value; App.render(); },
    paxSet(el) { setPax(cur(), parseInt(el.value, 10) || 0); }
  });

  function setPax(q, n) {
    n = Math.max(1, n); const old = q.pax;
    q.items.forEach(it => { if (it.sec === 'food' && Number(it.q1) === old && !/โต๊ะ|งาน/.test(it.u1 || '')) it.q1 = n; });
    q.pax = n; App.touch(q); App.render();
  }
  function setDisc(q, d) {
    q.discount = d; q.approval = null;
    const c = App.calc(q);
    if (c.pct > 10.0001) { q.approval = App.can('discount_over') ? 'approved' : 'pending'; }
    App.touch(q, 'ปรับส่วนลดเป็น ' + (d.type === 'amt' ? U.n(d.value) + ' ฿' : d.value + '%')); App.render();
  }

  function detailsModal(q, isNew) {
    App.modal(`<h2>ข้อมูลลูกค้าและงาน</h2>
      <div class="field"><label class="f">ชื่องาน</label><input class="input" name="title" value="${H(q.title)}"></div>
      <div class="grid g2" style="margin-top:14px"><div><label class="f">ลูกค้า (ชื่อย่อ)</label><input class="input" name="client" value="${H(q.client)}"></div><div><label class="f">ชื่อบริษัท/หน่วยงานเต็ม</label><input class="input" name="clientFull" value="${H(q.clientFull || '')}"></div>
      <div><label class="f">ผู้ติดต่อ</label><input class="input" name="contact" value="${H(q.contact || '')}"></div><div><label class="f">โทรศัพท์</label><input class="input" name="phone" type="tel" value="${H(q.phone || '')}"></div>
      <div><label class="f">อีเมล</label><input class="input" name="email" type="email" value="${H(q.email || '')}"></div><div><label class="f">สำเนาถึง (คั่นด้วย ,)</label><input class="input" name="cc" value="${H((q.cc || []).join(', '))}"></div>
      <div><label class="f">วันเริ่มงาน</label><input class="input" type="date" name="start" value="${H(q.start)}"></div><div><label class="f">วันสิ้นสุด</label><input class="input" type="date" name="end" value="${H(q.end)}"></div>
      <div><label class="f">จำนวนผู้เข้าร่วม</label><input class="input" name="pax" inputmode="numeric" value="${H(q.pax)}"></div><div><label class="f">เลขผู้เสียภาษี</label><input class="input" name="taxId" value="${H(q.taxId || '')}"></div>
      <div><label class="f">ยืนราคาถึง</label><input class="input" type="date" name="validUntil" value="${H(q.validUntil || '')}"></div><div><label class="f">กันห้องประชุมไว้ถึง</label><input class="input" type="date" name="holdUntil" value="${H(q.holdUntil || '')}"></div></div>
      <div class="field" style="margin-top:14px"><label class="f">ที่อยู่</label><input class="input" name="address" value="${H(q.address || '')}"></div>
      <div class="field"><label class="f">เงื่อนไขมัดจำ</label><input class="input" name="deposit" value="${H(q.deposit || '')}"></div>
      <div class="field"><label class="f">บันทึกติดตาม (แสดงในหน้าแรก)</label><input class="input" name="followUp" value="${H(q.followUp || '')}"></div>
      <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`, bg => {
      bg.querySelector('#ok').onclick = () => {
        const f = App.formData(bg);
        if (!f.start) { App.toast('กรุณาระบุวันงาน'); return; }
        if (!f.end || f.end < f.start) f.end = f.start;
        const oldPax = q.pax;
        Object.assign(q, { title: f.title, client: f.client, clientFull: f.clientFull, contact: f.contact, phone: f.phone, email: f.email, cc: f.cc.split(',').map(s => s.trim()).filter(Boolean),
          start: f.start, end: f.end, taxId: f.taxId, validUntil: f.validUntil, holdUntil: f.holdUntil, address: f.address, deposit: f.deposit, followUp: f.followUp });
        const days = U.daysOf(q.start, q.end).length;
        q.items.forEach(it => { it.sched = (it.sched || []).filter(s => s.day <= days); });
        App.closeModal();
        const np = parseInt(f.pax, 10) || q.pax;
        if (np !== oldPax) setPax(q, np); else { App.touch(q, 'แก้ไขข้อมูลลูกค้า'); App.render(); }
      };
    });
  }

  function itemModal(q, it) {
    App.modal(`<h2>แก้ไขรายการ</h2>
      <div class="field"><label class="f">ชื่อรายการ</label><input class="input" name="name" value="${H(it.name)}"></div>
      <div class="field"><label class="f">รายละเอียด</label><input class="input" name="desc" value="${H(it.desc || '')}"></div>
      <div class="grid g2" style="margin-top:14px">
        <div><label class="f">ราคาต่อหน่วย (฿)</label><input class="input" name="price" inputmode="decimal" value="${H(it.price)}"></div><div></div>
        <div><label class="f">จำนวน</label><input class="input" name="q1" inputmode="decimal" value="${H(it.q1)}"></div><div><label class="f">หน่วย</label><input class="input" name="u1" value="${H(it.u1 || '')}"></div>
        <div><label class="f">คูณ (เว้นว่างได้)</label><input class="input" name="q2" inputmode="decimal" value="${it.q2 == null ? '' : H(it.q2)}"></div><div><label class="f">หน่วย</label><input class="input" name="u2" value="${H(it.u2 || '')}"></div>
      </div>
      <label class="row" style="margin-top:16px">${App.sw('ฟรี', it.free, 'name="free"')} รายการแถม (ไม่คิดเงิน)</label>
      <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`, bg => {
      bg.querySelector('#ok').onclick = () => {
        const f = App.formData(bg), num = v => Math.max(0, parseFloat(String(v).replace(/,/g, '')) || 0);
        Object.assign(it, { name: f.name || it.name, desc: f.desc, price: num(f.price), q1: num(f.q1), u1: f.u1, q2: f.q2 === '' ? null : num(f.q2), u2: f.u2, free: f.free });
        App.closeModal(); App.touch(q); App.render();
      };
    });
  }

  function addItemModal(q, sec) {
    const days = U.daysOf(q.start, q.end).length, nights = Math.max(1, days - 1);
    let list = [];
    if (sec === 'meet') {
      list = D.MEETING_ROOMS.map(r => ({ k: 'meet:' + r.id, name: r.name, sub: r.desc, price: r.full }))
        .concat(D.SERVICES.map(s => ({ k: 'svc:' + s.id, name: s.name, sub: s.desc, price: s.price })));
    } else if (sec === 'food') {
      list = D.FOOD.map(f => ({ k: 'food:' + f.id, name: f.name, sub: (D.FOOD_CATS.find(c => c.id === f.cat) || {}).name + ' · ' + f.unit, price: f.price }));
    } else if (sec === 'room') {
      list = D.ROOM_TYPES.map(t => ({ k: 'room:' + t.id, name: t.name, sub: t.beds, price: t.group }));
    } else {
      list = D.SERVICES.map(s => ({ k: 'svc:' + s.id, name: s.name, sub: s.desc, price: s.price }));
    }
    App.modal(`<h2>เพิ่มรายการ</h2>
      <input class="input fill sm" placeholder="ค้นหา…" id="addQ" style="margin-bottom:12px">
      <div id="addList" style="max-height:46vh;overflow:auto">${list.map(x => `<button class="pick" data-k="${H(x.k)}" data-s="${H((x.name + ' ' + x.sub).toLowerCase())}"><span><b style="font-weight:500">${H(x.name)}</b><br><span class="small muted">${H(x.sub)}</span></span><b class="num">${U.n(x.price)}</b></button>`).join('')}</div>
      <div class="hr"></div><button class="link" id="custom">+ รายการกำหนดเอง</button>
      <div class="acts"><button class="btn ghost" data-act="closeModal">ปิด</button></div>`, bg => {
      bg.querySelector('#addQ').oninput = e => { const v = e.target.value.toLowerCase(); bg.querySelectorAll('#addList .pick').forEach(b => b.style.display = b.dataset.s.includes(v) ? '' : 'none'); };
      bg.querySelectorAll('.pick').forEach(b => b.onclick = () => {
        const [k, id] = b.dataset.k.split(':'); let it;
        if (k === 'meet') it = D.meetItem(id, days, 'classroom', { pax: q.pax });
        else if (k === 'svc') { it = D.svcItem(id, days); if (sec === 'other') it.sec = 'other'; }
        else if (k === 'food') {
          const f = D.FOOD.find(x => x.id === id);
          const times = f.cat === 'coffee' ? days * 2 : f.cat === 'lunch' ? days : f.cat === 'package' || f.cat === 'wedding' ? null : 1;
          it = D.foodItem(id, /โต๊ะ/.test(f.unit) ? Math.ceil(q.pax / 10) : /งาน/.test(f.unit) ? 1 : q.pax, times);
        } else { const t = D.ROOM_TYPES.find(x => x.id === id); it = D.roomItem(t.name + ' รวมอาหารเช้า', Math.ceil(q.pax / 2), nights, t.group, 'เช็กอิน ' + U.fmtD(q.start, { year: false }), id); }
        q.items.push(it); App.closeModal(); App.touch(q); App.render(); App.toast('เพิ่ม ' + it.name);
      });
      bg.querySelector('#custom').onclick = () => {
        const it = { id: U.uid('i'), sec, ref: '', name: 'รายการใหม่', desc: '', price: 0, q1: 1, u1: '', q2: null, u2: '', sched: [] };
        q.items.push(it); App.closeModal(); itemModal(q, it);
      };
    });
  }

  /* =====================================================================
     WIZARD (1d) — new quote
     ===================================================================== */
  const SLOTS = [
    { k: 'cbAM', name: 'Coffee break เช้า', cat: 'coffee', def: 'cb-morning', time: '10:00', end: '10:30', when: 'all', label: 'ทั้ง {n} วัน' },
    { k: 'lunch', name: 'อาหารกลางวัน', cat: 'lunch', def: 'ln-buffet', time: '12:00', end: '13:00', when: 'all', label: 'ทั้ง {n} วัน' },
    { k: 'cbPM', name: 'Coffee break บ่าย', cat: 'coffee', def: 'cb-morning', time: '14:30', end: '15:00', when: 'all', label: 'ทั้ง {n} วัน' },
    { k: 'gala', name: 'Gala dinner', cat: 'gala', def: 'gd-khantoke', time: '18:30', end: '21:30', when: 'first', label: 'วันที่ 1' },
    { k: 'cocktail', name: 'Cocktail ต้อนรับ', cat: 'cocktail', def: 'ck-sunset', time: '17:00', end: '19:00', when: 'last', label: 'วันที่ {n}' }
  ];
  const STEPS = ['ลูกค้า', 'ห้องประชุม', 'อาหาร', 'ห้องพัก', 'สรุป'];

  function newWiz() {
    const start = U.addDays(U.today(), 30);
    const w = {
      step: 1,
      q: { id: '', title: '', client: '', clientFull: '', contact: '', phone: '', email: '', cc: [], start, end: U.addDays(start, 1), pax: 50, status: 'draft',
        discount: { type: 'pct', value: 0 }, validUntil: U.addDays(U.today(), 14), holdUntil: '', deposit: '30% ภายใน 7 วันหลังยืนยัน', items: [], history: [], views: 0, dietary: [] },
      sel: { room: '', layout: 'classroom', half: false, av: true, slots: {}, rooms: {}, nights: 1 },
      changed: null
    };
    SLOTS.forEach(s => w.sel.slots[s.k] = { on: s.k !== 'cocktail' && s.k !== 'gala', ref: s.def });
    return w;
  }
  function wizBuild(w) {
    const q = w.q, s = w.sel, days = U.daysOf(q.start, q.end), n = days.length;
    const items = [];
    if (s.room) {
      const sched = days.map((d, i) => ({ day: i + 1, start: s.half ? '13:00' : '08:30', end: s.half ? '17:00' : '17:00', place: (D.MEETING_ROOMS.find(r => r.id === s.room) || {}).name }));
      items.push(D.meetItem(s.room, n, s.layout, { pax: q.pax, half: s.half, desc: s.half ? 'ครึ่งวัน 13:00–17:00' : 'เต็มวัน 08:00–17:00', sched }));
      if (s.av) items.push(D.svcItem('av', n));
    }
    for (const sl of SLOTS) {
      const v = s.slots[sl.k]; if (!v || !v.on) continue;
      const dl = sl.when === 'all' ? days.map((d, i) => i + 1) : sl.when === 'first' ? [1] : [n];
      const sched = dl.map(dd => ({ day: dd, start: sl.time, end: sl.end, place: '' }));
      const f = D.FOOD.find(x => x.id === v.ref);
      const it = D.foodItem(v.ref, /โต๊ะ/.test(f.unit) ? Math.ceil(q.pax / 10) : q.pax, /โต๊ะ|งาน/.test(f.unit) ? null : sched.length, sl.name + ' · ' + sl.time, sched);
      it.slot = sl.k;
      // merge duplicates of same menu (e.g. morning + afternoon coffee)
      const ex = items.find(x => x.ref === it.ref && x.sec === 'food');
      if (ex && ex.q2 != null) { ex.sched = ex.sched.concat(sched); ex.q2 = ex.sched.length; ex.desc = 'เช้า–บ่าย วันละ 2 รอบ'; } else items.push(it);
    }
    const nights = s.nights;
    for (const t of D.ROOM_TYPES) {
      const c = s.rooms[t.id] || 0;
      if (c > 0) items.push(D.roomItem(t.name + ' รวมอาหารเช้า', c, nights, t.group, 'เช็กอิน ' + U.fmtD(q.start, { year: false }) + ' · ' + nights + ' คืน', t.id));
    }
    const tot = Object.values(s.rooms).reduce((a, b) => a + b, 0);
    if (tot >= 20) items.push(D.roomItem('ห้องเจ้าหน้าที่ (ฟรีสำหรับผู้จัด)', Math.floor(tot / 20), nights, (D.ROOM_TYPES.find(t => (s.rooms[t.id] || 0) > 0) || D.ROOM_TYPES[0]).group, 'ทุก 20 ห้อง ฟรี 1 ห้อง', 'comp', { free: true }));
    q.items = items;
  }

  App.V.new = {
    nav: 'quote', title: 'ใบเสนอราคาใหม่', bleed: true,
    render() {
      const w = App.state.wiz || (App.state.wiz = newWiz());
      wizBuild(w);
      const q = w.q, c = App.calc(q), n = U.daysOf(q.start, q.end).length;
      const steps = STEPS.map((s, i) => {
        const k = i + 1, st = k < w.step ? 'done' : k === w.step ? 'cur' : '';
        return (i ? `<span class="ln ${k <= w.step ? 'done' : ''}"></span>` : '') + `<button class="s ${st}" data-act="wizGo" data-v="${k}"><i>${k < w.step ? ICON.check : k}</i>${s}</button>`;
      }).join('');
      let body = '';
      if (w.step === 1) body = `
        <h1>ข้อมูลลูกค้าและงาน</h1><div class="muted" style="margin:4px 0 20px">กรอกข้อมูลพื้นฐาน — แก้ไขได้ภายหลัง</div>
        <div class="card pad">
          <div class="field"><label class="f">ชื่องาน *</label><input class="input" data-change="wizF" data-f="title" value="${H(q.title)}" placeholder="เช่น สัมมนาประจำปี 2027"></div>
          <div class="grid g2" style="margin-top:14px">
            <div><label class="f">ลูกค้า / บริษัท *</label><input class="input" data-change="wizF" data-f="client" value="${H(q.client)}" placeholder="บจก. ตัวอย่าง"></div>
            <div><label class="f">ผู้ติดต่อ</label><input class="input" data-change="wizF" data-f="contact" value="${H(q.contact)}" placeholder="คุณ…"></div>
            <div><label class="f">โทรศัพท์</label><input class="input" type="tel" data-change="wizF" data-f="phone" value="${H(q.phone)}"></div>
            <div><label class="f">อีเมล</label><input class="input" type="email" data-change="wizF" data-f="email" value="${H(q.email)}"></div>
            <div><label class="f">วันเริ่มงาน</label><input class="input" type="date" data-change="wizF" data-f="start" value="${H(q.start)}"></div>
            <div><label class="f">วันสิ้นสุด</label><input class="input" type="date" data-change="wizF" data-f="end" value="${H(q.end)}"></div>
            <div><label class="f">จำนวนผู้เข้าร่วม</label><input class="input" inputmode="numeric" data-change="wizF" data-f="pax" value="${H(q.pax)}"></div>
          </div></div>`;
      if (w.step === 2) {
        const pax = q.pax;
        body = `<h1>ห้องประชุม</h1><div class="muted" style="margin:4px 0 16px">${U.n(pax)} ท่าน · ${n} วัน — เลือกรูปแบบและห้อง</div>
          <div class="seg white" style="margin-bottom:14px;flex-wrap:wrap">${D.LAYOUTS.map(l => `<button class="${w.sel.layout === l.id ? 'on' : ''}" data-act="wizSel" data-f="layout" data-v="${l.id}">${l.name}</button>`).join('')}</div>
          ${D.MEETING_ROOMS.map(r => {
            const [mn, mx] = r.cap[w.sel.layout]; const fit = !mx ? 'ไม่รองรับ' : pax > mx ? 'ไม่พอ' : pax < mn ? 'กว้างขวาง' : 'พอดี';
            const busy = U.daysOf(q.start, q.end).map(d => App.meetBooking(r.id, d)).find(Boolean);
            return `<button class="slot ${w.sel.room === r.id ? 'hl' : ''}" style="width:100%;text-align:left" data-act="wizSel" data-f="room" data-v="${r.id}">
              <span class="check ${w.sel.room === r.id ? 'on' : ''}">${w.sel.room === r.id ? ICON.check : ''}</span>
              <span style="flex:1"><b>${H(r.name)}</b><br><span class="small muted">${r.area} ตร.ม. · ${H(r.desc)}${busy ? ' · <span class="red">ไม่ว่าง (' + H(busy.client) + ')</span>' : ''}</span></span>
              <span style="text-align:right"><span class="small ${fit === 'ไม่พอ' || fit === 'ไม่รองรับ' ? 'red' : 'teal'}">${mx ? mn + '–' + mx : '—'} · ${fit}</span><br><b class="num">${U.n(w.sel.half ? r.half : r.full)} ฿</b></span></button>`;
          }).join('')}
          <div class="slot" style="margin-top:14px"><label class="row" style="flex:1">${App.sw('ครึ่งวัน', w.sel.half, 'data-change="wizSw" data-f="half"')} ครึ่งวัน (13:00–17:00)</label><label class="row">${App.sw('AV', w.sel.av, 'data-change="wizSw" data-f="av"')} AV Package (8,000/วัน)</label></div>`;
      }
      if (w.step === 3) {
        const dl = q.dietary || [];
        body = `<h1>อาหารและเครื่องดื่ม</h1><div class="muted" style="margin:4px 0 16px">${U.n(q.pax)} ท่าน · ${n} วัน — เลือกเมนูให้แต่ละช่วง ราคาอัปเดตทันที</div>
          ${SLOTS.map(sl => {
            const v = w.sel.slots[sl.k], f = D.FOOD.find(x => x.id === v.ref);
            const it = q.items.find(x => x.ref === v.ref && x.sec === 'food');
            const opts = D.FOOD.filter(x => x.cat === sl.cat);
            const share = it ? U.lineTotal(it) * (sl.when === 'all' ? n : 1) / Math.max(1, (it.sched || []).length) : 0;
            return `<div class="slot ${v.on ? (w.changed === sl.k ? 'hl' : '') : 'off'}">
              ${App.sw(sl.name, v.on, `data-change="wizSlot" data-k="${sl.k}"`)}
              <div style="flex:1;min-width:120px"><b>${sl.name}</b><div class="small muted">${sl.time} · ${sl.label.replace('{n}', n)}</div></div>
              <select class="input fill sm" data-change="wizMenu" data-k="${sl.k}" aria-label="เมนู">${opts.map(o => `<option value="${o.id}" ${o.id === v.ref ? 'selected' : ''}>${H(o.name)} · ${U.n(o.price)}</option>`).join('')}</select>
              <b class="num" style="min-width:80px;text-align:right">${v.on ? U.n(Math.round(share)) : '—'}</b></div>`;
          }).join('')}
          <h3 style="margin:22px 0 10px">ข้อจำกัดด้านอาหาร</h3>
          <div class="row wrap">${dl.map((x, i) => `<button class="chip on-teal" data-act="wizDiet" data-i="${i}">${H(x.k)} · ${x.n} ✕</button>`).join('')}<button class="chip" data-act="wizDietAdd" style="border-style:dashed">+ เพิ่ม</button></div>`;
      }
      if (w.step === 4) {
        const tot = Object.values(w.sel.rooms).reduce((a, b) => a + b, 0);
        body = `<h1>ห้องพัก</h1><div class="muted" style="margin:4px 0 16px">ราคากรุ๊ป · เช็กอิน ${U.fmtD(q.start)}</div>
          <div class="slot"><span style="flex:1"><b>จำนวนคืน</b></span><span class="stepper"><button class="minus" data-act="wizNights" data-v="-1">−</button><input value="${w.sel.nights}" readonly aria-label="จำนวนคืน"><button class="plus" data-act="wizNights" data-v="1">+</button></span></div>
          ${D.ROOM_TYPES.map(t => {
            const left = Math.min(...U.daysOf(q.start, U.addDays(q.start, Math.max(0, w.sel.nights - 1))).map(d => t.total - App.roomTypeBooked(t.id, d)));
            const c = w.sel.rooms[t.id] || 0;
            return `<div class="slot ${c ? 'hl' : ''}"><div style="flex:1"><b>${H(t.name)}</b><div class="small muted">${H(t.beds)} · เหลือ ${left} ห้อง</div></div><b class="num">${U.n(t.group)} ฿</b>
              <span class="stepper"><button class="minus" data-act="wizRoom" data-k="${t.id}" data-v="-1">−</button><input value="${c}" data-change="wizRoomSet" data-k="${t.id}" data-max="${left}" inputmode="numeric" aria-label="จำนวนห้อง"><button class="plus" data-act="wizRoom" data-k="${t.id}" data-v="1" data-max="${left}">+</button></span></div>`;
          }).join('')}
          <div class="small muted" style="margin-top:12px">รวม ${tot} ห้อง · รองรับ ${tot * 2} ท่าน${tot >= 20 ? ' · ฟรีห้องผู้จัด ' + Math.floor(tot / 20) + ' ห้อง' : ''}</div>`;
      }
      if (w.step === 5) {
        const d = q.discount;
        body = `<h1>สรุปและเงื่อนไข</h1><div class="muted" style="margin:4px 0 16px">ตรวจสอบก่อนสร้างใบเสนอราคา</div>
          <div class="card pad">
            <div class="sum-row"><span class="l">ส่วนลด</span><div class="seg">${[0, 5, 10].map(v => `<button class="${d.type === 'pct' && d.value === v ? 'on' : ''}" data-act="wizDisc" data-v="${v}">${v}%</button>`).join('')}</div></div>
            <div class="grid g2" style="margin-top:12px"><div><label class="f">ยืนราคาถึง</label><input class="input" type="date" data-change="wizF" data-f="validUntil" value="${H(q.validUntil)}"></div>
            <div><label class="f">กันห้องประชุมไว้ถึง (ถ้ามี)</label><input class="input" type="date" data-change="wizF" data-f="holdUntil" value="${H(q.holdUntil)}"></div></div>
            <div class="field" style="margin-top:14px"><label class="f">เงื่อนไขมัดจำ</label><input class="input" data-change="wizF" data-f="deposit" value="${H(q.deposit)}"></div>
          </div>
          ${!q.client || !q.title ? '<div class="note red" style="margin-top:14px">กรุณากรอกชื่องานและชื่อลูกค้าในขั้นที่ 1</div>' : ''}`;
      }
      const nextLabel = w.step < 5 ? 'ถัดไป: ' + STEPS[w.step] : 'สร้างใบเสนอราคา';
      const mix = ['meet', 'food', 'room'].map(s => ({ s, v: c.by[s] || 0 }));
      const mixTot = mix.reduce((a, b) => a + b.v, 0) || 1;
      const colors = { meet: 'var(--teal)', food: 'var(--red)', room: 'var(--yellow)' };
      const hlSec = { 2: 'meet', 3: 'food', 4: 'room' }[w.step];
      return `<div class="wiz">
  <div class="wiz-l">
    <div class="steps">${steps}</div>
    ${body}
    <div class="between" style="margin-top:26px">
      ${w.step > 1 ? `<button class="btn ghost" data-act="wizGo" data-v="${w.step - 1}">${ICON.back} ${STEPS[w.step - 2]}</button>` : `<button class="btn ghost" data-act="wizCancel">ยกเลิก</button>`}
      <button class="btn red" data-act="${w.step < 5 ? 'wizNext' : 'wizFinish'}">${nextLabel} ${ICON.arrow}</button>
    </div>
  </div>
  <div class="wiz-r">
    <div class="between" style="align-items:flex-end"><div><div class="small" style="opacity:.7">ยอดสุทธิ (รวม Service + VAT)</div><div style="font:700 clamp(40px,6vw,64px)/1.05 var(--fh);color:var(--yellow)" class="num">฿${U.n(Math.round(c.net))}</div></div>
      <div style="text-align:right"><div style="font:600 24px var(--fh)" class="num">${U.n(Math.round(c.perPax))}</div><div class="small" style="opacity:.7">บาท / ท่าน</div></div></div>
    <div class="mixbar">${mix.map(m => m.v ? `<i style="flex:${m.v / mixTot};background:${colors[m.s]}"></i>` : '').join('') || '<i style="flex:1;background:#2a3431"></i>'}</div>
    <div class="row wrap small" style="opacity:.8;gap:16px">${mix.map(m => `<span>● ${SEC.find(s => s.id === m.s).name} ${U.n(m.v)}</span>`).join('')}</div>
    <div class="prev-doc">
      <div class="between" style="align-items:flex-start"><div><img src="img/logo-wide-sm.png" alt="" style="width:100px"><div class="xs muted" style="margin-top:4px">ใบเสนอราคา (ร่าง)</div></div><div class="small muted" style="text-align:right">${q.contact ? 'เรียน ' + H(q.contact) + '<br>' : ''}${H(q.client)}</div></div>
      <h3 style="margin:14px 0 8px">${H(q.title || 'ชื่องาน')}</h3>
      ${q.items.filter(i => !i.free).map(i => `<div class="pl ${i.sec === hlSec ? 'hl' : ''}"><span>${H(i.name)}</span><b class="num">${U.n(U.lineTotal(i))}</b></div>`).join('') || '<div class="small muted" style="padding:10px 8px">ยังไม่มีรายการ</div>'}
      <div class="hr" style="margin:10px 0"></div>
      <div class="pl"><b>ยอดสุทธิ</b><b class="num">${U.n2(c.net)} ฿</b></div>
    </div>
    ${hlSec ? '<div class="small" style="opacity:.6;margin-top:14px">ไฮไลต์สีส้ม = รายการที่เพิ่งเปลี่ยนในขั้นนี้</div>' : ''}
  </div></div>`;
    }
  };
  const W = () => App.state.wiz;
  Object.assign(App.ACT, {
    wizGo(el) { W().step = Number(el.dataset.v); App.render(); },
    wizNext() {
      const w = W();
      if (w.step === 1 && (!w.q.title.trim() || !w.q.client.trim())) { App.toast('กรุณากรอกชื่องานและชื่อลูกค้า'); return; }
      if (w.step === 1) w.sel.nights = Math.max(1, U.daysOf(w.q.start, w.q.end).length - 1);
      w.step++; w.changed = null; App.render();
    },
    wizCancel() { App.confirm('ยกเลิกการสร้างใบเสนอราคานี้?', () => { App.state.wiz = null; App.go('#/quotes'); }, 'ยกเลิก'); },
    wizSel(el) { W().sel[el.dataset.f] = el.dataset.v; App.render(); },
    wizDisc(el) { W().q.discount = { type: 'pct', value: Number(el.dataset.v) }; App.render(); },
    wizNights(el) { const s = W().sel; s.nights = U.clamp(s.nights + Number(el.dataset.v), 1, 14); App.render(); },
    wizRoom(el) { const s = W().sel, k = el.dataset.k; s.rooms[k] = U.clamp((s.rooms[k] || 0) + Number(el.dataset.v), 0, Number(el.dataset.max || 999)); App.render(); },
    wizDiet(el) { W().q.dietary.splice(Number(el.dataset.i), 1); App.render(); },
    wizDietAdd() {
      App.modal(`<h2>เพิ่มข้อจำกัดด้านอาหาร</h2><div class="grid g2"><div><label class="f">ประเภท</label><input class="input" name="k" list="dietL" placeholder="เช่น ฮาลาล"><datalist id="dietL"><option>ฮาลาล</option><option>มังสวิรัติ</option><option>เจ</option><option>แพ้อาหารทะเล</option><option>แพ้ถั่ว</option><option>ไม่ทานเนื้อวัว</option></datalist></div>
        <div><label class="f">จำนวนที่</label><input class="input" name="n" inputmode="numeric" value="1"></div></div><div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">เพิ่ม</button></div>`,
      bg => bg.querySelector('#ok').onclick = () => { const f = App.formData(bg); if (f.k) W().q.dietary.push({ k: f.k, n: parseInt(f.n, 10) || 1 }); App.closeModal(); App.render(); });
    },
    wizFinish() {
      const w = W(), q = w.q;
      if (!q.title.trim() || !q.client.trim()) { w.step = 1; App.render(); App.toast('กรุณากรอกชื่องานและชื่อลูกค้า'); return; }
      wizBuild(w);
      q.id = Store.nextId('QT'); q.owner = App.state.uid; q.room = w.sel.room;
      q.createdAt = q.updatedAt = new Date().toISOString().slice(0, 19);
      q.history = [{ at: q.createdAt, t: 'สร้างใบเสนอราคา', by: App.me().name }];
      if (q.discount.value > 10) q.approval = App.can('discount_over') ? 'approved' : 'pending';
      Store.db.quotes.push(q); Store.save(); App.state.wiz = null; App.setCtx(q.id);
      App.toast('สร้าง ' + q.id + ' แล้ว'); App.go('#/quote/' + qid(q.id));
    }
  });
  Object.assign(App.CHG, {
    wizF(el) {
      const q = W().q, f = el.dataset.f;
      q[f] = f === 'pax' ? Math.max(1, parseInt(el.value, 10) || 1) : el.value;
      if (f === 'start' && (!q.end || q.end < q.start)) q.end = q.start;
      if (f === 'end' && q.end < q.start) q.end = q.start;
      App.render();
    },
    wizSw(el) { W().sel[el.dataset.f] = el.checked; App.render(); },
    wizSlot(el) { const w = W(); w.sel.slots[el.dataset.k].on = el.checked; w.changed = el.dataset.k; App.render(); },
    wizMenu(el) { const w = W(); w.sel.slots[el.dataset.k].ref = el.value; w.sel.slots[el.dataset.k].on = true; w.changed = el.dataset.k; App.render(); },
    wizRoomSet(el) { const s = W().sel; s.rooms[el.dataset.k] = U.clamp(parseInt(el.value, 10) || 0, 0, Number(el.dataset.max || 999)); App.render(); }
  });

  /* =====================================================================
     DOCUMENT (1i) — review & send
     ===================================================================== */
  App.V.doc = {
    nav: 'quote', title: p => 'ใบเสนอราคา ' + p[0], bleed: true,
    render(p) {
      const q = Store.quote(p[0]);
      if (!q) return '<div class="empty">ไม่พบใบเสนอราคา</div>';
      const ui = App.state.ui.doc && App.state.ui.doc.id === q.id ? App.state.ui.doc :
        (App.state.ui.doc = { id: q.id, ch: 'email', to: [q.email].concat(q.cc || []).filter(Boolean), unit: true, online: true, msg: null });
      if (ui.msg == null) ui.msg = `เรียน ${q.contact || 'ลูกค้า'}\nขอบคุณที่ไว้วางใจ ลักษวรรณ รีสอร์ต ค่ะ แนบใบเสนอราคาสำหรับงาน${q.title} ${U.fmtRange(q.start, q.end)}${q.room ? ' ห้อง ' + ((D.MEETING_ROOMS.find(r => r.id === q.room) || {}).name || '') : ''}${q.holdUntil ? ' กันไว้ให้ถึง ' + U.fmtD(q.holdUntil, { year: false }) : ''} ค่ะ`;
      const pending = q.approval === 'pending' && App.calc(q).pct > 10;
      const chLabel = { email: 'ส่งใบเสนอราคา', line: 'ส่งทาง LINE', link: 'คัดลอกลิงก์', pdf: 'พิมพ์ / บันทึก PDF' }[ui.ch];
      return `<div class="doc-wrap">
  <div class="doc-area"><div class="row no-print" style="max-width:820px;margin:0 auto 14px"><a class="link" href="#/quote/${qid(q.id)}">${ICON.back} กลับไปแก้ไข</a><span class="spacer"></span><button class="btn ghost sm" data-act="docPrint">พิมพ์ / PDF</button></div>
    <div class="doc-scroll">${App.V.doc.doc(q, ui)}</div></div>
  <aside class="doc-side no-print">
    <h2 style="font-size:22px;margin-bottom:18px">ส่งใบเสนอราคา</h2>
    <div class="grid g4" style="gap:8px;grid-template-columns:repeat(4,1fr)">${[['email', 'อีเมล'], ['line', 'LINE'], ['link', 'ลิงก์'], ['pdf', 'PDF']].map(([k, l]) => `<button class="btn ${ui.ch === k ? 'dark' : 'ghost'}" style="border-radius:12px;padding:0 6px" data-act="docCh" data-v="${k}">${l}</button>`).join('')}</div>
    ${ui.ch === 'email' ? `<label class="f" style="margin-top:20px">ถึง</label><div class="mail-to">${ui.to.map((e, i) => `<span>${H(e)}<button data-act="docToDel" data-i="${i}" aria-label="ลบ">${ICON.x}</button></span>`).join('')}<input type="email" placeholder="เพิ่มอีเมล…" data-change="docToAdd" aria-label="เพิ่มอีเมล"></div>` : ''}
    <label class="f" style="margin-top:18px">ข้อความ</label><textarea class="input" rows="5" data-change="docMsg">${H(ui.msg)}</textarea>
    <div class="perm" style="border:0;margin-top:8px"><span>ให้ลูกค้ากดยืนยันและเซ็นออนไลน์</span>${App.sw('ยืนยันออนไลน์', ui.online, 'data-change="docT" data-f="online"')}</div>
    <div class="perm" style="border:0"><span>แสดงราคาต่อหน่วย</span>${App.sw('ราคาต่อหน่วย', ui.unit, 'data-change="docT" data-f="unit"')}</div>
    <div class="perm" style="border:0"><span>ยืนราคาถึง</span><input type="date" class="input sm" style="width:auto" value="${H(q.validUntil || '')}" data-change="docValid"></div>
    ${pending ? '<div class="note red" style="margin:10px 0">ส่วนลดเกิน 10% รออนุมัติจากผู้จัดการฝ่ายขาย จึงยังส่งไม่ได้</div>' : ''}
    <button class="btn red block" style="margin-top:14px" data-act="docSend" ${pending && ui.ch !== 'pdf' ? 'disabled' : ''}>${chLabel}</button>
    <div class="hr"></div>
    <h3 style="margin-bottom:8px">ประวัติ</h3>
    ${(q.history || []).slice(0, 8).map((h, i) => `<div class="row small" style="padding:5px 0;align-items:flex-start"><i class="dot" style="margin-top:6px;background:${i ? '#cfccc2' : 'var(--teal)'}"></i><span class="muted">${H(h.t)} — ${U.fmtD(h.at.slice(0, 10), { year: false })} ${h.at.slice(11, 16)}</span></div>`).join('') || '<div class="small muted">ยังไม่มีประวัติ</div>'}
  </aside></div>`;
    },
    doc(q, ui) {
      const S = Store.db.settings, c = App.calc(q), owner = Store.db.users.find(u => u.id === q.owner) || App.me();
      const unit = ui.unit !== false;
      const period = it => {
        if (it.sec === 'room') return `IN ${U.d(q.start).getDate()} OUT ${U.fmtD(U.addDays(q.start, Number(it.q2) || 1))}<br>จำนวน ${it.q2 || 1} คืน`;
        if (it.sched && it.sched.length === 1 && q.end !== q.start) return 'คืนวันที่ ' + it.sched[0].day + ' ' + U.fmtD(U.addDays(q.start, it.sched[0].day - 1), { year: false });
        return U.fmtRange(q.start, q.end);
      };
      const unitTxt = it => it.sec === 'room' ? 'บาท/ห้อง' : it.sec === 'meet' ? 'บาท' : it.q2 != null ? 'บาท/ท่าน/' + (it.u2 || 'มื้อ') : 'บาท/' + (it.u1 || 'ท่าน');
      const qtyTxt = it => it.sec === 'room' ? it.q1 + ' ห้อง' : it.sec === 'meet' ? '' : U.n(it.q1) + ' ' + (it.u1 || 'ท่าน');
      const row = it => `<tr><td>- ${H(it.name)}${it.desc ? `<br><span style="color:#666">${H(it.desc)}</span>` : ''}</td><td class="c">${period(it)}</td><td class="c">${qtyTxt(it)}</td>
        <td class="c">${unit ? (it.sec === 'meet' ? U.n(it.price) + ' บาท × ' + it.q1 + ' ' + (it.u1 || 'วัน') : U.n(it.price) + ' ' + unitTxt(it)) : ''}</td><td class="r"><b>${it.free ? '<span style="color:#c00">แถม</span>' : U.n(U.lineTotal(it)) + ' บาท'}</b></td></tr>`;
      const groups = [['room', 'ห้องพัก'], ['food', 'อาหารและเครื่องดื่ม'], ['meet', 'ห้องประชุม ณ โรงแรมลักษวรรณรีสอร์ท'], ['other', 'บริการอื่น ๆ']].filter(([k]) => q.items.some(i => i.sec === k));
      const mainRoom = D.MEETING_ROOMS.find(r => q.items.some(i => i.sec === 'meet' && i.ref === r.id));
      return `<article class="doc">
  <div style="display:flex;justify-content:space-between;gap:20px;align-items:flex-start">
    <div><b>${H(S.company)}</b><br>${H(S.address)}<br>ติดต่อโรงแรม ${H(S.phone)} / ติดต่อสำนักงานใหญ่ ${H(S.hq)}<br>สาขา <span style="background:#cfe6fb;padding:1px 40px 1px 8px;color:#1a3fd8">ลักษวรรณ รีสอร์ต</span></div>
    <img src="img/logo-wide.png" alt="Luck Swan Resort" style="width:170px;flex:none">
  </div>
  <div class="blue b" style="margin:16px 0 10px;font-size:13px">Quotation<br>ใบเสนอราคา</div>
  <div style="display:grid;grid-template-columns:1.3fr 1fr 1fr 1fr;gap:6px 14px;margin-bottom:14px">
    <b>ผู้รับบริการ</b><b>ชื่อโครงการ/งาน</b><b>เสนอราคาโดย</b><b style="text-align:right">เลขที่ใบเสนอราคา</b>
    <span>เรียน : <span class="it">${H(q.contact || '-')}</span></span><span class="it">${H(q.title)}</span><span class="it">${H(owner ? owner.name : '')}</span><span class="it" style="text-align:right">${H(q.id)}</span>
    <span>ชื่อ บริษัท/หน่วยงาน : <span class="it">${H(q.clientFull || q.client)}</span></span><b>เบอร์ติดต่อ ผู้ประสานงาน</b><b style="text-align:center">วันที่จัดงาน</b><b style="text-align:right">วันที่เสนอราคา</b>
    <span>ที่อยู่ : <span class="it">${H(q.address || '-')}</span><br>เลขผู้เสียภาษี : <span class="it">${H(q.taxId || '-')}</span></span><span class="it">${H(q.phone || '-')}</span><span class="it" style="text-align:center">${U.fmtRange(q.start, q.end, { long: true })}</span><span class="it" style="text-align:right">${U.fmtD((q.updatedAt || U.today()).slice(0, 10), { long: true })}</span>
  </div>
  <table><thead><tr><th>คำอธิบาย/รายละเอียด</th><th class="c">ระยะเวลาการใช้บริการ</th><th class="c">จำนวน หน่วย</th><th class="c">${unit ? 'ราคาต่อหน่วย' : ''}</th><th class="r">ราคารวม</th></tr></thead><tbody>
  ${groups.map(([k, name], gi) => `<tr class="grp"><td colspan="5">${gi + 1}. ${name}</td></tr>` +
    (k === 'meet' && mainRoom && mainRoom.img ? `<tr><td colspan="5"><div style="display:flex;gap:14px;align-items:flex-start"><img src="${mainRoom.img}" alt="" style="width:150px;border-radius:2px"><div style="flex:1"><table>${q.items.filter(i => i.sec === k).map(it => `<tr><td style="padding:3px 0"><b>${H(it.name)}</b>${it.desc ? '<br>' + H(it.desc) : ''}</td><td class="r" style="padding:3px 0;white-space:nowrap">${unit ? U.n(it.price) + ' บาท × ' + it.q1 + ' ' + H(it.u1 || '') : ''}</td><td class="r" style="padding:3px 0 3px 12px;white-space:nowrap"><b>${U.n(U.lineTotal(it))} บาท</b></td></tr>`).join('')}</table>
      <div style="margin-top:6px;color:#444">ให้บริการ ดังนี้<br>- โต๊ะพร้อมเก้าอี้ (รวมผ้าคลุม) สำหรับผู้เข้าร่วมประชุม<br>- โต๊ะลงทะเบียนด้านหน้าห้อง<br>- ${H(mainRoom.amen.join(' / '))}<br>- น้ำดื่มบริการ และพนักงานบริการ</div></div></div></td></tr>` :
      q.items.filter(i => i.sec === k).map(row).join(''))).join('')}
  </tbody></table>
  <div style="display:flex;justify-content:flex-end;margin-top:12px"><table style="width:auto;min-width:260px">
    <tr><td>ยอดรวมก่อนส่วนลด</td><td class="r">${U.n2(c.sub)}</td></tr>
    ${c.disc ? `<tr><td>ส่วนลด ${q.discount.type === 'amt' ? '' : U.n(c.pct, c.pct % 1 ? 1 : 0) + '%'}</td><td class="r">−${U.n2(c.disc)}</td></tr>` : ''}
    <tr><td>ค่าบริการ (Service Charge) ${Math.round(S.service * 100)}%</td><td class="r">${U.n2(c.svc)}</td></tr>
    <tr><td>ภาษีมูลค่าเพิ่ม (VAT) ${Math.round(S.vat * 100)}%</td><td class="r">${U.n2(c.vat)}</td></tr></table></div>
  <div style="display:flex;justify-content:space-between;gap:12px;border-top:1px solid #999;margin-top:10px;padding-top:10px;flex-wrap:wrap"><b>ราคารวม ภาษีมูลค่าเพิ่ม(Vat) และค่าบริการ(Service Charge) แล้ว</b>
    <div style="text-align:right" class="blue"><b style="font-size:14px">${U.n2(c.net)} บาท</b><br><b>(${U.bahtText(c.net)})</b></div></div>
  <div style="display:grid;grid-template-columns:110px 1fr;gap:8px;margin-top:18px;color:#444;font-size:11px">
    <span>เงื่อนไขการจองห้องพัก และการชำระเงิน:</span>
    <div>- เพื่อเป็นการยืนยันการจัดงาน กรุณา เซ็นยืนยันใช้บริการในใบเสนอราคาและ/หรือ ชำระมัดจำ ${H(q.deposit || '30-50%')} จากยอดรวมตามใบเสนอราคาที่ ${H(S.bank)}<br>
    - เมื่อทำการชำระเงินเรียบร้อยแล้ว นำส่งหลักฐานการชำระเงินมาที่ ${H(S.email)} หรือ Line Official : ${H(S.line)}<br>
    - หากมีการเปลี่ยนแปลงรายละเอียดการจัดงาน สามารถแจ้งที่เจ้าหน้าที่ฝ่ายขายล่วงหน้าอย่างน้อย 7 วัน<br>
    - กรณียกเลิกกรุณาแจ้งล่วงหน้าอย่างน้อย 14 วัน ก่อนวันงาน หากแจ้งยกเลิกหลังจากนั้นต้องทำการชำระเงิน 100% ของยอดใบเสนอราคา<br>
    - ยืนราคาถึงวันที่ ${U.fmtD(q.validUntil, { long: true })}${q.holdUntil ? ' · กันห้องประชุมไว้ชั่วคราวถึงวันที่ ' + U.fmtD(q.holdUntil, { long: true }) : ''}${q.items.some(i => i.sec === 'room') ? ' · วันปล่อยห้องพัก (cut-off) ' + U.fmtD(U.addDays(q.start, -14), { long: true }) : ''}
    ${ui.online !== false ? '<br>- ลูกค้าสามารถยืนยันและลงนามออนไลน์ผ่านลิงก์ที่แนบในอีเมล' : ''}</div></div>
  <div class="sig">
    <div>(.................................................)<br>ผู้ยืนยันการจัดงาน/ผู้ชำระค่าบริการ<br>วันที่ ........./............./.................</div>
    <div>(.....${H(S.gm)}.....)<br>ผู้จัดการทั่วไป<br>ผู้อนุมัติใบเสนอราคา ในนามบจก.ลักษวรรณ<br>วันที่ ........./............./.................</div>
    <div>(${H(owner ? owner.name : '')})<br>พนักงานขายและการตลาด<br>Tel. ${H(owner && owner.phone || S.phone)}<br>ผู้เสนอราคา ในนามบจก.ลักษวรรณ<br>${U.fmtD((q.updatedAt || U.today()).slice(0, 10), { long: true })}</div>
  </div>
</article>`;
    }
  };
  const DU = () => App.state.ui.doc;
  Object.assign(App.ACT, {
    docCh(el) { DU().ch = el.dataset.v; App.render(); },
    docToDel(el) { DU().to.splice(Number(el.dataset.i), 1); App.render(); },
    docPrint() { window.print(); },
    docSend() {
      const q = Store.quote(App.state.params[0]), ui = DU(), c = App.calc(q);
      const subject = `ใบเสนอราคา ${q.id} · ${q.title} · ลักษวรรณ รีสอร์ต`;
      const summary = `${ui.msg}\n\nเลขที่: ${q.id}\nงาน: ${q.title}\nวันที่: ${U.fmtRange(q.start, q.end)} · ${q.pax} ท่าน\nยอดสุทธิ: ${U.n2(c.net)} บาท (รวม Service และ VAT)\nยืนราคาถึง: ${U.fmtD(q.validUntil)}`;
      let note = '';
      if (ui.ch === 'email') {
        if (!ui.to.length) { App.toast('กรุณาระบุอีเมลผู้รับ'); return; }
        const [to, ...cc] = ui.to;
        location.href = `mailto:${encodeURIComponent(to)}?${cc.length ? 'cc=' + encodeURIComponent(cc.join(',')) + '&' : ''}subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(summary + '\n\n(แนบไฟล์ PDF ใบเสนอราคา)')}`;
        note = 'ส่งอีเมลถึง ' + ui.to.join(', ');
        App.toast('เปิดโปรแกรมอีเมลแล้ว — แนบไฟล์ PDF จากปุ่ม “พิมพ์ / PDF”');
      } else if (ui.ch === 'line') {
        window.open('https://line.me/R/msg/text/?' + encodeURIComponent(summary), '_blank', 'noopener');
        note = 'ส่งทาง LINE';
      } else if (ui.ch === 'link') {
        const url = location.href.split('#')[0] + '#/doc/' + qid(q.id);
        (navigator.clipboard ? navigator.clipboard.writeText(url) : Promise.reject()).then(() => App.toast('คัดลอกลิงก์แล้ว'), () => App.modal(`<h2>ลิงก์ใบเสนอราคา</h2><input class="input" value="${H(url)}" readonly onclick="this.select()"><div class="acts"><button class="btn dark" data-act="closeModal">ปิด</button></div>`));
        note = 'คัดลอกลิงก์ใบเสนอราคา';
      } else { window.print(); note = 'พิมพ์ / บันทึก PDF'; }
      if (q.status === 'draft') q.status = 'sent';
      const ver = (q.history || []).filter(h => /^เวอร์ชัน/.test(h.t)).length + 1;
      App.touch(q, `เวอร์ชัน ${ver} · ${note}`);
      setTimeout(() => App.render(), 50);
    }
  });
  Object.assign(App.CHG, {
    docToAdd(el) { const v = el.value.trim(); if (v && /.+@.+/.test(v)) { DU().to.push(v); App.render(); } else if (v) App.toast('อีเมลไม่ถูกต้อง'); },
    docMsg(el) { DU().msg = el.value; },
    docT(el) { DU()[el.dataset.f] = el.checked; App.render(); },
    docValid(el) { const q = Store.quote(App.state.params[0]); q.validUntil = el.value; App.touch(q); App.render(); }
  });
})(window);
