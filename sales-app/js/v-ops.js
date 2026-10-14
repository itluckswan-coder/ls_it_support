/* Luck Swan Sales — operations: function sheet (1j), calendar (1k), staff (1l), users (1m) */
(function (g) {
  'use strict';
  const U = g.U, D = g.DATA, Store = g.Store, App = g.App, ICON = App.ICON;
  const H = U.esc;
  const dept = id => D.DEPTS.find(d => d.id === id) || { name: id, short: id, color: '#999', bg: '#eee', fg: '#333' };
  const tag = id => { const d = dept(id); return `<span class="dept-tag" style="background:${d.bg};color:${d.fg}">${H(d.short)}</span>`; };

  /* =====================================================================
     FUNCTION SHEET / BEO (1j)
     ===================================================================== */
  App.makeBEO = function (q) {
    const days = U.daysOf(q.start, q.end), schedule = {};
    days.forEach((d, i) => schedule[i + 1] = []);
    const push = (day, time, task, place, dp) => { (schedule[day] = schedule[day] || []).push({ time, task, place, dept: dp }); };
    const mr = q.items.find(i => i.sec === 'meet' && D.MEETING_ROOMS.some(r => r.id === i.ref));
    const room = mr && D.MEETING_ROOMS.find(r => r.id === mr.ref);
    for (const it of q.items) for (const s of it.sched || []) {
      if (it.sec === 'meet' && room && it === mr) {
        push(s.day, U.fromHm(Math.max(5, U.hm(s.start) - 2.5)), `จัดห้อง ${(D.LAYOUTS.find(l => l.id === it.layout) || {}).name || ''} ${q.pax} ที่ + โต๊ะลงทะเบียน`, room.name, 'banquet');
        push(s.day, U.fromHm(Math.max(5, U.hm(s.start) - 1.5)), 'ทดสอบเสียงและภาพ', room.name, 'av');
        push(s.day, U.fromHm(Math.max(5, U.hm(s.start) - 1)), 'เปิดลงทะเบียน', 'หน้าห้อง', 'banquet');
      } else if (it.sec === 'food') push(s.day, s.start, it.name + ` · ${U.n(it.q1)} ที่`, s.place || 'หน้าห้อง', 'kitchen');
    }
    const rooms = q.items.filter(i => i.sec === 'room' && !i.free);
    if (rooms.length) {
      push(1, '14:00', `เช็กอินกรุ๊ป ${rooms.reduce((a, r) => a + Number(r.q1), 0)} ห้อง`, 'ล็อบบี้', 'front');
      push(Math.min(days.length, 1 + (Number(rooms[0].q2) || 1)), '12:00', 'เช็กเอาต์กรุ๊ป · เก็บกุญแจ', 'ล็อบบี้', 'front');
    }
    Object.values(schedule).forEach(l => l.sort((a, b) => a.time.localeCompare(b.time)));
    const foods = q.items.filter(i => i.sec === 'food');
    const c = App.calc(q);
    const depts = {
      kitchen: { items: foods.map(f => `${f.name} · ${App.qtyLabel(f)}`).concat((q.dietary || []).length ? [(q.dietary || []).map(x => x.k + ' ' + x.n).join(' · ')] : []), owner: '', ack: false },
      banquet: { items: mr ? [`${(D.LAYOUTS.find(l => l.id === mr.layout) || {}).name || ''} ${q.pax} ที่`, 'โต๊ะลงทะเบียน + ป้ายต้อนรับ', 'น้ำดื่ม กระดาษ ปากกา ทุกที่นั่ง'] : [], owner: '', ack: false },
      av: { items: q.items.filter(i => i.ref && i.ref.startsWith('svc:')).map(i => i.name + ' · ' + i.desc), owner: '', ack: false },
      front: { items: rooms.map(r => `${r.name} ${r.q1} ห้อง × ${r.q2} คืน`), owner: '', ack: false },
      hk: { items: rooms.length ? [`เตรียม ${rooms.reduce((a, r) => a + Number(r.q1), 0)} ห้องก่อน 13:00`, 'ทำความสะอาดห้องประชุมหลังเลิก'] : ['ทำความสะอาดห้องประชุมหลังเลิก'], owner: '', ack: false },
      acc: { items: [`ยอดสุทธิ ${U.n2(c.net)} ฿`, 'มัดจำ ' + (q.deposit || '-'), 'ใบกำกับภาษีในนาม ' + (q.clientFull || q.client)], owner: '', ack: false }
    };
    const b = { id: Store.nextId('FS'), quoteId: q.id, rev: 0, status: 'draft', confirmPax: q.pax + ' ท่าน', payment: q.status === 'confirmed' ? 'มัดจำแล้ว' : 'รอมัดจำ', schedule, depts,
      note: (q.dietary || []).length ? 'ข้อจำกัดอาหาร: ' + q.dietary.map(x => x.k + ' ' + x.n).join(' / ') : '' };
    Store.db.beos.unshift(b); Store.save(); App.toast('สร้างใบฟังก์ชัน ' + b.id); return b;
  };

  App.V.beo = {
    title: 'ใบฟังก์ชัน',
    render(p) {
      const ui = App.state.ui.beo || (App.state.ui.beo = { f: 'soon', day: 1 });
      const today = U.today();
      const all = Store.db.beos.map(b => ({ b, q: Store.quote(b.quoteId) })).filter(x => x.q).sort((a, b) => a.q.start.localeCompare(b.q.start));
      const list = all.filter(({ b, q }) => ui.f === 'all' || (ui.f === 'draft' ? b.status === 'draft' : q.end >= today && U.diffDays(today, q.start) <= 60));
      const sel = all.find(x => x.b.id === p[0]) || list[0] || all[0];
      if (sel && !p[0]) { history.replaceState(null, '', '#/beo/' + encodeURIComponent(sel.b.id)); }
      const ackN = b => Object.values(b.depts || {}).filter(d => d.ack).length, depN = b => Object.keys(b.depts || {}).length;
      const left = `<aside>
        <div class="between" style="margin-bottom:14px"><h2 style="font-size:22px">ใบฟังก์ชัน</h2><button class="btn red sm" data-act="beoNew">+ สร้างจากใบเสนอ</button></div>
        <div class="seg" style="width:100%;margin-bottom:14px">${[['soon', 'ใกล้ถึง'], ['all', 'ทั้งหมด'], ['draft', 'ร่าง']].map(([k, l]) => `<button style="flex:1" class="${ui.f === k ? 'on' : ''}" data-act="beoF" data-v="${k}">${l}</button>`).join('')}</div>
        ${list.map(({ b, q }) => `<a class="beo-item ${sel && sel.b.id === b.id ? 'on' : ''}" href="#/beo/${encodeURIComponent(b.id)}" style="color:inherit">
          <div class="between xs"><span class="mono muted">${H(b.id)}</span><span class="muted">${U.fmtRange(q.start, q.end, { year: false })}</span></div>
          <div class="b" style="margin:4px 0 0">${H(q.title)}</div><div class="small muted" style="margin-bottom:6px">${H(q.client)}</div>
          <div class="between small"><span class="muted">${H(((D.MEETING_ROOMS.find(r => r.id === q.room) || {}).name) || '')} · ${q.pax} ท่าน</span>
          ${apprBadge(b) || (b.status === 'draft' ? '<span class="badge st-draft">ร่าง</span>' : ackN(b) === depN(b) ? '<span class="badge st-confirmed">ครบทุกแผนก</span>' : `<span class="badge st-sent">รอรับทราบ ${depN(b) - ackN(b)}</span>`)}</div></a>`).join('') || '<div class="empty small">ไม่มีใบฟังก์ชัน</div>'}
      </aside>`;
      if (!sel) return `<div class="split left">${left}<div class="empty">ยังไม่มีใบฟังก์ชัน — สร้างจากใบเสนอราคาที่ยืนยันแล้ว</div></div>`;
      const { b, q } = sel, days = U.daysOf(q.start, q.end);
      if (ui.day > days.length) ui.day = 1;
      const sched = (b.schedule[ui.day] || []).slice().sort((a, c) => a.time.localeCompare(c.time));
      const owner = Store.db.users.find(u => u.id === q.owner);
      const pending = Object.values(b.depts).filter(d => !d.ack).length;
      return `<div class="split left side260">${left}
  <div style="min-width:0">
    <div class="page-hd" style="align-items:flex-start;margin-bottom:16px"><div>
      <div class="row wrap small" style="gap:8px"><span class="mono muted">${H(b.id)}</span><a class="link" href="#/quote/${encodeURIComponent(q.id)}">อ้างอิง ${H(q.id)}</a>${b.rev ? `<span class="badge st-tentative">แก้ไขครั้งที่ ${b.rev}${b.revNote ? ' · ' + H(b.revNote) : ''}</span>` : ''}</div>
      <h1 style="margin-top:6px">${H(q.title)} · ${H(q.client)}</h1></div>
      <div class="row"><button class="btn ghost" data-act="docPrint">พิมพ์ / PDF</button><button class="btn red" data-act="beoSend" ${g.Approval.state(b) === 'approved' ? '' : 'disabled title="ต้องได้รับอนุมัติก่อน"'}>${b.status === 'draft' ? 'ส่งให้ ' + Object.keys(D.DEPTS).length + ' แผนก' : 'ส่งฉบับแก้ไข'}</button></div></div>
    ${apprBanner(b)}
    <section class="card facts" style="margin-bottom:16px">
      <div><div class="k">วันงาน</div><div class="v">${U.fmtRange(q.start, q.end)}</div></div>
      <div><div class="k">ยืนยันจำนวน</div><div class="v" data-act="beoEditF" data-f="confirmPax" style="cursor:pointer">${H(b.confirmPax)}</div></div>
      <div><div class="k">ห้องจัดงาน</div><div class="v">${H(((D.MEETING_ROOMS.find(r => r.id === q.room) || {}).name) || '-')}${(q.items.find(i => i.sec === 'meet' && i.layout) || {}).layout ? ' · ' + D.LAYOUTS.find(l => l.id === q.items.find(i => i.sec === 'meet' && i.layout).layout).name : ''}</div></div>
      <div><div class="k">ผู้ติดต่อลูกค้า</div><div class="v">${H(q.contact || '-')} ${q.phone ? `<a href="${App.tel(q.phone)}">${H(q.phone)}</a>` : ''}</div></div>
      <div><div class="k">ฝ่ายขาย</div><div class="v">${H(owner ? App.firstName(owner) : '-')}${owner && owner.phone ? ' · ' + H(owner.phone) : ''}</div></div>
      <div><div class="k">การชำระเงิน</div><div class="v" data-act="beoEditF" data-f="payment" style="cursor:pointer">${H(b.payment)}</div></div>
    </section>
    <section class="card" style="margin-bottom:16px;overflow:hidden">
      <div class="card-hd"><h2>กำหนดการ · วันที่ ${ui.day} ${U.DS[U.d(days[ui.day - 1]).getDay()]} ${U.fmtD(days[ui.day - 1], { year: false })}</h2>
        <div class="row">${days.length > 1 ? `<div class="seg">${days.map((d, i) => `<button class="${ui.day === i + 1 ? 'on' : ''}" data-act="beoDay" data-v="${i + 1}">วันที่ ${i + 1}</button>`).join('')}</div>` : ''}<button class="btn ghost sm no-print" data-act="beoImport">นำเข้าจากไฟล์</button><button class="btn ghost sm no-print" data-act="beoRow" data-i="-1">+ เพิ่ม</button></div></div>
      <table class="tbl sched resp"><tbody>${sched.map(s => { const i = b.schedule[ui.day].indexOf(s); return `<tr class="click" data-act="beoRow" data-i="${i}"><td class="b num" style="width:80px">${H(s.time)}</td><td>${H(s.task)}</td><td class="muted hide-m">${H(s.place)}</td><td class="r">${tag(s.dept)}</td></tr>`; }).join('') || '<tr><td><div class="empty small">ยังไม่มีกำหนดการ — กด “+ เพิ่ม”</div></td></tr>'}</tbody></table>
    </section>
    <div class="dcards" style="margin-bottom:16px">${D.DEPTS.map(dp => {
      const x = b.depts[dp.id] || { items: [], owner: '', ack: false };
      return `<section class="card dcard"><div class="between"><div class="row"><span class="dot" style="background:${dp.color};border-radius:3px;width:11px;height:11px"></span><h3 style="font-size:15px">${H(dp.name)}</h3></div>
        <button class="small b ${x.ack ? 'teal' : 'red'}" style="white-space:nowrap;flex:none" data-act="beoAck" data-v="${dp.id}" title="สลับสถานะรับทราบ">${x.ack ? ICON.check + ' รับทราบแล้ว' : 'รอรับทราบ'}</button></div>
        <ul>${x.items.map(t => `<li>${H(t)}</li>`).join('') || '<li class="muted">—</li>'}</ul>
        <div class="own between"><span>ผู้รับผิดชอบ: ${H(x.owner || '-')}</span><button class="link no-print" data-act="beoDept" data-v="${dp.id}">แก้ไข</button></div></section>`;
    }).join('')}</div>
    <div class="note" style="cursor:pointer" data-act="beoNote"><b>หมายเหตุพิเศษ:</b> ${H(b.note || 'แตะเพื่อเพิ่มหมายเหตุ')}</div>
    ${b.status !== 'draft' && pending ? `<div class="small muted" style="margin-top:10px">ส่งแล้ว · รอรับทราบ ${pending} แผนก</div>` : ''}
  </div></div>`;
    }
  };
  const curB = () => Store.db.beos.find(b => b.id === App.state.params[0]) || Store.db.beos[0];

  /* ---------- approval ---------- */
  const fmtAt = iso => iso ? U.fmtD(iso.slice(0, 10), { year: false }) + ' ' + iso.slice(11, 16) : '';
  function apprBadge(b) {
    const st = g.Approval.state(b);
    if (st === 'pending') return '<span class="badge st-sent">รออนุมัติ</span>';
    if (st === 'rejected') return '<span class="badge st-tentative">ถูกตีกลับ</span>';
    return '';
  }
  function apprBanner(b) {
    const a = b.approval || {}, st = g.Approval.state(b), me = App.me();
    if (st === 'pending') {
      const mine = g.Approval.canDecide(b, me);
      return `<div class="note" style="margin-bottom:16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap"><div style="flex:1;min-width:220px"><b>รออนุมัติ</b> — แก้ไขโดย ${H(a.requestedByName || '-')} ${fmtAt(a.requestedAt)}<div class="xs">${mine ? 'คุณเป็นผู้อนุมัติได้ — ตรวจการเปลี่ยนแปลงแล้วกดอนุมัติหรือตีกลับ' : 'ส่งให้แผนกได้หลังผู้มีตำแหน่งสูงกว่าอนุมัติ'}</div></div>
        <button class="btn ghost sm" data-act="beoDiff">ดูการเปลี่ยนแปลง</button>${mine ? '<button class="btn ghost sm" data-act="beoReject" style="color:var(--red)">ตีกลับ</button><button class="btn teal sm" data-act="beoApprove">อนุมัติ</button>' : ''}</div>`;
    }
    if (st === 'rejected') return `<div class="note red" style="margin-bottom:16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap"><div style="flex:1;min-width:220px"><b>ถูกตีกลับ</b> โดย ${H(a.byName || '-')} ${fmtAt(a.at)}${a.comment ? ' — “' + H(a.comment) + '”' : ''}<div class="xs">แก้ไขข้อมูลแล้วระบบจะส่งขออนุมัติใหม่อัตโนมัติ</div></div><button class="btn ghost sm" data-act="beoDiff">ดูการเปลี่ยนแปลง</button></div>`;
    if (a.by && !a.imported) return `<div class="small teal" style="margin:-6px 0 12px">${ICON.check} อนุมัติแล้วโดย ${H(a.byName || '-')} ${fmtAt(a.at)}${a.comment ? ' — “' + H(a.comment) + '”' : ''}</div>`;
    return '';
  }
  function beoDiff(base, cur) {
    const rows = [], add = (t, cls) => rows.push(`<li class="${cls}">${t}</li>`);
    if (!base) return '<p class="muted">ใบฟังก์ชันใหม่ ยังไม่เคยได้รับอนุมัติ — ตรวจข้อมูลทั้งหมดในหน้านี้</p>';
    for (const [k, label] of [['confirmPax', 'ยืนยันจำนวน'], ['payment', 'การชำระเงิน'], ['note', 'หมายเหตุพิเศษ']])
      if ((base[k] || '') !== (cur[k] || '')) add(`${label}: <s>${H(base[k] || '—')}</s> → <b>${H(cur[k] || '—')}</b>`, 'chg');
    const days = new Set([...Object.keys(base.schedule || {}), ...Object.keys(cur.schedule || {})]);
    const key = r => [r.time, r.task, r.place, r.dept].join('|');
    for (const d of [...days].sort()) {
      const A = (base.schedule || {})[d] || [], B = (cur.schedule || {})[d] || [];
      const ka = new Set(A.map(key)), kb = new Set(B.map(key));
      B.filter(r => !ka.has(key(r))).forEach(r => add(`วันที่ ${d} · เพิ่ม ${H(r.time)} ${H(r.task)} (${H(dept(r.dept).short)}${r.place ? ' · ' + H(r.place) : ''})`, 'plus'));
      A.filter(r => !kb.has(key(r))).forEach(r => add(`วันที่ ${d} · ลบ ${H(r.time)} ${H(r.task)} (${H(dept(r.dept).short)})`, 'minus'));
    }
    for (const dp of D.DEPTS) {
      const A = (base.depts || {})[dp.id] || { items: [], owner: '' }, B = (cur.depts || {})[dp.id] || { items: [], owner: '' };
      B.items.filter(x => !A.items.includes(x)).forEach(x => add(`${H(dp.short)} · เพิ่ม “${H(x)}”`, 'plus'));
      A.items.filter(x => !B.items.includes(x)).forEach(x => add(`${H(dp.short)} · ลบ “${H(x)}”`, 'minus'));
      if ((A.owner || '') !== (B.owner || '')) add(`${H(dp.short)} · ผู้รับผิดชอบ <s>${H(A.owner || '—')}</s> → <b>${H(B.owner || '—')}</b>`, 'chg');
    }
    return rows.length ? `<ul class="diff">${rows.join('')}</ul>` : '<p class="muted">ไม่มีความแตกต่างจากฉบับที่อนุมัติล่าสุด</p>';
  }
  async function decide(b, decision, comment) {
    const me = App.me();
    if (g.API && g.API.online) {
      try { const j = await g.API.req('POST', 'beo/' + encodeURIComponent(b.id) + '/' + decision, { comment }); g.API.applyDoc(j.doc); Store.write(); }
      catch (e) { App.toast(e.message); return; }
    } else {
      const a = b.approval || {}, at = U.nowTH();
      b.approval = decision === 'approve' ? { state: 'approved', by: me.id, byName: me.name, at, comment, requestedBy: a.requestedBy, requestedByName: a.requestedByName }
        : Object.assign({}, a, { state: 'rejected', by: me.id, byName: me.name, at, comment });
      g.API.beoPrev[b.id] = JSON.stringify(b); Store.save();
    }
    App.closeModal(); App.render(); App.toast(decision === 'approve' ? 'อนุมัติแล้ว — ส่งให้แผนกได้' : 'ตีกลับแล้ว');
  }
  Object.assign(App.ACT, {
    beoDiff() {
      const b = curB(), a = b.approval || {};
      App.modal(`<h2>การเปลี่ยนแปลงที่รออนุมัติ</h2><div class="small muted" style="margin-bottom:12px">${H(b.id)} · เทียบกับฉบับที่อนุมัติล่าสุด</div>${beoDiff(a.base, g.Approval.content(b))}
        <div class="acts">${g.Approval.canDecide(b, App.me()) ? '<button class="btn ghost" data-act="beoReject" style="color:var(--red)">ตีกลับ</button><button class="btn teal" data-act="beoApprove">อนุมัติ</button>' : '<button class="btn dark" data-act="closeModal">ปิด</button>'}</div>`);
    },
    beoApprove() {
      const b = curB();
      App.modal(`<h2>อนุมัติใบฟังก์ชัน ${H(b.id)}</h2><label class="f">ความเห็น (ไม่บังคับ)</label><textarea class="input" name="c" rows="3"></textarea>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn teal" id="ok">อนุมัติ</button></div>`, bg => bg.querySelector('#ok').onclick = () => decide(b, 'approve', App.formData(bg).c.trim()));
    },
    beoReject() {
      const b = curB();
      App.modal(`<h2>ตีกลับใบฟังก์ชัน ${H(b.id)}</h2><label class="f">เหตุผล / สิ่งที่ต้องแก้ *</label><textarea class="input" name="c" rows="3"></textarea><div class="small red" id="rjErr"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">ตีกลับ</button></div>`, bg => bg.querySelector('#ok').onclick = () => {
        const c = App.formData(bg).c.trim(); if (!c) { bg.querySelector('#rjErr').textContent = 'กรุณาระบุเหตุผล'; return; } decide(b, 'reject', c);
      });
    }
  });

  /* ---------- import schedule from .xlsx / .csv ---------- */
  const DEPT_WORDS = [
    ['kitchen', /ครัว|อาหาร|f\s*&\s*b|fnb|kitchen|chef|เชฟ/i],
    ['banquet', /จัดเลี้ยง|banquet|จัดห้อง/i],
    ['av', /ช่าง|เสียง|ภาพ|\bav\b|a\/v|technic|engineer/i],
    ['front', /ส่วนหน้า|ต้อนรับ|front|reception|fo\b/i],
    ['hk', /แม่บ้าน|housekeep|\bhk\b/i],
    ['acc', /บัญชี|การเงิน|account|finance/i]
  ];
  const HEAD = { day: /^(วัน(ที่)?|day|date)$/i, time: /เวลา|time/i, task: /รายละเอียด|งาน|task|activity|กิจกรรม/i, place: /สถานที่|place|location|ห้อง|venue/i, dept: /แผนก|ฝ่าย|dept|department/i };
  function parseDept(v) {
    const t = String(v || '').trim(); if (!t) return null;
    const d = D.DEPTS.find(x => x.id === t.toLowerCase() || x.name === t || x.short === t);
    if (d) return d.id;
    const m = DEPT_WORDS.find(([, re]) => re.test(t)); return m ? m[0] : null;
  }
  function parseTime(v) {
    if (typeof v === 'number') { const f = v % 1 || (v < 1 ? v : 0); if (v >= 0 && v < 1 || f) return U.fromHm(Math.round(f * 24 * 60) / 60); if (v >= 0 && v <= 24) return U.fromHm(v); return null; }
    const t = String(v || '').trim().replace(/\s*(น\.|นาฬิกา|hrs?)$/i, '').split(/\s*[-–~]\s*/)[0];
    let m = /^(\d{1,2})[:.](\d{2})$/.exec(t) || /^(\d{2})(\d{2})$/.exec(t) || /^(\d{1,2})$/.exec(t);
    if (!m) return null;
    const h = Number(m[1]), mi = Number(m[2] || 0);
    return h <= 24 && mi < 60 ? String(h).padStart(2, '0') + ':' + String(mi).padStart(2, '0') : null;
  }
  function parseDay(v, days, fallback) {
    if (v === '' || v == null) return fallback;
    if (typeof v === 'number') {
      if (v >= 1 && v <= days.length && Number.isInteger(v)) return v;
      if (v > 20000) { // Excel date serial
        const d = new Date(Math.round(v - 25569) * 864e5), i = days.indexOf(U.iso(new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())));
        return i >= 0 ? i + 1 : null;
      }
      return null;
    }
    const t = String(v).trim();
    let m = /^(?:วันที่|day)?\s*(\d{1,2})$/i.exec(t);
    if (m) { const n = Number(m[1]); return n >= 1 && n <= days.length ? n : null; }
    m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t);
    let y, mo, d;
    if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
    else if ((m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/.exec(t))) { d = +m[1]; mo = +m[2]; y = +m[3]; if (y < 100) y += 2500; }
    else {
      m = /^(\d{1,2})\s*([ก-๙.]+)\s*(\d{2,4})?$/.exec(t);
      if (m) { const k = U.MS.findIndex(x => x.replace(/\./g, '') === m[2].replace(/\./g, '')); const kl = U.ML.indexOf(m[2]); const mi = k >= 0 ? k : kl; if (mi < 0) return null; d = +m[1]; mo = mi + 1; y = m[3] ? +m[3] : U.d(days[0]).getFullYear(); if (y < 100) y += 2500; }
      else return null;
    }
    if (y > 2400) y -= 543;
    const i = days.indexOf(U.iso(new Date(y, mo - 1, d)));
    return i >= 0 ? i + 1 : null;
  }
  function parseSchedule(rows, days, fallbackDay) {
    rows = rows.filter(r => r && r.some(c => String(c).trim() !== ''));
    let map = { day: 0, time: 1, task: 2, place: 3, dept: 4 }, start = 0;
    const hi = rows.findIndex((r, i) => i < 5 && r.some(c => HEAD.time.test(String(c))) && r.some(c => HEAD.task.test(String(c))));
    if (hi >= 0) {
      map = {}; start = hi + 1;
      rows[hi].forEach((c, i) => { const t = String(c).trim(); for (const k of ['day', 'time', 'dept', 'place', 'task']) if (map[k] == null && HEAD[k].test(t)) { map[k] = i; break; } });
    }
    return rows.slice(start).map((r, i) => {
      const get = k => map[k] == null ? '' : (r[map[k]] == null ? '' : r[map[k]]);
      const out = { line: start + i + 1, raw: r, day: parseDay(get('day'), days, fallbackDay), time: parseTime(get('time')), task: String(get('task')).trim(), place: String(get('place')).trim(), dept: parseDept(get('dept')), errs: [] };
      if (!out.task) out.errs.push('ไม่มีรายละเอียดงาน');
      if (!out.time) out.errs.push('เวลาไม่ถูกต้อง');
      if (!out.day) out.errs.push('วันไม่อยู่ในช่วงงาน');
      if (!out.dept) out.errs.push(String(get('dept')).trim() ? 'ไม่รู้จักแผนก "' + String(get('dept')).trim() + '"' : 'ไม่ได้ระบุแผนก');
      return out;
    });
  }
  function templateRows(b) {
    const q = Store.quote(b.quoteId), days = U.daysOf(q.start, q.end);
    const rows = [['วัน', 'เวลา', 'รายละเอียดงาน', 'สถานที่', 'แผนก']];
    const cur = Object.entries(b.schedule).flatMap(([d, l]) => l.map(s => [Number(d), s.time, s.task, s.place, dept(s.dept).short]));
    if (cur.length) return rows.concat(cur.sort((a, c) => a[0] - c[0] || a[1].localeCompare(c[1])));
    const room = ((D.MEETING_ROOMS.find(r => r.id === q.room) || {}).name) || 'ห้องประชุม';
    return rows.concat([[1, '06:00', 'จัดห้อง ' + q.pax + ' ที่ + โต๊ะลงทะเบียน', room, 'จัดเลี้ยง'], [1, '07:00', 'ทดสอบเสียงและภาพ', room, 'ช่าง/AV'], [1, '10:00', 'Coffee break ' + q.pax + ' ที่', 'หน้าห้อง', 'ครัว'],
      [1, '12:00', 'อาหารกลางวัน', 'ห้องอาหาร', 'ครัว'], [1, '14:00', 'เช็กอินกรุ๊ป', 'ล็อบบี้', 'ส่วนหน้า']].concat(days.length > 1 ? [[2, '12:00', 'เช็กเอาต์กรุ๊ป', 'ล็อบบี้', 'ส่วนหน้า']] : []));
  }
  const IMPORT_HELP = `
    <ol style="margin:0 0 14px;padding-left:20px;line-height:1.8">
      <li>ดาวน์โหลดแม่แบบ (Excel หรือ CSV) — มีกำหนดการของงานนี้ใส่ไว้เป็นตัวอย่าง</li>
      <li>กรอกหรือแก้ใน Excel / Google Sheets ทีละ 1 แถวต่อ 1 รายการ แล้วบันทึกไฟล์</li>
      <li>กด “เลือกไฟล์” แล้วตรวจตัวอย่าง แถวสีแดงคือข้อมูลที่ต้องแก้ (จะไม่ถูกนำเข้า)</li>
      <li>เลือก “เพิ่มต่อจากเดิม” หรือ “แทนที่ทั้งหมด” แล้วกดนำเข้า</li>
    </ol>
    <div style="overflow-x:auto;border:1px solid var(--line);border-radius:10px;margin-bottom:8px"><table class="tbl compact" style="font-size:12.5px"><thead><tr><th>วัน</th><th>เวลา</th><th>รายละเอียดงาน</th><th>สถานที่</th><th>แผนก</th></tr></thead><tbody>
      <tr><td>1</td><td>06:00</td><td>จัดห้อง Classroom 80 ที่</td><td>Thikhampon</td><td>จัดเลี้ยง</td></tr>
      <tr><td>1</td><td>10:00</td><td>Coffee break 80 ที่</td><td>หน้าห้อง</td><td>ครัว</td></tr>
      <tr><td>2</td><td>12:00</td><td>เช็กเอาต์กรุ๊ป</td><td>ล็อบบี้</td><td>ส่วนหน้า</td></tr></tbody></table></div>
    <div class="xs muted" style="line-height:1.7"><b>วัน</b> = ลำดับวันของงาน (1, 2, …) หรือวันที่จริง เช่น 12/11/2569, 2026-11-12, 12 พ.ย. 2569 · เว้นว่าง = วันที่เปิดอยู่<br>
    <b>เวลา</b> = 06:00, 6.00 หรือ 0600 (ถ้าเป็นช่วง เช่น 10:00–10:30 จะใช้เวลาเริ่ม)<br>
    <b>แผนก</b> = ครัว, จัดเลี้ยง, ช่าง/AV, ส่วนหน้า, แม่บ้าน, บัญชี (ชื่อเต็มหรือภาษาอังกฤษก็ได้)<br>
    ลำดับคอลัมน์สลับกันได้ ถ้ามีแถวหัวตาราง · รองรับไฟล์ .xlsx และ .csv</div>`;

  App.ACT.beoImport = () => {
    const b = curB(), q = Store.quote(b.quoteId), days = U.daysOf(q.start, q.end), fallback = App.state.ui.beo.day || 1;
    let parsed = null;
    const bg = App.modal(`<h2>นำเข้ากำหนดการจากไฟล์</h2>
      <div id="impBody">${IMPORT_HELP}
        <div class="row wrap" style="margin:16px 0 12px"><button class="btn ghost sm" id="tplX">ดาวน์โหลดแม่แบบ Excel</button><button class="btn ghost sm" id="tplC">ดาวน์โหลดแม่แบบ CSV</button></div>
        <label class="btn dark block" style="cursor:pointer">เลือกไฟล์ (.xlsx / .csv)<input type="file" id="impF" accept=".xlsx,.csv,.txt,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" class="sr"></label>
        <div id="impMsg" class="small red" style="margin-top:10px" role="alert"></div></div>
      <div class="acts" id="impActs"><button class="btn ghost" data-act="closeModal">ปิด</button></div>`);
    bg.querySelector('.modal').style.maxWidth = '760px';
    const fname = 'กำหนดการ-' + b.id;
    bg.querySelector('#tplX').onclick = () => {
      const blob = g.Sheet.writeXlsx(templateRows(b), { sheetName: 'กำหนดการ', widths: [8, 10, 46, 22, 14] });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fname + '.xlsx'; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    };
    bg.querySelector('#tplC').onclick = () => U.download(fname + '.csv', '﻿' + templateRows(b).map(r => r.map(c => /[",\n]/.test(String(c)) ? '"' + String(c).replace(/"/g, '""') + '"' : c).join(',')).join('\r\n') + '\r\n', 'text/csv');
    bg.querySelector('#impF').onchange = async e => {
      const f = e.target.files[0]; if (!f) return;
      const msg = bg.querySelector('#impMsg'); msg.textContent = 'กำลังอ่านไฟล์…';
      try { parsed = parseSchedule(await g.Sheet.readFile(f), days, fallback); }
      catch (err) { msg.textContent = err.message || 'อ่านไฟล์ไม่สำเร็จ'; e.target.value = ''; return; }
      if (!parsed.length) { msg.textContent = 'ไม่พบข้อมูลในไฟล์'; return; }
      const ok = parsed.filter(r => !r.errs.length), bad = parsed.length - ok.length;
      bg.querySelector('#impBody').innerHTML = `
        <div class="between" style="margin-bottom:10px;flex-wrap:wrap"><b>${U.esc(f.name)}</b><span class="small"><span class="teal b">พร้อมนำเข้า ${ok.length} รายการ</span>${bad ? ` · <span class="red b">มีปัญหา ${bad} แถว</span>` : ''}</span></div>
        <div style="max-height:46vh;overflow:auto;border:1px solid var(--line);border-radius:12px"><table class="tbl compact" style="font-size:12.5px"><thead><tr><th>แถว</th><th>วัน</th><th>เวลา</th><th>รายละเอียดงาน</th><th>สถานที่</th><th>แผนก</th></tr></thead><tbody>
        ${parsed.map(r => `<tr style="${r.errs.length ? 'background:var(--red-tint)' : ''}"><td class="muted">${r.line}</td><td>${r.day || '<span class="red">?</span>'}</td><td>${r.time || '<span class="red">?</span>'}</td>
          <td>${U.esc(r.task) || '<span class="red">—</span>'}${r.errs.length ? `<div class="xs red">${r.errs.map(U.esc).join(' · ')}</div>` : ''}</td><td>${U.esc(r.place)}</td><td>${r.dept ? tag(r.dept) : '<span class="red">?</span>'}</td></tr>`).join('')}
        </tbody></table></div>
        <div class="row wrap" style="margin-top:14px;gap:18px"><label class="row small"><input type="radio" name="mode" value="append" checked style="accent-color:var(--teal);width:18px;height:18px"> เพิ่มต่อจากกำหนดการเดิม</label>
          <label class="row small"><input type="radio" name="mode" value="replace" style="accent-color:var(--red);width:18px;height:18px"> แทนที่กำหนดการเดิมทั้งหมด</label></div>`;
      bg.querySelector('#impActs').innerHTML = `<button class="btn ghost" id="impBack">เลือกไฟล์ใหม่</button><button class="btn red" id="impGo" ${ok.length ? '' : 'disabled'}>นำเข้า ${ok.length} รายการ</button>`;
      bg.querySelector('#impBack').onclick = () => { App.closeModal(); App.ACT.beoImport(); };
      bg.querySelector('#impGo').onclick = () => {
        const mode = (bg.querySelector('input[name=mode]:checked') || {}).value;
        if (mode === 'replace') Object.keys(b.schedule).forEach(k => b.schedule[k] = []);
        for (const r of ok) (b.schedule[r.day] = b.schedule[r.day] || []).push({ time: r.time, task: r.task, place: r.place, dept: r.dept });
        Object.values(b.schedule).forEach(l => l.sort((x, y) => x.time.localeCompare(y.time)));
        Store.save(); App.closeModal(); App.render();
        App.toast(`นำเข้า ${ok.length} รายการแล้ว${bad ? ' · ข้าม ' + bad + ' แถว' : ''}`);
      };
    };
  };
  Object.assign(App.ACT, {
    beoF(el) { App.state.ui.beo.f = el.dataset.v; App.go('#/beo'); },
    beoDay(el) { App.state.ui.beo.day = Number(el.dataset.v); App.render(); },
    beoNew() {
      App.pickQuoteModal('สร้างใบฟังก์ชันจากใบเสนอราคา', id => {
        const ex = Store.db.beos.find(b => b.quoteId === id);
        const b = ex || App.makeBEO(Store.quote(id)); App.go('#/beo/' + encodeURIComponent(b.id));
      }, q => q.status !== 'cancelled' && q.status !== 'draft');
    },
    beoSend() {
      const b = curB(), q = Store.quote(b.quoteId);
      if (g.Approval.state(b) !== 'approved') { App.toast('ใบฟังก์ชันนี้ต้องได้รับอนุมัติก่อนส่งให้แผนก'); return; }
      if (b.status !== 'draft') { b.rev = (b.rev || 0) + 1; b.revNote = ''; }
      b.status = 'sent'; Object.values(b.depts).forEach(d => d.ack = false);
      Store.save(); App.render();
      App.toast('ส่งใบฟังก์ชันให้ทุกแผนกแล้ว');
      const text = `ใบฟังก์ชัน ${b.id}${b.rev ? ' (แก้ไขครั้งที่ ' + b.rev + ')' : ''}\n${q.title} · ${q.client}\n${U.fmtRange(q.start, q.end)} · ${b.confirmPax}`;
      if (navigator.share) navigator.share({ title: b.id, text, url: location.href }).catch(() => { /* cancelled */ });
    },
    beoAck(el) { const b = curB(); const d = b.depts[el.dataset.v] || (b.depts[el.dataset.v] = { items: [], owner: '', ack: false }); d.ack = !d.ack; Store.save(); App.render(); },
    beoRow(el) {
      const b = curB(), ui = App.state.ui.beo, i = Number(el.dataset.i), list = b.schedule[ui.day] = b.schedule[ui.day] || [];
      const s = i >= 0 ? list[i] : { time: '08:00', task: '', place: '', dept: 'banquet' };
      App.modal(`<h2>${i >= 0 ? 'แก้ไข' : 'เพิ่ม'}กำหนดการ</h2>
        <div class="grid g2"><div><label class="f">เวลา</label><input class="input" type="time" name="time" value="${H(s.time)}"></div><div><label class="f">แผนก</label><select class="input" name="dept">${D.DEPTS.map(d => `<option value="${d.id}" ${s.dept === d.id ? 'selected' : ''}>${H(d.name)}</option>`).join('')}</select></div></div>
        <div class="field" style="margin-top:14px"><label class="f">รายละเอียดงาน</label><input class="input" name="task" value="${H(s.task)}"></div>
        <div class="field"><label class="f">สถานที่</label><input class="input" name="place" value="${H(s.place)}"></div>
        <div class="acts">${i >= 0 ? '<button class="btn ghost" id="del" style="color:var(--red);margin-right:auto">ลบ</button>' : ''}<button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`, bg => {
        bg.querySelector('#ok').onclick = () => { const f = App.formData(bg); if (!f.task) return; Object.assign(s, f); if (i < 0) list.push(s); Store.save(); App.closeModal(); App.render(); };
        const del = bg.querySelector('#del'); if (del) del.onclick = () => { list.splice(i, 1); Store.save(); App.closeModal(); App.render(); };
      });
    },
    beoDept(el) {
      const b = curB(), id = el.dataset.v, d = b.depts[id] || (b.depts[id] = { items: [], owner: '', ack: false });
      App.modal(`<h2>${H(dept(id).name)}</h2><div class="field"><label class="f">รายการสั่งงาน (บรรทัดละ 1 รายการ)</label><textarea class="input" name="items" rows="6">${H(d.items.join('\n'))}</textarea></div>
        <div class="field"><label class="f">ผู้รับผิดชอบ</label><input class="input" name="owner" value="${H(d.owner)}"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`,
      bg => bg.querySelector('#ok').onclick = () => { const f = App.formData(bg); d.items = f.items.split('\n').map(s => s.trim()).filter(Boolean); d.owner = f.owner; Store.save(); App.closeModal(); App.render(); });
    },
    beoNote() {
      const b = curB();
      App.modal(`<h2>หมายเหตุพิเศษ</h2><textarea class="input" name="note" rows="5">${H(b.note || '')}</textarea><div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`,
        bg => bg.querySelector('#ok').onclick = () => { b.note = App.formData(bg).note; Store.save(); App.closeModal(); App.render(); });
    },
    beoEditF(el) {
      const b = curB(), f = el.dataset.f;
      App.modal(`<h2>${f === 'payment' ? 'การชำระเงิน' : 'ยืนยันจำนวน'}</h2><input class="input" name="v" value="${H(b[f])}"><div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`,
        bg => bg.querySelector('#ok').onclick = () => { b[f] = App.formData(bg).v; Store.save(); App.closeModal(); App.render(); });
    }
  });

  /* =====================================================================
     CALENDAR (1k)
     ===================================================================== */
  const ST_COLOR = {
    draft: ['#EDECE6', '#4A524F', '#9a9a92'], sent: ['#FFF1C2', '#7a5c00', '#D4A21A'], tentative: ['#F8E1E1', '#8E1F25', '#A8262D'], confirmed: ['#DDF3EF', '#0B7C6F', '#0E9F8E']
  };
  App.V.cal = {
    title: 'ปฏิทิน',
    render() {
      const t = U.today();
      const ui = App.state.ui.cal || (App.state.ui.cal = { ym: t.slice(0, 7), view: 'month', area: 'all', sel: t });
      const [y, m] = ui.ym.split('-').map(Number);
      const first = new Date(y, m - 1, 1), startD = U.addDays(U.iso(first), -first.getDay());
      const qs = Store.db.quotes.filter(q => q.status !== 'cancelled' && (ui.area === 'all' || q.room === ui.area || q.items.some(i => i.ref === ui.area)));
      const inMonth = qs.filter(q => q.start.slice(0, 7) === ui.ym || q.end.slice(0, 7) === ui.ym);
      const stat = s => { const l = inMonth.filter(q => q.status === s); return [l.length, l.reduce((a, q) => a + App.calc(q).net, 0)]; };
      const evOn = d => qs.filter(q => d >= q.start && d <= q.end);
      const evBtn = q => { const c = ST_COLOR[q.status] || ST_COLOR.draft; return `<button class="e" style="background:${c[0]};color:${c[1]}" data-act="openQuote" data-id="${H(q.id)}" title="${H(q.client + ' · ' + q.title)}">${H(q.title.length > 14 ? q.client.replace(/^(บจก\.|บมจ\.|คุณ)\s*/, '') + ' ' + q.title.split(' ')[0] : q.title)} · ${H(((D.MEETING_ROOMS.find(r => r.id === q.room) || {}).name) || '')}</button>`; };
      const areas = [['all', 'ทั้งหมด']].concat(D.MEETING_ROOMS.map(r => [r.id, r.id === 'pool' ? 'สระว่ายน้ำ' : r.name]));
      let main = '';
      if (ui.view === 'month') {
        const cells = [];
        for (let i = 0; i < 42; i++) {
          const d = U.addDays(startD, i); if (i >= 35 && d.slice(0, 7) !== ui.ym) break;
          const ev = evOn(d);
          cells.push(`<div class="d ${d.slice(0, 7) !== ui.ym ? 'out' : ''} ${d === ui.sel ? 'sel' : ''} ${d === t ? 'today' : ''}" data-act="calSel" data-v="${d}"><span class="dn">${U.d(d).getDate()}</span>${ev.slice(0, 3).map(evBtn).join('')}${ev.length > 3 ? `<div class="xs muted" style="margin-top:3px">+${ev.length - 3} งาน</div>` : ''}</div>`);
        }
        main = `<div class="cal">${U.DS.map(d => `<div class="dh">${d}</div>`).join('')}${cells.join('')}</div>`;
      } else if (ui.view === 'week') {
        const ws = U.addDays(ui.sel, -U.d(ui.sel).getDay());
        main = `<div style="padding:8px 0">${[0, 1, 2, 3, 4, 5, 6].map(i => { const d = U.addDays(ws, i), ev = evOn(d); return `<div class="row" style="padding:12px 20px;border-bottom:1px solid var(--line);align-items:flex-start;${d === ui.sel ? 'background:var(--red-tint)' : ''}"><div style="width:90px;flex:none"><b>${U.DS[i]} ${U.d(d).getDate()}</b><div class="xs muted">${U.MS[U.d(d).getMonth()]}</div></div><div style="flex:1;display:grid;gap:6px">${ev.map(q => `<button class="pick" data-act="openQuote" data-id="${H(q.id)}" style="padding:8px 12px"><span><b style="font-weight:500">${H(q.client)}</b> · ${H(q.title)}<br><span class="small muted">${H(((D.MEETING_ROOMS.find(r => r.id === q.room) || {}).name) || '')} · ${q.pax} ท่าน</span></span>${App.statusBadge(q.status)}</button>`).join('') || '<span class="small muted">—</span>'}</div></div>`; }).join('')}</div>`;
      } else {
        main = `<div class="grid g4" style="padding:16px;align-items:start">${['draft', 'sent', 'tentative', 'confirmed'].map(s => { const l = qs.filter(q => q.status === s && q.end >= t).sort((a, b) => a.start.localeCompare(b.start)); return `<div><div class="between" style="margin-bottom:10px"><b>${D.STATUS[s]}</b><span class="small muted">${l.length} · ${U.short(l.reduce((a, q) => a + App.calc(q).net, 0))}</span></div>${l.map(q => `<button class="pick" style="display:block;border-left:4px solid ${ST_COLOR[s][2]}" data-act="openQuote" data-id="${H(q.id)}"><b style="font-weight:500">${H(q.client)}</b><div class="small muted">${H(q.title)} · ${U.fmtRange(q.start, q.end, { year: false })}</div><div class="small b num">฿${U.n(Math.round(App.calc(q).net))}</div></button>`).join('') || '<div class="small muted">—</div>'}</div>`; }).join('')}</div>`;
      }
      const selEv = evOn(ui.sel);
      return `<div class="split left side260">
  <aside>
    <h1 style="margin-bottom:16px">ปฏิทินการจอง</h1>
    ${['draft', 'sent', 'tentative', 'confirmed'].map(s => { const [n, v] = stat(s); return `<div class="stat-row"><span class="sq" style="border-color:${ST_COLOR[s][2]};background:${ST_COLOR[s][0]}"></span><span style="flex:1">${D.STATUS[s]}</span><b style="font:600 18px var(--fh);width:36px;text-align:center">${n}</b><span class="xs muted" style="width:44px;text-align:right">${U.short(v)}</span></div>`; }).join('')}
    <div class="small muted" style="margin:18px 0 8px">พื้นที่</div>
    <div class="row wrap" style="gap:6px">${areas.map(([k, l]) => `<button class="chip ${ui.area === k ? 'on' : ''}" data-act="calArea" data-v="${k}">${H(l)}</button>`).join('')}</div>
    ${selEv.length ? `<div class="card pad" style="margin-top:14px"><b>${U.fmtD(ui.sel, { dow: true })}</b>${selEv.map(q => `<a class="row" style="padding:8px 0;color:inherit;border-top:1px solid var(--line);margin-top:8px" href="#/quote/${encodeURIComponent(q.id)}"><i class="dot" style="background:${(ST_COLOR[q.status] || ST_COLOR.draft)[2]}"></i><span class="small" style="flex:1">${H(q.client)} · ${H(q.title)}</span></a>`).join('')}</div>` : ''}
  </aside>
  <section class="card" style="overflow:hidden;min-width:0">
    <div class="between" style="padding:16px 20px;border-bottom:1px solid var(--line);flex-wrap:wrap">
      <div class="row"><button class="icon-btn" data-act="calNav" data-v="-1" aria-label="เดือนก่อนหน้า">${ICON.left}</button><h2 style="font-size:20px;min-width:170px;text-align:center">${U.ML[m - 1]} ${U.be(y)}</h2><button class="icon-btn" data-act="calNav" data-v="1" aria-label="เดือนถัดไป">${ICON.right}</button><button class="link" data-act="calToday">วันนี้</button></div>
      <div class="seg">${[['week', 'สัปดาห์'], ['month', 'เดือน'], ['pipe', 'ไปป์ไลน์']].map(([k, l]) => `<button class="${ui.view === k ? 'on' : ''}" data-act="calView" data-v="${k}">${l}</button>`).join('')}</div>
    </div>
    ${main}
  </section></div>`;
    }
  };
  Object.assign(App.ACT, {
    calNav(el) { const ui = App.state.ui.cal; const [y, m] = ui.ym.split('-').map(Number); const d = new Date(y, m - 1 + Number(el.dataset.v), 1); ui.ym = U.iso(d).slice(0, 7); ui.sel = ui.ym + '-01'; App.render(); },
    calToday() { const ui = App.state.ui.cal; ui.sel = U.today(); ui.ym = ui.sel.slice(0, 7); App.render(); },
    calView(el) { App.state.ui.cal.view = el.dataset.v; App.render(); },
    calArea(el) { App.state.ui.cal.area = el.dataset.v; App.render(); },
    calSel(el, e) { if (e.target.closest('.e')) return; const ui = App.state.ui.cal; ui.sel = el.dataset.v; if (ui.sel.slice(0, 7) !== ui.ym) ui.ym = ui.sel.slice(0, 7); App.render(); }
  });

  /* =====================================================================
     STAFF SHIFTS (1l)
     ===================================================================== */
  const PENCIL = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>';
  const TRASH = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>';
  const EMP = id => Store.db.employees.find(e => e.id === id);
  // staffing rules by guest count: { dept: [{ upTo: 50, n: 1 }, …, { upTo: null, n: 4 }] } — n people per shift
  const DEFAULT_RULES = { kitchen: [{ upTo: 50, n: 1 }, { upTo: 100, n: 2 }, { upTo: 150, n: 3 }, { upTo: null, n: 4 }] };
  const staffRules = () => (Store.db.settings && Store.db.settings.staffRules) || DEFAULT_RULES;
  const ruleNeed = (dp, pax) => {
    const tiers = staffRules()[dp]; if (!tiers || !tiers.length) return null;
    const t = tiers.find(x => x.upTo == null || pax <= x.upTo) || tiers[tiers.length - 1];
    return Math.max(0, Number(t.n) || 0);
  };
  function genPlan(q) {
    const plan = genPlanBase(q);
    for (const day of Object.values(plan.days)) for (const [dp, shifts] of Object.entries(day)) { const n = ruleNeed(dp, q.pax); if (n != null) shifts.forEach(s => { s.need = n; }); }
    return plan;
  }
  function genPlanBase(q) {
    const days = {}, pax = q.pax;
    const need = { kitchen: Math.max(2, Math.ceil(pax / 10)), banquet: Math.max(2, Math.ceil(pax / 12)), av: q.items.some(i => i.ref === 'svc:av') ? 2 : 1, front: q.items.some(i => i.sec === 'room') ? 4 : 2, hk: q.items.some(i => i.sec === 'room') ? Math.max(2, Math.ceil(q.items.filter(i => i.sec === 'room').reduce((a, r) => a + Number(r.q1), 0) / 6)) : 2 };
    U.daysOf(q.start, q.end).forEach(d => {
      let n = 0; const sid = () => 'g' + U.uid();
      days[d] = {
        kitchen: [{ id: sid(), type: 'morning', label: 'รอบเช้า', start: '06:00', end: '14:00', need: Math.ceil(need.kitchen / 2), staff: [] }, { id: sid(), type: 'evening', label: 'รอบบ่าย', start: '14:00', end: '22:00', need: Math.floor(need.kitchen / 2), staff: [] }],
        banquet: [{ id: sid(), type: 'morning', label: 'รอบเช้า', start: '06:00', end: '14:00', need: Math.ceil(need.banquet / 2), staff: [] }, { id: sid(), type: 'evening', label: 'รอบบ่าย', start: '14:00', end: '22:00', need: Math.floor(need.banquet / 2), staff: [] }],
        av: [{ id: sid(), type: 'morning', label: 'รอบเต็มวัน', start: '07:00', end: '17:00', need: need.av, staff: [] }],
        front: [{ id: sid(), type: 'morning', label: 'รอบเช้า', start: '06:00', end: '14:00', need: Math.ceil(need.front / 2), staff: [] }, { id: sid(), type: 'evening', label: 'รอบบ่าย', start: '14:00', end: '22:00', need: Math.floor(need.front / 2), staff: [] }],
        hk: [{ id: sid(), type: 'morning', label: 'รอบกลางวัน', start: '08:00', end: '16:00', need: need.hk, staff: [] }]
      }; n++;
    });
    return { days, ot: [] };
  }
  const otHours = (empId, plan) => (plan ? plan.ot : []).filter(o => o.status !== 'rejected' && o.staff.includes(empId)).reduce((a, o) => a + o.hours, 0);

  App.V.staff = {
    title: 'พนักงาน',
    render(p) {
      const tab = App.state.ui.staffTab || 'shifts';
      const tabs = `<div class="seg white" style="margin-bottom:18px" role="tablist"><button role="tab" class="${tab === 'shifts' ? 'on' : ''}" data-act="stTab" data-v="shifts">รอบพนักงาน</button><button role="tab" class="${tab === 'emps' ? 'on' : ''}" data-act="stTab" data-v="emps">พนักงานทั้งหมด</button><button role="tab" class="${tab === 'rules' ? 'on' : ''}" data-act="stTab" data-v="rules">เงื่อนไขจำนวนพนักงาน</button></div>`;
      return tabs + (tab === 'emps' ? App.V.staff.renderEmps() : tab === 'rules' ? App.V.staff.renderRules() : App.V.staff.renderShifts(p));
    },
    /* ---- staffing rules: people needed per shift by guest count ---- */
    renderRules() {
      const canEdit = App.can('staff'), ui = App.state.ui;
      if (!ui.rulesDraft) ui.rulesDraft = JSON.parse(JSON.stringify(staffRules()));
      const R = ui.rulesDraft, dis = canEdit ? '' : 'disabled';
      const card = dp => {
        const tiers = R[dp.id];
        const body = !tiers ? '<p class="small muted" style="margin:6px 0 0">ใช้การคำนวณอัตโนมัติของระบบ (ตามจำนวนลูกค้าและรายการในใบเสนอ)</p>' : `
          <table class="tbl" style="margin-top:8px"><thead><tr><th>จำนวนลูกค้าต่องาน</th><th class="r">พนักงานต่อรอบ</th>${canEdit ? '<th></th>' : ''}</tr></thead><tbody>
          ${tiers.map((t, i) => { const from = i ? (Number(tiers[i - 1].upTo) || 0) + 1 : 1, last = i === tiers.length - 1; return `<tr>
            <td><div class="row" style="gap:6px;flex-wrap:nowrap">${last ? `<span>${from} คนขึ้นไป</span>` : `<span style="white-space:nowrap">${from} –</span><input class="input" style="width:76px;height:36px" type="number" min="${from}" value="${t.upTo}" data-change="ruleUp" data-d="${dp.id}" data-i="${i}" ${dis} aria-label="ถึงจำนวนลูกค้า"><span>คน</span>`}</div></td>
            <td class="r"><span class="row" style="gap:6px;flex-wrap:nowrap;justify-content:flex-end"><input class="input" style="width:70px;height:36px" type="number" min="0" value="${t.n}" data-change="ruleN" data-d="${dp.id}" data-i="${i}" ${dis} aria-label="จำนวนพนักงาน"><span>คน</span></span></td>
            ${canEdit ? `<td class="r">${tiers.length > 1 ? `<button class="icon-btn us-edit emp-del" data-act="ruleDel" data-d="${dp.id}" data-i="${i}" title="ลบช่วงนี้" aria-label="ลบช่วงนี้">${TRASH}</button>` : ''}</td>` : ''}</tr>`; }).join('')}
          </tbody></table>${canEdit ? `<button class="link" style="margin-top:8px" data-act="ruleAdd" data-d="${dp.id}">+ เพิ่มช่วงจำนวนลูกค้า</button>` : ''}`;
        return `<section class="card pad"><div class="between"><div class="row" style="gap:8px"><i class="dot" style="background:${dp.color};border-radius:3px"></i><h3>${H(dp.name)}</h3></div>
          <label class="row small" style="gap:8px;cursor:pointer">ใช้เงื่อนไข${App.sw('ใช้เงื่อนไขของ ' + dp.short, !!tiers, `data-change="ruleOn" data-d="${dp.id}" ${dis}`)}</label></div>${body}</section>`;
      };
      return `
<div class="page-hd"><div><h1>เงื่อนไขจำนวนพนักงาน</h1><div class="small muted">จำนวนพนักงานที่ต้องใช้ในแต่ละรอบ ตามจำนวนลูกค้าต่องาน — ใช้ตอนกด “สร้างแผนจากใบเสนอราคา” หรือ “ใช้เงื่อนไขกับแผนนี้”</div></div>
  ${canEdit ? '<div class="row wrap"><button class="btn ghost" data-act="ruleReset">ยกเลิกการแก้ไข</button><button class="btn red" data-act="ruleSave">บันทึกเงื่อนไข</button></div>' : ''}</div>
<div class="grid g2">${D.DEPTS.filter(d => d.id !== 'acc').map(card).join('')}</div>
${canEdit ? '' : '<div class="note" style="margin-top:14px">คุณดูเงื่อนไขได้อย่างเดียว — การแก้ไขต้องใช้สิทธิ์ “จัดรอบพนักงาน / อนุมัติต่อรอบ”</div>'}`;
    },
    /* ---- employee directory: add / edit / delete per department ---- */
    renderEmps() {
      const db = Store.db, canEdit = App.can('staff');
      const ui = App.state.ui.emps || (App.state.ui.emps = { q: '', dept: 'all', page: 1 });
      const list = db.employees.filter(e => (ui.dept === 'all' || e.dept === ui.dept) && (!ui.q || (e.name + ' ' + (e.role || '') + ' ' + (e.phone || '')).toLowerCase().includes(ui.q.toLowerCase())))
        .sort((a, b) => D.DEPTS.findIndex(d => d.id === a.dept) - D.DEPTS.findIndex(d => d.id === b.dept) || a.name.localeCompare(b.name, 'th'));
      const per = 15, pages = Math.max(1, Math.ceil(list.length / per)); ui.page = Math.min(ui.page, pages);
      const rows = list.slice((ui.page - 1) * per, ui.page * per);
      const cnt = d => db.employees.filter(e => d === 'all' || e.dept === d).length;
      return `
<div class="page-hd"><div><h1>พนักงานทั้งหมด</h1><div class="small muted">${db.employees.length} คน · ${D.DEPTS.filter(d => cnt(d.id)).map(d => d.short + ' ' + cnt(d.id)).join(' · ')}</div></div>
  ${canEdit ? '<div class="row wrap"><button class="btn red" data-act="empNew">+ เพิ่มพนักงาน</button></div>' : ''}</div>
<div class="row wrap" style="margin-bottom:16px">
  <input class="input" style="max-width:260px;border-radius:999px" placeholder="ค้นหาชื่อ ตำแหน่ง เบอร์โทร…" value="${H(ui.q)}" data-change="empQ" aria-label="ค้นหาพนักงาน">
  <button class="chip ${ui.dept === 'all' ? 'on' : ''}" data-act="empDept" data-v="all">ทั้งหมด <span class="n">${cnt('all')}</span></button>
  ${D.DEPTS.map(d => `<button class="chip ${ui.dept === d.id ? 'on' : ''}" data-act="empDept" data-v="${d.id}">${H(d.short)} <span class="n">${cnt(d.id)}</span></button>`).join('')}
</div>
<section class="card" style="overflow:hidden">
  <table class="tbl resp"><thead><tr><th>พนักงาน</th><th>ตำแหน่ง</th><th>แผนก</th><th>เบอร์โทร</th><th class="r">ชม.สะสม/สัปดาห์</th>${canEdit ? '<th class="r">จัดการ</th>' : ''}</tr></thead><tbody>
  ${rows.map(e => { const dp = dept(e.dept); return `<tr>
    <td><div class="row"><span class="avatar sm" style="background:${dp.bg};color:${dp.fg}">${H(U.initials(e.name))}</span><div style="min-width:0"><div style="font-weight:500">${H(e.name)}</div><div class="sub">${H(e.role || '-')} · ${H(dp.short)}</div></div></div></td>
    <td class="hide-m">${H(e.role || '-')}</td>
    <td class="hide-m"><span class="role-badge" style="background:${dp.bg};color:${dp.fg}">${H(dp.short)}</span></td>
    <td class="hide-m small">${H(e.phone || '-')}</td>
    <td class="r hide-m">${Number(e.hours) || 0} ชม.</td>
    ${canEdit ? `<td class="r"><div class="row" style="gap:6px;justify-content:flex-end;flex-wrap:nowrap"><button class="icon-btn us-edit" data-act="empEdit" data-v="${e.id}" title="แก้ไข" aria-label="แก้ไข ${H(e.name)}">${PENCIL}</button><button class="icon-btn us-edit emp-del" data-act="empDel" data-v="${e.id}" title="ลบ" aria-label="ลบ ${H(e.name)}">${TRASH}</button></div></td>` : ''}</tr>`; }).join('') || `<tr><td colspan="6"><div class="empty">${db.employees.length ? 'ไม่พบพนักงาน' : 'ยังไม่มีรายชื่อพนักงาน — กด “+ เพิ่มพนักงาน”'}</div></td></tr>`}
  </tbody></table>
  <div class="between" style="padding:14px 22px;border-top:1px solid var(--line)"><span class="small muted">แสดง ${list.length ? (ui.page - 1) * per + 1 : 0}–${Math.min(ui.page * per, list.length)} จาก ${list.length} คน</span>
    <div class="pager"><button data-act="empPage" data-v="${ui.page - 1}" ${ui.page <= 1 ? 'disabled' : ''} aria-label="ก่อนหน้า">‹</button>${Array.from({ length: pages }, (_, i) => `<button class="${ui.page === i + 1 ? 'on' : ''}" data-act="empPage" data-v="${i + 1}">${i + 1}</button>`).join('')}<button data-act="empPage" data-v="${ui.page + 1}" ${ui.page >= pages ? 'disabled' : ''} aria-label="ถัดไป">›</button></div></div>
</section>
${canEdit ? '' : '<div class="note" style="margin-top:14px">คุณดูรายชื่อได้อย่างเดียว — การเพิ่ม/แก้ไข/ลบพนักงานต้องใช้สิทธิ์ “จัดรอบพนักงาน / อนุมัติต่อรอบ”</div>'}`;
    },
    renderShifts(p) {
      const db = Store.db;
      const candidates = db.quotes.filter(q => q.status !== 'cancelled' && q.status !== 'draft').sort((a, b) => a.start.localeCompare(b.start));
      const q = Store.quote(p[0]) || Store.quote(App.state.ui.staffQ) || candidates.find(x => db.staff[x.id]) || candidates[0];
      if (!q) return '<div class="empty">ยังไม่มีงานที่ต้องจัดรอบพนักงาน</div>';
      App.state.ui.staffQ = q.id;
      const plan = db.staff[q.id];
      const head = (extra) => `<div class="page-hd"><div><h1>รอบพนักงาน</h1><div class="small muted"><button class="link" data-act="stPickQ">${H(q.id)} · ${H(q.title)} · ${H(q.client)} · ${q.pax} ท่าน ▾</button></div></div>${extra || ''}</div>`;
      if (!plan) return head() + `<div class="card pad" style="text-align:center"><p class="muted">ยังไม่มีแผนรอบพนักงานสำหรับงานนี้</p><button class="btn red" data-act="stGen">สร้างแผนจากใบเสนอราคา</button></div>`;
      const dates = Object.keys(plan.days).sort();
      const ui = App.state.ui.staff && App.state.ui.staff.q === q.id ? App.state.ui.staff : (App.state.ui.staff = { q: q.id, date: dates[0], ot: null });
      if (!plan.days[ui.date]) ui.date = dates[0];
      const day = plan.days[ui.date], date = ui.date, dayIdx = dates.indexOf(date) + 1;
      // KPIs
      let need = 0, have = 0; const short = {};
      for (const [dp, shifts] of Object.entries(day)) for (const s of shifts) { need += s.need; have += Math.min(s.staff.length, s.need); if (s.staff.length < s.need) short[dp] = (short[dp] || 0) + s.need - s.staff.length; }
      const otToday = plan.ot.filter(o => o.date === date && o.status !== 'rejected');
      const otH = otToday.reduce((a, o) => a + o.hours * o.staff.length, 0), otN = new Set(otToday.flatMap(o => o.staff)).size;
      // timeline
      const A = 6, B = 24, pos = t => (U.hm(t) - A) / (B - A) * 100;
      const evRow = q.items.flatMap(it => (it.sched || []).filter(s => s.day === dayIdx).map(s => ({ it, s }))).filter(x => x.it.sec === 'meet' && !String(x.it.ref).startsWith('svc:') || (D.FOOD.find(f => f.id === x.it.ref) || {}).cat === 'gala');
      const lanes = D.DEPTS.filter(d => day[d.id]).map(dp => {
        const shifts = day[dp.id];
        const got = shifts.reduce((a, s) => a + s.staff.length, 0), nd = shifts.reduce((a, s) => a + s.need, 0);
        const blocks = shifts.map(s => {
          const l = pos(s.start), w = pos(s.end) - l;
          if (!s.staff.length && s.need) return `<div class="blk short" style="left:${l}%;width:${w}%" data-act="stAssign" data-dept="${dp.id}" data-s="${s.id}"><b>ขาด ${s.need} คน</b><br>ยังไม่มอบหมาย</div>`;
          if (!s.need && !s.staff.length) return '';
          const lack = s.need - s.staff.length;
          return `<div class="blk ${s.type === 'evening' ? 'evening' : 'morning'} ${lack > 0 ? 'short-edge' : ''}" style="left:${l}%;width:${w}%" data-act="stAssign" data-dept="${dp.id}" data-s="${s.id}"><b>${H(s.label)}${s.type === 'morning' && s.label === 'รอบเต็มวัน' ? ' ' + s.start + '–' + s.end : ''}</b><br>${lack > 0 ? `${s.staff.length} / ${s.need} คน · ขาด ${lack}` : s.staff.length + ' คน'}</div>`;
        }).join('');
        const ots = plan.ot.filter(o => o.date === date && o.dept === dp.id && o.status !== 'rejected').map(o => {
          const s = shifts.find(x => x.id === o.shiftId) || shifts[shifts.length - 1];
          const l = pos(s.end), w = Math.min(pos(U.fromHm(Math.min(24, U.hm(s.end) + o.hours))) - l, 100 - l);
          return `<div class="blk ot ${o.status === 'approved' ? 'ok' : ''}" style="left:${l}%;width:${w}%" data-act="stOtOpen" data-id="${o.id}"><b>+${o.hours} ชม.</b><br>${o.staff.length} คน${o.status === 'approved' ? ' ✓' : ''}</div>`;
        }).join('');
        return `<div class="lab"><div class="row" style="align-items:flex-start;gap:8px"><i class="dot" style="background:${dp.color};border-radius:3px;margin-top:6px"></i><div><b style="font-weight:600">${H(dp.name)}</b><div class="xs muted">จัดแล้ว ${got} / ${nd} คน</div></div></div></div><div class="lane">${blocks}${ots}</div>`;
      }).join('');
      // OT panel state
      let ot = ui.ot && plan.ot.find(o => o.id === ui.ot && o.date === date);
      if (!ot && ui.ot !== '__new') ot = plan.ot.find(o => o.date === date && o.status === 'draft') || null;
      const otPanel = App.V.staff.otPanel(q, plan, day, date, ot);
      // todos
      const todos = [];
      for (const [dp, shifts] of Object.entries(day)) for (const s of shifts) if (s.staff.length < s.need) {
        const free = Store.db.employees.filter(e => e.dept === dp && !Object.values(day).flat().some(x => x.staff.includes(e.id)));
        todos.push({ c: 'var(--red)', t: `${dept(dp).short} ขาด ${s.need - s.staff.length} คน · ${s.label} ${s.start}–${s.end}`, s: free.length ? `มีพนักงานว่าง ${free.length} คน — แนะนำ: ${free[0].name}` : 'ไม่มีพนักงานว่างในแผนก', a: `<button class="btn dark sm" data-act="stAssign" data-dept="${dp}" data-s="${s.id}">มอบหมาย</button>` });
      }
      for (const o of plan.ot.filter(o => o.status === 'pending')) todos.push({ c: 'var(--yellow)', t: `ต่อรอบ ${dept(o.dept).short} รออนุมัติ · ${o.staff.length} คน`, s: `${U.fmtD(o.date, { year: false })} · +${o.hours} ชม. · ${o.reason}`, a: App.can('staff') ? `<button class="btn ghost sm" data-act="stOtApprove" data-id="${o.id}">อนุมัติ</button>` : '<span class="small muted">รอผู้จัดการ</span>' });
      for (const e of Store.db.employees) { const h = e.hours + otHours(e.id, plan); if (h >= 48 && Object.values(day).flat().some(s => s.staff.includes(e.id))) todos.push({ c: 'var(--muted2)', t: `${e.name} ชั่วโมงสะสม ${h} ชม.`, s: 'ครบ 48 ชม./สัปดาห์ ควรจัดวันหยุด', a: `<button class="btn ghost sm" data-act="stOff" data-e="${e.id}">จัดวันหยุด</button>` }); }
      return head(`<div class="row wrap"><div class="seg white">${dates.map((d, i) => `<button class="${d === date ? 'on' : ''}" data-act="stDate" data-v="${d}">${U.DS[U.d(d).getDay()]} ${U.fmtD(d, { year: false })}</button>`).join('')}</div>
        ${App.can('staff') ? '<button class="btn ghost" data-act="stApplyRules">ใช้เงื่อนไขกับแผนนี้</button>' : ''}<button class="btn ghost" data-act="stShare">ส่งตารางให้หัวหน้าแผนก</button><button class="btn dark" data-act="stShift">+ เพิ่มรอบ</button></div>`) + `
<div class="grid g4" style="margin-bottom:16px">
  <div class="card kpi"><div class="small muted">พนักงานที่ต้องใช้</div><div class="v">${need} <small>คน</small></div><div class="small muted">${Object.keys(day).length} แผนก · วันที่ ${dayIdx} (${U.DS[U.d(date).getDay()]} ${U.fmtD(date, { year: false })})</div></div>
  <div class="card kpi"><div class="small muted">จัดรอบแล้ว</div><div class="v">${have} <small>/ ${need} คน</small></div><div class="bar"><i style="width:${need ? have / need * 100 : 0}%"></i></div><div class="small muted" style="margin-top:6px">${need ? Math.round(have / need * 100) : 0}% ของที่ต้องใช้</div></div>
  <div class="card kpi"><div class="small muted">ขาดคน</div><div class="v ${need - have ? 'red' : ''}">${need - have} <small>คน</small></div><div class="small muted">${Object.entries(short).map(([k, v]) => dept(k).short + ' ' + v).join(' · ') || 'ครบทุกแผนก'}</div></div>
  <div class="card kpi"><div class="small muted">ต่อรอบ (QT)</div><div class="v" style="color:var(--yellow-d)">${otH} <small>ชม.</small></div><div class="small muted">รวมชั่วโมงทำงานเกินกำหนดการ · ${otN} คน</div></div>
</div>
<div class="split" style="margin-bottom:16px">
  <section class="card" style="overflow:hidden;min-width:0">
    <div class="card-hd"><h2>ไทม์ไลน์รอบ · ${U.DS[U.d(date).getDay()]} ${U.fmtD(date, { year: false })}</h2><div class="row wrap xs muted" style="gap:10px"><span>▢ รอบเช้า</span><span>■ รอบบ่าย</span><span style="color:var(--yellow-d)">▢ ต่อรอบ</span><span class="red">▢ ขาดคน</span></div></div>
    <div class="st-scroll"><div class="st-tl">
      <div></div><div class="st-axis">${[6, 8, 10, 12, 14, 16, 18, 20, 22, 24].map(h => `<span style="left:${(h - A) / (B - A) * 100}%">${String(h).padStart(2, '0')}</span>`).join('')}</div>
      <div class="lab b">กำหนดการงาน</div><div class="lane">${evRow.map(({ it, s }) => `<div class="blk event ${it.sec === 'food' ? 'gala' : ''}" style="left:${pos(s.start)}%;width:${pos(s.end) - pos(s.start)}%"><b>${it.sec === 'food' ? 'งานเลี้ยง' : 'งาน' + (q.title.split(' ')[0].length < 12 ? q.title.split(' ')[0] : '')} ${s.start}–${s.end}</b></div>`).join('')}</div>
      ${lanes}
    </div></div>
  </section>
  <aside class="card pad sticky">${otPanel}</aside>
</div>
<div class="grid g2">
  <section class="card" style="overflow:hidden"><div class="card-hd"><div class="row"><h2>ต้องดำเนินการ</h2><span class="badge" style="background:var(--red);color:#fff">${todos.length}</span></div></div>
    ${todos.map(t => `<div class="todo"><i class="dot" style="background:${t.c}"></i><div style="flex:1;min-width:0"><b style="font-weight:600">${H(t.t)}</b><div class="small muted">${H(t.s)}</div></div>${t.a}</div>`).join('') || '<div class="empty small">ไม่มีงานค้าง</div>'}</section>
  <section class="card" style="overflow:hidden"><div class="card-hd"><h2>สรุปรายแผนก</h2></div>
    <table class="tbl"><thead><tr><th>แผนก</th><th>จัดแล้ว / ต้องใช้</th><th class="hide-m">ต่อรอบ</th><th class="r">สถานะ</th></tr></thead><tbody>
    ${D.DEPTS.filter(d => day[d.id]).map(dp => { const sh = day[dp.id], got = sh.reduce((a, s) => a + Math.min(s.staff.length, s.need), 0), nd = sh.reduce((a, s) => a + s.need, 0), o = otToday.filter(x => x.dept === dp.id); const lack = nd - got;
      return `<tr><td><span class="row"><i class="dot" style="background:${dp.color};border-radius:3px"></i>${H(dp.short)}</span></td><td><div class="row"><div class="bar ${lack ? 'red' : ''}" style="flex:1;min-width:60px"><i style="width:${nd ? got / nd * 100 : 100}%"></i></div><span class="small num">${got} / ${nd}</span></div></td>
        <td class="small muted hide-m">${o.length ? o.reduce((a, x) => a + x.staff.length, 0) + ' คน · ' + o.reduce((a, x) => a + x.hours * x.staff.length, 0) + ' ชม.' : '—'}</td><td class="r">${lack ? `<span class="badge st-tentative">ขาด ${lack}</span>` : '<span class="badge st-confirmed">ครบ</span>'}</td></tr>`; }).join('')}
    </tbody></table></section>
</div>`;
    },
    otPanel(q, plan, day, date, ot) {
      const shifts = Object.entries(day).flatMap(([dp, l]) => l.filter(s => s.staff.length).map(s => ({ dp, s })));
      if (!shifts.length) return '<h2>ต่อรอบ</h2><div class="small muted" style="margin-top:8px">ยังไม่มีรอบที่มอบหมายพนักงาน</div>';
      const d = App.state.ui.staffOtDraft && App.state.ui.staffOtDraft.date === date && !ot ? App.state.ui.staffOtDraft : null;
      const cur = ot || d || (App.state.ui.staffOtDraft = { id: null, date, dept: shifts[0].dp, shiftId: shifts[0].s.id, hours: 2, reason: 'เก็บงานเลี้ยง + ล้างครัว', staff: [], status: 'new' });
      const sh = (day[cur.dept] || []).find(s => s.id === cur.shiftId) || shifts[0].s;
      const emps = sh.staff.map(EMP).filter(Boolean);
      const extra = Store.db.employees.filter(e => e.dept === cur.dept && !sh.staff.includes(e.id));
      const editable = cur.status === 'new' || cur.status === 'draft';
      const stLabel = { new: 'ใหม่', draft: 'ร่าง · ยังไม่ส่ง', pending: 'รออนุมัติ', approved: 'อนุมัติแล้ว', rejected: 'ไม่อนุมัติ' }[cur.status];
      return `<div class="row" style="margin-bottom:14px"><h2>ต่อรอบ</h2><span class="badge ${cur.status === 'approved' ? 'st-confirmed' : 'st-sent'}">${stLabel}</span>${plan.ot.some(o => o.date === date) ? `<span class="spacer"></span><button class="link" data-act="stOtNew">+ ใหม่</button>` : ''}</div>
        <div style="background:var(--bg);border-radius:12px;padding:12px 14px;margin-bottom:14px">
          <select class="input sm" style="background:transparent;border:0;padding:0;height:auto;font-size:12px;color:var(--muted)" data-change="stOtShift" ${editable ? '' : 'disabled'} aria-label="รอบ">${shifts.map(({ dp, s }) => `<option value="${dp}|${s.id}" ${s.id === sh.id ? 'selected' : ''}>${H(dept(dp).name)} · ${H(s.label)}</option>`).join('')}</select>
          <div style="font:600 20px var(--fh);margin-top:4px">${sh.start}–${sh.end} <span class="muted" style="font-size:14px">→</span> <span style="color:var(--yellow-d)">${sh.start}–${U.fromHm(Math.min(24, U.hm(sh.end) + cur.hours))}</span></div></div>
        <div class="small muted" style="margin-bottom:6px">ต่อเพิ่ม</div>
        <div class="grid g4" style="gap:6px;grid-template-columns:repeat(4,1fr)">${[1, 2, 3].map(h => `<button class="btn ${cur.hours === h ? 'dark' : 'ghost'} sm" style="border-radius:12px;height:44px" data-act="stOtH" data-v="${h}" ${editable ? '' : 'disabled'}>+${h} ชม.</button>`).join('')}<button class="btn ${[1, 2, 3].includes(cur.hours) ? 'ghost' : 'dark'} sm" style="border-radius:12px;height:44px;padding:0 4px" data-act="stOtH" data-v="custom" ${editable ? '' : 'disabled'}>${[1, 2, 3].includes(cur.hours) ? 'กำหนดเอง' : '+' + cur.hours}</button></div>
        <div class="small muted" style="margin:14px 0 6px">เหตุผล</div>
        <select class="input" data-change="stOtReason" ${editable ? '' : 'disabled'}>${['เก็บงานเลี้ยง + ล้างครัว', 'เก็บงานเลี้ยงขันโตก + ล้างครัว', 'เก็บสถานที่งานเลี้ยง', 'งานเลิกช้ากว่ากำหนด', 'เตรียมงานวันถัดไป', 'ทดแทนพนักงานลา'].concat(cur.reason ? [cur.reason] : []).filter((v, i, a) => a.indexOf(v) === i).map(r => `<option ${r === cur.reason ? 'selected' : ''}>${H(r)}</option>`).join('')}</select>
        <div class="between small" style="margin:14px 0 8px"><span class="muted">เลือกพนักงาน (${cur.staff.length} / ${emps.length})</span>${editable ? '<button class="link" data-act="stOtAll">เลือกทั้งหมดที่ว่าง</button>' : ''}</div>
        ${emps.concat(cur.staff.map(EMP).filter(e => e && !emps.includes(e))).map(e => {
          const h = e.hours + otHours(e.id, plan) - (ot && ot.staff.includes(e.id) ? ot.hours : 0), on = cur.staff.includes(e.id), full = h + cur.hours > 48 && !on;
          return `<div class="emp ${full ? 'dis' : ''}" ${editable && !full ? `data-act="stOtEmp" data-e="${e.id}"` : ''}><span class="check ${on ? 'on' : ''}">${on ? ICON.check : ''}</span><span class="avatar sm" style="background:var(--red-soft);color:var(--red-d)">${H(U.initials(e.name))}</span><div style="flex:1;min-width:0"><b style="font-weight:500">${H(e.name)}</b><div class="xs muted">${H(e.role)} · ${H(sh.label)}</div></div><span class="xs ${h >= 48 ? 'red' : h >= 44 ? '' : 'muted'}" style="${h >= 44 && h < 48 ? 'color:var(--yellow-d)' : ''}">${h >= 48 ? 'ครบ 48 ชม.' : h + ' / 48 ชม.'}</span></div>`;
        }).join('')}
        ${editable && extra.length ? `<button class="link" style="margin-top:8px" data-act="stOtMore">+ พนักงานอื่นในแผนก (${extra.length})</button>` : ''}
        <div style="background:var(--bg);border-radius:12px;padding:14px;margin:14px 0"><div class="small muted">${cur.staff.length} คน × ${cur.hours} ชม.</div>
          <div class="between" style="margin-top:4px"><b>รวมชั่วโมงทำงานเกินกำหนดการ</b><b style="font:700 20px var(--fh)" class="num">${cur.staff.length * cur.hours} ชม.</b></div></div>
        <div class="xs muted" style="margin-bottom:14px">ต่อรอบเกิน 2 ชม./วัน หรือสะสมเกิน 48 ชม./สัปดาห์ ต้องให้ผู้จัดการแผนกอนุมัติ</div>
        ${editable ? `<button class="btn red block" data-act="stOtSubmit" ${cur.staff.length ? '' : 'disabled'}>ขออนุมัติต่อรอบ (${cur.staff.length} คน)</button>${cur.status === 'draft' ? '<button class="link red" style="display:block;margin:12px auto 0" data-act="stOtDel">ลบร่าง</button>' : ''}`
          : cur.status === 'pending' && App.can('staff') ? `<div class="row"><button class="btn ghost" style="flex:1" data-act="stOtReject" data-id="${cur.id}">ไม่อนุมัติ</button><button class="btn teal" style="flex:1" data-act="stOtApprove" data-id="${cur.id}">อนุมัติ</button></div>`
          : `<div class="small muted" style="text-align:center">${cur.status === 'pending' ? 'ส่งขออนุมัติแล้ว — รอผู้จัดการแผนก' : 'ดำเนินการแล้ว'}</div>`}`;
    }
  };
  const SP = () => Store.db.staff[App.state.ui.staffQ];
  const curOt = () => {
    const ui = App.state.ui.staff, p = SP();
    if (ui.ot === '__new') return App.state.ui.staffOtDraft;
    return (ui.ot && p.ot.find(o => o.id === ui.ot && o.date === ui.date)) || p.ot.find(o => o.date === ui.date && o.status === 'draft') || App.state.ui.staffOtDraft;
  };
  Object.assign(App.ACT, {
    stPickQ() { App.pickQuoteModal('เลือกงาน', id => { App.state.ui.staff = null; App.go('#/staff/' + encodeURIComponent(id)); }, q => q.status !== 'cancelled' && q.status !== 'draft'); },
    stGen() { const q = Store.quote(App.state.ui.staffQ); Store.db.staff[q.id] = genPlan(q); Store.save(); App.render(); App.toast('สร้างแผนแล้ว — กดที่รอบเพื่อมอบหมายพนักงาน'); },
    stDate(el) { App.state.ui.staff.date = el.dataset.v; App.state.ui.staff.ot = null; App.render(); },
    stShare() {
      const q = Store.quote(App.state.ui.staffQ), p = SP(), date = App.state.ui.staff.date, day = p.days[date];
      const lines = Object.entries(day).map(([dp, l]) => dept(dp).name + ': ' + l.map(s => `${s.label} ${s.start}–${s.end} (${s.staff.map(id => (EMP(id) || {}).name).filter(Boolean).join(', ') || '-'})`).join(' / ')).join('\n');
      const text = `ตารางรอบพนักงาน ${q.title} · ${U.fmtD(date)}\n${lines}`;
      if (navigator.share) navigator.share({ title: 'ตารางรอบพนักงาน', text }).catch(() => { /* cancelled */ });
      else (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject()).then(() => App.toast('คัดลอกตารางแล้ว — วางใน LINE กลุ่มหัวหน้าแผนกได้เลย'), () => App.toast('แชร์ไม่สำเร็จ'));
    },
    stShift() {
      const p = SP(), date = App.state.ui.staff.date;
      App.modal(`<h2>เพิ่มรอบ</h2><div class="grid g2">
        <div><label class="f">แผนก</label><select class="input" name="dept">${D.DEPTS.filter(d => d.id !== 'acc').map(d => `<option value="${d.id}">${H(d.name)}</option>`).join('')}</select></div>
        <div><label class="f">ชื่อรอบ</label><input class="input" name="label" value="รอบเสริม"></div>
        <div><label class="f">เริ่ม</label><input class="input" type="time" name="start" value="16:00"></div><div><label class="f">สิ้นสุด</label><input class="input" type="time" name="end" value="22:00"></div>
        <div><label class="f">จำนวนที่ต้องใช้</label><input class="input" name="need" inputmode="numeric" value="2"></div>
        <div><label class="f">ประเภท</label><select class="input" name="type"><option value="morning">กลางวัน</option><option value="evening">กลางคืน/บ่าย</option></select></div></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">เพิ่มรอบ</button></div>`,
      bg => bg.querySelector('#ok').onclick = () => { const f = App.formData(bg); (p.days[date][f.dept] = p.days[date][f.dept] || []).push({ id: 'x' + U.uid(), type: f.type, label: f.label, start: f.start, end: f.end, need: parseInt(f.need, 10) || 1, staff: [] }); Store.save(); App.closeModal(); App.render(); });
    },
    stAssign(el) {
      const p = SP(), date = App.state.ui.staff.date, day = p.days[date], s = day[el.dataset.dept].find(x => x.id === el.dataset.s);
      const busy = new Set(Object.values(day).flat().filter(x => x !== s && !(U.hm(x.end) <= U.hm(s.start) || U.hm(x.start) >= U.hm(s.end))).flatMap(x => x.staff));
      const emps = Store.db.employees.filter(e => e.dept === el.dataset.dept);
      App.modal(`<h2>${H(dept(el.dataset.dept).name)} · ${H(s.label)} ${s.start}–${s.end}</h2>
        <div class="grid g3" style="margin-bottom:12px"><div><label class="f">ต้องใช้ (คน)</label><input class="input" name="need" inputmode="numeric" value="${s.need}"></div><div><label class="f">เริ่ม</label><input class="input" type="time" name="start" value="${s.start}"></div><div><label class="f">สิ้นสุด</label><input class="input" type="time" name="end" value="${s.end}"></div></div>
        <div class="small muted" style="margin-bottom:8px">เลือกพนักงาน</div>
        <div style="max-height:44vh;overflow:auto">${emps.map(e => { const b = busy.has(e.id), on = s.staff.includes(e.id); return `<label class="emp ${b ? 'dis' : ''}"><input type="checkbox" name="e_${e.id}" ${on ? 'checked' : ''} ${b ? 'disabled' : ''} style="width:20px;height:20px;accent-color:var(--teal)"><div style="flex:1"><b style="font-weight:500">${H(e.name)}</b><div class="xs muted">${H(e.role)}${b ? ' · อยู่รอบอื่นแล้ว' : ''}</div></div><span class="xs muted">${e.hours + otHours(e.id, p)} / 48 ชม.</span></label>`; }).join('')}</div>
        <div class="acts"><button class="btn ghost" id="del" style="color:var(--red);margin-right:auto">ลบรอบ</button><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">บันทึก</button></div>`, bg => {
        bg.querySelector('#ok').onclick = () => {
          const f = App.formData(bg);
          s.staff = emps.filter(e => f['e_' + e.id]).map(e => e.id); s.need = parseInt(f.need, 10) || 0; s.start = f.start; s.end = f.end;
          Store.save(); App.closeModal(); App.render(); App.toast('บันทึกการมอบหมายแล้ว');
        };
        bg.querySelector('#del').onclick = () => { const l = day[el.dataset.dept]; l.splice(l.indexOf(s), 1); p.ot = p.ot.filter(o => o.shiftId !== s.id); Store.save(); App.closeModal(); App.render(); };
      });
    },
    stOtOpen(el) { App.state.ui.staff.ot = el.dataset.id; App.render(); },
    stOtNew() { App.state.ui.staffOtDraft = null; App.state.ui.staff.ot = '__new'; App.render(); },
    stOtH(el) {
      const o = curOt();
      if (el.dataset.v === 'custom') {
        App.modal(`<h2>ชั่วโมงต่อรอบ</h2><input class="input" name="h" inputmode="decimal" value="${o.hours}"><div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">ตกลง</button></div>`,
          bg => bg.querySelector('#ok').onclick = () => { o.hours = U.clamp(parseFloat(App.formData(bg).h) || 1, 0.5, 8); persistOt(o); App.closeModal(); App.render(); });
        return;
      }
      o.hours = Number(el.dataset.v); persistOt(o); App.render();
    },
    stOtEmp(el) { const o = curOt(), i = o.staff.indexOf(el.dataset.e); if (i >= 0) o.staff.splice(i, 1); else o.staff.push(el.dataset.e); persistOt(o); App.render(); },
    stOtAll() {
      const o = curOt(), p = SP(), sh = p.days[o.date][o.dept].find(s => s.id === o.shiftId);
      o.staff = sh.staff.filter(id => { const e = EMP(id); return e && e.hours + otHours(id, p) - (o.id && o.staff.includes(id) ? o.hours : 0) + o.hours <= 48; });
      persistOt(o); App.render();
    },
    stOtMore() {
      const o = curOt(), p = SP(), emps = Store.db.employees.filter(e => e.dept === o.dept && !o.staff.includes(e.id));
      App.modal(`<h2>เพิ่มพนักงาน</h2>${emps.map(e => `<button class="pick" data-e="${e.id}"><span>${H(e.name)}<br><span class="xs muted">${H(e.role)}</span></span><span class="xs muted">${e.hours + otHours(e.id, p)} / 48</span></button>`).join('')}<div class="acts"><button class="btn ghost" data-act="closeModal">ปิด</button></div>`,
        bg => bg.querySelectorAll('.pick').forEach(b => b.onclick = () => { o.staff.push(b.dataset.e); persistOt(o); App.closeModal(); App.render(); }));
    },
    stOtSubmit() {
      const o = curOt(); persistOt(o, true);
      o.status = App.can('staff') ? 'approved' : 'pending'; Store.save();
      App.toast(o.status === 'approved' ? 'บันทึกและอนุมัติต่อรอบแล้ว' : 'ส่งขออนุมัติแล้ว'); App.state.ui.staffOtDraft = null; App.state.ui.staff.ot = o.id; App.render();
    },
    stOtApprove(el) { const o = SP().ot.find(x => x.id === el.dataset.id); o.status = 'approved'; Store.save(); App.render(); App.toast('อนุมัติแล้ว'); },
    stOtReject(el) { const o = SP().ot.find(x => x.id === el.dataset.id); o.status = 'rejected'; Store.save(); App.state.ui.staff.ot = null; App.render(); },
    stOtDel() { const p = SP(), o = curOt(); p.ot = p.ot.filter(x => x !== o); Store.save(); App.state.ui.staff.ot = null; App.render(); },
    stOff(el) {
      const e = EMP(el.dataset.e), p = SP(), date = App.state.ui.staff.date;
      App.confirm(`จัดวันหยุดให้ ${e.name} และนำออกจากรอบวันที่ ${U.fmtD(date, { year: false })}?`, () => {
        Object.values(p.days[date]).flat().forEach(s => s.staff = s.staff.filter(x => x !== e.id));
        p.ot.forEach(o => { if (o.date === date) o.staff = o.staff.filter(x => x !== e.id); });
        Store.save(); App.render(); App.toast('จัดวันหยุดแล้ว');
      }, 'จัดวันหยุด');
    }
  });
  const POSITIONS = { kitchen: ['เชฟ', 'ผู้ช่วยเชฟ', 'พนักงานครัว'], banquet: ['หัวหน้าจัดเลี้ยง', 'พนักงานจัดเลี้ยง'], av: ['หัวหน้าช่าง', 'ช่างเทคนิค'], front: ['หัวหน้าส่วนหน้า', 'พนักงานต้อนรับ'], hk: ['หัวหน้าแม่บ้าน', 'แม่บ้าน'], acc: ['หัวหน้าบัญชี', 'พนักงานบัญชี'] };
  function empForm(e) {
    const isNew = !e.id;
    App.modal(`<h2>${isNew ? 'เพิ่มพนักงาน' : 'แก้ไขข้อมูลพนักงาน'}</h2>${isNew ? '' : `<p class="muted" style="margin-top:-8px">${H(e.name)} · รหัส ${H(e.id)}</p>`}
      <div class="field"><label class="f" for="emName">ชื่อ-นามสกุล</label><input class="input" id="emName" name="name" value="${H(e.name || '')}"></div>
      <div class="grid g2" style="margin-top:14px">
        <div><label class="f" for="emDept">แผนก</label><select class="input" id="emDept" name="dept">${D.DEPTS.map(d => `<option value="${d.id}" ${d.id === e.dept ? 'selected' : ''}>${H(d.name)}</option>`).join('')}</select></div>
        <div><label class="f" for="emRole">ตำแหน่ง</label><input class="input" id="emRole" name="role" value="${H(e.role || '')}" list="emRoles" placeholder="เช่น พนักงานครัว"><datalist id="emRoles"></datalist></div>
      </div>
      <div class="grid g2" style="margin-top:14px">
        <div><label class="f" for="emPhone">เบอร์โทร</label><input class="input" id="emPhone" type="tel" name="phone" value="${H(e.phone || '')}"></div>
        <div><label class="f" for="emHours">ชั่วโมงทำงานสะสม/สัปดาห์</label><input class="input" id="emHours" type="number" min="0" max="168" step="1" name="hours" value="${Number(e.hours) || 0}"></div>
      </div>
      <div class="small red" id="emErr" style="min-height:20px;margin-top:6px" role="alert"></div>
      <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="emOk">${isNew ? 'เพิ่มพนักงาน' : 'บันทึก'}</button></div>`, bg => {
      const sel = bg.querySelector('#emDept'), dl = bg.querySelector('#emRoles');
      const fill = () => {
        const used = Store.db.employees.filter(x => x.dept === sel.value).map(x => x.role).filter(Boolean);
        dl.innerHTML = [...new Set((POSITIONS[sel.value] || []).concat(used))].map(r => `<option value="${H(r)}">`).join('');
      };
      fill(); sel.onchange = fill;
      bg.querySelector('#emOk').onclick = () => {
        const f = App.formData(bg), err = bg.querySelector('#emErr');
        f.name = f.name.trim(); f.role = f.role.trim(); f.phone = f.phone.trim();
        if (!f.name) { err.textContent = 'กรุณากรอกชื่อพนักงาน'; return; }
        if (!f.role) { err.textContent = 'กรุณากรอกตำแหน่ง'; return; }
        if (Store.db.employees.some(x => x.id !== e.id && x.name.replace(/\s+/g, ' ') === f.name.replace(/\s+/g, ' ') && x.dept === f.dept)) { err.textContent = 'มีพนักงานชื่อนี้ในแผนกนี้อยู่แล้ว'; return; }
        const hours = Math.max(0, Math.min(168, Math.round(Number(f.hours) || 0)));
        const db = Store.db;
        if (isNew) {
          const ids = db.employees.map(x => parseInt(String(x.id).slice(1), 10) || 0);
          db.employees.push({ id: 'e' + (Math.max(0, ...ids) + 1), name: f.name, dept: f.dept, role: f.role, phone: f.phone, hours });
        } else {
          const cur = db.employees.find(x => x.id === e.id); if (!cur) { App.closeModal(); return; }
          // moving to another department takes the person off that department's old shifts
          if (cur.dept !== f.dept) unassign(cur.id);
          Object.assign(cur, { name: f.name, dept: f.dept, role: f.role, phone: f.phone, hours });
        }
        Store.save(); App.closeModal(); App.render(); App.toast(isNew ? 'เพิ่ม ' + f.name + ' แล้ว' : 'บันทึกแล้ว');
      };
    });
  }
  // remove an employee from every shift and OT request in all plans
  function unassign(id) {
    for (const plan of Object.values(Store.db.staff)) {
      for (const day of Object.values(plan.days || {})) for (const s of Object.values(day).flat()) s.staff = s.staff.filter(x => x !== id);
      for (const o of plan.ot || []) o.staff = o.staff.filter(x => x !== id);
    }
  }
  Object.assign(App.ACT, {
    stTab(el) { App.state.ui.staffTab = el.dataset.v; App.state.ui.rulesDraft = null; App.render(); },
    ruleAdd(el) {
      const t = App.state.ui.rulesDraft[el.dataset.d], last = t[t.length - 1], prev = t.length > 1 ? Number(t[t.length - 2].upTo) || 0 : 0;
      // the open-ended last tier gets a limit; a new open-ended tier goes after it
      last.upTo = prev + 50; t.push({ upTo: null, n: (Number(last.n) || 0) + 1 }); App.render();
    },
    ruleDel(el) { const t = App.state.ui.rulesDraft[el.dataset.d]; t.splice(Number(el.dataset.i), 1); t[t.length - 1].upTo = null; App.render(); },
    ruleReset() { App.state.ui.rulesDraft = null; App.render(); },
    ruleSave() {
      const R = App.state.ui.rulesDraft;
      for (const [dp, t] of Object.entries(R)) {
        if (!t) continue;
        let prev = 0;
        for (let i = 0; i < t.length - 1; i++) {
          const up = Number(t[i].upTo);
          if (!Number.isInteger(up) || up <= prev) { App.toast(`${dept(dp).short}: ช่วงจำนวนลูกค้าต้องเรียงจากน้อยไปมาก (ช่วงที่ ${i + 1})`); return; }
          prev = up;
        }
        if (t.some(x => !Number.isInteger(Number(x.n)) || Number(x.n) < 0)) { App.toast(`${dept(dp).short}: จำนวนพนักงานต้องเป็นจำนวนเต็ม 0 ขึ้นไป`); return; }
        t.forEach((x, i) => { x.n = Number(x.n); x.upTo = i === t.length - 1 ? null : Number(x.upTo); });
      }
      Store.db.settings.staffRules = JSON.parse(JSON.stringify(R)); Store.save();
      App.state.ui.rulesDraft = null; App.render(); App.toast('บันทึกเงื่อนไขแล้ว — ใช้กับแผนที่สร้างใหม่ หรือกด “ใช้เงื่อนไขกับแผนนี้” ในแต่ละงาน');
    },
    stApplyRules() {
      const q = Store.quote(App.state.ui.staffQ), p = SP(); if (!q || !p) return;
      App.confirm(`ตั้งจำนวนพนักงานที่ต้องใช้ของงาน ${q.id} (${q.pax} ท่าน) ใหม่ตามเงื่อนไข?`, () => {
        let n = 0;
        for (const day of Object.values(p.days)) for (const [dp, shifts] of Object.entries(day)) { const k = ruleNeed(dp, q.pax); if (k != null) shifts.forEach(s => { s.need = k; n++; }); }
        Store.save(); App.render(); App.toast(n ? 'ปรับจำนวนที่ต้องใช้ตามเงื่อนไขแล้ว (คนที่มอบหมายไว้ยังอยู่)' : 'ยังไม่มีแผนกที่ตั้งเงื่อนไขไว้');
      }, 'ใช้เงื่อนไข');
    },
    empDept(el) { App.state.ui.emps.dept = el.dataset.v; App.state.ui.emps.page = 1; App.render(); },
    empPage(el) { App.state.ui.emps.page = Number(el.dataset.v); App.render(); },
    empNew() {
      if (!App.can('staff')) return;
      const d = App.state.ui.emps && App.state.ui.emps.dept;
      empForm({ dept: d && d !== 'all' ? d : D.DEPTS[0].id, hours: 0 });
    },
    empEdit(el) { const e = Store.db.employees.find(x => x.id === el.dataset.v); if (e && App.can('staff')) empForm(e); },
    empDel(el) {
      const e = Store.db.employees.find(x => x.id === el.dataset.v); if (!e || !App.can('staff')) return;
      const used = Object.values(Store.db.staff).some(p => Object.values(p.days || {}).some(day => Object.values(day).flat().some(s => s.staff.includes(e.id))));
      App.modal(`<h2>ลบพนักงานคนนี้?</h2><p class="muted" style="margin-top:-8px">${H(e.name)} · ${H(e.role || '')} · ${H(dept(e.dept).short)}</p>
        <p class="small">${used ? 'พนักงานคนนี้ถูกจัดอยู่ในรอบงาน — เมื่อลบ ชื่อจะถูกเอาออกจากทุกรอบและคำขอต่อรอบด้วย ' : ''}ลบแล้วกู้คืนไม่ได้</p>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="emDelOk">ลบพนักงาน</button></div>`, bg => bg.querySelector('#emDelOk').onclick = () => {
        unassign(e.id); Store.db.employees = Store.db.employees.filter(x => x.id !== e.id);
        Store.save(); App.closeModal(); App.render(); App.toast('ลบ ' + e.name + ' แล้ว');
      });
    }
  });
  const RD = el => App.state.ui.rulesDraft[el.dataset.d][Number(el.dataset.i)];
  App.CHG.ruleUp = el => { RD(el).upTo = el.value; App.render(); };
  App.CHG.ruleN = el => { RD(el).n = el.value; };
  App.CHG.ruleOn = el => { const R = App.state.ui.rulesDraft; R[el.dataset.d] = el.checked ? (R[el.dataset.d] || [{ upTo: 50, n: 1 }, { upTo: null, n: 2 }]) : null; if (!el.checked) delete R[el.dataset.d]; App.render(); };
  App.CHG.empQ = el => { App.state.ui.emps.q = el.value; App.state.ui.emps.page = 1; App.render(); };
  App.CHG.stOtShift = el => { const o = curOt(); const [dp, sid] = el.value.split('|'); o.dept = dp; o.shiftId = sid; o.staff = []; persistOt(o); App.render(); };
  App.CHG.stOtReason = el => { const o = curOt(); o.reason = el.value; persistOt(o); };
  function persistOt(o, force) {
    const p = SP();
    if (!o.id && (force || o.staff.length)) { o.id = 'ot' + U.uid(); o.status = 'draft'; p.ot.push(o); App.state.ui.staff.ot = o.id; App.state.ui.staffOtDraft = null; }
    Store.save();
  }

  /* =====================================================================
     USERS (1m)
     ===================================================================== */
  const ROLE_CLS = { sales: 'r-sales', manager: 'r-manager', admin: 'r-admin', head: 'r-head', acc: 'r-acc' };
  const ST = { active: ['ใช้งานอยู่', 'st-confirmed'], invited: ['รอ', 'st-sent'], disabled: ['ปิดใช้', 'st-draft'] };
  const depts = ['ฝ่ายขายและการตลาด', 'ผู้บริหาร', 'ครัว / F&B', 'จัดเลี้ยง', 'ช่าง / AV', 'ส่วนหน้า', 'แม่บ้าน', 'บัญชี'];
  const fmtLast = iso => {
    if (!iso) return 'ยังไม่เคยเข้า';
    const d = iso.slice(0, 10), t = U.today();
    if (d === t) return 'วันนี้ ' + iso.slice(11, 16);
    if (d === U.addDays(t, -1)) return 'เมื่อวาน ' + iso.slice(11, 16);
    return U.fmtD(d, { year: false });
  };
  App.V.users = {
    title: 'ผู้ใช้งาน',
    render() {
      const db = Store.db, canEdit = App.can('users');
      const ui = App.state.ui.users || (App.state.ui.users = { q: '', role: 'all', page: 1, sel: App.state.uid, draft: null });
      const list = db.users.filter(u => (ui.role === 'all' || u.role === ui.role) && (!ui.q || (u.name + ' ' + u.email).toLowerCase().includes(ui.q.toLowerCase())));
      const per = 7, pages = Math.max(1, Math.ceil(list.length / per)); ui.page = Math.min(ui.page, pages);
      const rows = list.slice((ui.page - 1) * per, ui.page * per);
      const u = db.users.find(x => x.id === ui.sel) || db.users[0];
      if (!ui.draft || ui.draft.id !== u.id) ui.draft = JSON.parse(JSON.stringify(u));
      const dr = ui.draft;
      const cnt = r => db.users.filter(x => r === 'all' || x.role === r).length;
      return `
<div class="page-hd"><div><h1>ผู้ใช้งาน</h1><div class="small muted">${db.users.length} คน · ใช้งานอยู่ ${db.users.filter(x => x.status === 'active').length} · รอ ${db.users.filter(x => x.status === 'invited').length} · ปิดใช้ ${db.users.filter(x => x.status === 'disabled').length}</div></div>
  ${canEdit ? '<div class="row wrap"><button class="btn ghost" data-act="usCsv">นำเข้าจากไฟล์ CSV</button><button class="btn red" data-act="usInvite">+ เพิ่มผู้ใช้งาน</button></div>' : ''}</div>
<div class="row wrap" style="margin-bottom:16px">
  <input class="input" style="max-width:260px;border-radius:999px" placeholder="ค้นหาชื่อ อีเมล…" value="${H(ui.q)}" data-change="usQ" aria-label="ค้นหาผู้ใช้">
  <button class="chip ${ui.role === 'all' ? 'on' : ''}" data-act="usRole" data-v="all">ทั้งหมด <span class="n">${cnt('all')}</span></button>
  ${D.ROLES.map(r => `<button class="chip ${ui.role === r.id ? 'on' : ''}" data-act="usRole" data-v="${r.id}">${r.name} <span class="n">${cnt(r.id)}</span></button>`).join('')}
</div>
<div class="split wide-side">
  <section class="card" style="overflow:hidden">
    <table class="tbl resp"><thead><tr><th>ผู้ใช้</th><th>บทบาท</th><th>แผนก</th><th>เข้าใช้ล่าสุด</th><th class="r">สถานะ</th></tr></thead><tbody>
    ${rows.map(x => `<tr class="click" data-act="usSel" data-v="${x.id}" style="${x.id === u.id ? 'background:var(--red-tint);box-shadow:inset 3px 0 var(--red)' : ''};${x.status === 'disabled' ? 'opacity:.55' : ''}">
      <td><div class="row"><span class="avatar sm" style="background:${x.role === 'admin' ? 'var(--ink)' : x.role === 'head' ? 'var(--yellow-soft)' : x.role === 'acc' ? 'var(--navy-soft)' : x.role === 'manager' ? 'var(--red-soft)' : 'var(--teal-soft)'};color:${x.role === 'admin' ? 'var(--yellow)' : 'var(--ink)'}">${H(U.initials(x.name))}</span><div style="min-width:0"><div class="row" style="gap:6px;flex-wrap:nowrap"><span style="font-weight:500">${H(x.name)}</span>${canEdit ? `<button class="icon-btn us-edit" data-act="usEdit" data-v="${x.id}" title="แก้ไขข้อมูลผู้ใช้" aria-label="แก้ไขข้อมูล ${H(x.name)}">${PENCIL}</button>` : ''}</div><div class="sub" style="overflow:hidden;text-overflow:ellipsis">${H(x.email)}</div></div></div></td>
      <td class="hide-m"><span class="role-badge ${ROLE_CLS[x.role]}">${H(D.ROLES.find(r => r.id === x.role).name)}</span></td><td class="hide-m">${H(x.dept)}</td><td class="hide-m small">${fmtLast(x.lastLogin)}</td>
      <td class="r"><span class="badge ${ST[x.status][1]}">${ST[x.status][0]}</span></td></tr>`).join('') || '<tr><td colspan="5"><div class="empty">ไม่พบผู้ใช้</div></td></tr>'}
    </tbody></table>
    <div class="between" style="padding:14px 22px;border-top:1px solid var(--line)"><span class="small muted">แสดง ${list.length ? (ui.page - 1) * per + 1 : 0}–${Math.min(ui.page * per, list.length)} จาก ${list.length} คน</span>
      <div class="pager"><button data-act="usPage" data-v="${ui.page - 1}" ${ui.page <= 1 ? 'disabled' : ''} aria-label="ก่อนหน้า">‹</button>${Array.from({ length: pages }, (_, i) => `<button class="${ui.page === i + 1 ? 'on' : ''}" data-act="usPage" data-v="${i + 1}">${i + 1}</button>`).join('')}<button data-act="usPage" data-v="${ui.page + 1}" ${ui.page >= pages ? 'disabled' : ''} aria-label="ถัดไป">›</button></div></div>
  </section>
  <aside class="card pad sticky" id="usPanel">
    <div class="row" style="align-items:flex-start;gap:14px"><span class="avatar lg">${H(U.initials(u.name))}</span><div style="flex:1;min-width:0"><h2 style="font-size:20px">${H(u.name)}</h2><div class="small muted" style="word-break:break-all">${H(u.email)}${u.phone ? ' · ' + H(u.phone) : ''}</div></div><span class="badge ${ST[u.status][1]}">${ST[u.status][0]}</span></div>
    <div class="small muted" style="margin:18px 0 8px">บทบาท</div>
    <div class="row wrap" style="gap:6px">${D.ROLES.map(r => `<button class="chip ${dr.role === r.id ? 'on' : ''}" data-act="usDraftRole" data-v="${r.id}" ${canEdit ? '' : 'disabled'}>${r.name}</button>`).join('')}</div>
    <div class="small muted" style="margin:16px 0 8px">แผนก</div>
    <select class="input" data-change="usDept" ${canEdit ? '' : 'disabled'}>${depts.concat(depts.includes(dr.dept) ? [] : [dr.dept]).map(d => `<option ${d === dr.dept ? 'selected' : ''}>${H(d)}</option>`).join('')}</select>
    <div class="small muted" style="margin:16px 0 4px">สิทธิ์การใช้งาน</div>
    ${D.PERMS.map(p => `<div class="perm"><div><div>${H(p.name)}</div>${p.note ? `<div class="xs muted">${H(p.note)}</div>` : ''}</div>${App.sw(p.name, dr.perms[p.id], `data-change="usPerm" data-p="${p.id}" ${canEdit ? '' : 'disabled'}`)}</div>`).join('')}
    <div style="background:var(--bg);border-radius:12px;padding:12px 14px;margin:14px 0" class="small">
      <div class="between"><span class="muted">เข้าสู่ระบบล่าสุด</span><span>${fmtLast(u.lastLogin)}${u.lastDevice ? ' · ' + H(u.lastDevice) : ''}</span></div>
      <div class="between" style="margin-top:6px"><span class="muted">ยืนยันตัวตน 2 ขั้นตอน</span>${canEdit ? `<button class="link" data-act="us2fa">${dr.twoFA ? 'เปิดอยู่' : 'ปิดอยู่'}</button>` : `<b class="teal">${dr.twoFA ? 'เปิดอยู่' : 'ปิดอยู่'}</b>`}</div></div>
    ${canEdit ? `<div class="note" style="margin-bottom:14px">การเปลี่ยนบทบาทมีผลทันที ใบเสนอราคาที่ส่งไปแล้วจะไม่เปลี่ยน</div>
    <div class="row"><button class="btn ghost" style="flex:1" data-act="usResetPw">เปลี่ยนรหัสผ่าน</button><button class="btn red" style="flex:1" data-act="usSave">บันทึก</button></div>
    ${u.id !== App.state.uid ? `<button class="link red" style="display:block;margin:14px auto 0" data-act="usToggle">${u.status === 'disabled' ? 'เปิดใช้งานบัญชีนี้' : 'ปิดใช้งานบัญชีนี้'}</button>` : ''}
    ${u.id !== App.state.uid && App.me().role === 'admin' ? '<button class="btn ghost red" style="width:100%;margin-top:14px" data-act="usDelete">ลบบัญชีนี้</button>' : ''}`
    : '<div class="note">คุณดูข้อมูลได้อย่างเดียว — การแก้ไขผู้ใช้ต้องใช้สิทธิ์ “จัดการผู้ใช้งาน”</div>'}
  </aside>
</div>`;
    }
  };
  const UU = () => App.state.ui.users;
  Object.assign(App.ACT, {
    usRole(el) { UU().role = el.dataset.v; UU().page = 1; App.render(); },
    usPage(el) { UU().page = Number(el.dataset.v); App.render(); },
    usSel(el) { UU().sel = el.dataset.v; UU().draft = null; App.render(); if (window.innerWidth <= 1100) document.getElementById('usPanel')?.scrollIntoView({ behavior: 'smooth' }); },
    usDraftRole(el) { const d = UU().draft; d.role = el.dataset.v; d.perms = Object.fromEntries(D.PERMS.map(p => [p.id, D.ROLE_PERMS[d.role].includes(p.id)])); App.render(); },
    us2fa() { const d = UU().draft; d.twoFA = !d.twoFA; App.render(); },
    usSave() {
      const d = UU().draft, u = Store.db.users.find(x => x.id === d.id);
      if (u.id === App.state.uid && !d.perms.users) { App.toast('ไม่สามารถถอดสิทธิ์จัดการผู้ใช้ของตัวเองได้'); return; }
      Object.assign(u, { role: d.role, dept: d.dept, perms: d.perms, twoFA: d.twoFA }); Store.save(); UU().draft = null; App.render(); App.toast('บันทึกแล้ว');
    },
    usToggle() {
      const u = Store.db.users.find(x => x.id === UU().sel);
      const run = () => { u.status = u.status === 'disabled' ? 'active' : 'disabled'; Store.save(); UU().draft = null; App.render(); };
      if (u.status !== 'disabled') App.confirm('ปิดใช้งานบัญชี ' + u.name + '?', run, 'ปิดใช้งาน'); else run();
    },
    usEdit(el) {
      const u = Store.db.users.find(x => x.id === el.dataset.v), me = App.me();
      if (!u) return;
      if (u.role === 'admin' && me.role !== 'admin') { App.toast('เฉพาะผู้ดูแลระบบที่แก้ไขข้อมูลของผู้ดูแลระบบได้'); return; }
      UU().sel = u.id; UU().draft = null; App.render();
      const self = u.id === me.id, d = JSON.parse(JSON.stringify(u));
      const roles = D.ROLES.filter(r => r.id !== 'admin' || me.role === 'admin' || u.role === 'admin');
      App.modal(`<h2>แก้ไขข้อมูลผู้ใช้งาน</h2><p class="muted" style="margin-top:-8px">${H(u.name)} · รหัส ${H(u.id)}</p>
        <div class="field"><label class="f" for="ueName">ชื่อ-นามสกุล</label><input class="input" id="ueName" name="name" value="${H(u.name)}"></div>
        <div class="grid g2" style="margin-top:14px">
          <div><label class="f" for="ueEmail">อีเมล (ใช้เข้าสู่ระบบ)</label><input class="input" id="ueEmail" type="email" name="email" value="${H(u.email)}" autocomplete="off"></div>
          <div><label class="f" for="uePhone">เบอร์โทร</label><input class="input" id="uePhone" type="tel" name="phone" value="${H(u.phone || '')}"></div>
        </div>
        <div class="grid g2" style="margin-top:14px">
          <div><label class="f" for="ueRole">บทบาท</label><select class="input" id="ueRole" name="role">${roles.map(r => `<option value="${r.id}" ${r.id === u.role ? 'selected' : ''}>${r.name}</option>`).join('')}</select></div>
          <div><label class="f" for="ueDept">แผนก</label><input class="input" id="ueDept" name="dept" value="${H(u.dept || '')}" list="ueDepts"><datalist id="ueDepts">${depts.map(x => `<option value="${H(x)}">`).join('')}</datalist></div>
        </div>
        <div class="grid g2" style="margin-top:14px">
          <div><label class="f" for="ueStatus">สถานะ</label><select class="input" id="ueStatus" name="status" ${self ? 'disabled' : ''}>${Object.entries(ST).map(([k, v]) => `<option value="${k}" ${k === u.status ? 'selected' : ''}>${v[0]}</option>`).join('')}</select></div>
          <div><label class="f">ยืนยันตัวตน 2 ขั้นตอน</label><div style="padding-top:8px">${App.sw('ยืนยันตัวตน 2 ขั้นตอน', d.twoFA, 'id="ue2fa"')}</div></div>
        </div>
        <p class="xs muted" style="margin:10px 0 0">เปลี่ยนบทบาทแล้ว สิทธิ์การใช้งานจะตั้งตามบทบาทใหม่ (ปรับรายข้อได้ที่แผงรายละเอียดผู้ใช้)</p>
        <button type="button" class="link" id="uePw" style="margin-top:12px">เปลี่ยนรหัสผ่าน…</button>
        <div class="small red" id="ueErr" style="min-height:20px;margin-top:6px" role="alert"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ueOk">บันทึก</button></div>`, bg => {
        const err = bg.querySelector('#ueErr');
        bg.querySelector('#ue2fa').onchange = e => { d.twoFA = e.target.checked; };
        bg.querySelector('#uePw').onclick = () => { App.closeModal(); App.ACT.usResetPw(); };
        bg.querySelector('#ueOk').onclick = async e => {
          const f = App.formData(bg), btn = e.currentTarget;
          f.name = (f.name || '').trim(); f.email = (f.email || '').trim(); f.dept = (f.dept || '').trim();
          if (!f.name) { err.textContent = 'กรุณากรอกชื่อ'; return; }
          if (!/.+@.+\..+/.test(f.email)) { err.textContent = 'อีเมลไม่ถูกต้อง'; return; }
          if (Store.db.users.some(x => x.id !== u.id && x.email.toLowerCase() === f.email.toLowerCase())) { err.textContent = 'อีเมลนี้มีผู้ใช้อยู่แล้ว'; return; }
          // a new role brings that role's default permissions; same role keeps the current ones
          const perms = f.role === u.role ? d.perms : Object.fromEntries(D.PERMS.map(p => [p.id, D.ROLE_PERMS[f.role].includes(p.id)]));
          if (self && !perms.users) { err.textContent = 'ไม่สามารถเปลี่ยนบทบาทของตัวเองเป็นบทบาทที่ไม่มีสิทธิ์จัดการผู้ใช้ได้'; return; }
          const next = Object.assign({}, u, { name: f.name, email: f.email, phone: (f.phone || '').trim(), role: f.role, dept: f.dept || u.dept, status: self ? u.status : f.status, perms, twoFA: d.twoFA });
          btn.disabled = true;
          try {
            if (g.API && g.API.online) {
              const j = await g.API.req('PUT', 'users/' + encodeURIComponent(u.id), { data: next, version: u._v || 0 });
              g.API.applyUser(j.user); Store.write();
            } else { Object.assign(u, next); Store.save(); }
          } catch (x) { err.textContent = x.message; btn.disabled = false; return; }
          App.closeModal(); UU().draft = null; App.render(); App.toast('บันทึกข้อมูลของ ' + f.name + ' แล้ว');
        };
      });
    },
    usDelete() {
      const u = Store.db.users.find(x => x.id === UU().sel), me = App.me();
      if (!u || me.role !== 'admin') { App.toast('เฉพาะผู้ดูแลระบบที่ลบบัญชีผู้ใช้ได้'); return; }
      if (u.id === me.id) { App.toast('ไม่สามารถลบบัญชีของตัวเองได้'); return; }
      App.modal(`<h2>ลบบัญชีนี้?</h2><p class="muted" style="margin-top:-8px">${H(u.name)} · ${H(u.email)}</p>
        <p class="small">ผู้ใช้นี้จะเข้าสู่ระบบไม่ได้อีกและกู้คืนไม่ได้ ส่วนใบเสนอราคาและใบฟังก์ชันที่เคยทำไว้ยังอยู่ครบ — ถ้าต้องการระงับชั่วคราว ให้ใช้ “ปิดใช้งานบัญชีนี้” แทน</p>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="delYes">ลบบัญชี</button></div>`, bg => bg.querySelector('#delYes').onclick = async e => {
        e.currentTarget.disabled = true;
        try {
          if (g.API && g.API.online) { await g.API.req('DELETE', 'users/' + encodeURIComponent(u.id)); delete g.API.userSnap[u.id]; }
        } catch (x) { App.closeModal(); App.toast(x.message); return; }
        Store.db.users = Store.db.users.filter(x => x.id !== u.id);
        if (g.API && g.API.online) Store.write(); else Store.save();
        App.closeModal(); UU().sel = me.id; UU().draft = null; App.render(); App.toast('ลบบัญชี ' + u.name + ' แล้ว');
      });
    },
    usResetPw() {
      const u = Store.db.users.find(x => x.id === UU().sel), me = App.me();
      if (u.role === 'admin' && me.role !== 'admin') { App.toast('เฉพาะผู้ดูแลระบบที่เปลี่ยนรหัสผ่านของผู้ดูแลระบบได้'); return; }
      App.modal(`<h2>เปลี่ยนรหัสผ่าน</h2><p class="muted" style="margin-top:-8px">${H(u.name)} · ${H(u.email)}</p>
        <div class="field"><label class="f" for="cpPw">รหัสผ่านใหม่ (อย่างน้อย 6 ตัวอักษร)</label><div class="pw-wrap"><input class="input" id="cpPw" type="password" autocomplete="new-password"><button type="button" id="cpShow">แสดง</button></div></div>
        <div class="field"><label class="f" for="cpPw2">ยืนยันรหัสผ่านใหม่</label><input class="input" id="cpPw2" type="password" autocomplete="new-password"></div>
        <button type="button" class="link" id="cpGen" style="margin-top:8px">สุ่มรหัสผ่านให้</button>
        ${u.id === me.id ? '' : '<p class="small muted" style="margin:10px 0 0">ผู้ใช้คนนี้จะถูกออกจากระบบทุกเครื่อง และต้องเข้าสู่ระบบใหม่ด้วยรหัสผ่านนี้</p>'}
        <div class="small red" id="cpErr" style="min-height:20px;margin-top:6px" role="alert"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="cpOk">บันทึกรหัสผ่าน</button></div>`, bg => {
        const pw = bg.querySelector('#cpPw'), pw2 = bg.querySelector('#cpPw2'), show = bg.querySelector('#cpShow'), err = bg.querySelector('#cpErr');
        const setShown = on => { pw.type = pw2.type = on ? 'text' : 'password'; show.textContent = on ? 'ซ่อน' : 'แสดง'; };
        show.onclick = () => setShown(pw.type === 'password');
        bg.querySelector('#cpGen').onclick = () => {
          const abc = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789', r = crypto.getRandomValues(new Uint32Array(10));
          pw.value = pw2.value = Array.from(r, x => abc[x % abc.length]).join(''); setShown(true); err.textContent = '';
        };
        bg.querySelector('#cpOk').onclick = async e => {
          const btn = e.currentTarget;
          if (pw.value.length < 6) { err.textContent = 'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร'; return; }
          if (pw.value !== pw2.value) { err.textContent = 'รหัสผ่านทั้งสองช่องไม่ตรงกัน'; return; }
          btn.disabled = true;
          try {
            if (g.API && g.API.online) await g.API.req('POST', 'users/' + encodeURIComponent(u.id) + '/password', { password: pw.value });
            else { u.pw = await g.hashPw(pw.value); Store.save(); }
          } catch (x) { err.textContent = x.message; btn.disabled = false; return; }
          App.closeModal(); App.toast('เปลี่ยนรหัสผ่านของ ' + u.name + ' แล้ว');
        };
      });
    },
    usInvite() {
      App.modal(`<h2>เพิ่มผู้ใช้งาน</h2>
        <div class="field"><label class="f">ชื่อ-นามสกุล</label><input class="input" name="name"></div>
        <div class="field"><label class="f">อีเมล</label><input class="input" type="email" name="email" placeholder="name@luckswangroup.com" autocomplete="off"></div>
        <div class="field"><label class="f">เบอร์โทร</label><input class="input" type="tel" name="phone"></div>
        <div class="grid g2" style="margin-top:14px"><div><label class="f">บทบาท</label><select class="input" name="role">${D.ROLES.map(r => `<option value="${r.id}">${r.name}</option>`).join('')}</select></div><div><label class="f">แผนก</label><input class="input" name="dept" value="ฝ่ายขายและการตลาด"></div></div>
        <div class="grid g2" style="margin-top:14px">
          <div><label class="f" for="nuPw">รหัสผ่าน (อย่างน้อย 6 ตัวอักษร)</label><div class="pw-wrap"><input class="input" id="nuPw" name="pw" type="password" autocomplete="new-password"><button type="button" id="nuShow">แสดง</button></div></div>
          <div><label class="f" for="nuPw2">ยืนยันรหัสผ่าน</label><input class="input" id="nuPw2" name="pw2" type="password" autocomplete="new-password"></div>
        </div>
        <button type="button" class="link" id="nuGen" style="margin-top:8px">สุ่มรหัสผ่านให้</button>
        <div class="small red" id="invErr" style="min-height:20px;margin-top:6px" role="alert"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ok">เพิ่มผู้ใช้งาน</button></div>`, bg => {
        const pw = bg.querySelector('#nuPw'), pw2 = bg.querySelector('#nuPw2'), show = bg.querySelector('#nuShow'), err = bg.querySelector('#invErr');
        const setShown = on => { pw.type = pw2.type = on ? 'text' : 'password'; show.textContent = on ? 'ซ่อน' : 'แสดง'; };
        show.onclick = () => setShown(pw.type === 'password');
        bg.querySelector('#nuGen').onclick = () => {
          const abc = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789', r = crypto.getRandomValues(new Uint32Array(10));
          pw.value = pw2.value = Array.from(r, x => abc[x % abc.length]).join(''); setShown(true); err.textContent = '';
        };
        bg.querySelector('#ok').onclick = async e => {
          const f = App.formData(bg), btn = e.currentTarget;
          f.name = f.name.trim(); f.email = f.email.trim();
          if (!f.name || !/.+@.+\..+/.test(f.email)) { err.textContent = 'กรุณากรอกชื่อและอีเมลให้ถูกต้อง'; return; }
          if (Store.db.users.some(u => u.email.toLowerCase() === f.email.toLowerCase())) { err.textContent = 'อีเมลนี้มีอยู่แล้ว'; return; }
          if (f.pw.length < 6) { err.textContent = 'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร'; return; }
          if (f.pw !== f.pw2) { err.textContent = 'รหัสผ่านทั้งสองช่องไม่ตรงกัน'; return; }
          const u = makeUser(f);
          btn.disabled = true;
          try {
            if (g.API && g.API.online) {
              // create on the server directly so the password goes with it (never stored in the local cache)
              const j = await g.API.req('PUT', 'users/' + encodeURIComponent(u.id), { data: u, version: 0, password: f.pw });
              g.API.applyUser(j.user); Store.write();
            } else { u.pw = await g.hashPw(f.pw); Store.db.users.push(u); Store.save(); }
          } catch (x) { err.textContent = x.message; btn.disabled = false; return; }
          App.closeModal(); UU().sel = u.id; UU().draft = null; App.render();
          const link = location.href.split('#')[0];
          App.modal(`<h2>เพิ่ม ${H(u.name)} แล้ว</h2><p class="muted">แจ้งข้อมูลนี้ให้ผู้ใช้ และแจ้งรหัสผ่านที่ตั้งไว้แยกต่างหาก (เช่น บอกด้วยตัวเองหรือทาง LINE ส่วนตัว)</p><div class="card pad small">ลิงก์: ${H(link)}<br>อีเมล: ${H(u.email)}</div>
            <div class="acts"><a class="btn ghost" href="mailto:${encodeURIComponent(u.email)}?subject=${encodeURIComponent('บัญชีเข้าใช้ระบบฝ่ายขาย ลักษวรรณ รีสอร์ต')}&body=${encodeURIComponent('ลิงก์: ' + link + '\nอีเมล: ' + u.email + '\nรหัสผ่าน: ตามที่ผู้ดูแลระบบแจ้ง\nกรุณาเปลี่ยนรหัสผ่านหลังเข้าสู่ระบบ (เมนูรูปโปรไฟล์ → เปลี่ยนรหัสผ่าน)')}">ส่งอีเมลแจ้งผู้ใช้</a><button class="btn dark" data-act="closeModal">ปิด</button></div>`);
        };
      });
    },
    usCsv() {
      App.modal(`<h2>นำเข้าผู้ใช้จาก CSV</h2><p class="small muted">คอลัมน์: ชื่อ, อีเมล, บทบาท (sales/manager/head/acc/admin), แผนก, เบอร์โทร — แถวแรกเป็นหัวตารางได้</p>
        <input type="file" accept=".csv,text/csv" class="input" style="padding-top:10px" id="csvF"><div class="small" id="csvMsg" style="margin-top:10px"></div>
        <div class="acts"><button class="btn ghost" data-act="usCsvTpl">ดาวน์โหลดตัวอย่าง</button><button class="btn dark" data-act="closeModal">ปิด</button></div>`, bg => {
        bg.querySelector('#csvF').onchange = e => {
          const file = e.target.files[0]; if (!file) return;
          file.text().then(t => {
            const rows = U.parseCSV(t.replace(/^﻿/, '')); let n = 0, skip = 0;
            for (const r of rows) {
              const [name, email, role, dp, phone] = r.map(x => (x || '').trim());
              if (!email || !/@/.test(email) || /email|อีเมล/i.test(email)) { skip++; continue; }
              if (Store.db.users.some(u => u.email.toLowerCase() === email.toLowerCase())) { skip++; continue; }
              const rid = D.ROLES.find(x => x.id === role || x.name === role) || D.ROLES[0];
              addUser({ name: name || email, email, role: rid.id, dept: dp || 'ฝ่ายขายและการตลาด', phone }); n++;
            }
            Store.save(); bg.querySelector('#csvMsg').textContent = `นำเข้า ${n} คน · ข้าม ${skip} แถว`; App.render();
          });
        };
      });
    },
    usCsvTpl() { U.download('users-template.csv', '﻿ชื่อ,อีเมล,บทบาท,แผนก,เบอร์โทร\nสมหญิง ใจดี,somying@luckswangroup.com,sales,ฝ่ายขายและการตลาด,081-000-0000\n', 'text/csv'); }
  });
  function makeUser(f) {
    const ids = Store.db.users.map(u => parseInt(u.id.slice(1), 10) || 0);
    return { id: 'u' + (Math.max(0, ...ids) + 1), name: f.name, email: f.email, role: f.role, dept: f.dept, status: 'invited', phone: f.phone || '', lastLogin: null,
      perms: Object.fromEntries(D.PERMS.map(p => [p.id, D.ROLE_PERMS[f.role].includes(p.id)])), twoFA: false, pw: null };
  }
  function addUser(f) { const u = makeUser(f); Store.db.users.push(u); return u; }
  Object.assign(App.CHG, {
    usQ(el) { UU().q = el.value; UU().page = 1; App.render(); },
    usDept(el) { UU().draft.dept = el.value; },
    usPerm(el) { UU().draft.perms[el.dataset.p] = el.checked; }
  });
})(window);
