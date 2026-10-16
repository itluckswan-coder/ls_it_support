/* Luck Swan Sales — รายละเอียดฟังก์ชัน: the hotel's "Banquet Function Event" form for a function sheet.
 * Every field starts from the quote and the sheet; anything typed in is kept on the sheet (b.fn) and prints
 * on A4 landscape like the paper form. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  const curB = () => Store.db.beos.find(b => b.id === App.state.params[0]);
  const canEdit = () => App.can('create_quote') || App.can('staff');
  const lines = t => String(t || '').split('\n').map(s => s.trim()).filter(Boolean);
  const nl = t => H(t || '').replace(/\n/g, '<br>');
  const money = v => U.n(Math.round(Number(v) || 0));

  // RS<number>/<2-digit Buddhist year>, continuing from the highest number used this year
  function nextNo() {
    const yy = String(U.be(U.d(U.today()).getFullYear())).slice(-2);
    const max = Store.db.beos.reduce((m, b) => { const r = /^RS\s*(\d+)\s*\/\s*(\d{2})$/i.exec((b.fn && b.fn.no) || ''); return r && r[2] === yy ? Math.max(m, Number(r[1])) : m; }, 0);
    return 'RS' + (max + 1) + '/' + yy;
  }

  /* ---------- values taken from the quote and the function sheet ---------- */
  function auto(b, q) {
    const c = App.calc(q), meet = q.items.filter(i => i.sec === 'meet'), food = q.items.filter(i => i.sec === 'food'), rooms = q.items.filter(i => i.sec === 'room');
    const room = D.MEETING_ROOMS.find(r => r.id === q.room) || D.MEETING_ROOMS.find(r => meet.some(i => i.ref === r.id));
    const times = q.items.flatMap(i => i.sched || []).filter(s => s.start);
    const t0 = times.map(s => s.start).sort()[0], t1 = times.map(s => s.end || s.start).sort().slice(-1)[0];
    const layout = (meet.find(i => i.layout) || {}).layout;
    const isParty = layout === 'banquet' || layout === 'cocktail' || food.some(i => (D.FOOD.find(f => f.id === i.ref) || {}).cat === 'gala');
    const dept = id => ((b.depts || {})[id] || {}).items || [];
    const unitTxt = i => i.u1 || 'ท่าน';
    // menu: one block per food line, titled by its serving time
    const menu = food.map(i => {
      const f = D.FOOD.find(x => x.id === i.ref), s = (i.sched || [])[0];
      const dishes = f && Array.isArray(f.items) ? f.items.map(D.itemText).filter(Boolean) : [];
      return { time: s ? s.start + (s.end ? ' - ' + s.end : '') + 'น.' : '', title: i.name, items: dishes.join('\n') };
    });
    const cost = q.items.filter(i => U.lineTotal(i)).map(i => `${i.name} ${money(i.price)} × ${i.q1}${i.q2 != null && i.q2 !== '' ? ' × ' + i.q2 : ''} = ${money(U.lineTotal(i))} บาท`);
    if (c.disc) cost.push(`ส่วนลด −${money(c.disc)} บาท`);
    if (c.svc) cost.push(`ค่าบริการ 10% ${money(c.svc)} บาท`);
    if (c.vat) cost.push(`VAT 7% ${money(c.vat)} บาท`);
    const sched = Object.entries(b.schedule || {}).flatMap(([d, rows]) => rows.slice().sort((a, x) => a.time.localeCompare(x.time)).map(r => (Object.keys(b.schedule).length > 1 ? 'วันที่ ' + d + ' ' : '') + r.time + ' ' + r.task));
    const owner = Store.db.users.find(u => u.id === q.owner);
    return {
      no: '', issued: U.today(), org: q.clientFull || q.client || '', coord: [q.contact, q.phone].filter(Boolean).join(' '), quote: q.id,
      date: U.fmtRange(q.start, q.end, { long: true }),
      venue: [isParty ? 'งานเลี้ยง' : 'ประชุม', room ? '@' + (room.th || room.name).replace(/^ห้อง/, '') : '', t0 ? t0 + 'น. - ' + t1 + 'น.' : ''].filter(Boolean).join(' '),
      type: isParty ? 'งานเลี้ยง' : 'ประชุม / สัมมนา',
      pax: b.confirmPax || (q.pax + ' ท่าน'),
      pricing: q.items.filter(i => i.sec !== 'room' && i.price).map(i => `${i.name}@${money(i.price)}/${unitTxt(i)}`).join('\n'),
      layout: dept('banquet').map(x => '- ' + x).join('\n'),
      fyi: [b.note, ...sched].filter(Boolean).join('\n'),
      stage: '', vinyl: '',
      menu, kitchen: dept('kitchen').join('\n'),
      hk: dept('hk').join('\n'), av: dept('av').map(x => '/ ' + x).join('\n'),
      total: money(c.net) + ' บาท', cost: cost.join('\n'), bill: 'แจ้งภายหลัง',
      checkin: rooms.length ? U.fmtD(q.start) : '-', checkout: rooms.length ? U.fmtD(U.addDays(q.end, 1)) : '-',
      roomType: rooms.length ? rooms.map(i => i.name).join('\n') : '-', roomQty: rooms.length ? rooms.map(i => i.q1 + ' ห้อง × ' + (i.q2 || 1) + ' คืน').join('\n') : '-',
      front: rooms.length ? dept('front').join('\n') : '** ไม่มีห้องพัก **',
      by: owner ? owner.name : '', byTitle: 'Sales and Marketing',
      appr: b.signoff ? b.signoff.byName : '', apprTitle: b.signoff ? 'อนุมัติ ' + U.fmtD(b.signoff.at.slice(0, 10)) : ''
    };
  }
  // what shows: typed-in value if there is one, else the value from the system
  function values(b, q) {
    const a = auto(b, q), f = b.fn || {}, v = {};
    for (const k of Object.keys(a)) v[k] = f[k] != null ? f[k] : a[k];
    v.menu = Array.isArray(f.menu) ? f.menu : a.menu;
    v.appr = a.appr; v.apprTitle = a.apprTitle; // always the real sign-off
    return { v, a };
  }

  /* ---------- extra details added with "+" (b.fn.extra) ---------- */
  const XSEC = [['event', 'ข้อมูลงาน'], ['banquet', 'แผนกจัดเลี้ยง / ห้องอาหาร'], ['hk', 'แผนกแม่บ้าน'], ['av', 'แผนกช่างเครื่องเสียง'], ['acc', 'แผนกบัญชี / การเงิน'], ['front', 'แผนกต้อนรับ'], ['other', 'รายละเอียดเพิ่มเติม']];
  const xName = k => (XSEC.find(x => x[0] === k) || XSEC[XSEC.length - 1])[1];

  /* ---------- the paper form ---------- */
  function paper(v, extra = []) {
    const cell = (t, cls = '') => `<td class="${cls}">${nl(t)}</td>`;
    const xs = k => extra.filter(x => (XSEC.some(s => s[0] === x.sec) ? x.sec : 'other') === k);
    const xth = x => `<th>${H(x.label || '')}${x.label ? ' :' : ''}</th>`;
    const xIn = k => xs(k).map(x => `<div class="fnf-x">${x.label ? `<b>${H(x.label)} :</b> ` : ''}${nl(x.text)}</div>`).join('');
    const acc = xs('acc'), front = xs('front'), mc = Math.max(1, v.menu.length);
    const menuCols = v.menu.length ? v.menu.map(m => `<td class="fnf-menu"><div class="fnf-mt">${H(m.time)}</div><div class="fnf-mh">${H(m.title)}</div>${nl(m.items)}</td>`).join('') : '<td class="fnf-menu">-</td>';
    return `<div class="fnf-paper" id="fnPaper">
  <div class="fnf-head">
    <img src="img/logo-wide.png" alt="Luck Swan Resort">
    <div class="fnf-title"><b>ลักษวรรณ รีสอร์ท</b><br>Banquet Function Event</div>
    <table class="fnf-meta"><tr><th>Function No.</th><td class="fnf-no">${H(v.no || '—')}</td></tr><tr><th>วันที่ออกเอกสาร</th><td>${H(v.issued)}</td></tr><tr><th>Quotation No. :</th><td>${H(v.quote || '-')}</td></tr></table>
  </div>
  <div class="fnf-org"><span><b>สำนักงาน/องค์กร :</b> ${H(v.org)}</span><span><b>ผู้ประสานงาน :</b> ${H(v.coord)}</span></div>
  <table class="fnf-t">
    <tr><th>วันที่จัดงาน :</th>${cell(v.date, 'c')}<th>สถานที่จัดงาน :</th>${cell(v.venue, 'c')}<th>ประเภทงาน :</th>${cell(v.type, 'c')}</tr>
    <tr><th>จำนวนการันตี :</th>${cell(v.pax, 'c')}<td colspan="4" class="c">${nl(v.pricing)}</td></tr>
    ${xs('event').map(x => `<tr>${xth(x)}<td colspan="5">${nl(x.text)}</td></tr>`).join('')}
  </table>
  <table class="fnf-t">
    <tr><th colspan="2" class="sec">แผนกจัดเลี้ยง / ห้องอาหาร</th><th colspan="${Math.max(1, v.menu.length)}" class="sec">เมนูอาหาร</th></tr>
    <tr><th>รูปแบบการจัดห้อง :</th><td>${nl(v.layout)}${v.fyi ? `<div class="fnf-fyi"><b>FYI</b><br>${nl(v.fyi)}</div>` : ''}</td>${menuCols}</tr>
    <tr><th>ข้อความบนเวที :${v.vinyl ? `<div class="fnf-small">${nl(v.vinyl)}</div>` : ''}</th><td class="c">${nl(v.stage || '-')}</td><td colspan="${mc}" class="fnf-small">${v.kitchen ? '<b>ครัว:</b> ' + nl(v.kitchen) : ''}</td></tr>
    ${xs('banquet').map(x => `<tr>${xth(x)}<td colspan="${1 + mc}">${nl(x.text)}</td></tr>`).join('')}
  </table>
  <table class="fnf-t">
    <tr><th class="sec" style="width:50%">แผนกแม่บ้าน</th><th class="sec">แผนกช่างเครื่องเสียง</th></tr>
    <tr><td class="c">${nl(v.hk || '-')}${xIn('hk')}</td><td>${nl(v.av || '-')}${xIn('av')}</td></tr>
  </table>
  <table class="fnf-t">
    <tr><th colspan="2" class="sec" style="width:50%">แผนกบัญชี / การเงิน</th><th colspan="4" class="sec">แผนกต้อนรับ</th></tr>
    <tr><th>ยอดรวม :</th><td class="c fnf-total">${H(v.total)}</td><th>Check-in :</th>${cell(v.checkin, 'c')}<th>Check-out :</th>${cell(v.checkout, 'c')}</tr>
    <tr><th>รายละเอียดค่าใช้จ่าย :</th><td class="c">${nl(v.cost)}</td><th>ห้องพัก (ประเภท) :</th>${cell(v.roomType, 'c')}<th>จำนวนห้อง :</th>${cell(v.roomQty, 'c')}</tr>
    <tr><th>ที่อยู่ออกบิล :</th>${cell(v.bill, 'c')}<td colspan="4" class="c">${nl(v.front)}</td></tr>
    ${Array.from({ length: Math.max(acc.length, front.length) }, (_, i) => `<tr>${acc[i] ? xth(acc[i]) + `<td>${nl(acc[i].text)}</td>` : '<td colspan="2"></td>'}${front[i] ? xth(front[i]) + `<td colspan="3">${nl(front[i].text)}</td>` : '<td colspan="4"></td>'}</tr>`).join('')}
  </table>
  ${xs('other').length ? `<table class="fnf-t"><tr><th colspan="2" class="sec">รายละเอียดเพิ่มเติม</th></tr>${xs('other').map(x => `<tr>${xth(x)}<td>${nl(x.text)}</td></tr>`).join('')}</table>` : ''}
  <div class="fnf-sign">
    <div><span>ออกโดย :</span><div class="fnf-line"></div><div>(${H(v.by || '................................')})</div><div>${H(v.byTitle)}</div></div>
    <div><span>อนุมัติโดย :</span><div class="fnf-line"></div><div>(${H(v.appr || '................................')})</div><div>${H(v.apprTitle || 'ยังไม่อนุมัติ')}</div></div>
  </div>
</div>`;
  }

  /* ---------- the page ---------- */
  const FIELDS = [
    ['หัวเอกสาร', [['no', 'Function No.', 'input'], ['issued', 'วันที่ออกเอกสาร', 'date'], ['org', 'สำนักงาน / องค์กร', 'input'], ['coord', 'ผู้ประสานงาน', 'input']]],
    ['ข้อมูลงาน', [['date', 'วันที่จัดงาน', 'input'], ['venue', 'สถานที่จัดงาน (+ เวลา)', 'input'], ['type', 'ประเภทงาน', 'input'], ['pax', 'จำนวนการันตี', 'input'], ['pricing', 'ราคาต่อหน่วย (บรรทัดละรายการ)', 'area']]],
    ['แผนกจัดเลี้ยง / ห้องอาหาร', [['layout', 'รูปแบบการจัดห้อง', 'area'], ['fyi', 'FYI (กำหนดการ / หมายเหตุ)', 'area'], ['stage', 'ข้อความบนเวที', 'area'], ['vinyl', 'ไวนิล / ป้าย (ขนาด, ตำแหน่ง)', 'input'], ['kitchen', 'หมายเหตุครัว', 'area']]],
    ['แม่บ้าน / ช่างเครื่องเสียง', [['hk', 'แผนกแม่บ้าน', 'area'], ['av', 'แผนกช่างเครื่องเสียง', 'area']]],
    ['บัญชี / การเงิน', [['total', 'ยอดรวม', 'input'], ['cost', 'รายละเอียดค่าใช้จ่าย', 'area'], ['bill', 'ที่อยู่ออกบิล', 'area']]],
    ['แผนกต้อนรับ', [['checkin', 'Check-in', 'input'], ['checkout', 'Check-out', 'input'], ['roomType', 'ห้องพัก (ประเภท)', 'area'], ['roomQty', 'จำนวนห้อง', 'area'], ['front', 'หมายเหตุต้อนรับ', 'area']]],
    ['ลงชื่อ', [['by', 'ออกโดย (ชื่อ)', 'input'], ['byTitle', 'ตำแหน่ง', 'input']]]
  ];
  App.V.fn = {
    title: 'รายละเอียดฟังก์ชัน', nav: 'beo',
    render() {
      const b = curB(), q = b && Store.quote(b.quoteId);
      if (!b || !q) return '<div class="empty">ไม่พบใบฟังก์ชัน — <a href="#/beo">กลับ</a></div>';
      const { v, a } = values(b, q), edit = canEdit(), ui = App.state.ui.fn || (App.state.ui.fn = { tab: 'preview' });
      const tab = edit ? ui.tab : 'preview', own = b.fn || {};
      const head = `<div class="page-hd no-print"><div><div class="small"><a class="link" href="#/beo/${encodeURIComponent(b.id)}">← ใบฟังก์ชัน ${H(b.id)}</a></div><h1>รายละเอียดฟังก์ชัน</h1>
        <div class="small muted">${H(q.title)} · ${H(q.client)} · ${H(v.no || 'ยังไม่มี Function No.')}${b.signoff ? ' · อนุมัติแล้ว' : ''}</div></div>
        <div class="row wrap">${edit ? `<div class="seg white"><button class="${tab === 'preview' ? 'on' : ''}" data-act="fnTab" data-v="preview">ตัวอย่างพิมพ์</button><button class="${tab === 'form' ? 'on' : ''}" data-act="fnTab" data-v="form">แก้ไขข้อมูล</button></div>` : ''}
        ${edit ? '<button class="btn ghost fnf-plus" data-act="fnX" title="เพิ่มรายละเอียดลงในใบรายละเอียดฟังก์ชัน">＋ เพิ่มรายละเอียด</button>' : ''}
        <button class="btn red" data-act="fnPrint">พิมพ์ / PDF</button></div></div>`;
      if (tab === 'preview') return head + `${!v.no && edit ? '<div class="note no-print" style="margin-bottom:12px">ยังไม่มี Function No. — <button class="link" data-act="fnNo">ออกเลขอัตโนมัติ (' + H(nextNo()) + ')</button> หรือกรอกเองในแท็บ “แก้ไขข้อมูล”</div>' : ''}<div class="fnf-wrap">${paper(v, own.extra || [])}</div>`;
      const field = ([k, label, kind]) => {
        const val = v[k] == null ? '' : v[k], mine = own[k] != null && own[k] !== a[k];
        const reset = mine ? `<button class="link xs" data-act="fnReset" data-k="${k}">ใช้ค่าจากระบบ</button>` : `<span class="xs muted">${a[k] ? 'จากระบบ' : ''}</span>`;
        const ctl = kind === 'area' ? `<textarea class="input" rows="${Math.min(8, Math.max(2, lines(val).length + 1))}" data-change="fnSet" data-k="${k}" id="fn_${k}" style="height:auto;padding-top:10px">${H(val)}</textarea>`
          : `<input class="input" ${kind === 'date' ? 'type="date"' : ''} value="${H(val)}" data-change="fnSet" data-k="${k}" id="fn_${k}">`;
        return `<div class="field"><div class="between"><label class="f" for="fn_${k}">${H(label)}</label>${reset}</div>${ctl}${k === 'no' && !val ? `<button class="link small" data-act="fnNo">ออกเลขอัตโนมัติ (${H(nextNo())})</button>` : ''}</div>`;
      };
      const menuEd = `<section class="card pad"><div class="between"><h3>เมนูอาหาร (แยกตามช่วงเวลา)</h3>${Array.isArray(own.menu) ? '<button class="link xs" data-act="fnReset" data-k="menu">ใช้ค่าจากระบบ</button>' : '<span class="xs muted">จากใบเสนอราคา</span>'}</div>
        ${v.menu.map((m, i) => `<div class="fnf-mrow"><div class="grid g2"><div><label class="f">ช่วงเวลา</label><input class="input" value="${H(m.time)}" data-change="fnMenu" data-i="${i}" data-k="time"></div><div><label class="f">หัวข้อ</label><input class="input" value="${H(m.title)}" data-change="fnMenu" data-i="${i}" data-k="title"></div></div>
          <label class="f" style="margin-top:8px">รายการอาหาร (บรรทัดละรายการ)</label><textarea class="input" rows="4" style="height:auto;padding-top:10px" data-change="fnMenu" data-i="${i}" data-k="items">${H(m.items)}</textarea>
          <button class="link red small" data-act="fnMenuDel" data-i="${i}">ลบช่วงนี้</button></div>`).join('')}
        <button class="btn ghost sm" data-act="fnMenuAdd">+ เพิ่มช่วงเวลา</button></section>`;
      const xEd = `<section class="card pad"><div class="between"><h3>รายละเอียดที่เพิ่มเอง</h3><button class="btn ghost sm" data-act="fnX">＋ เพิ่ม</button></div>
        ${(own.extra || []).length ? own.extra.map((x, i) => `<div class="fnf-mrow"><div class="xs muted">${H(xName(x.sec))}</div><b>${H(x.label || '(ไม่มีหัวข้อ)')}</b><div class="small">${nl(x.text)}</div>
          <button class="link small" data-act="fnX" data-i="${i}">แก้ไข</button> · <button class="link red small" data-act="fnXDel" data-i="${i}">ลบ</button></div>`).join('') : '<div class="small muted">ยังไม่มี — กด “＋ เพิ่ม” เพื่อเพิ่มหัวข้อ/รายละเอียดลงในแผนกที่ต้องการ</div>'}</section>`;
      return head + `<div class="note" style="margin-bottom:14px">ทุกช่องดึงจากใบเสนอราคาและใบฟังก์ชันให้แล้ว แก้ได้เลย — ช่องที่แก้เองจะมีปุ่ม “ใช้ค่าจากระบบ” เพื่อกลับไปใช้ค่าอัตโนมัติ · “อนุมัติโดย” มาจากผู้ดูแลระบบที่กดอนุมัติในหน้า Approve</div>
        <div class="fnf-form">${FIELDS.map(([t, fs]) => `<section class="card pad"><h3>${H(t)}</h3>${fs.map(field).join('')}</section>`).join('')}${menuEd}${xEd}</div>`;
    }
  };

  function setFn(fn) {
    const b = curB(); if (!b || !canEdit()) return;
    b.fn = Object.assign({}, b.fn); fn(b.fn, b);
    for (const k of Object.keys(b.fn)) if (b.fn[k] === undefined) delete b.fn[k];
    Store.save(); App.render();
  }
  const menuOf = (f, b) => Array.isArray(f.menu) ? f.menu.map(m => Object.assign({}, m)) : values(b, Store.quote(b.quoteId)).a.menu;
  Object.assign(App.ACT, {
    fnTab(el) { App.state.ui.fn.tab = el.dataset.v; App.render(); },
    fnPrint() { const ui = App.state.ui.fn || (App.state.ui.fn = {}); if (ui.tab !== 'preview') { ui.tab = 'preview'; App.render(); } setTimeout(() => window.print(), 50); },
    fnNo() { setFn(f => { f.no = nextNo(); }); App.toast('ออกเลข ' + curB().fn.no); },
    fnReset(el) { setFn(f => { delete f[el.dataset.k]; }); },
    fnMenuAdd() { setFn((f, b) => { f.menu = menuOf(f, b).concat({ time: '', title: '', items: '' }); }); },
    // "+": a small form laid out like the paper sheet — pick the department, give a heading and the details
    fnX(el) {
      const b = curB(); if (!b || !canEdit()) return;
      const i = el.dataset.i != null ? Number(el.dataset.i) : -1, x = i >= 0 ? (b.fn.extra || [])[i] : { sec: 'banquet', label: '', text: '' };
      if (!x) return;
      App.modal(`<h2>${i >= 0 ? 'แก้ไขรายละเอียด' : 'เพิ่มรายละเอียด'}</h2>
        <div class="field"><label class="f" for="fnxSec">เพิ่มลงในส่วน</label><select class="input" id="fnxSec">${XSEC.map(([k, t]) => `<option value="${k}" ${k === x.sec ? 'selected' : ''}>${H(t)}</option>`).join('')}</select></div>
        <div class="fnf-xform"><table class="fnf-t"><tr><th colspan="2" class="sec" id="fnxHd">${H(xName(x.sec))}</th></tr>
          <tr><th><input class="input" id="fnxLabel" maxlength="80" placeholder="หัวข้อ เช่น ดอกไม้" value="${H(x.label)}"></th>
          <td><textarea class="input" id="fnxText" rows="5" maxlength="2000" placeholder="รายละเอียด (ขึ้นบรรทัดใหม่ได้)">${H(x.text)}</textarea></td></tr></table></div>
        <div class="small red" id="fnxErr"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="fnxOk">${i >= 0 ? 'บันทึก' : 'เพิ่มลงใบ'}</button></div>`, bg => {
        bg.classList.add('wide');
        const sec = bg.querySelector('#fnxSec');
        sec.onchange = () => { bg.querySelector('#fnxHd').textContent = xName(sec.value); };
        bg.querySelector('#fnxOk').onclick = () => {
          const label = bg.querySelector('#fnxLabel').value.trim(), text = bg.querySelector('#fnxText').value.trim();
          if (!label && !text) { bg.querySelector('#fnxErr').textContent = 'กรอกหัวข้อหรือรายละเอียดอย่างน้อยหนึ่งช่อง'; return; }
          App.closeModal();
          setFn(f => { const list = (f.extra || []).map(o => Object.assign({}, o)), item = { sec: sec.value, label, text }; if (i >= 0) list[i] = item; else list.push(item); f.extra = list; });
          App.toast(i >= 0 ? 'บันทึกแล้ว' : 'เพิ่มลงใบรายละเอียดฟังก์ชันแล้ว');
        };
      });
    },
    fnXDel(el) {
      App.confirm('ลบรายละเอียดนี้?', () => setFn(f => { const list = (f.extra || []).slice(); list.splice(Number(el.dataset.i), 1); f.extra = list.length ? list : undefined; }), 'ลบ');
    },
    fnMenuDel(el) { setFn((f, b) => { const m = menuOf(f, b); m.splice(Number(el.dataset.i), 1); f.menu = m; }); }
  });
  Object.assign(App.CHG, {
    fnSet(el) {
      const k = el.dataset.k, val = el.value;
      setFn((f, b) => { const a = auto(b, Store.quote(b.quoteId)); if (val === a[k]) delete f[k]; else f[k] = val.slice(0, 4000); });
    },
    fnMenu(el) { setFn((f, b) => { const m = menuOf(f, b); m[Number(el.dataset.i)][el.dataset.k] = el.value.slice(0, 2000); f.menu = m; }); }
  });
  App.fnNo = b => (b.fn && b.fn.no) || '';
})(window);
