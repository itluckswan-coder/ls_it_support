/* Luck Swan Sales — รายละเอียดฟังก์ชัน: the hotel's "Banquet Function Event" form for a function sheet.
 * Every field starts from the quote and the sheet; anything typed in is kept on the sheet (b.fn) and prints
 * on A4 landscape like the paper form. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  const curB = () => Store.db.beos.find(b => b.id === App.state.params[0]);
  const canEdit = () => App.can('create_quote') || App.can('staff');
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
    v.more = f.more || {}; // lines added inside a cell with its corner "+": { field: [{ t: 'h'|'d', text }] }
    v.appr = a.appr; v.apprTitle = a.apprTitle; // always the real sign-off
    return { v, a };
  }

  /* ---------- extra details added with "+" (b.fn.extra) ---------- */
  const XSEC = [['event', 'ข้อมูลงาน'], ['banquet', 'แผนกจัดเลี้ยง / ห้องอาหาร'], ['hk', 'แผนกแม่บ้าน'], ['av', 'แผนกช่างเครื่องเสียง'], ['acc', 'แผนกบัญชี / การเงิน'], ['front', 'แผนกต้อนรับ'], ['other', 'รายละเอียดเพิ่มเติม']];
  const xName = k => (XSEC.find(x => x[0] === k) || XSEC[XSEC.length - 1])[1];

  /* ---------- the paper form ----------
   * ed = the system values while editing: every value cell becomes a box you type into, right on the sheet */
  function paper(v, extra = [], ed = null) {
    // lines added inside a cell (heading or detail) + the cell's corner "+" that adds them
    const more = (key, list = []) => ed
      ? list.map((it, j) => `<div class="fnf-ed fnf-more${it.t === 'h' ? ' fnf-hbar' : ''}"><button class="fnf-rv red" data-act="fnDCDel" data-key="${key}" data-j="${j}" title="ลบบรรทัดนี้">✕</button><textarea class="fnf-in${it.t === 'h' ? ' b' : ''}" rows="1" data-input="fnDC" data-key="${key}" data-j="${j}" placeholder="${it.t === 'h' ? 'หัวข้อ' : 'รายละเอียดข้อมูล'}">${H(it.text || '')}</textarea></div>`).join('')
        + `<button class="fnf-cplus" data-act="fnCPlus" data-key="${key}" title="เพิ่มหัวข้อ หรือ รายละเอียดข้อมูลในช่องนี้">＋</button>`
      : list.filter(it => (it.text || '').trim()).map(it => it.t === 'h' ? `<div class="fnf-mh2">${H(it.text)}</div>` : `<div class="fnf-md">${nl(it.text)}</div>`).join('');
    const mv = v.more || {};
    const T = (k, dash = '', ph = '', plus = true) => {
      if (!ed) return nl(v[k] || dash) + (plus ? more(k, mv[k]) : '');
      const val = v[k] == null ? '' : String(v[k]), rv = ed[k] != null && val !== String(ed[k]);
      return `<div class="fnf-ed${plus ? ' fnf-cell' : ''}">${rv ? `<button class="fnf-rv" data-act="fnDRv" data-k="${k}" title="ใช้ค่าจากระบบ">↺</button>` : ''}<textarea class="fnf-in" rows="1" data-input="fnD" data-k="${k}" placeholder="${H(ph)}" aria-label="${H(ph || k)}">${H(val)}</textarea>${plus ? more(k, mv[k]) : ''}</div>`;
    };
    const cell = (k, cls = '', dash = '', ph = '') => `<td class="${cls}">${T(k, dash, ph)}</td>`;
    const box = (attr, val, ph, cls = '') => `<textarea class="fnf-in ${cls}" rows="1" ${attr} placeholder="${H(ph)}">${H(val || '')}</textarea>`;
    const xs = k => extra.map((x, i) => Object.assign({ i }, x)).filter(x => (XSEC.some(s => s[0] === x.sec) ? x.sec : 'other') === k);
    const xdel = x => `<button class="fnf-rv red" data-act="fnDXDel" data-i="${x.i}" title="ลบรายการนี้">✕</button>`;
    const xth = x => ed ? `<th><div class="fnf-ed">${xdel(x)}${box(`data-input="fnDX" data-i="${x.i}" data-k="label"`, x.label, 'หัวข้อ')}</div></th>`
      : `<th>${H(x.label || '')}${x.label ? ' :' : ''}</th>`;
    const xtd = (x, span = 1) => `<td${span > 1 ? ` colspan="${span}"` : ''}>${ed ? `<div class="fnf-ed fnf-cell">${box(`data-input="fnDX" data-i="${x.i}" data-k="text"`, x.text, 'รายละเอียด')}${more('x.' + x.i, x.more)}</div>` : nl(x.text) + more('x.' + x.i, x.more)}</td>`;
    const xIn = k => xs(k).map(x => ed ? `<div class="fnf-x fnf-ed">${xdel(x)}${box(`data-input="fnDX" data-i="${x.i}" data-k="label"`, x.label, 'หัวข้อ', 'b')}${box(`data-input="fnDX" data-i="${x.i}" data-k="text"`, x.text, 'รายละเอียด')}</div>`
      : `<div class="fnf-x">${x.label ? `<b>${H(x.label)} :</b> ` : ''}${nl(x.text)}</div>`).join('');
    const acc = xs('acc'), front = xs('front'), mc = Math.max(1, v.menu.length);
    // "+" inside each section while editing: adds a heading + detail row shaped like the rows above it
    const add = sec => ed ? `<button class="fnf-add" data-act="fnDXAdd" data-sec="${sec}" title="เพิ่มหัวข้อ/ช่องข้อมูลในส่วนนี้">＋ เพิ่มช่อง</button>` : '';
    const addRow = (sec, span) => ed ? `<tr class="fnf-addrow"><td colspan="${span}">${add(sec)}</td></tr>` : '';
    const menuCols = v.menu.length ? v.menu.map((m, i) => ed
      ? `<td class="fnf-menu"><div class="fnf-ed"><button class="fnf-rv red" data-act="fnDMDel" data-i="${i}" title="ลบช่วงนี้">✕</button>${box(`data-input="fnDM" data-i="${i}" data-k="time"`, m.time, 'ช่วงเวลา', 'b')}</div>${box(`data-input="fnDM" data-i="${i}" data-k="title"`, m.title, 'หัวข้อ', 'b')}${box(`data-input="fnDM" data-i="${i}" data-k="items"`, m.items, 'รายการอาหาร (บรรทัดละรายการ)')}<div class="fnf-ed fnf-cell">${more('menu.' + i, m.more)}</div></td>`
      : `<td class="fnf-menu"><div class="fnf-mt">${H(m.time)}</div><div class="fnf-mh">${H(m.title)}</div>${nl(m.items)}${more('menu.' + i, m.more)}</td>`).join('') : '<td class="fnf-menu">-</td>';
    const kv = (k, ph) => ed ? `<span class="fnf-kvin">${T(k, '', ph, false)}</span>` : H(v[k]);
    return `<div class="fnf-paper${ed ? ' editing' : ''}" id="fnPaper">
  <div class="fnf-head">
    <img src="img/logo-wide.png" alt="Luck Swan Resort">
    <div class="fnf-title"><b>ลักษวรรณ รีสอร์ท</b><br>Banquet Function Event</div>
    <table class="fnf-meta"><tr><th>Function No.</th><td class="fnf-no">${ed ? T('no', '', 'RS…/', false) + (!v.no ? `<button class="link xs" data-act="fnNo">ออกเลขอัตโนมัติ (${H(nextNo())})</button>` : '') : H(v.no || '—')}</td></tr><tr><th>วันที่ออกเอกสาร</th><td>${kv('issued', 'วันที่')}</td></tr><tr><th>Quotation No. :</th><td>${kv('quote', 'เลขใบเสนอราคา')}</td></tr></table>
  </div>
  <div class="fnf-org"><span><b>สำนักงาน/องค์กร :</b> ${kv('org', 'ชื่อสำนักงาน / องค์กร')}</span><span><b>ผู้ประสานงาน :</b> ${kv('coord', 'ชื่อ / เบอร์โทร')}</span></div>
  <table class="fnf-t">
    <tr><th>วันที่จัดงาน :</th>${cell('date', 'c')}<th>สถานที่จัดงาน :</th>${cell('venue', 'c')}<th>ประเภทงาน :</th>${cell('type', 'c')}</tr>
    <tr><th>จำนวนการันตี :</th>${cell('pax', 'c')}<td colspan="4" class="c">${T('pricing', '', 'ราคาต่อหน่วย (บรรทัดละรายการ)')}</td></tr>
    ${xs('event').map(x => `<tr>${xth(x)}${xtd(x, 5)}</tr>`).join('')}
    ${addRow('event', 6)}
  </table>
  <table class="fnf-t">
    <tr><th colspan="2" class="sec">แผนกจัดเลี้ยง / ห้องอาหาร</th><th colspan="${mc}" class="sec">เมนูอาหาร${ed ? ' <button class="link xs" data-act="fnDMAdd">+ เพิ่มช่วงเวลา</button>' : ''}</th></tr>
    <tr><th>รูปแบบการจัดห้อง :</th><td>${T('layout', '', 'รูปแบบการจัดห้อง')}${v.fyi || ed ? `<div class="fnf-fyi"><b>FYI</b><br>${T('fyi', '', 'กำหนดการ / หมายเหตุ')}</div>` : ''}</td>${menuCols}</tr>
    <tr><th>ข้อความบนเวที :${v.vinyl || ed ? `<div class="fnf-small">${T('vinyl', '', 'ไวนิล / ป้าย')}</div>` : ''}</th><td class="c">${T('stage', '-', 'ข้อความบนเวที')}</td><td colspan="${mc}" class="fnf-small">${v.kitchen || ed ? '<b>ครัว:</b> ' + T('kitchen', '', 'หมายเหตุครัว') : ''}</td></tr>
    ${xs('banquet').map(x => `<tr>${xth(x)}${xtd(x, 1 + mc)}</tr>`).join('')}
    ${addRow('banquet', 2 + mc)}
  </table>
  <table class="fnf-t">
    <tr><th class="sec" style="width:50%">แผนกแม่บ้าน</th><th class="sec">แผนกช่างเครื่องเสียง</th></tr>
    <tr><td class="c">${T('hk', '-', 'แผนกแม่บ้าน')}${xIn('hk')}${add('hk')}</td><td>${T('av', '-', 'แผนกช่างเครื่องเสียง')}${xIn('av')}${add('av')}</td></tr>
  </table>
  <table class="fnf-t">
    <tr><th colspan="2" class="sec" style="width:50%">แผนกบัญชี / การเงิน</th><th colspan="4" class="sec">แผนกต้อนรับ</th></tr>
    <tr><th>ยอดรวม :</th><td class="c fnf-total">${ed ? T('total') : H(v.total)}</td><th>Check-in :</th>${cell('checkin', 'c')}<th>Check-out :</th>${cell('checkout', 'c')}</tr>
    <tr><th>รายละเอียดค่าใช้จ่าย :</th>${cell('cost', 'c')}<th>ห้องพัก (ประเภท) :</th>${cell('roomType', 'c')}<th>จำนวนห้อง :</th>${cell('roomQty', 'c')}</tr>
    <tr><th>ที่อยู่ออกบิล :</th>${cell('bill', 'c')}<td colspan="4" class="c">${T('front', '', 'หมายเหตุต้อนรับ')}</td></tr>
    ${Array.from({ length: Math.max(acc.length, front.length) }, (_, i) => `<tr>${acc[i] ? xth(acc[i]) + xtd(acc[i]) : '<td colspan="2"></td>'}${front[i] ? xth(front[i]) + xtd(front[i], 3) : '<td colspan="4"></td>'}</tr>`).join('')}
    ${ed ? `<tr class="fnf-addrow"><td colspan="2">${add('acc')}</td><td colspan="4">${add('front')}</td></tr>` : ''}
  </table>
  ${xs('other').length || ed ? `<table class="fnf-t"><tr><th colspan="2" class="sec">รายละเอียดเพิ่มเติม</th></tr>${xs('other').map(x => `<tr>${xth(x)}${xtd(x)}</tr>`).join('')}${addRow('other', 2)}</table>` : ''}
  <div class="fnf-sign">
    <div><span>ออกโดย :</span><div class="fnf-line"></div><div>${ed ? T('by', '', 'ชื่อผู้ออกเอกสาร') : `(${H(v.by || '................................')})`}</div><div>${ed ? T('byTitle', '', 'ตำแหน่ง') : H(v.byTitle)}</div></div>
    <div><span>อนุมัติโดย :</span><div class="fnf-line"></div><div>(${H(v.appr || '................................')})</div><div>${H(v.apprTitle || 'ยังไม่อนุมัติ')}</div></div>
  </div>
</div>`;
  }

  /* ---------- the page ---------- */
  const EDIT_KEYS = ['no', 'issued', 'quote', 'org', 'coord', 'date', 'venue', 'type', 'pax', 'pricing', 'layout', 'fyi', 'stage', 'vinyl', 'kitchen', 'hk', 'av', 'total', 'cost', 'bill', 'checkin', 'checkout', 'roomType', 'roomQty', 'front', 'by', 'byTitle'];
  const clone = o => JSON.parse(JSON.stringify(o));
  const fnUi = () => App.state.ui.fn || (App.state.ui.fn = { tab: 'preview' });
  // the unsaved copy being edited on the sheet; nothing is stored until "บันทึกแก้ไขข้อมูล"
  function draft(b) {
    const ui = fnUi();
    if (!ui.draft || ui.draft.bid !== b.id) {
      const { v } = values(b, Store.quote(b.quoteId)), d = { bid: b.id, v: {}, menu: clone(v.menu), extra: clone((b.fn && b.fn.extra) || []), more: clone(v.more) };
      for (const k of EDIT_KEYS) d.v[k] = v[k] == null ? '' : String(v[k]);
      ui.draft = d; ui.dirty = false;
    }
    return ui.draft;
  }
  const dropDraft = () => { const ui = fnUi(); ui.draft = null; ui.dirty = false; };
  const markDirty = () => { fnUi().dirty = true; const s = document.getElementById('fnDirty'); if (s) { s.textContent = 'มีการแก้ไขที่ยังไม่บันทึก'; s.className = 'small red'; } };
  App.V.fn = {
    title: 'รายละเอียดฟังก์ชัน', nav: 'beo',
    render() {
      const b = curB(), q = b && Store.quote(b.quoteId);
      if (!b || !q) return '<div class="empty">ไม่พบใบฟังก์ชัน — <a href="#/beo">กลับ</a></div>';
      const { v, a } = values(b, q), edit = canEdit(), ui = fnUi();
      const tab = edit ? ui.tab : 'preview', own = b.fn || {};
      const head = `<div class="page-hd no-print"><div><div class="small"><a class="link" href="#/beo/${encodeURIComponent(b.id)}">← ใบฟังก์ชัน ${H(b.id)}</a></div><h1>รายละเอียดฟังก์ชัน</h1>
        <div class="small muted">${H(q.title)} · ${H(q.client)} · ${H(v.no || 'ยังไม่มี Function No.')}${b.signoff ? ' · อนุมัติแล้ว' : ''}</div></div>
        <div class="row wrap">${edit ? `<div class="seg white"><button class="${tab === 'preview' ? 'on' : ''}" data-act="fnTab" data-v="preview">ตัวอย่างพิมพ์</button><button class="${tab === 'form' ? 'on' : ''}" data-act="fnTab" data-v="form">แก้ไขข้อมูล</button></div>` : ''}
        ${edit ? '<button class="btn ghost fnf-plus" data-act="fnX" title="เพิ่มรายละเอียดลงในใบรายละเอียดฟังก์ชัน">＋ เพิ่มรายละเอียด</button>' : ''}
        <button class="btn red" data-act="fnPrint">พิมพ์ / PDF</button></div></div>`;
      if (tab === 'preview') return head + `${!v.no && edit ? '<div class="note no-print" style="margin-bottom:12px">ยังไม่มี Function No. — <button class="link" data-act="fnNo">ออกเลขอัตโนมัติ (' + H(nextNo()) + ')</button> หรือกรอกเองในแท็บ “แก้ไขข้อมูล”</div>' : ''}<div class="fnf-wrap">${paper(v, own.extra || [])}</div>`;
      const d = draft(b), dv = Object.assign({}, d.v, { menu: d.menu, more: d.more, appr: v.appr, apprTitle: v.apprTitle });
      return head + `<div class="fnf-bar no-print"><div><b>แก้ไขข้อมูลบนใบ</b> — พิมพ์แก้ในช่องได้เลย ปุ่ม ↺ = ใช้ค่าจากระบบ<div id="fnDirty" class="${ui.dirty ? 'small red' : 'small muted'}">${ui.dirty ? 'มีการแก้ไขที่ยังไม่บันทึก' : 'ยังไม่มีการแก้ไข'}</div></div>
        <div class="row"><button class="btn ghost" data-act="fnDCancel">ยกเลิก</button><button class="btn red" data-act="fnDSave">บันทึกแก้ไขข้อมูล</button></div></div>
        <div class="fnf-wrap">${paper(dv, d.extra, a)}</div>`;
    },
    after() {
      document.querySelectorAll('.fnf-in').forEach(fit);
      const ui = fnUi();
      if (ui.focusC) { const t = document.querySelector(`[data-input="fnDC"][data-key="${ui.focusC.key}"][data-j="${ui.focusC.j}"]`); ui.focusC = null; if (t) t.focus(); }
      if (ui.focusX != null) { const t = document.querySelector(`[data-input="fnDX"][data-i="${ui.focusX}"][data-k="label"]`); ui.focusX = null; if (t) t.focus(); }
    }
  };
  function fit(t) { t.style.height = 'auto'; t.style.height = t.scrollHeight + 2 + 'px'; }

  function setFn(fn) {
    const b = curB(); if (!b || !canEdit()) return;
    b.fn = Object.assign({}, b.fn); fn(b.fn, b);
    for (const k of Object.keys(b.fn)) if (b.fn[k] === undefined) delete b.fn[k];
    Store.save(); App.render();
  }
  Object.assign(App.ACT, {
    fnTab(el) {
      const ui = fnUi(), go = () => { if (el.dataset.v === 'preview') dropDraft(); ui.tab = el.dataset.v; App.render(); };
      if (ui.tab === 'form' && el.dataset.v === 'preview' && ui.dirty) App.confirm('ทิ้งการแก้ไขที่ยังไม่บันทึก?', go, 'ทิ้งการแก้ไข'); else go();
    },
    fnDCancel() { App.ACT.fnTab({ dataset: { v: 'preview' } }); },
    fnDSave() {
      const b = curB(), d = fnUi().draft; if (!b || !d || !canEdit()) return;
      const a = auto(b, Store.quote(b.quoteId));
      setFn(f => {
        for (const k of EDIT_KEYS) { const val = d.v[k]; if (val === (a[k] == null ? '' : String(a[k]))) delete f[k]; else f[k] = val.slice(0, 4000); }
        const keep = l => (l || []).filter(it => (it.text || '').trim());
        const menu = d.menu.map(m => { const o = Object.assign({}, m), l = keep(m.more); if (l.length) o.more = l; else delete o.more; return o; });
        f.menu = JSON.stringify(menu) === JSON.stringify(a.menu) ? undefined : menu;
        const mo = {}; for (const k of Object.keys(d.more)) { const l = keep(d.more[k]); if (l.length) mo[k] = l; }
        f.more = Object.keys(mo).length ? mo : undefined;
        const ex = d.extra.filter(x => (x.label || '').trim() || (x.text || '').trim() || keep(x.more).length).map(x => { const o = Object.assign({}, x), l = keep(x.more); if (l.length) o.more = l; else delete o.more; return o; });
        f.extra = ex.length ? ex : undefined;
      });
      dropDraft(); fnUi().tab = 'preview'; App.render(); App.toast('บันทึกแก้ไขข้อมูลแล้ว');
    },
    fnDRv(el) { const b = curB(), d = fnUi().draft; if (!b || !d) return; const a = auto(b, Store.quote(b.quoteId)); d.v[el.dataset.k] = a[el.dataset.k] == null ? '' : String(a[el.dataset.k]); fnUi().dirty = true; App.render(); },
    fnDMAdd() { const d = fnUi().draft; if (!d) return; d.menu.push({ time: '', title: '', items: '' }); fnUi().dirty = true; App.render(); },
    fnDMDel(el) { const d = fnUi().draft; if (!d) return; d.menu.splice(Number(el.dataset.i), 1); fnUi().dirty = true; App.render(); },
    fnDXAdd(el) { const d = fnUi().draft; if (!d) return; d.extra.push({ sec: el.dataset.sec, label: '', text: '' }); fnUi().dirty = true; fnUi().focusX = d.extra.length - 1; App.render(); },
    // a cell's corner "+": choose heading or detail, then a new line appears inside that cell
    fnCPlus(el) {
      document.querySelectorAll('.fnf-pick').forEach(x => x.remove());
      const pk = document.createElement('div'); pk.className = 'fnf-pick';
      pk.innerHTML = `<div class="xs muted">เพิ่มในช่องนี้</div><button data-act="fnCAdd" data-key="${H(el.dataset.key)}" data-t="h"><b>หัวข้อ</b></button><button data-act="fnCAdd" data-key="${H(el.dataset.key)}" data-t="d">รายละเอียดข้อมูล</button>`;
      el.parentNode.appendChild(pk);
      setTimeout(() => document.addEventListener('mousedown', function off(e) { if (!pk.contains(e.target)) { pk.remove(); document.removeEventListener('mousedown', off); } }));
    },
    fnCAdd(el) { const l = cellList(el.dataset.key, true); if (!l) return; l.push({ t: el.dataset.t, text: '' }); fnUi().dirty = true; fnUi().focusC = { key: el.dataset.key, j: l.length - 1 }; App.render(); },
    fnDCDel(el) { const l = cellList(el.dataset.key); if (!l) return; l.splice(Number(el.dataset.j), 1); fnUi().dirty = true; App.render(); },
    fnDXDel(el) { const d = fnUi().draft; if (!d) return; d.extra.splice(Number(el.dataset.i), 1); fnUi().dirty = true; App.render(); },
    fnPrint() { const ui = fnUi(); if (ui.tab === 'form' && ui.dirty) { App.toast('กด “บันทึกแก้ไขข้อมูล” ก่อนพิมพ์'); return; } if (ui.tab !== 'preview') { dropDraft(); ui.tab = 'preview'; App.render(); } setTimeout(() => window.print(), 50); },
    fnNo() { const d = fnUi().tab === 'form' && fnUi().draft; if (d) { d.v.no = nextNo(); fnUi().dirty = true; App.render(); return; } setFn(f => { f.no = nextNo(); }); App.toast('ออกเลข ' + curB().fn.no); },
    // "+": a small form laid out like the paper sheet — pick the department, give a heading and the details
    fnX(el) {
      const b = curB(); if (!b || !canEdit()) return;
      const i = el.dataset.i != null ? Number(el.dataset.i) : -1, x = i >= 0 ? ((b.fn || {}).extra || [])[i] : { sec: 'banquet', label: '', text: '' };
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
          const item = { sec: sec.value, label, text }, d = fnUi().tab === 'form' && fnUi().draft;
          if (d) { if (i >= 0) d.extra[i] = item; else d.extra.push(item); fnUi().dirty = true; App.render(); return; }
          setFn(f => { const list = (f.extra || []).map(o => Object.assign({}, o)); if (i >= 0) list[i] = item; else list.push(item); f.extra = list; });
          App.toast(i >= 0 ? 'บันทึกแล้ว' : 'เพิ่มลงใบรายละเอียดฟังก์ชันแล้ว');
        };
      });
    }
  });
  App.fnNo = b => (b.fn && b.fn.no) || '';
  // the list of added lines for a cell key: a field name, "menu.<i>" or "x.<i>" (an added row)
  function cellList(key, make) {
    const d = fnUi().draft; if (!d) return null;
    const m = /^(menu|x)\.(\d+)$/.exec(key), o = m ? (m[1] === 'menu' ? d.menu : d.extra)[Number(m[2])] : d.more;
    if (!o) return null;
    const k = m ? 'more' : key;
    if (!Array.isArray(o[k])) { if (!make) return null; o[k] = []; }
    return o[k];
  }
  Object.assign(App.INP, {
    fnD(el) { const d = fnUi().draft; if (!d) return; d.v[el.dataset.k] = el.value.slice(0, 4000); fit(el); markDirty(); },
    fnDM(el) { const d = fnUi().draft, m = d && d.menu[Number(el.dataset.i)]; if (!m) return; m[el.dataset.k] = el.value.slice(0, 2000); fit(el); markDirty(); },
    fnDC(el) { const l = cellList(el.dataset.key), it = l && l[Number(el.dataset.j)]; if (!it) return; it.text = el.value.slice(0, 2000); fit(el); markDirty(); },
    fnDX(el) { const d = fnUi().draft, x = d && d.extra[Number(el.dataset.i)]; if (!x) return; x[el.dataset.k] = el.value.slice(0, el.dataset.k === 'label' ? 80 : 2000); fit(el); markDirty(); }
  });
})(window);
