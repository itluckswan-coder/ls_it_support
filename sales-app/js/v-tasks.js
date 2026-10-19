/* Luck Swan Sales — Assignment & Approve: admins assign work to users or departments and sign off
 * confirmed function sheets; everyone else sees and confirms the work assigned to them. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  // a user's department, from the free-text dept on their account (mirrors DEPT_WORDS in worker/src/index.js)
  const DEPT_WORDS = { kitchen: ['ครัว', 'f&b'], banquet: ['จัดเลี้ยง', 'banquet'], av: ['ช่าง', 'av'], front: ['ส่วนหน้า', 'front'], hk: ['แม่บ้าน', 'housekeeping'], acc: ['บัญชี', 'account'] };
  const deptOf = u => { const d = String((u && u.dept) || '').toLowerCase(); return Object.keys(DEPT_WORDS).find(k => DEPT_WORDS[k].some(w => d.includes(w))) || null; };
  const dept = id => D.DEPTS.find(d => d.id === id) || { name: id, short: id, color: '#999', bg: '#eee', fg: '#333' };
  const isAdmin = () => { const me = App.me(); return !!me && me.role === 'admin'; };
  const tasks = () => Store.db.tasks || (Store.db.tasks = []);
  const mine = (t, me) => !!me && !!t.to && (t.to.type === 'user' ? t.to.id === me.id : t.to.type === 'dept' && t.to.dept === deptOf(me));
  const toText = t => t.to.type === 'user' ? t.to.name || ((Store.db.users.find(u => u.id === t.to.id) || {}).name) || '—' : 'แผนก' + dept(t.to.dept).short;
  const when = iso => iso ? U.fmtD(iso.slice(0, 10), { year: false }) + ' ' + iso.slice(11, 16) : '—';
  const empName = id => (Store.db.employees.find(e => e.id === id) || {}).name;
  // function sheets waiting for the admin's sign-off: the quote is confirmed and no edit is waiting for approval
  const toApprove = () => Store.db.beos.filter(b => { const q = Store.quote(b.quoteId); return q && q.status === 'confirmed' && !b.signoff && g.Approval && g.Approval.state(b) !== 'pending'; });
  const ICON_BELL = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 8a6 6 0 1 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.9 1.9 0 0 0 3.4 0"/></svg>';

  /* ---------------- notifications (bell) ---------------- */
  const SEEN = () => 'lsSales.notifSeen.' + App.state.uid;
  const seenAt = () => { try { return localStorage.getItem(SEEN()) || ''; } catch (e) { return ''; } };
  App.notifications = function () {
    const me = App.me(); if (!me) return [];
    const out = [], admin = me.role === 'admin';
    for (const t of tasks()) {
      if (t.status !== 'done' && mine(t, me)) out.push({ at: t.at, title: 'งานใหม่ที่ได้รับมอบหมาย', sub: t.title + (t.due ? ' · กำหนด ' + U.fmtD(t.due, { year: false }) : ''), href: '#/tasks', tone: t.priority === 'high' ? 'red' : 'teal' });
      if (admin && t.status === 'done') out.push({ at: t.doneAt, title: (t.doneByName || 'ผู้รับงาน') + ' ยืนยันงานแล้ว', sub: t.title, href: '#/tasks', tone: 'teal' });
    }
    if (admin) for (const b of toApprove()) { const q = Store.quote(b.quoteId); out.push({ at: q.updatedAt || '', title: 'ใบฟังก์ชันรออนุมัติ', sub: b.id + ' · ' + q.client, href: '#/tasks/approve', tone: 'yellow' }); }
    for (const b of Store.db.beos) {
      const a = b.approval || {};
      if (g.Approval && g.Approval.canDecide(b, me)) out.push({ at: a.requestedAt, title: 'การแก้ไขใบฟังก์ชันรออนุมัติ', sub: b.id + ' · แก้ไขโดย ' + (a.requestedByName || ''), href: '#/beo/' + encodeURIComponent(b.id), tone: 'yellow' });
      if (a.state === 'rejected' && a.requestedBy === me.id) out.push({ at: a.at, title: 'ใบฟังก์ชันถูกตีกลับ', sub: b.id + (a.comment ? ' · ' + a.comment : ''), href: '#/beo/' + encodeURIComponent(b.id), tone: 'red' });
      if (b.signoff && (Store.quote(b.quoteId) || {}).owner === me.id) out.push({ at: b.signoff.at, title: 'ใบฟังก์ชันได้รับอนุมัติ', sub: b.id + ' · โดย ' + b.signoff.byName, href: '#/beo/' + encodeURIComponent(b.id), tone: 'teal' });
    }
    const seen = seenAt();
    return out.map(n => Object.assign(n, { unread: (n.at || '') > seen })).sort((a, b) => (b.at || '').localeCompare(a.at || '')).slice(0, 30);
  };
  App.bellButton = function () {
    const n = App.notifications().filter(x => x.unread).length;
    return `<button class="bell-btn" data-act="notifs" aria-label="การแจ้งเตือน${n ? ' ' + n + ' รายการใหม่' : ''}">${ICON_BELL}${n ? `<span class="bell-n">${n > 9 ? '9+' : n}</span>` : ''}</button>`;
  };
  App.ACT.notifs = function () {
    const list = App.notifications(), tone = { red: 'var(--red)', yellow: 'var(--yellow)', teal: 'var(--teal)' };
    try { localStorage.setItem(SEEN(), U.nowTH()); } catch (e) { /* ignore */ }
    App.modal(`<div class="between" style="margin-bottom:8px"><h2 style="margin:0">การแจ้งเตือน</h2><a class="link" href="#/tasks" data-act="closeModal">ดูงานทั้งหมด</a></div>
      ${list.length ? `<div class="notif-list">${list.map(n => `<a class="notif ${n.unread ? 'new' : ''}" href="${n.href}" data-act="closeModal"><i style="background:${tone[n.tone] || 'var(--muted2)'}"></i><span><b>${H(n.title)}</b><span class="small muted">${H(n.sub)}</span></span><span class="xs muted" style="white-space:nowrap">${H(when(n.at))}</span></a>`).join('')}</div>` : '<div class="empty">ยังไม่มีการแจ้งเตือน</div>'}
      <div class="acts"><button class="btn ghost" data-act="closeModal">ปิด</button></div>`, () => { const b = document.querySelector('.bell-btn .bell-n'); if (b) b.remove(); });
  };

  /* ---------------- home page card ----------------
   * everyone but the admin: "งานที่ได้รับมอบหมายวันนี้" — work still open, plus what they confirmed today;
   * the admin: "งานที่มอบหมาย" — what was assigned, its status, and when it was confirmed */
  const badge = t => { const late = t.status !== 'done' && t.due && t.due < U.today(); return t.status === 'done' ? '<span class="badge st-confirmed">ยืนยันแล้ว</span>' : `<span class="badge ${late ? 'st-tentative' : 'st-sent'}">${late ? 'เลยกำหนด' : 'รอดำเนินการ'}</span>`; };
  App.homeTasks = function () {
    const me = App.me(); if (!me) return '';
    const today = U.today(), admin = me.role === 'admin';
    const row = (t, sub) => `<button class="hm-task" data-act="tkView" data-id="${H(t.id)}"><span style="min-width:0;flex:1;text-align:left"><b>${t.priority === 'high' ? '<span class="badge st-tentative" style="margin-right:4px">ด่วน</span>' : ''}${H(t.title)}</b><span class="small muted">${sub}</span></span>${badge(t)}</button>`;
    if (admin) {
      const list = tasks().slice().sort((a, b) => ((b.status === 'done' ? b.doneAt : b.at) || '').localeCompare((a.status === 'done' ? a.doneAt : a.at) || ''));
      const open = list.filter(t => t.status !== 'done').length;
      return `<section class="card pad"><div class="between" style="margin-bottom:10px"><h2>งานที่มอบหมาย</h2><a class="link small" href="#/tasks">ดูทั้งหมด</a></div>
        <div class="small muted" style="margin:-4px 0 8px">รอดำเนินการ ${open} · ยืนยันแล้ว ${list.length - open}</div>
        ${list.slice(0, 6).map(t => row(t, `ถึง${H(toText(t))} · ${t.status === 'done' ? 'ยืนยันโดย ' + H(t.doneByName || '') + ' ' + H(when(t.doneAt)) : 'มอบหมาย ' + H(when(t.at))}`)).join('') || '<div class="muted small">ยังไม่มีงานที่มอบหมาย</div>'}</section>`;
    }
    const list = tasks().filter(t => mine(t, me) && (t.status !== 'done' || (t.doneAt || '').slice(0, 10) === today))
      .sort((a, b) => (a.status === 'done') - (b.status === 'done') || (a.due || '9').localeCompare(b.due || '9'));
    return `<section class="card pad"><div class="between" style="margin-bottom:10px"><h2>งานที่ได้รับมอบหมายวันนี้</h2><a class="link small" href="#/tasks">ดูทั้งหมด</a></div>
      ${list.map(t => row(t, (t.due ? 'กำหนด ' + U.fmtD(t.due, { year: false }) + ' · ' : '') + 'จาก ' + H(t.byName || ''))).join('') || '<div class="muted small">ไม่มีงานที่ได้รับมอบหมาย 🎉</div>'}</section>`;
  };
  // one assignment in full
  function taskView(t) {
    const me = App.me(), canDo = t.status !== 'done' && mine(t, me), q = t.quoteId && Store.quote(t.quoteId);
    const kv = (k, v) => v ? `<div class="tk-kv"><span class="muted">${k}</span><span>${v}</span></div>` : '';
    App.modal(`<div class="between" style="align-items:flex-start;gap:10px"><h2 style="margin:0">${H(t.title)}</h2>${badge(t)}</div>
      ${t.priority === 'high' ? '<div style="margin-top:6px"><span class="badge st-tentative">ด่วน</span></div>' : ''}
      ${t.detail ? `<p style="white-space:pre-line;margin:12px 0 4px">${H(t.detail)}</p>` : '<p class="muted small" style="margin:12px 0 4px">ไม่มีรายละเอียดเพิ่มเติม</p>'}
      <div class="tk-kvs">
        ${kv('มอบหมายให้', H(toText(t)))}
        ${t.to.type === 'dept' && (t.to.emps || []).length ? kv('พนักงาน', H(t.to.emps.map(empName).filter(Boolean).join(', '))) : ''}
        ${kv('มอบหมายโดย', H(t.byName || '—'))}
        ${kv('มอบหมายเมื่อ', H(when(t.at)))}
        ${kv('กำหนดเสร็จ', t.due ? H(U.fmtD(t.due)) : '—')}
        ${q ? kv('ใบเสนอราคา', `<a href="#/quote/${encodeURIComponent(q.id)}" data-act="closeModal">${H(q.id)} · ${H(q.client)}</a>`) : ''}
        ${t.status === 'done' ? kv('ยืนยันโดย', H(t.doneByName || '—')) + kv('ยืนยันเมื่อ', H(when(t.doneAt))) + (t.doneNote ? kv('หมายเหตุ', '“' + H(t.doneNote) + '”') : '') : ''}
      </div>
      <div class="acts"><a class="btn ghost" href="#/tasks" data-act="closeModal">ไปหน้างานทั้งหมด</a>${canDo ? `<button class="btn teal" data-act="tkDone" data-id="${H(t.id)}">ยืนยันงานนี้</button>` : '<button class="btn dark" data-act="closeModal">ปิด</button>'}</div>`);
  }

  /* ---------------- the page ---------------- */
  App.V.tasks = {
    title: 'Assignment & Approve',
    render(p) {
      const admin = isAdmin(), me = App.me();
      const tab = admin && p[0] === 'approve' ? 'approve' : 'assign';
      const head = `<div class="page-hd"><div><h1>Assignment &amp; Approve</h1><div class="small muted">${admin ? 'มอบหมายงานให้ผู้ใช้งานหรือแผนก และอนุมัติใบฟังก์ชันที่ยืนยันแล้ว' : 'งานที่ได้รับมอบหมายให้คุณ' + (deptOf(me) ? ' และแผนก' + H(dept(deptOf(me)).short) : '')}</div></div>
        ${admin && tab === 'assign' ? '<button class="btn red" data-act="tkNew">+ มอบหมายงาน</button>' : ''}</div>
        ${admin ? `<div class="seg white" style="margin-bottom:16px" role="tablist"><a role="tab" href="#/tasks" class="${tab === 'assign' ? 'on' : ''}">Assignment</a><a role="tab" href="#/tasks/approve" class="${tab === 'approve' ? 'on' : ''}">Approve${toApprove().length ? ` <span class="seg-n">${toApprove().length}</span>` : ''}</a></div>` : ''}`;
      return head + (tab === 'approve' ? this.approve() : admin ? this.assignAdmin() : this.assignMine(me));
    },

    assignAdmin() {
      const ui = App.state.ui.tasks || (App.state.ui.tasks = { st: 'all' });
      const all = tasks().slice().sort((a, b) => (b.at || '').localeCompare(a.at || ''));
      const list = all.filter(t => ui.st === 'all' || (ui.st === 'done' ? t.status === 'done' : t.status !== 'done'));
      const cnt = s => all.filter(t => s === 'all' || (s === 'done' ? t.status === 'done' : t.status !== 'done')).length;
      return `
<div class="row wrap" style="margin-bottom:14px">${[['all', 'ทั้งหมด'], ['open', 'รอดำเนินการ'], ['done', 'ยืนยันแล้ว']].map(([k, l]) => `<button class="chip ${ui.st === k ? 'on' : ''}" data-act="tkSt" data-v="${k}">${l} <span class="n">${cnt(k)}</span></button>`).join('')}</div>
<section class="card" style="overflow:hidden">
  <table class="tbl resp"><thead><tr><th>งาน</th><th>มอบหมายให้</th><th>มอบหมายเมื่อ</th><th>กำหนดเสร็จ</th><th>สถานะ</th><th class="r">จัดการ</th></tr></thead><tbody>
  ${list.map(t => { const late = t.status !== 'done' && t.due && t.due < U.today(); return `<tr>
    <td><div style="font-weight:600">${t.priority === 'high' ? '<span class="badge st-tentative" style="margin-right:6px">ด่วน</span>' : ''}${H(t.title)}</div><div class="sub">${H(t.detail || '')}${t.quoteId ? ` · <a href="#/quote/${encodeURIComponent(t.quoteId)}">${H(t.quoteId)}</a>` : ''}</div>
      <div class="sub show-m">${H(toText(t))} · ${H(when(t.at))}</div></td>
    <td class="hide-m"><div>${H(toText(t))}</div>${t.to.type === 'dept' && (t.to.emps || []).length ? `<div class="xs muted">${H(t.to.emps.map(empName).filter(Boolean).join(', '))}</div>` : ''}</td>
    <td class="hide-m small">${H(when(t.at))}<div class="xs muted">โดย ${H(t.byName || '')}</div></td>
    <td class="hide-m small ${late ? 'red' : ''}">${t.due ? U.fmtD(t.due, { year: false }) + (late ? ' · เลยกำหนด' : '') : '—'}</td>
    <td>${t.status === 'done' ? `<span class="badge st-confirmed">ยืนยันแล้ว</span><div class="xs muted" style="margin-top:4px">${H(t.doneByName || '')} · ${H(when(t.doneAt))}</div>${t.doneNote ? `<div class="xs muted">“${H(t.doneNote)}”</div>` : ''}` : '<span class="badge st-sent">รอดำเนินการ</span>'}</td>
    <td class="r"><div class="row" style="gap:6px;justify-content:flex-end;flex-wrap:nowrap"><button class="icon-btn us-edit" data-act="tkEdit" data-id="${H(t.id)}" aria-label="แก้ไข ${H(t.title)}">${ICON_PEN}</button><button class="icon-btn us-edit emp-del" data-act="tkDel" data-id="${H(t.id)}" aria-label="ลบ ${H(t.title)}">${ICON_BIN}</button></div></td></tr>`; }).join('') || '<tr><td colspan="6"><div class="empty">ยังไม่มีงานที่มอบหมาย — กด “+ มอบหมายงาน”</div></td></tr>'}
  </tbody></table>
</section>`;
    },

    assignMine(me) {
      const list = tasks().filter(t => mine(t, me)).sort((a, b) => (a.status === 'done') - (b.status === 'done') || (a.due || '9').localeCompare(b.due || '9'));
      const open = list.filter(t => t.status !== 'done'), done = list.filter(t => t.status === 'done');
      const card = t => { const late = t.status !== 'done' && t.due && t.due < U.today(); return `<div class="card pad tk-card">
        <div class="between" style="align-items:flex-start;gap:10px"><div style="min-width:0"><div style="font:600 16px var(--fh)">${t.priority === 'high' ? '<span class="badge st-tentative" style="margin-right:6px">ด่วน</span>' : ''}${H(t.title)}</div>
          <div class="small muted" style="margin-top:2px">มอบหมายโดย ${H(t.byName || '')} · ${H(when(t.at))} · ถึง${H(toText(t))}</div></div>
          ${t.status === 'done' ? '<span class="badge st-confirmed">ยืนยันแล้ว</span>' : `<span class="badge ${late ? 'st-tentative' : 'st-sent'}">${late ? 'เลยกำหนด' : 'รอดำเนินการ'}</span>`}</div>
        ${t.detail ? `<p style="margin:10px 0 0;white-space:pre-line">${H(t.detail)}</p>` : ''}
        <div class="row wrap small" style="gap:14px;margin-top:10px;color:var(--ink2)">${t.due ? `<span>กำหนดเสร็จ <b class="${late ? 'red' : ''}">${U.fmtD(t.due)}</b></span>` : ''}${t.quoteId ? `<a href="#/quote/${encodeURIComponent(t.quoteId)}">${H(t.quoteId)}</a>` : ''}${t.to.type === 'dept' && (t.to.emps || []).length ? `<span>พนักงาน: ${H(t.to.emps.map(empName).filter(Boolean).join(', '))}</span>` : ''}</div>
        ${t.status === 'done' ? `<div class="small muted" style="margin-top:10px">ยืนยันโดย ${H(t.doneByName || '')} · ${H(when(t.doneAt))}${t.doneNote ? ' · “' + H(t.doneNote) + '”' : ''}</div>` : `<button class="btn teal" style="margin-top:12px" data-act="tkDone" data-id="${H(t.id)}">ยืนยันงานนี้</button>`}
      </div>`; };
      return `<h2 style="margin:0 0 10px">รอดำเนินการ <span class="muted" style="font-weight:400">${open.length}</span></h2>
        <div class="grid g2" style="margin-bottom:22px">${open.map(card).join('') || '<div class="empty card">ไม่มีงานค้าง</div>'}</div>
        ${done.length ? `<h2 style="margin:0 0 10px">ยืนยันแล้ว <span class="muted" style="font-weight:400">${done.length}</span></h2><div class="grid g2">${done.map(card).join('')}</div>` : ''}`;
    },

    approve() {
      const wait = toApprove(), done = Store.db.beos.filter(b => b.signoff).sort((a, b) => b.signoff.at.localeCompare(a.signoff.at));
      const editing = Store.db.beos.filter(b => g.Approval && g.Approval.state(b) === 'pending');
      const row = (b, btn) => { const q = Store.quote(b.quoteId) || {}, ack = Object.values(b.depts || {}), ok = ack.filter(d => d.ack).length; return `<tr>
        <td><a href="#/beo/${encodeURIComponent(b.id)}" style="font-weight:600">${H(b.id)}</a><div class="sub">${H(q.client || '')} · ${H(q.title || '')}</div><div class="sub show-m">${q.start ? U.fmtRange(q.start, q.end) : ''} · ${U.n(q.pax || 0)} ท่าน</div></td>
        <td class="hide-m">${q.start ? U.fmtRange(q.start, q.end) : '—'}</td><td class="hide-m num">${U.n(q.pax || 0)}</td>
        <td class="hide-m small">${b.status === 'sent' ? `ส่งแผนกแล้ว · รับทราบ ${ok}/${ack.length}` : 'ยังไม่ส่งแผนก'}</td>
        <td class="r">${btn}</td></tr>`; };
      return `
<section class="card" style="overflow:hidden;margin-bottom:16px">
  <div class="card-hd"><div><h2>รออนุมัติ <span class="muted" style="font-weight:400">${wait.length}</span></h2><div class="small muted">ใบฟังก์ชันของงานที่ลูกค้ายืนยันแล้ว</div></div></div>
  <table class="tbl resp"><thead><tr><th>ใบฟังก์ชัน</th><th>วันงาน</th><th>ท่าน</th><th>สถานะแผนก</th><th class="r"></th></tr></thead><tbody>
  ${wait.map(b => row(b, `<div class="row" style="gap:6px;justify-content:flex-end;flex-wrap:nowrap"><a class="btn ghost sm" href="#/beo/${encodeURIComponent(b.id)}">ดู</a><button class="btn teal sm" data-act="apApprove" data-id="${H(b.id)}">อนุมัติ</button></div>`)).join('') || '<tr><td colspan="5"><div class="empty">ไม่มีใบฟังก์ชันรออนุมัติ</div></td></tr>'}
  </tbody></table>
  ${editing.length ? `<div class="note" style="margin:0 22px 18px">มีการแก้ไขใบฟังก์ชันรออนุมัติ ${editing.length} ใบ (${editing.map(b => `<a href="#/beo/${encodeURIComponent(b.id)}">${H(b.id)}</a>`).join(', ')}) — อนุมัติการแก้ไขในหน้าใบฟังก์ชันก่อน แล้วจะขึ้นในรายการนี้</div>` : ''}
</section>
<section class="card" style="overflow:hidden">
  <div class="card-hd"><div><h2>อนุมัติแล้ว <span class="muted" style="font-weight:400">${done.length}</span></h2><div class="small muted">ถ้ามีการแก้ไขใบฟังก์ชันภายหลัง จะกลับไปรออนุมัติใหม่</div></div></div>
  <table class="tbl resp"><thead><tr><th>ใบฟังก์ชัน</th><th>วันงาน</th><th>ท่าน</th><th>สถานะแผนก</th><th class="r">อนุมัติโดย</th></tr></thead><tbody>
  ${done.map(b => row(b, `<span class="badge st-confirmed">อนุมัติแล้ว</span><div class="xs muted" style="margin-top:4px">${H(b.signoff.byName)} · ${H(when(b.signoff.at))}</div>`)).join('') || '<tr><td colspan="5"><div class="empty">ยังไม่มีรายการที่อนุมัติ</div></td></tr>'}
  </tbody></table>
</section>`;
    }
  };
  const ICON_PEN = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>';
  const ICON_BIN = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>';

  /* ---------------- assign form ---------------- */
  function taskForm(t) {
    const isNew = !t.id, users = Store.db.users.filter(u => u.status !== 'disabled');
    const quotes = Store.db.quotes.filter(q => q.status !== 'cancelled').sort((a, b) => a.start.localeCompare(b.start));
    App.modal(`<h2>${isNew ? 'มอบหมายงาน' : 'แก้ไขงานที่มอบหมาย'}</h2>
      <div class="field"><label class="f" for="tkTitle">งานที่ต้องทำ</label><input class="input" id="tkTitle" name="title" value="${H(t.title || '')}" placeholder="เช่น โทรยืนยันจำนวนผู้เข้าร่วม"></div>
      <div class="field"><label class="f" for="tkDetail">รายละเอียด</label><textarea class="input" id="tkDetail" name="detail" rows="3" style="height:auto;padding-top:10px">${H(t.detail || '')}</textarea></div>
      <div class="grid g2" style="margin-top:14px">
        <div><label class="f" for="tkDue">กำหนดเสร็จ</label><input class="input" id="tkDue" type="date" name="due" value="${H(t.due || '')}"></div>
        <div><label class="f" for="tkPri">ความสำคัญ</label><select class="input" id="tkPri" name="priority"><option value="normal">ปกติ</option><option value="high" ${t.priority === 'high' ? 'selected' : ''}>ด่วน</option></select></div>
      </div>
      <div class="field" style="margin-top:14px"><label class="f" for="tkQ">เกี่ยวกับงาน (ไม่บังคับ)</label><select class="input" id="tkQ" name="quoteId"><option value="">— ไม่ระบุ —</option>${quotes.map(q => `<option value="${H(q.id)}" ${q.id === t.quoteId ? 'selected' : ''}>${H(q.id)} · ${H(q.client)} · ${U.fmtD(q.start, { year: false })}</option>`).join('')}</select></div>
      <div class="f" style="margin-top:14px">มอบหมายให้</div>
      <div class="seg white" role="radiogroup" style="margin:6px 0 10px"><button type="button" class="${t.to.type === 'user' ? 'on' : ''}" data-to="user" role="radio" aria-checked="${t.to.type === 'user'}">ผู้ใช้งาน</button><button type="button" class="${t.to.type === 'dept' ? 'on' : ''}" data-to="dept" role="radio" aria-checked="${t.to.type === 'dept'}">แผนก</button></div>
      <div id="tkToUser" ${t.to.type === 'user' ? '' : 'hidden'}><label class="f" for="tkUser">ผู้ใช้งาน</label><select class="input" id="tkUser">${users.map(u => `<option value="${H(u.id)}" ${u.id === t.to.id ? 'selected' : ''}>${H(u.name)} · ${H(App.roleName(u))} · ${H(u.dept || '')}</option>`).join('')}</select></div>
      <div id="tkToDept" ${t.to.type === 'dept' ? '' : 'hidden'}><label class="f" for="tkDept">แผนก</label><select class="input" id="tkDept">${D.DEPTS.map(d => `<option value="${d.id}" ${d.id === t.to.dept ? 'selected' : ''}>${H(d.name)}</option>`).join('')}</select>
        <div class="small muted" style="margin:10px 0 6px">พนักงานในแผนก (เลือกได้หลายคน หรือไม่เลือกก็ได้ = ทั้งแผนก)</div><div id="tkEmps" class="tk-emps"></div>
        <div class="xs muted" style="margin-top:6px">ผู้ใช้งานที่อยู่แผนกนี้ (เช่น หัวหน้าแผนก) จะเห็นงานและกดยืนยันได้</div></div>
      <div class="small red" id="tkErr" style="min-height:20px;margin-top:8px" role="alert"></div>
      <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="tkOk">${isNew ? 'มอบหมายงาน' : 'บันทึก'}</button></div>`, bg => {
      let type = t.to.type, picked = new Set(t.to.emps || []);
      const sel = bg.querySelector('#tkDept'), box = bg.querySelector('#tkEmps');
      const fillEmps = () => {
        const emps = Store.db.employees.filter(e => e.dept === sel.value);
        box.innerHTML = emps.map(e => `<label class="tk-emp"><input type="checkbox" value="${H(e.id)}" ${picked.has(e.id) ? 'checked' : ''}> ${H(e.name)} <span class="xs muted">${H(e.role || '')}</span></label>`).join('') || '<span class="small muted">ยังไม่มีพนักงานในแผนกนี้</span>';
        box.querySelectorAll('input').forEach(i => { i.onchange = () => { if (i.checked) picked.add(i.value); else picked.delete(i.value); }; });
      };
      fillEmps(); sel.onchange = () => { picked = new Set(); fillEmps(); };
      bg.querySelectorAll('[data-to]').forEach(b => { b.onclick = () => {
        type = b.dataset.to;
        bg.querySelectorAll('[data-to]').forEach(x => { x.classList.toggle('on', x === b); x.setAttribute('aria-checked', x === b); });
        bg.querySelector('#tkToUser').hidden = type !== 'user'; bg.querySelector('#tkToDept').hidden = type !== 'dept';
      }; });
      bg.querySelector('#tkOk').onclick = () => {
        const f = App.formData(bg), err = bg.querySelector('#tkErr');
        f.title = f.title.trim();
        if (!f.title) { err.textContent = 'กรุณาระบุงานที่ต้องทำ'; return; }
        let to;
        if (type === 'user') { const u = Store.db.users.find(x => x.id === bg.querySelector('#tkUser').value); if (!u) { err.textContent = 'เลือกผู้ใช้งาน'; return; } to = { type: 'user', id: u.id, name: u.name }; }
        else to = { type: 'dept', dept: sel.value, name: dept(sel.value).name, emps: [...picked] };
        const me = App.me(), data = { title: f.title, detail: f.detail.trim(), due: f.due, priority: f.priority, quoteId: f.quoteId, to };
        if (isNew) tasks().push(Object.assign({ id: 'TK-' + Date.now().toString(36).toUpperCase(), by: me.id, byName: me.name, at: U.nowTH(), status: 'open' }, data));
        else {
          const cur = tasks().find(x => x.id === t.id); if (!cur) { App.closeModal(); return; }
          // a new assignee starts over
          const moved = JSON.stringify([cur.to.type, cur.to.id, cur.to.dept]) !== JSON.stringify([to.type, to.id, to.dept]);
          Object.assign(cur, data);
          if (moved) { cur.status = 'open'; delete cur.doneBy; delete cur.doneByName; delete cur.doneAt; delete cur.doneNote; cur.at = U.nowTH(); }
        }
        Store.save(); App.closeModal(); App.render(); App.toast(isNew ? 'มอบหมายงานให้' + (to.type === 'user' ? to.name : 'แผนก' + dept(to.dept).short) + 'แล้ว' : 'บันทึกแล้ว');
      };
    });
  }

  Object.assign(App.ACT, {
    tkSt(el) { App.state.ui.tasks.st = el.dataset.v; App.render(); },
    tkNew() { if (isAdmin()) taskForm({ to: { type: 'user' }, priority: 'normal' }); },
    tkEdit(el) { const t = tasks().find(x => x.id === el.dataset.id); if (t && isAdmin()) taskForm(JSON.parse(JSON.stringify(t))); },
    tkDel(el) {
      const t = tasks().find(x => x.id === el.dataset.id); if (!t || !isAdmin()) return;
      App.confirm('ลบงาน “' + t.title + '”?', () => { Store.db.tasks = tasks().filter(x => x !== t); Store.save(); App.render(); App.toast('ลบแล้ว'); }, 'ลบงาน');
    },
    tkView(el) { const t = tasks().find(x => x.id === el.dataset.id), me = App.me(); if (t && (isAdmin() || mine(t, me))) taskView(t); },
    tkDone(el) {
      const t = tasks().find(x => x.id === el.dataset.id), me = App.me(); if (!t || !mine(t, me)) return;
      App.modal(`<h2>ยืนยันงานนี้?</h2><p class="muted" style="margin-top:-8px">${H(t.title)}</p>
        <div class="field"><label class="f" for="tkNote">หมายเหตุถึงผู้ดูแลระบบ (ไม่บังคับ)</label><textarea class="input" id="tkNote" rows="2" style="height:auto;padding-top:10px" placeholder="เช่น ดำเนินการเรียบร้อยแล้ว"></textarea></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn teal" id="tkDoneOk">ยืนยันงาน</button></div>`, bg => bg.querySelector('#tkDoneOk').onclick = () => {
        Object.assign(t, { status: 'done', doneBy: me.id, doneByName: me.name, doneAt: U.nowTH(), doneNote: bg.querySelector('#tkNote').value.trim() });
        Store.save(); App.closeModal(); App.render(); App.toast('ยืนยันงานแล้ว — แจ้งผู้ดูแลระบบแล้ว');
      });
    },
    apApprove(el) {
      const b = Store.db.beos.find(x => x.id === el.dataset.id), me = App.me(); if (!b || !isAdmin()) return;
      const q = Store.quote(b.quoteId) || {};
      App.confirm(`อนุมัติใบฟังก์ชัน ${b.id} · ${q.client || ''}?`, () => {
        b.signoff = { by: me.id, byName: me.name, at: U.nowTH() };
        if (g.API && g.API.beoPrev) g.API.beoPrev[b.id] = JSON.stringify(b); // sign-off is not a content edit
        Store.save(); App.render(); App.toast('อนุมัติ ' + b.id + ' แล้ว');
      }, 'อนุมัติ');
    }
  });
})(window);
