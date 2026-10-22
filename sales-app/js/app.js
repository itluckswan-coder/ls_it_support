/* Luck Swan Sales — core: router, shell, login, home, shared helpers */
(function (g) {
  'use strict';
  const U = g.U, D = g.DATA, Store = g.Store;
  const H = U.esc;

  const ICON = {
    arrow: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>',
    back: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5M11 6l-6 6 6 6"/></svg>',
    menu: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 7h16M4 12h16M4 17h16"/></svg>',
    search: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
    home: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M3 11 12 4l9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/></svg>',
    doc: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M6 3h8l5 5v13H6z"/><path d="M14 3v5h5M9 13h7M9 17h7"/></svg>',
    cal: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3.5" y="5" width="17" height="15.5" rx="2"/><path d="M3.5 10h17M8 3v4M16 3v4"/></svg>',
    sheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="4" y="3.5" width="16" height="17" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/></svg>',
    more: '<svg viewBox="0 0 24 24" fill="currentColor"><circle cx="5" cy="12" r="2"/><circle cx="12" cy="12" r="2"/><circle cx="19" cy="12" r="2"/></svg>',
    check: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="m5 12 5 5 9-10"/></svg>',
    x: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M6 6l12 12M18 6 6 18"/></svg>',
    pen: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20h4L19 9l-4-4L4 16v4zM13.5 6.5l4 4"/></svg>',
    trash: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>',
    left: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="m15 6-6 6 6 6"/></svg>',
    right: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="m9 6 6 6-6 6"/></svg>'
  };

  const NAV = [
    ['home', 'หน้าหลัก', '#/home'], ['meet', 'ห้องประชุม', '#/meet'], ['room', 'ห้องพัก', '#/rooms'], ['food', 'จัดเลี้ยง', '#/food'], ['quote', 'ใบเสนอราคา', '#/quotes'],
    ['beo', 'ใบฟังก์ชัน', '#/beo'], ['cal', 'ปฏิทิน', '#/cal'], ['staff', 'พนักงาน', '#/staff'], ['users', 'ผู้ใช้งาน', '#/users'], ['leave', 'การแจ้งใบลา', '#/leave'], ['tasks', 'Assignment & Approve', '#/tasks'], ['dash', 'Dashboard', '#/dash']
  ];

  const App = {
    ICON, NAV, V: {}, ACT: {}, CHG: {}, INP: {},
    state: { route: 'home', params: [], ui: {} },
    deferredInstall: null,

    /* ---------- session ---------- */
    me() { const id = App.state.uid; return id ? Store.db.users.find(u => u.id === id) : null; },
    // sections hidden for the signed-in user's role (settings.roles.views[roleId]); admins see everything
    navHidden() {
      const me = App.me(); if (!me || me.role === 'admin' || !App.roleOf) return [];
      const v = ((Store.db.settings && Store.db.settings.roles) || {}).views || {}, l = v[App.roleOf(me).id];
      return Array.isArray(l) ? l.filter(x => x !== 'home') : [];
    },
    // the built-in admin, sales manager and sales roles (not roles added on those levels)
    // the floor-plan designer is for admins and the F&B department (its head and staff — the user's department is F&B / kitchen)
    canPlan(me) { me = me === undefined ? App.me() : me; if (!me) return false; if (me.role === 'admin') return true; const k = (D.DEPTS || []).find(d => d.id === 'kitchen') || {}; const dept = String(me.dept || '').trim().toLowerCase(); return !!dept && (/f\s*&\s*b|ครัว|อาหารและเครื่องดื่ม|kitchen/.test(dept) || [k.name, k.short].filter(Boolean).some(n => n.toLowerCase() === dept)); },
    salesRole(me) { me = me === undefined ? App.me() : me; return !!me && ['admin', 'manager', 'sales'].includes(me.role) && !(App.roleOf && App.roleOf(me).custom); },
    can(p) { const m = App.me(); return !!(m && m.perms && m.perms[p]); },
    firstName(u) { return (u && u.name || '').split(' ')[0]; },
    restoreSession() {
      let id = null;
      try { id = localStorage.getItem('lsSales.session') || sessionStorage.getItem('lsSales.session'); } catch (e) { /* ignore */ }
      if (id && Store.db.users.some(u => u.id === id && u.status !== 'disabled')) App.state.uid = id;
    },
    async login(ident, pw, remember) {
      if (g.API && g.API.online) {
        try { const u = await g.API.login(ident, pw, remember); App.state.uid = u.id; localStorage.setItem('lsSales.cloudUid', u.id); return null; }
        catch (e) { return e.message; }
      }
      ident = String(ident || '').trim().toLowerCase();
      const u = Store.db.users.find(x => x.email.toLowerCase() === ident || x.email.split('@')[0].toLowerCase() === ident || x.id.toLowerCase() === ident);
      if (!u) return 'ไม่พบผู้ใช้นี้ในระบบ';
      if (u.status === 'disabled') return 'บัญชีนี้ถูกปิดใช้งาน กรุณาติดต่อผู้ดูแลระบบ';
      const ok = u.pw ? (await g.hashPw(pw)) === u.pw : pw === D.DEFAULT_PW;
      if (!ok) return 'รหัสผ่านไม่ถูกต้อง';
      if (u.status === 'invited') u.status = 'active';
      u.lastLogin = U.nowTH();
      u.lastDevice = App.deviceName();
      Store.save();
      App.state.uid = u.id;
      try {
        (remember ? localStorage : sessionStorage).setItem('lsSales.session', u.id);
        (remember ? sessionStorage : localStorage).removeItem('lsSales.session');
      } catch (e) { /* ignore */ }
      return null;
    },
    logout() {
      if (g.API && g.API.online) g.API.logout();
      try { localStorage.removeItem('lsSales.session'); sessionStorage.removeItem('lsSales.session'); localStorage.removeItem('lsSales.cloudUid'); } catch (e) { /* ignore */ }
      App.state.uid = null; App.go('#/login');
    },
    deviceName() {
      const ua = navigator.userAgent;
      const b = /Edg\//.test(ua) ? 'Edge' : /Chrome\//.test(ua) ? 'Chrome' : /Safari\//.test(ua) ? 'Safari' : /Firefox\//.test(ua) ? 'Firefox' : 'Browser';
      const d = /iPad/.test(ua) ? 'iPad' : /iPhone/.test(ua) ? 'iPhone' : /Android/.test(ua) ? 'Android' : /Mac/.test(ua) ? 'Mac' : /Windows/.test(ua) ? 'Windows' : '';
      return b + (d ? ' · ' + d : '');
    },

    /* ---------- routing ---------- */
    go(h) { if (location.hash === h) App.render(); else location.hash = h; },
    parse() {
      const h = (location.hash || '#/home').replace(/^#\/?/, '');
      const [route, ...params] = h.split('/').map(decodeURIComponent);
      return { route: route || 'home', params };
    },
    render() {
      const { route, params } = App.parse();
      const cloud = g.API && g.API.online;
      if (cloud && g.API.setup && route !== 'setup') { location.replace('#/setup'); return; }
      if (route === 'setup' && !(cloud && g.API.setup)) { location.replace(App.state.uid ? '#/home' : '#/login'); return; }
      if (!App.state.uid && route !== 'login' && route !== 'setup') { location.replace('#/login'); return; }
      if (App.state.uid && route === 'login') { location.replace('#/home'); return; }
      // a section the user's role can't see goes back to the home page
      const sec = (App.V[route] || App.V.home).nav || route;
      if (App.state.uid && sec !== 'home' && App.navHidden().includes(sec)) { location.replace('#/home'); return; }
      App.state.route = route; App.state.params = params;
      const view = App.V[route] || App.V.home;
      const root = document.getElementById('app');
      const keepScroll = App.state._sameRoute === location.hash;
      const y = window.scrollY;
      let body;
      try { body = view.render(params); } catch (e) { console.error(e); body = '<div class="empty">เกิดข้อผิดพลาด: ' + H(e.message) + '</div>'; }
      if (view.bare) root.innerHTML = body;
      else root.innerHTML = App.shell(view.nav || route, body, view);
      document.title = (view.title ? (typeof view.title === 'function' ? view.title(params) : view.title) + ' · ' : '') + 'Luck Swan Sales';
      if (view.after) try { view.after(params); } catch (e) { console.error(e); }
      if (keepScroll) window.scrollTo(0, y); else window.scrollTo(0, 0);
      App.state._sameRoute = location.hash;
    },
    rerender() { App.render(); },
    // re-render after background changes, but never under an open dialog or while someone is typing
    softRender() {
      const a = document.activeElement;
      if (document.getElementById('modalBg') || (a && (/^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) || a.isContentEditable))) { App._pendingRender = true; return; }
      App._pendingRender = false; App.render();
    },

    // menu groups: one button that opens its sections (shown only when the user may see at least one of them)
    NAV_GROUPS: [{ id: 'rooms', name: 'การจัดการห้อง', kids: ['meet', 'room'] }],
    // the menu as rendered: plain entries, with each group in place of its first visible section
    navTree() {
      const out = [], done = new Set();
      for (const n of App.navItems()) {
        const grp = App.NAV_GROUPS.find(gr => gr.kids.includes(n[0]));
        if (!grp) { out.push(n); continue; }
        if (done.has(grp.id)) continue; done.add(grp.id);
        out.push({ grp, items: App.navItems().filter(x => grp.kids.includes(x[0])) });
      }
      return out;
    },
    // menu entries this user may see (Dashboard is for admins only)
    closeNavPop() { document.querySelectorAll('.nav-pop:not(.lang-pop):not(.theme-pop)').forEach(x => x.remove()); document.querySelectorAll('[data-act=navGrp]').forEach(b => b.setAttribute('aria-expanded', 'false')); },
    navItems() {
      const me = App.me();
      const hidden = App.navHidden();
      return NAV.filter(n => (n[0] !== 'users' || App.can('users') || App.can('view_beo')) && (n[0] !== 'dash' || (me && me.role === 'admin')) && !hidden.includes(n[0]));
    },
    /* ---------- shell ---------- */
    shell(active, body, view) {
      const me = App.me();
      const tabs = [['home', 'หน้าหลัก', '#/home', ICON.home], ['quote', 'ใบเสนอ', '#/quotes', ICON.doc], ['cal', 'ปฏิทิน', '#/cal', ICON.cal], ['beo', 'ใบฟังก์ชัน', '#/beo', ICON.sheet]];
      return `
<header class="topbar">
  <button class="menu-btn" data-act="drawer" aria-label="เมนู">${ICON.menu}</button>
  <a class="brand" href="#/home"><img src="img/logo-wide-sm.png" alt="Luck Swan Resort"><span class="sep"></span><span class="t">ลักษวรรณ รีสอร์ต<small>ฝ่ายขาย</small></span></a>
  <nav class="nav">${App.navTree().map(n => n.grp
    ? `<button type="button" class="nav-grp ${n.items.some(x => x[0] === active) ? 'on' : ''}" data-act="navGrp" data-g="${n.grp.id}" aria-haspopup="true" aria-expanded="false">${n.grp.name} <span class="caret" aria-hidden="true">▾</span></button>`
    : `<a href="${n[2]}" class="${n[0] === active ? 'on' : ''}">${n[1]}</a>`).join('')}</nav>
  <div class="tb-right">
    ${g.I18N ? g.I18N.button() : ''}
    ${g.THEME ? g.THEME.button() : ''}
    ${App.bellButton ? App.bellButton() : ''}
    <button class="search-btn" data-act="toggleSearch" aria-label="ค้นหา">${ICON.search}</button>
    <div class="search" id="searchBox"><input type="search" placeholder="ค้นหาลูกค้า, ใบเสนอราคา…" data-input="search" aria-label="ค้นหา" autocomplete="off"><div class="search-res hidden" id="searchRes"></div></div>
    <button class="avatar" data-act="userMenu" aria-label="บัญชีผู้ใช้">${H(U.initials(me && me.name))}</button>
  </div>
</header>
<main class="${view && view.bleed ? 'bleed' : ''}">${body}</main>
<nav class="tabbar">${tabs.map(t => `<a href="${t[2]}" class="${t[0] === active ? 'on' : ''}">${t[3]}<span>${t[1]}</span></a>`).join('')}<a href="#" data-act="drawer" class="${['food', 'meet', 'room', 'staff', 'users', 'tasks', 'dash'].includes(active) ? 'on' : ''}">${ICON.more}<span>เพิ่มเติม</span></a></nav>`;
    },

    /* ---------- ui helpers ---------- */
    toast(msg) {
      const old = document.querySelector('.toast'); if (old) old.remove();
      const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t);
      setTimeout(() => t.remove(), 2600);
    },
    modal(html, after) {
      App.closeModal();
      const bg = document.createElement('div'); bg.className = 'modal-bg'; bg.id = 'modalBg';
      bg.innerHTML = `<div class="modal" role="dialog" aria-modal="true">${html}</div>`;
      bg.addEventListener('mousedown', e => { if (e.target === bg) App.closeModal(); });
      document.body.appendChild(bg);
      const f = bg.querySelector('input:not([type=checkbox]),select,textarea'); if (f && window.innerWidth > 760) f.focus();
      if (after) after(bg);
      return bg;
    },
    closeModal() { const m = document.getElementById('modalBg'); if (m) m.remove(); if (App._pendingRender) setTimeout(() => App.softRender()); },
    confirm(msg, yes, label = 'ยืนยัน') {
      App.modal(`<h2>${H(msg)}</h2><div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="cfYes">${H(label)}</button></div>`,
        bg => bg.querySelector('#cfYes').onclick = () => { App.closeModal(); yes(); });
    },
    formData(root) {
      const o = {};
      root.querySelectorAll('[name]').forEach(el => { o[el.name] = el.type === 'checkbox' ? el.checked : el.value; });
      return o;
    },
    statusBadge(s) { return `<span class="badge st-${s}">${H(D.STATUS[s] || s)}</span>`; },
    save() { Store.save(); },
    sw(name, on, attrs = '') { return `<label class="switch"><input type="checkbox" ${on ? 'checked' : ''} ${attrs} aria-label="${H(name)}"><span></span></label>`; },

    /* ---------- domain helpers ---------- */
    calc(q) { return U.calc(q, Store.db.settings); },
    holding(q) { return q.status === 'tentative' || q.status === 'confirmed'; },
    // meeting room bookings on a date (other quotes)
    /* meeting-room time slots a quote holds: [{ room, date, from, to }] — from each day's schedule, else the times in
       the line's description (e.g. "เต็มวัน 08:00–17:00", "18:00–22:00"), else a full day 08:00–17:00 */
    meetSlots(q) {
      const out = [], days = U.daysOf(q.start, q.end), hm = t => String(t || '').replace('.', ':');
      for (const it of q.items || []) {
        if (it.sec !== 'meet' || !D.MEETING_ROOMS.some(r => r.id === it.ref)) continue;
        const sch = (it.sched || []).filter(s => s.start);
        if (sch.length) { for (const s of sch) { const d = days[(Number(s.day) || 1) - 1]; if (d) out.push({ room: it.ref, date: d, from: hm(s.start), to: hm(s.end) || '23:59' }); } continue; }
        const m = /(\d{1,2}[:.]\d{2})\s*[–\-]\s*(\d{1,2}[:.]\d{2})/.exec(it.desc || ''), pad = t => t.length === 4 ? '0' + t : t;
        const from = m ? pad(hm(m[1])) : '08:00', to = m ? pad(hm(m[2])) : '17:00';
        for (const d of days.slice(0, Math.max(1, Math.min(days.length, Number(it.q1) || days.length)))) out.push({ room: it.ref, date: d, from, to });
      }
      return out;
    },
    // the same room on the same day at overlapping times as another held booking (จองชั่วคราว / ยืนยันแล้ว)
    meetConflicts(q) {
      const mine = App.meetSlots(q), out = [];
      if (!mine.length) return out;
      for (const o of Store.db.quotes) {
        if (o.id === q.id || !App.holding(o)) continue;
        for (const b of App.meetSlots(o)) for (const a of mine)
          if (a.room === b.room && a.date === b.date && a.from < b.to && b.from < a.to) out.push({ room: a.room, date: a.date, mine: a.from + '–' + a.to, other: o, theirs: b.from + '–' + b.to });
      }
      return out;
    },
    overlapKey: list => JSON.stringify(list.map(c => [c.room, c.date, c.mine, c.other.id, c.theirs]).sort()),
    // overlapping bookings need the admin's OK before the quote holds the room; a change of times needs it again
    overlapOk(q) { const c = App.meetConflicts(q); return !c.length || !!(q.overlapOk && q.overlapOk.key === App.overlapKey(c)); },
    meetBooking(roomId, date, excludeId) {
      return Store.db.quotes.find(q => q.id !== excludeId && App.holding(q) && date >= q.start && date <= q.end &&
        q.items.some(it => it.sec === 'meet' && it.ref === roomId));
    },
    roomTypeBooked(typeId, date, excludeId) {
      let n = 0;
      for (const q of Store.db.quotes) {
        if (q.id === excludeId || !App.holding(q)) continue;
        for (const it of q.items) {
          if (it.sec !== 'room' || it.free) continue;
          const nights = Number(it.q2) || 1, ci = it.ci || q.start, co = U.addDays(ci, nights);
          if (!(date >= ci && date < co)) continue;
          if (it.block && it.block[typeId]) n += it.block[typeId];
          else if (it.ref === typeId) n += Number(it.q1) || 0;
        }
      }
      return n;
    },
    roomsLeft(date, excludeId) {
      const total = D.ROOM_TYPES.reduce((a, t) => a + t.total, 0);
      const used = D.ROOM_TYPES.reduce((a, t) => a + App.roomTypeBooked(t.id, date, excludeId), 0);
      return { total, used, left: total - used };
    },
    qtyLabel(it) {
      if (it.q2 == null || it.q2 === '') return `${U.n(it.q1)} ${it.u1 || ''}`.trim();
      return `${U.n(it.q1)}${it.u1 && it.sec !== 'food' ? ' ' + it.u1 : ''} × ${U.n(it.q2)} ${it.u2 || ''}`.trim();
    },
    // which quote catalog pages add to
    ctxQuote() {
      const db = Store.db;
      let id = App.state.ctx || (function () { try { return sessionStorage.getItem('lsSales.ctx'); } catch (e) { return null; } })();
      let q = id && db.quotes.find(x => x.id === id && x.status !== 'cancelled');
      if (!q) q = db.quotes.filter(x => x.status === 'draft' || x.status === 'sent' || x.status === 'tentative').sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''))[0];
      return q || null;
    },
    setCtx(id) { App.state.ctx = id; try { sessionStorage.setItem('lsSales.ctx', id); } catch (e) { /* ignore */ } },
    touch(q, note) {
      q.updatedAt = U.nowTH();
      if (note) q.history.unshift({ at: q.updatedAt, t: note, by: App.me() && App.me().name });
      Store.save();
    },
    pickQuoteModal(title, cb, filter) {
      const qs = Store.db.quotes.filter(filter || (q => q.status !== 'cancelled')).sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
      App.modal(`<h2>${H(title)}</h2>
        <div>${qs.map(q => `<button class="pick" data-id="${H(q.id)}"><span><b>${H(q.client)}</b><br><span class="small muted">${H(q.id)} · ${H(q.title)} · ${U.fmtRange(q.start, q.end)}</span></span>${App.statusBadge(q.status)}</button>`).join('') || '<div class="empty">ไม่มีใบเสนอราคา</div>'}</div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ปิด</button><a class="btn red" href="#/new" data-act="closeModal">+ ใบเสนอราคาใหม่</a></div>`,
      bg => bg.querySelectorAll('.pick').forEach(b => b.onclick = () => { App.closeModal(); cb(b.dataset.id); }));
    },
    ctxBar(q, what) {
      if (!q) return `<div class="row wrap"><span class="muted">ยังไม่ได้เลือกใบเสนอราคา</span><button class="link" data-act="pickCtx">เลือกใบเสนอ</button></div>`;
      return `<div class="row wrap" style="gap:6px"><span class="muted">${what || 'กำลังเพิ่มให้'}</span><b>${H(q.id)} · ${U.n(q.pax)} ท่าน</b><span class="muted small">(${H(q.client)})</span><button class="link" data-act="pickCtx">เปลี่ยน</button></div>`;
    },
    dock(q) {
      if (!q) return '';
      const c = App.calc(q);
      return `<div class="dock no-print"><span class="lbl small">${H(q.id)} · ${q.items.length} รายการ</span><span class="v num">฿${U.n(Math.round(c.net))}</span><a class="btn red sm" href="#/quote/${encodeURIComponent(q.id)}">กลับไปใบเสนอ</a></div>`;
    },
    tel(p) { return 'tel:' + String(p || '').replace(/[^\d+]/g, ''); }
  };

  /* ---------- event delegation ---------- */
  document.addEventListener('click', e => {
    const el = e.target.closest('[data-act]');
    if (!el) {
      if (!e.target.closest('.pop') && !e.target.closest('[data-act="userMenu"]')) { const p = document.querySelector('.pop'); if (p) p.remove(); }
      if (!e.target.closest('#searchBox')) { const r = document.getElementById('searchRes'); if (r) r.classList.add('hidden'); }
      return;
    }
    const fn = App.ACT[el.dataset.act];
    if (fn) { if (el.tagName === 'A' && el.getAttribute('href') === '#') e.preventDefault(); fn(el, e); }
  });
  document.addEventListener('change', e => { const el = e.target.closest('[data-change]'); if (el && App.CHG[el.dataset.change]) App.CHG[el.dataset.change](el, e); });
  document.addEventListener('input', e => { const el = e.target.closest('[data-input]'); if (el && App.INP[el.dataset.input]) App.INP[el.dataset.input](el, e); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') { App.closeModal(); const d = document.querySelector('.drawer'); if (d) App.ACT.closeDrawer(); } });

  Object.assign(App.ACT, {
    closeModal() { App.closeModal(); },
    drawer() {
      const active = App.state.route === 'quote' || App.state.route === 'quotes' || App.state.route === 'new' || App.state.route === 'doc' ? 'quote' : (App.V[App.state.route] || {}).nav || App.state.route;
      const bg = document.createElement('div'); bg.className = 'drawer-bg'; bg.dataset.act = 'closeDrawer';
      const d = document.createElement('nav'); d.className = 'drawer';
      d.innerHTML = `<div style="padding:4px 10px 16px"><img src="img/logo-wide-sm.png" alt="" style="width:120px"></div>` + App.navTree().map(n => {
        const link = (x, sub) => `<a href="${x[2]}" class="${x[0] === active ? 'on' : ''}${sub ? ' sub' : ''}" data-act="closeDrawer">${x[1]}</a>`;
        if (!n.grp) return link(n);
        const open = n.items.some(x => x[0] === active);
        return `<button type="button" class="drawer-grp ${open ? 'open' : ''}" data-act="drawerGrp" aria-expanded="${open}">${n.grp.name} <span class="caret" aria-hidden="true">▾</span></button><div class="drawer-kids" ${open ? '' : 'hidden'}>${n.items.map(x => link(x, true)).join('')}</div>`;
      }).join('');
      document.body.append(bg, d);
    },
    drawerGrp(el) { const k = el.nextElementSibling, open = k.hidden; k.hidden = !open; el.classList.toggle('open', open); el.setAttribute('aria-expanded', open); },
    // a menu group's sections drop down under its button (fixed, so the scrolling menu bar does not clip them)
    navGrp(el) {
      const old = document.querySelector('.nav-pop'), same = old && old.dataset.g === el.dataset.g;
      App.closeNavPop(); if (same) return;
      const n = App.navTree().find(x => x.grp && x.grp.id === el.dataset.g); if (!n) return;
      const active = (App.V[App.state.route] || {}).nav || App.state.route, r = el.getBoundingClientRect();
      const pop = document.createElement('div'); pop.className = 'nav-pop'; pop.dataset.g = n.grp.id; pop.setAttribute('role', 'menu');
      pop.style.top = (r.bottom + 6) + 'px'; pop.style.left = Math.max(8, Math.min(r.left, innerWidth - 228)) + 'px';
      pop.innerHTML = n.items.map(x => `<a role="menuitem" href="${x[2]}" class="${x[0] === active ? 'on' : ''}">${x[1]}</a>`).join('');
      document.body.appendChild(pop); el.setAttribute('aria-expanded', 'true');
    },
    closeDrawer() { document.querySelectorAll('.drawer,.drawer-bg').forEach(x => x.remove()); },
    toggleSearch() { const s = document.getElementById('searchBox'); s.classList.toggle('open'); if (s.classList.contains('open')) s.querySelector('input').focus(); },
    userMenu(el) {
      const old = document.querySelector('.pop'); if (old) { old.remove(); return; }
      const me = App.me(), cloud = g.API && g.API.online;
      const p = document.createElement('div'); p.className = 'pop';
      p.innerHTML = `<div class="hd"><b>${H(me.name)}</b><div class="small muted">${H(me.email)}</div><div class="small muted">${H(App.roleName(me))}</div>
        <div class="xs" style="margin-top:6px;color:${cloud ? (g.API.offline ? 'var(--red)' : 'var(--teal-d)') : 'var(--muted)'}">${cloud ? (g.API.offline ? '● ออฟไลน์ — รอส่งข้อมูล' : '● เชื่อมต่อฐานข้อมูลกลางแล้ว') : '● ข้อมูลเก็บในเครื่องนี้'}</div></div>
        <button data-act="changePw">เปลี่ยนรหัสผ่าน</button>
        ${App.deferredInstall ? '<button data-act="install">ติดตั้งแอปบนเครื่องนี้</button>' : ''}
        <button data-act="howInstall">วิธีติดตั้งบนมือถือ / แท็บเล็ต</button>
        <button data-act="exportData">สำรองข้อมูล (ดาวน์โหลด)</button>
        ${!cloud || me.role === 'admin' ? `<button data-act="importData">${cloud ? 'นำเข้าข้อมูลสำรองขึ้นคลาวด์' : 'นำเข้าข้อมูลสำรอง'}</button>` : ''}
        ${me.role === 'admin' && !cloud ? '<button data-act="resetData">รีเซ็ตเป็นข้อมูลตัวอย่าง</button>' : ''}
        <button data-act="logout" style="color:var(--red)">ออกจากระบบ</button>`;
      document.body.appendChild(p);
    },
    logout() { document.querySelector('.pop')?.remove(); App.logout(); },
    install() { document.querySelector('.pop')?.remove(); const p = App.deferredInstall; if (p) { p.prompt(); p.userChoice.finally(() => { App.deferredInstall = null; }); } },
    howInstall() {
      document.querySelector('.pop')?.remove();
      App.modal(`<h2>ติดตั้งแอปบนอุปกรณ์</h2>
        <div class="stack">
          <div><b>iPhone / iPad (Safari)</b><div class="muted">แตะปุ่มแชร์ <b>⬆︎</b> → เลือก “เพิ่มไปยังหน้าจอโฮม” (Add to Home Screen)</div></div>
          <div><b>Android (Chrome)</b><div class="muted">แตะเมนู <b>⋮</b> → “ติดตั้งแอป” หรือ “เพิ่มลงในหน้าจอหลัก”</div></div>
          <div><b>คอมพิวเตอร์ (Chrome / Edge)</b><div class="muted">คลิกไอคอนติดตั้ง <b>⊕</b> ที่แถบที่อยู่ หรือเมนู → “ติดตั้ง Luck Swan Sales”</div></div>
          <div class="note">${g.API && g.API.online ? 'ข้อมูลเก็บในฐานข้อมูลกลาง ทุกเครื่องเห็นข้อมูลเดียวกัน — ถ้าออฟไลน์จะดูข้อมูลล่าสุดได้ และส่งการแก้ไขเมื่อกลับมาออนไลน์' : 'ใช้งานแบบออฟไลน์ได้หลังเปิดครั้งแรก ข้อมูลถูกเก็บในเครื่องนี้ — ใช้ “สำรองข้อมูล” เพื่อย้ายไปเครื่องอื่น'}</div>
        </div><div class="acts"><button class="btn dark" data-act="closeModal">เข้าใจแล้ว</button></div>`);
    },
    exportData() { document.querySelector('.pop')?.remove(); U.download('luckswan-sales-backup-' + U.today() + '.json', Store.exportJSON()); App.toast('ดาวน์โหลดไฟล์สำรองแล้ว'); },
    importData() {
      document.querySelector('.pop')?.remove();
      const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.json,application/json';
      inp.onchange = () => {
        const f = inp.files[0]; if (!f) return;
        f.text().then(async t => {
          try {
            if (g.API && g.API.online) {
              const d = JSON.parse(t);
              await g.API.req('POST', 'import', { data: d, users: true });
              await g.API.loadState(); App.toast('นำเข้าข้อมูลขึ้นฐานข้อมูลกลางแล้ว (ผู้ใช้ใหม่ใช้รหัสผ่านเริ่มต้น)'); App.render(); return;
            }
            Store.importJSON(t); App.toast('นำเข้าข้อมูลเรียบร้อย'); App.restoreSession(); App.render();
          } catch (e) { App.toast('นำเข้าไม่สำเร็จ: ' + e.message); }
        });
      };
      inp.click();
    },
    resetData() {
      document.querySelector('.pop')?.remove();
      App.confirm('รีเซ็ตข้อมูลทั้งหมดกลับเป็นข้อมูลตัวอย่าง?', () => { const uid = App.state.uid; Store.reset(); App.state.uid = Store.db.users.some(u => u.id === uid) ? uid : null; App.render(); App.toast('รีเซ็ตแล้ว'); }, 'รีเซ็ต');
    },
    changePw() {
      document.querySelector('.pop')?.remove();
      App.modal(`<h2>เปลี่ยนรหัสผ่าน</h2>
        <div class="field"><label class="f">รหัสผ่านปัจจุบัน</label><input class="input" type="password" name="cur" autocomplete="current-password"></div>
        <div class="field"><label class="f">รหัสผ่านใหม่ (อย่างน้อย 6 ตัวอักษร)</label><input class="input" type="password" name="nw" autocomplete="new-password"></div>
        <div class="field"><label class="f">ยืนยันรหัสผ่านใหม่</label><input class="input" type="password" name="nw2" autocomplete="new-password"></div>
        <div class="small red" id="pwErr"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="pwOk">บันทึก</button></div>`, bg => {
        bg.querySelector('#pwOk').onclick = async () => {
          const f = App.formData(bg), me = App.me(), err = bg.querySelector('#pwErr');
          if (g.API && g.API.online) {
            if (f.nw.length < 6) { err.textContent = 'รหัสผ่านใหม่สั้นเกินไป'; return; }
            if (f.nw !== f.nw2) { err.textContent = 'รหัสผ่านใหม่ไม่ตรงกัน'; return; }
            try { await g.API.req('POST', 'me/password', { current: f.cur, next: f.nw }); App.closeModal(); App.toast('เปลี่ยนรหัสผ่านแล้ว'); } catch (e) { err.textContent = e.message; }
            return;
          }
          const ok = me.pw ? (await g.hashPw(f.cur)) === me.pw : f.cur === D.DEFAULT_PW;
          if (!ok) { err.textContent = 'รหัสผ่านปัจจุบันไม่ถูกต้อง'; return; }
          if (f.nw.length < 6) { err.textContent = 'รหัสผ่านใหม่สั้นเกินไป'; return; }
          if (f.nw !== f.nw2) { err.textContent = 'รหัสผ่านใหม่ไม่ตรงกัน'; return; }
          me.pw = await g.hashPw(f.nw); Store.save(); App.closeModal(); App.toast('เปลี่ยนรหัสผ่านแล้ว');
        };
      });
    },
    pickCtx() { App.pickQuoteModal('เลือกใบเสนอราคาที่จะเพิ่มรายการ', id => { App.setCtx(id); App.render(); }, q => q.status !== 'cancelled'); }
  });

  /* ---------- global search ---------- */
  App.INP.search = U.debounce(el => {
    const v = el.value.trim().toLowerCase(), box = document.getElementById('searchRes');
    if (!v) { box.classList.add('hidden'); return; }
    const qs = Store.db.quotes.filter(q => [q.id, q.client, q.title, q.contact, q.email].join(' ').toLowerCase().includes(v)).slice(0, 8);
    const bs = Store.db.beos.filter(b => b.id.toLowerCase().includes(v)).slice(0, 3);
    box.innerHTML = qs.map(q => `<a href="#/quote/${encodeURIComponent(q.id)}"><b>${H(q.client)}</b> ${App.statusBadge(q.status)}<div class="small muted">${H(q.id)} · ${H(q.title)} · ${U.fmtRange(q.start, q.end)}</div></a>`).join('') +
      bs.map(b => `<a href="#/beo/${encodeURIComponent(b.id)}"><b>${H(b.id)}</b><div class="small muted">ใบฟังก์ชัน</div></a>`).join('') || '<div class="empty small">ไม่พบผลลัพธ์</div>';
    box.classList.remove('hidden');
  }, 120);

  /* =====================================================================
     LOGIN (1a)
     ===================================================================== */
  App.V.login = {
    bare: true, title: 'เข้าสู่ระบบ',
    render() {
      return `<div class="login">
  <div class="login-hero" style="background-image:url(img/meet-kinnaree.jpg)">
    <div class="tags"><span><i class="dot" style="background:var(--red)"></i>จัดเลี้ยง</span><span><i class="dot" style="background:var(--teal)"></i>ห้องประชุม</span><span><i class="dot" style="background:var(--yellow)"></i>ห้องพัก</span></div>
    <h2>ทำใบเสนอราคาสัมมนา<br>ครบในที่เดียว</h2>
    <p style="margin:0;opacity:.9">เลือกห้อง เมนู และห้องพัก ระบบรวมยอดพร้อมค่าบริการและ VAT ให้ทันที</p>
  </div>
  <div class="login-side"><div class="login-lang">${g.I18N ? g.I18N.button() : ''}${g.THEME ? g.THEME.button() : ''}</div><form class="login-box" id="loginForm" autocomplete="on">
    <div class="brand"><img src="img/logo-wide-sm.png" alt="Luck Swan Resort" style="width:116px;height:48px"><span class="sep" style="display:block"></span><span class="t" style="display:block">ลักษวรรณ รีสอร์ต<small>ฝ่ายขาย</small></span></div>
    <h1>เข้าสู่ระบบ</h1>
    <div class="muted" style="margin-bottom:22px">สำหรับพนักงานฝ่ายขายและการตลาด</div>
    <div class="card pad">
      <div class="field"><label class="f" for="lgU">อีเมลหรือรหัสพนักงาน</label><input class="input" id="lgU" name="u" type="text" autocomplete="username" required placeholder="name@luckswangroup.com"></div>
      <div class="field"><label class="f" for="lgP">รหัสผ่าน</label><div class="pw-wrap"><input class="input fill" id="lgP" name="p" type="password" autocomplete="current-password" required placeholder="อย่างน้อย 6 ตัวอักษร"><button type="button" data-act="togglePw">แสดง</button></div></div>
      <div class="between" style="margin:16px 0"><label class="row small" style="cursor:pointer">${App.sw('จำฉันไว้ในเครื่องนี้', true, 'name="r"')}จำฉันไว้ในเครื่องนี้</label><button type="button" class="link" data-act="forgotPw">ลืมรหัสผ่าน?</button></div>
      <div class="small red" id="lgErr" style="min-height:20px;margin-bottom:6px" role="alert"></div>
      <button class="btn red block" type="submit">เข้าสู่ระบบ ${ICON.arrow}</button>
    </div>
    <p class="small muted" style="text-align:center;margin-top:20px">ติดปัญหาเข้าสู่ระบบ ติดต่อผู้ดูแลระบบ โทร 053-750-855<br>ข้อมูลลูกค้าและราคาเป็นความลับของ บจก. ลักษวรรณ</p>
  </form></div></div>`;
    },
    after() {
      const f = document.getElementById('loginForm');
      f.onsubmit = async e => {
        e.preventDefault();
        const d = App.formData(f);
        const err = await App.login(d.u, d.p, d.r);
        if (err) { document.getElementById('lgErr').textContent = err; return; }
        App.go('#/home');
      };
    }
  };
  /* first run on the shared database: create the administrator */
  App.V.setup = {
    bare: true, title: 'ตั้งค่าครั้งแรก',
    render() {
      return `<div class="login"><div class="login-hero" style="background-image:url(img/meet-kinnaree.jpg)"><h2>ตั้งค่าระบบ<br>ฐานข้อมูลกลาง</h2><p style="margin:0;opacity:.9">สร้างบัญชีผู้ดูแลระบบคนแรก แล้วเพิ่มผู้ใช้งานคนอื่นจากหน้า “ผู้ใช้งาน”</p></div>
  <div class="login-side"><form class="login-box" id="setupForm">
    <div class="brand"><img src="img/logo-wide-sm.png" alt="Luck Swan Resort" style="width:116px;height:48px"></div>
    <h1>ตั้งค่าครั้งแรก</h1><div class="muted" style="margin-bottom:22px">ทำครั้งเดียว — หลังจากนี้หน้านี้จะไม่แสดงอีก</div>
    <div class="card pad">
      <div class="field"><label class="f" for="suN">ชื่อ-นามสกุล ผู้ดูแลระบบ</label><input class="input" id="suN" name="name" required></div>
      <div class="field"><label class="f" for="suE">อีเมล</label><input class="input" id="suE" name="email" type="email" required autocomplete="username" placeholder="name@luckswangroup.com"></div>
      <div class="field"><label class="f" for="suP">รหัสผ่าน (อย่างน้อย 8 ตัวอักษร)</label><input class="input" id="suP" name="password" type="password" required minlength="8" autocomplete="new-password"></div>
      <div class="field"><label class="f" for="suP2">ยืนยันรหัสผ่าน</label><input class="input" id="suP2" name="password2" type="password" required autocomplete="new-password"></div>
      <label class="row small" style="margin:16px 0">${App.sw('ข้อมูลตัวอย่าง', true, 'name="sample"')} ใส่ข้อมูลตัวอย่าง (ใบเสนอ ใบฟังก์ชัน พนักงาน) — ลบหรือแก้ทีหลังได้</label>
      <div class="small red" id="suErr" style="min-height:20px" role="alert"></div>
      <button class="btn red block" type="submit">สร้างผู้ดูแลระบบ ${ICON.arrow}</button>
    </div></form></div></div>`;
    },
    after() {
      const f = document.getElementById('setupForm');
      f.onsubmit = async e => {
        e.preventDefault();
        const d = App.formData(f), err = document.getElementById('suErr');
        if (d.password.length < 8) { err.textContent = 'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร'; return; }
        if (d.password !== d.password2) { err.textContent = 'รหัสผ่านไม่ตรงกัน'; return; }
        let sample = null;
        if (d.sample) { const s = Store.seedData(); sample = { quotes: s.quotes, beos: s.beos, staff: s.staff, employees: s.employees, settings: s.settings }; }
        try {
          const u = await g.API.setupAdmin(d, sample);
          App.state.uid = u.id; localStorage.setItem('lsSales.cloudUid', u.id);
          App.go('#/home'); App.toast('ตั้งค่าเรียบร้อย — เพิ่มผู้ใช้งานได้ที่หน้า “ผู้ใช้งาน”');
        } catch (x) { err.textContent = x.message; }
      };
    }
  };

  App.ACT.togglePw = el => { const i = document.getElementById('lgP'); i.type = i.type === 'password' ? 'text' : 'password'; el.textContent = i.type === 'password' ? 'แสดง' : 'ซ่อน'; };
  App.ACT.forgotPw = () => App.modal(`<h2>ลืมรหัสผ่าน</h2><p class="muted">กรุณาติดต่อผู้ดูแลระบบหรือผู้จัดการฝ่ายขาย เพื่อรีเซ็ตรหัสผ่านจากหน้า “ผู้ใช้งาน”<br>โทร 053-750-855</p><div class="acts"><button class="btn dark" data-act="closeModal">ตกลง</button></div>`);

  /* =====================================================================
     HOME (1e)
     ===================================================================== */
  App.V.home = {
    title: 'หน้าหลัก',
    render() {
      const db = Store.db, me = App.me(), now = new Date(), today = U.today();
      const hr = now.getHours();
      const greet = hr < 12 ? 'สวัสดีตอนเช้า' : hr < 17 ? 'สวัสดีตอนบ่าย' : 'สวัสดีตอนเย็น';
      const quotes = db.quotes.slice().sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
      const ym = today.slice(0, 7);
      const closed = db.quotes.filter(q => q.status === 'confirmed' && q.start.slice(0, 7) === ym).reduce((a, q) => a + App.calc(q).net, 0);
      const target = db.settings.target || 1;
      const pct = Math.round(closed / target * 100);
      const pkgFrom = Math.min(...D.FOOD.filter(f => f.cat === 'package').map(f => f.price));
      const fu = App.V.home.followUps();
      const fuCard = `<section class="card pad">
      <h2 style="margin-bottom:14px">ต้องติดตามวันนี้</h2>
      ${fu.length ? fu.slice(0, 5).map(f => `<div class="fu"><i class="dot" style="background:${f.color}"></i><div style="flex:1"><b style="font-weight:600;font-size:13.5px">${H(f.title)}</b><div class="small muted">${H(f.sub)}</div></div>${f.action}</div>`).join('') : '<div class="muted small">ไม่มีงานค้าง 🎉</div>'}
    </section>`;
      // the full sales home is for admins, sales managers and sales staff (and roles added on those levels);
      // everyone else gets just their follow-ups
      if (!App.V.home.salesHome(me)) return `
<div class="page-hd"><div><div class="eyebrow">${U.DL[now.getDay()]} ${U.fmtD(today)}</div><h1>${greet} คุณ${H(App.firstName(me))}</h1></div></div>
<div class="stack" style="max-width:720px">${App.homeTasks ? App.homeTasks() : ''}${fuCard}</div>`;
      return `
<div class="page-hd"><div><div class="eyebrow">${U.DL[now.getDay()]} ${U.fmtD(today)}</div><h1>${greet} คุณ${H(App.firstName(me))}</h1></div>
  <a class="btn red" href="#/new">+ ใบเสนอราคาใหม่</a></div>
<div class="tiles">
  <a class="tile teal" href="#/meet"><span class="k">${D.MEETING_ROOMS.length - 1} ห้อง + สระว่ายน้ำ · สูงสุด ${Math.max(...D.MEETING_ROOMS.map(r => Math.max(r.cap.cocktail[1], r.cap.theatre[1])))} ท่าน</span><h3>ห้องประชุม<br>และจัดงาน</h3><span class="go">${ICON.arrow}</span></a>
  <a class="tile yellow" href="#/rooms"><span class="k">${D.ROOM_TYPES.length} ประเภท · ${D.ROOM_TYPES.reduce((a, t) => a + t.total, 0)} ห้อง</span><h3>ห้องพัก<br>กลุ่ม</h3><span class="go">${ICON.arrow}</span></a>
  <a class="tile red" href="#/food"><span class="k">${D.FOOD.filter(f => f.cat === 'package').length} แพ็กเกจสัมมนา</span><h3>จัดเลี้ยง<br>และอาหาร</h3><span class="go">${ICON.arrow}</span></a>
  <a class="tile dark" href="#/food"><span class="k">เริ่มเร็วที่สุด</span><h3>แพ็กเกจสัมมนา<br>พร้อมห้องพัก</h3><span class="small" style="opacity:.8">Basic · Standard · Premium เริ่ม ${U.n(pkgFrom)} ฿ Net/ท่าน</span></a>
</div>
<div class="split">
  <section class="card" style="overflow:hidden">
    <div class="card-hd"><h2>ใบเสนอราคาล่าสุด</h2><a class="link" href="#/quotes">ดูทั้งหมด</a></div>
    <table class="tbl resp"><thead><tr><th>เลขที่</th><th>ลูกค้า / งาน</th><th>วันงาน</th><th>ท่าน</th><th class="r">ยอด (฿)</th><th class="r">สถานะ</th></tr></thead><tbody>
    ${quotes.slice(0, 6).map(q => `<tr class="click" data-act="openQuote" data-id="${H(q.id)}">
      <td class="mono small muted hide-m">${H(q.id)}</td>
      <td><b style="font-weight:500">${H(q.client)}</b><div class="sub">${H(q.title)}<span class="mono" style="display:none"> · ${H(q.id)}</span></div></td>
      <td class="hide-m">${U.fmtRange(q.start, q.end, { year: false })}</td><td class="hide-m num">${U.n(q.pax)}</td>
      <td class="r num b">${U.n(Math.round(App.calc(q).net))}<div class="sub show-m" style="font-weight:400">${U.fmtRange(q.start, q.end, { year: false })}</div></td>
      <td class="r">${App.statusBadge(q.status)}</td></tr>`).join('')}
    </tbody></table>
  </section>
  <aside class="stack">
    <section class="card pad">
      <div class="between"><h2>เป้าเดือน${U.ML[now.getMonth()]}</h2><span class="small muted">ปิดได้ ${pct}%</span></div>
      <div style="margin:12px 0 14px"><span style="font:700 30px var(--fh)">฿${U.short(closed)}</span> <span class="muted">/ ${U.short(target)}</span></div>
      <div class="bar"><i style="width:${U.clamp(pct, 0, 100)}%"></i></div>
    </section>
    ${App.homeTasks ? App.homeTasks() : ''}
    ${fuCard}
  </aside>
</div>`;
    },
    // the quote side of the home page: only the built-in admin, sales manager and sales roles —
    // roles added on top of those levels (e.g. an F&B staff role on the sales level) get the short home page
    salesHome(me) { return App.salesRole(me); },
    followUps() {
      const out = [], today = U.today(), me = App.me();
      for (const b of Store.db.beos) {
        const st = g.Approval ? g.Approval.state(b) : 'approved', q = Store.quote(b.quoteId), lnk = `<a class="link" href="#/beo/${encodeURIComponent(b.id)}">`;
        if (st === 'pending' && g.Approval.canDecide(b, me))
          out.push({ p: 0, color: 'var(--yellow)', title: `ใบฟังก์ชัน ${b.id} รออนุมัติ`, sub: `${q ? q.title + ' · ' : ''}แก้ไขโดย ${b.approval.requestedByName || '-'}`, action: lnk + 'ตรวจ</a>' });
        if (st === 'rejected' && b.approval.requestedBy === App.state.uid)
          out.push({ p: 0, color: 'var(--red)', title: `ใบฟังก์ชัน ${b.id} ถูกตีกลับ`, sub: b.approval.comment || '', action: lnk + 'แก้ไข</a>' });
      }
      for (const q of App.V.home.salesHome(me) ? Store.db.quotes : []) {
        const open = `<a class="link" href="#/quote/${encodeURIComponent(q.id)}">`;
        if (q.status === 'sent' && q.views > 0)
          out.push({ p: 1, color: 'var(--red)', title: `${q.client} เปิดอ่านใบเสนอ ${q.views} ครั้ง`, sub: 'ยังไม่ตอบ' + (q.holdUntil ? ' · ห้องกันไว้ถึง ' + U.fmtD(q.holdUntil, { year: false }) : ''), action: q.phone ? `<a class="link" href="${App.tel(q.phone)}">โทร</a>` : open + 'เปิด</a>' });
        if (q.followUp && q.status !== 'confirmed' && q.status !== 'cancelled')
          out.push({ p: 2, color: 'var(--yellow)', title: `${q.client}`, sub: q.followUp, action: open + 'แก้ใบเสนอ</a>' });
        if (q.status === 'tentative' && q.holdUntil && U.diffDays(today, q.holdUntil) <= 7 && U.diffDays(today, q.holdUntil) >= 0 && !q.followUp)
          out.push({ p: 2, color: 'var(--red)', title: `${q.client} จองชั่วคราวใกล้หมดอายุ`, sub: 'หมดอายุ ' + U.fmtD(q.holdUntil), action: open + 'ติดตาม</a>' });
        if (q.status === 'draft') {
          const age = U.diffDays((q.updatedAt || today).slice(0, 10), today);
          if (age >= 2) out.push({ p: 3, color: 'var(--muted2)', title: q.client, sub: `ร่างค้างไว้ ${age} วัน`, action: open + 'ทำต่อ</a>' });
        }
      }
      return out.sort((a, b) => a.p - b.p);
    }
  };
  App.ACT.openQuote = el => App.go('#/quote/' + encodeURIComponent(el.dataset.id));

  /* ---------- boot ---------- */
  async function boot() {
    const root = document.getElementById('app');
    if (root) root.innerHTML = '<div class="empty" style="padding-top:30vh">กำลังโหลด…</div>';
    const cloud = g.API ? await g.API.detect() : false;
    if (cloud) {
      try { const raw = localStorage.getItem(Store.KEY); Store.db = raw ? JSON.parse(raw) : null; } catch (e) { Store.db = null; }
      if (!Store.db) Store.db = { v: 1, settings: Store.seedData().settings, users: [], quotes: [], beos: [], employees: [], staff: {}, tasks: [], seq: {} };
      g.DATA.applyCatalog(Store.db.catalog && Store.db.catalog.food);
      g.API.rememberBeos();
      if (g.API.token && !g.API.setup) {
        try { App.state.uid = await g.API.loadState(); }
        catch (e) { if (e.network && Store.db.users.length) App.state.uid = localStorage.getItem('lsSales.cloudUid'); else g.API.clearToken(); }
      }
    } else { Store.load(); App.restoreSession(); }
    window.addEventListener('hashchange', () => { App.closeNavPop(); App.render(); });
    document.addEventListener('click', e => { if (document.querySelector('.nav-pop:not(.lang-pop):not(.theme-pop)') && !e.target.closest('.nav-pop,[data-act=navGrp]')) App.closeNavPop(); }, true);
    window.addEventListener('resize', () => App.closeNavPop());
    window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); App.deferredInstall = e; });
    if (!location.hash) location.replace(App.state.uid ? '#/home' : '#/login');
    App.render();
    if ('serviceWorker' in navigator && /^https?:/.test(location.protocol)) {
      // when a new version is published, the new service worker takes over and the page reloads once by itself
      const hadController = !!navigator.serviceWorker.controller;
      let reloaded = false;
      navigator.serviceWorker.addEventListener('controllerchange', () => { if (hadController && !reloaded) { reloaded = true; location.reload(); } });
      navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then(reg => {
        reg.update();
        // a page left open all day still picks up a new version: check on return, on focus and every few minutes
        document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') reg.update(); });
        window.addEventListener('focus', () => reg.update());
        setInterval(() => { if (document.visibilityState === 'visible') reg.update(); }, 5 * 60e3);
      }).catch(() => { /* offline support unavailable */ });
    }
  }
  g.App = App;
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => setTimeout(boot)); else setTimeout(boot);
})(window);
