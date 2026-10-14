/* Luck Swan Sales — cloud sync (Cloudflare Workers + D1) and function-sheet approval rules.
 * When the app is served by the worker, /api/health answers and the app runs in shared "online" mode.
 * Anywhere else (GitHub Pages, a local file) it keeps working on this device's own data. */
(function (g) {
  'use strict';
  const U = g.U, Store = g.Store;

  /* ---------------- approval rules (mirrors worker/src/index.js) ---------------- */
  const RANK = { sales: 1, acc: 1, head: 2, manager: 3, admin: 4 };
  const Approval = {
    RANK,
    content(b) {
      const { approval, status, rev, revNote, ...c } = b || {};
      c.depts = Object.fromEntries(Object.entries((b && b.depts) || {}).map(([k, v]) => [k, { items: v.items || [], owner: v.owner || '' }]));
      return c;
    },
    state(b) { return (b && b.approval && b.approval.state) || 'approved'; },
    rank(u) { return (u && RANK[u.role]) || 1; },
    // called when a function sheet was saved; prev = the version before this edit
    apply(b, prev, user) {
      const changed = !prev || JSON.stringify(this.content(prev)) !== JSON.stringify(this.content(b));
      if (!changed) { if (prev && prev.approval) b.approval = prev.approval; return; }
      const prevAppr = prev ? (prev.approval || { state: 'approved' }) : null;
      if (this.rank(user) >= RANK.manager) b.approval = { state: 'approved', by: user.id, byName: user.name, at: new Date().toISOString().slice(0, 19), auto: true };
      else b.approval = { state: 'pending', requestedBy: user.id, requestedByName: user.name, requestedAt: new Date().toISOString().slice(0, 19),
        base: prevAppr && prevAppr.state === 'approved' ? this.content(prev) : (prevAppr ? prevAppr.base || null : null) };
    },
    canDecide(b, me) {
      const a = b && b.approval;
      if (!a || a.state !== 'pending' || !me) return false;
      const req = Store.db.users.find(u => u.id === a.requestedBy);
      return this.rank(me) > (req ? this.rank(req) : 1);
    }
  };

  /* ---------------- sync ---------------- */
  const TOKEN = 'lsSales.token';
  const kindsOf = db => ({
    quote: db.quotes.map(q => [q.id, q]),
    beo: db.beos.map(b => [b.id, b]),
    staff: Object.entries(db.staff || {}),
    employees: [['all', { list: db.employees || [] }]],
    settings: [['main', db.settings || {}]],
    catalog: db.catalog && Array.isArray(db.catalog.food) ? [['food', { list: db.catalog.food }]] : []
  });
  const userView = u => JSON.stringify({ name: u.name, email: u.email, role: u.role, dept: u.dept, phone: u.phone || '', status: u.status, twoFA: !!u.twoFA, perms: u.perms || {} });

  const API = {
    online: false, setup: false, token: null, serverTime: null, offline: false,
    ver: {}, snap: {}, userSnap: {}, beoPrev: {}, pushing: false, again: false, timer: null, poller: null,

    async detect() {
      if (!/^https?:/.test(location.protocol) || /\.github\.io$/.test(location.hostname)) return false;
      try {
        const r = await fetch('api/health', { cache: 'no-store' });
        if (!r.ok || !/json/.test(r.headers.get('content-type') || '')) return false;
        const j = await r.json();
        this.online = !!j.ok; this.setup = !!j.setup;
      } catch (e) { this.online = false; }
      if (this.online) {
        Store.KEY = 'lsSales.cloud';
        try { this.token = localStorage.getItem(TOKEN) || sessionStorage.getItem(TOKEN); } catch (e) { /* ignore */ }
      }
      return this.online;
    },
    async req(method, path, body) {
      const headers = { 'content-type': 'application/json', 'x-device': g.App ? encodeURIComponent(g.App.deviceName()) : '' };
      if (this.token) headers.authorization = 'Bearer ' + this.token;
      let r;
      try { r = await fetch('api/' + path, { method, headers, body: body === undefined ? undefined : JSON.stringify(body), cache: 'no-store' }); }
      catch (e) { this.setOffline(true); const err = new Error('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้'); err.network = true; throw err; }
      this.setOffline(false);
      const j = await r.json().catch(() => ({}));
      if (r.status === 401 && path !== 'login') { this.expired(j.error); const err = new Error(j.error || 'กรุณาเข้าสู่ระบบ'); err.status = 401; throw err; }
      if (!r.ok) { const err = new Error(j.message || j.error || 'เกิดข้อผิดพลาด (' + r.status + ')'); err.status = r.status; err.body = j; throw err; }
      return j;
    },
    setOffline(v) {
      if (this.offline === v) return;
      this.offline = v;
      if (g.App) g.App.toast(v ? 'ออฟไลน์ — การแก้ไขจะถูกส่งเมื่อกลับมาออนไลน์' : 'กลับมาออนไลน์แล้ว');
      if (!v) this.queuePush();
    },
    saveToken(token, remember) {
      this.token = token;
      try { localStorage.removeItem(TOKEN); sessionStorage.removeItem(TOKEN); (remember ? localStorage : sessionStorage).setItem(TOKEN, token); } catch (e) { /* ignore */ }
    },
    clearToken() { this.token = null; try { localStorage.removeItem(TOKEN); sessionStorage.removeItem(TOKEN); localStorage.removeItem(Store.KEY); } catch (e) { /* ignore */ } },
    expired(msg) {
      this.clearToken(); this.stopPolling();
      if (g.App && g.App.state.uid) { g.App.state.uid = null; g.App.toast(msg || 'กรุณาเข้าสู่ระบบใหม่'); g.App.go('#/login'); }
    },

    async login(ident, pw, remember) {
      const j = await this.req('POST', 'login', { email: ident, password: pw, remember: !!remember });
      this.saveToken(j.token, remember);
      await this.loadState();
      return j.user;
    },
    async setupAdmin(form, sample) {
      const j = await this.req('POST', 'setup', { name: form.name, email: form.email, password: form.password, sample });
      this.saveToken(j.token, true); this.setup = false;
      await this.loadState();
      return j.user;
    },
    async logout() { try { await this.req('POST', 'logout', {}); } catch (e) { /* ignore */ } this.clearToken(); this.stopPolling(); },

    /* full download */
    async loadState() {
      const j = await this.req('GET', 'state');
      const seed = Store.seedData();
      const db = { v: 1, settings: seed.settings, users: [], quotes: [], beos: [], employees: [], staff: {}, seq: {} };
      this.ver = {}; this.snap = {}; this.userSnap = {};
      Store.db = db;
      g.DATA.applyCatalog(null); // built-in menu unless the shared database has an edited one
      for (const d of j.docs) this.applyDoc(d, true);
      for (const u of j.users) this.applyUser(u);
      this.serverTime = j.serverTime;
      this.rememberBeos();
      Store.write();
      this.startPolling();
      return j.me;
    },
    rememberBeos() { this.beoPrev = {}; for (const b of Store.db.beos) this.beoPrev[b.id] = JSON.stringify(b); },

    applyDoc(d, fresh) {
      const key = d.kind + ':' + d.id, db = Store.db;
      if (d.deleted) {
        if (d.kind === 'quote') db.quotes = db.quotes.filter(x => x.id !== d.id);
        if (d.kind === 'beo') db.beos = db.beos.filter(x => x.id !== d.id);
        if (d.kind === 'staff') delete db.staff[d.id];
        delete this.ver[key]; delete this.snap[key]; return;
      }
      const data = d.data;
      const replace = (arr) => { const i = arr.findIndex(x => x.id === d.id); if (i >= 0) arr[i] = data; else arr.push(data); };
      if (d.kind === 'quote') replace(db.quotes);
      else if (d.kind === 'beo') { replace(db.beos); this.beoPrev[d.id] = JSON.stringify(data); }
      else if (d.kind === 'staff') db.staff[d.id] = data;
      else if (d.kind === 'employees') db.employees = data.list || [];
      else if (d.kind === 'settings') db.settings = Object.assign({}, Store.seedData().settings, data);
      else if (d.kind === 'catalog' && d.id === 'food') { db.catalog = { food: data.list || [] }; g.DATA.applyCatalog(db.catalog.food); }
      this.ver[key] = d.version;
      this.snap[key] = JSON.stringify(d.kind === 'employees' ? { list: db.employees } : d.kind === 'settings' ? db.settings : d.kind === 'catalog' ? { list: db.catalog.food } : data);
    },
    applyUser(u) {
      const db = Store.db, i = db.users.findIndex(x => x.id === u.id);
      if (i >= 0) db.users[i] = u; else db.users.push(u);
      this.userSnap[u.id] = userView(u);
    },

    /* upload what changed since the last sync */
    queuePush() {
      if (!this.online || !this.token) return;
      clearTimeout(this.timer);
      this.timer = setTimeout(() => this.push(), 350);
    },
    dirty() {
      const out = [], seen = new Set();
      for (const [kind, list] of Object.entries(kindsOf(Store.db))) for (const [id, data] of list) {
        const key = kind + ':' + id; seen.add(key);
        const s = JSON.stringify(data);
        if (s !== this.snap[key]) out.push({ kind, id, data, key });
      }
      for (const key of Object.keys(this.snap)) if (!seen.has(key) && !/^(catalog|employees|settings):/.test(key)) { const [kind, ...rest] = key.split(':'); out.push({ kind, id: rest.join(':'), del: true, key }); }
      const users = Store.db.users.filter(u => userView(u) !== this.userSnap[u.id]);
      return { docs: out, users };
    },
    async push() {
      if (this.pushing) { this.again = true; return; }
      this.pushing = true;
      let rerender = false;
      try {
        const { docs, users } = this.dirty();
        for (const d of docs) {
          try {
            if (d.del) { await this.req('DELETE', 'doc/' + d.kind + '/' + encodeURIComponent(d.id)); delete this.snap[d.key]; delete this.ver[d.key]; continue; }
            const j = await this.req('PUT', 'doc/' + d.kind + '/' + encodeURIComponent(d.id), { data: d.data, version: this.ver[d.key] || 0 });
            const before = JSON.stringify(d.data);
            this.applyDoc(j.doc);
            if (JSON.stringify(j.doc.data) !== before) rerender = true;
          } catch (e) {
            if (e.network || e.status === 401) throw e;
            rerender = true;
            if (e.status === 409 && e.body && 'doc' in e.body) {
              if (e.body.doc) this.applyDoc(e.body.doc); else this.dropLocal(d);
            } else this.revert(d);
            if (g.App) g.App.toast(e.message);
          }
        }
        for (const u of users) {
          try {
            const j = await this.req('PUT', 'users/' + encodeURIComponent(u.id), { data: u, version: u._v || 0 });
            this.applyUser(j.user); rerender = true;
          } catch (e) {
            if (e.network || e.status === 401) throw e;
            const i = Store.db.users.indexOf(u); if (this.userSnap[u.id] && i >= 0) Object.assign(u, JSON.parse(this.userSnap[u.id])); else if (i >= 0) Store.db.users.splice(i, 1);
            rerender = true; if (g.App) g.App.toast(e.message);
          }
        }
        Store.write();
      } catch (e) { /* offline or signed out: retry later */ }
      finally {
        this.pushing = false;
        if (rerender && g.App && g.App.state.uid) g.App.softRender();
        if (this.again) { this.again = false; this.queuePush(); }
      }
    },
    revert(d) { if (this.snap[d.key]) this.applyDoc({ kind: d.kind, id: d.id, version: this.ver[d.key], data: JSON.parse(this.snap[d.key]) }); else this.dropLocal(d); },
    dropLocal(d) { this.applyDoc({ kind: d.kind, id: d.id, deleted: true }); },

    /* pull other people's changes */
    startPolling() {
      this.stopPolling();
      this.poller = setInterval(() => this.pull(), 30000);
      if (!this._vis) { this._vis = true; document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && this.online && this.token) this.pull(); }); }
    },
    stopPolling() { clearInterval(this.poller); this.poller = null; },
    async pull() {
      if (!this.token || this.pushing) return;
      try {
        const j = await this.req('GET', 'state?since=' + encodeURIComponent(this.serverTime || ''));
        const { docs: dirty } = this.dirty();
        const busy = new Set(dirty.map(d => d.key));
        let changed = false;
        for (const d of j.docs) {
          const key = d.kind + ':' + d.id;
          if (busy.has(key) || (this.ver[key] && this.ver[key] >= d.version && !d.deleted)) continue;
          this.applyDoc(d); changed = true;
        }
        for (const u of j.users) if (!this.userSnap[u.id] || u._v !== (Store.db.users.find(x => x.id === u.id) || {})._v) { this.applyUser(u); changed = true; }
        this.serverTime = j.serverTime;
        if (changed) { Store.write(); if (g.App) g.App.softRender(); }
      } catch (e) { /* retry next tick */ }
    },

    /* hooks used by Store.save() */
    beforeSave() {
      const me = g.App && g.App.me();
      if (!me) return;
      for (const b of Store.db.beos) {
        const prevS = this.beoPrev[b.id], cur = JSON.stringify(b);
        if (prevS === cur) continue;
        Approval.apply(b, prevS ? JSON.parse(prevS) : null, me);
        this.beoPrev[b.id] = JSON.stringify(b);
      }
    }
  };

  // wrap Store.save so every edit goes through the approval rules and, when online, to the server
  Store.save = function () { API.beforeSave(); Store.write(); API.queuePush(); };

  g.API = API;
  g.Approval = Approval;
})(window);
