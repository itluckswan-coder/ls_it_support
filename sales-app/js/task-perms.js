/* Luck Swan Sales — who may assign work to whom (Assignment & Approve ⚙). The admin sets it one of two ways:
 *  - org chart (แบบผังองค์กร): every role and department sits under a role; a role assigns to everything below it
 *  - form (แบบฟอร์มสำหรับเลือก): for each role, tick the roles and departments it may assign to
 * Saved in settings.taskPerms = { mode: 'org' | 'form', parent: { node: parentNode }, rules: { roleId: { roles: [], depts: [] } } }
 * where a node is 'role:<id>' or 'dept:<id>'. The admin always assigns to everyone. The worker checks the same rules. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  const TP = () => (Store.db.settings && Store.db.settings.taskPerms) || {};
  const roleList = () => App.roles ? App.roles() : D.ROLES.map(r => ({ id: r.id, name: r.name, base: r.id }));
  const roleById = id => roleList().find(r => r.id === id);
  const BASE_PARENT = { manager: 'admin', sales: 'manager', head: 'admin', acc: 'admin' };
  // where a node sits when the admin hasn't moved it: sales under the sales manager, the rest under the admin
  function defParent(node) {
    if (node === 'role:admin') return null;
    if (node.startsWith('dept:')) return 'role:admin';
    const r = roleById(node.slice(5));
    if (r && r.custom) return 'role:' + (BASE_PARENT[r.base] || 'admin');
    return 'role:' + (BASE_PARENT[r ? r.base : node.slice(5)] || 'admin');
  }
  const nodes = () => roleList().map(r => 'role:' + r.id).concat(D.DEPTS.map(d => 'dept:' + d.id));
  const parentIn = (map, n) => n === 'role:admin' ? null : (map[n] && map[n].startsWith('role:') && map[n] !== n ? map[n] : defParent(n));
  // is `node` somewhere below `top` in the chart (walking up its parents; a loop counts as not below)
  function below(map, node, top) {
    const seen = new Set([node]);
    for (let p = parentIn(map, node); p; p = parentIn(map, p)) { if (p === top) return true; if (seen.has(p)) return false; seen.add(p); }
    return false;
  }

  // what `me` may assign to: { all } for the admin, else the role ids and department ids allowed
  App.taskScope = function (me) {
    me = me === undefined ? App.me() : me;
    if (!me) return { roles: new Set(), depts: new Set(), any: false };
    if (me.role === 'admin') return { all: true, any: true };
    const tp = TP(), key = App.roleOf ? App.roleOf(me).id : me.role;
    let roles, depts;
    if (tp.mode === 'org') {
      const map = tp.parent || {}, top = 'role:' + key;
      roles = new Set(roleList().filter(r => below(map, 'role:' + r.id, top)).map(r => r.id));
      depts = new Set(D.DEPTS.filter(d => below(map, 'dept:' + d.id, top)).map(d => d.id));
    } else {
      const r = (tp.rules || {})[key] || {};
      roles = new Set((r.roles || []).filter(x => x !== 'admin')); depts = new Set(r.depts || []);
    }
    return { roles, depts, any: roles.size > 0 || depts.size > 0 };
  };
  App.canAssignTo = function (to, me) {
    const s = App.taskScope(me); if (s.all) return true;
    if (to.type === 'dept') return s.depts.has(to.dept);
    const u = Store.db.users.find(x => x.id === to.id);
    return !!u && s.roles.has(App.roleOf ? App.roleOf(u).id : u.role);
  };

  /* ---------------- the settings window (admin) ---------------- */
  const GEAR = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12.2 2h-.4a2 2 0 0 0-2 2v.2a2 2 0 0 1-1 1.7l-.4.3a2 2 0 0 1-2 0l-.2-.1a2 2 0 0 0-2.7.7l-.2.4a2 2 0 0 0 .7 2.7l.2.1a2 2 0 0 1 1 1.7v.5a2 2 0 0 1-1 1.7l-.2.1a2 2 0 0 0-.7 2.7l.2.4a2 2 0 0 0 2.7.7l.2-.1a2 2 0 0 1 2 0l.4.3a2 2 0 0 1 1 1.7v.2a2 2 0 0 0 2 2h.4a2 2 0 0 0 2-2v-.2a2 2 0 0 1 1-1.7l.4-.3a2 2 0 0 1 2 0l.2.1a2 2 0 0 0 2.7-.7l.2-.4a2 2 0 0 0-.7-2.7l-.2-.1a2 2 0 0 1-1-1.7v-.5a2 2 0 0 1 1-1.7l.2-.1a2 2 0 0 0 .7-2.7l-.2-.4a2 2 0 0 0-2.7-.7l-.2.1a2 2 0 0 1-2 0l-.4-.3a2 2 0 0 1-1-1.7V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/></svg>';
  App.taskPermsButton = () => App.me() && App.me().role === 'admin' ? `<button class="icon-btn tp-gear" data-act="tpOpen" title="ตั้งค่าสิทธิ์การสั่งงาน" aria-label="ตั้งค่าสิทธิ์การสั่งงาน">${GEAR}</button>` : '';

  App.ACT.tpOpen = function () {
    if (!App.me() || App.me().role !== 'admin') return;
    const cur = TP(), st0 = Store.db.settings || {}, rs0 = st0.roles || {};
    const draft = {
      mode: cur.mode === 'org' ? 'org' : 'form', parent: Object.assign({}, cur.parent), rules: JSON.parse(JSON.stringify(cur.rules || {})),
      // roles and departments can be renamed, added and removed on the chart; they are written back on save
      roles: { names: Object.assign({}, rs0.names), custom: JSON.parse(JSON.stringify(rs0.custom || [])) },
      depts: D.DEPTS.map(d => ({ id: d.id, name: d.name, short: d.short })), deptsChanged: false
    };
    const ui = { zoom: 1, edit: null, add: null, addType: 'role', del: null, msg: '' };
    const dRoles = () => D.ROLES.map(r => ({ id: r.id, name: String(draft.roles.names[r.id] || r.name), base: r.id, custom: false }))
      .concat(draft.roles.custom.map(c => ({ id: c.id, name: String(c.name || ''), base: c.base, custom: true })));
    const dNodes = () => dRoles().map(r => 'role:' + r.id).concat(draft.depts.map(d => 'dept:' + d.id));
    const nodeName = n => { if (n.startsWith('dept:')) { const d = draft.depts.find(x => x.id === n.slice(5)); return 'แผนก' + (d ? d.short : n.slice(5)); } const r = dRoles().find(x => x.id === n.slice(5)); return r ? r.name : n.slice(5); };
    const rawName = n => n.startsWith('dept:') ? ((draft.depts.find(x => x.id === n.slice(5)) || {}).name || '') : ((dRoles().find(x => x.id === n.slice(5)) || {}).name || '');
    const others = () => dRoles().filter(r => r.id !== 'admin');
    const builtInDept = id => D.DEPTS_DEFAULT.some(x => x.id === id);
    const usedBy = n => n.startsWith('dept:') ? Store.db.employees.filter(e => e.dept === n.slice(5)).length : Store.db.users.filter(u => u.roleId === n.slice(5)).length;
    const canDel = n => n.startsWith('dept:') ? !builtInDept(n.slice(5)) : dRoles().some(r => 'role:' + r.id === n && r.custom);
    const LEVELS = D.ROLES.filter(r => r.id !== 'admin');
    const bg = App.modal('<div id="tpBody"></div>'), modal = bg.querySelector('.modal');
    modal.style.maxWidth = '1100px';

    function chart() {
      const map = draft.parent, all = dNodes(), kids = n => all.filter(x => parentIn(map, x) === n);
      const tool = (act, n, label, icon, cls = '') => `<button type="button" class="oc-tool ${cls}" data-tp-act="${act}" data-n="${H(n)}" title="${H(label)}" aria-label="${H(label)}">${icon}</button>`;
      const card = n => {
        const opts = dRoles().filter(r => 'role:' + r.id !== n && !below(map, 'role:' + r.id, n)), isRole = n.startsWith('role:');
        const head = ui.edit === n
          ? `<div class="oc-edit"><input class="input sm" id="tpRen" value="${H(rawName(n))}" maxlength="40" aria-label="ชื่อใหม่"><div class="row" style="gap:4px;justify-content:center"><button type="button" class="btn red sm" data-tp-act="renOk" data-n="${H(n)}">บันทึก</button><button type="button" class="btn ghost sm" data-tp-act="cancel">ยกเลิก</button></div></div>`
          : `<div class="oc-tools">${tool('edit', n, 'แก้ไขชื่อ', '✏️')}${isRole ? tool('add', n, 'เพิ่มบทบาทหรือแผนกใต้กล่องนี้', '＋') : ''}${canDel(n) ? tool('del', n, 'ลบ', '🗑', 'red') : ''}</div><b>${H(nodeName(n))}</b>`;
        const del = ui.del === n ? `<div class="oc-ask"><span class="xs">${usedBy(n) ? (n.startsWith('dept:') ? `มีพนักงาน ${usedBy(n)} คนในแผนกนี้ — ย้ายพนักงานก่อนจึงจะลบได้` : `มีผู้ใช้ ${usedBy(n)} คนในบทบาทนี้ — เปลี่ยนบทบาทของผู้ใช้ก่อนจึงจะลบได้`) : 'ลบ ' + H(nodeName(n)) + '?'}</span>
          <div class="row" style="gap:4px;justify-content:center">${usedBy(n) ? '' : `<button type="button" class="btn red sm" data-tp-act="delOk" data-n="${H(n)}">ลบ</button>`}<button type="button" class="btn ghost sm" data-tp-act="cancel">${usedBy(n) ? 'ปิด' : 'ยกเลิก'}</button></div></div>` : '';
        const add = ui.add === n ? `<div class="oc-add"><div class="seg white" style="padding:3px"><button type="button" class="${ui.addType === 'role' ? 'on' : ''}" data-tp-act="addType" data-v="role">บทบาท</button><button type="button" class="${ui.addType === 'dept' ? 'on' : ''}" data-tp-act="addType" data-v="dept">แผนก</button></div>
          <input class="input sm" id="tpNew" maxlength="40" placeholder="${ui.addType === 'role' ? 'ชื่อบทบาทใหม่' : 'ชื่อแผนกใหม่'}" aria-label="${ui.addType === 'role' ? 'ชื่อบทบาทใหม่' : 'ชื่อแผนกใหม่'}">
          ${ui.addType === 'role' ? `<label class="xs muted">ระดับสิทธิ์ของระบบ<select class="input sm" id="tpBase">${LEVELS.map(r => `<option value="${r.id}">เทียบเท่า${H(r.name)}</option>`).join('')}</select></label>` : ''}
          <div class="row" style="gap:4px;justify-content:center"><button type="button" class="btn red sm" data-tp-act="addOk" data-n="${H(n)}">เพิ่ม</button><button type="button" class="btn ghost sm" data-tp-act="cancel">ยกเลิก</button></div></div>` : '';
        return `<div class="oc-card ${isRole ? 'role' : 'dept'}${n === 'role:admin' ? ' root' : ''}">${head}
          ${n === 'role:admin' ? '<span class="xs muted">สั่งงานได้ทุกคน</span>' : `<label class="xs muted">อยู่ใต้<select class="input sm" data-tp-par="${H(n)}">${opts.map(r => `<option value="role:${H(r.id)}" ${parentIn(map, n) === 'role:' + r.id ? 'selected' : ''}>${H(r.name)}</option>`).join('')}</select></label>`}${del}${add}</div>`;
      };
      const tree = n => { const k = kids(n); return `<li>${card(n)}${k.length ? `<ul>${k.map(tree).join('')}</ul>` : ''}</li>`; };
      return `<p class="small muted" style="margin:0 0 10px">บทบาทที่อยู่ด้านบนสั่งงานได้ทุกบทบาทและทุกแผนกที่อยู่ใต้สาย · เปลี่ยน “อยู่ใต้” เพื่อย้ายตำแหน่งในผัง · ✏️ แก้ไขชื่อ ＋ เพิ่ม 🗑 ลบ</p>
        <div class="oc-zoom no-print"><button type="button" class="icon-btn" data-tp-zoom="-" title="ย่อ" aria-label="ย่อ">−</button><span class="small num" id="tpZoomV">${Math.round(ui.zoom * 100)}%</span><button type="button" class="icon-btn" data-tp-zoom="+" title="ขยาย" aria-label="ขยาย">＋</button><button type="button" class="btn ghost sm" data-tp-zoom="fit">พอดีจอ</button><button type="button" class="btn ghost sm" data-tp-zoom="1">100%</button></div>
        <div class="xs ${/^!/.test(ui.msg) ? 'red' : 'teal'}" style="min-height:16px;margin:4px 0" role="status">${H(ui.msg.replace(/^!/, ''))}</div>
        <div class="oc-wrap" id="tpWrap"><ul class="oc" id="tpChart" style="zoom:${ui.zoom}">${tree('role:admin')}</ul></div>`;
    }
    function form() {
      const chip = (on, attrs, label) => `<button type="button" class="chip ${on ? 'on' : ''}" ${attrs}>${H(label)}</button>`;
      return `<p class="small muted" style="margin:0 0 10px">เลือกว่าแต่ละบทบาทสั่งงานให้บทบาทไหน และแผนกไหนได้ (ผู้ดูแลระบบสั่งงานได้ทุกคนเสมอ)</p>
        <div class="tp-rows">${others().map(r => { const x = draft.rules[r.id] || {}, rs = x.roles || [], ds = x.depts || []; return `<div class="tp-row">
          <div class="tp-who"><b>${H(r.name)}</b><span class="xs muted">สั่งงานให้ได้</span></div>
          <div><div class="xs muted" style="margin-bottom:4px">ผู้ใช้งานบทบาท</div><div class="row wrap" style="gap:6px">${others().map(t => chip(rs.includes(t.id), `data-tp-role="${H(r.id)}" data-v="${H(t.id)}"`, t.name)).join('')}</div>
            <div class="xs muted" style="margin:8px 0 4px">แผนก</div><div class="row wrap" style="gap:6px">${draft.depts.map(d => chip(ds.includes(d.id), `data-tp-dept="${H(r.id)}" data-v="${H(d.id)}"`, d.short)).join('')}</div></div>
        </div>`; }).join('')}</div>`;
    }
    function draw() {
      const y = modal.scrollTop, w = modal.querySelector('#tpWrap'), sx = w ? w.scrollLeft : 0, sy = w ? w.scrollTop : 0;
      modal.querySelector('#tpBody').innerHTML = `<h2>ตั้งค่าสิทธิ์การสั่งงาน</h2>
        <div class="seg white" role="tablist" style="margin-bottom:6px"><button type="button" class="${draft.mode === 'org' ? 'on' : ''}" data-tp-mode="org">แบบผังองค์กร</button><button type="button" class="${draft.mode === 'form' ? 'on' : ''}" data-tp-mode="form">แบบฟอร์มสำหรับเลือก</button></div>
        <div class="xs muted" style="margin-bottom:14px">กดบันทึกแล้วระบบจะใช้แบบที่เลือกอยู่</div>
        ${draft.mode === 'org' ? chart() : form()}
        <div class="acts" style="flex-wrap:wrap"><button type="button" class="btn ghost" data-tp-reset style="margin-right:auto">คืนค่าเริ่มต้น</button><button type="button" class="btn ghost" data-act="closeModal">ยกเลิก</button><button type="button" class="btn red" data-tp-save>บันทึก</button></div>`;
      modal.scrollTop = y;
      const w2 = modal.querySelector('#tpWrap'); if (w2) { w2.scrollLeft = sx; w2.scrollTop = sy; }
      const f = modal.querySelector('#tpRen') || modal.querySelector('#tpNew'); if (f) { f.focus(); if (f.id === 'tpRen') f.select(); }
    }
    const setZoom = z => { ui.zoom = Math.max(0.3, Math.min(1.6, Math.round(z * 100) / 100)); const c = modal.querySelector('#tpChart'), v = modal.querySelector('#tpZoomV'); if (c) c.style.zoom = ui.zoom; if (v) v.textContent = Math.round(ui.zoom * 100) + '%'; };
    const clean = v => String(v || '').trim().replace(/\s+/g, ' ');
    const nameTaken = (n, v) => n.startsWith('dept:') ? draft.depts.some(d => 'dept:' + d.id !== n && d.name === v) : dRoles().some(r => 'role:' + r.id !== n && r.name === v);
    function act(a, n, b) {
      if (a === 'cancel') { ui.edit = ui.add = ui.del = null; ui.msg = ''; return draw(); }
      if (a === 'edit') { ui.edit = n; ui.add = ui.del = null; ui.msg = ''; return draw(); }
      if (a === 'add') { ui.add = n; ui.edit = ui.del = null; ui.msg = ''; return draw(); }
      if (a === 'del') { ui.del = n; ui.edit = ui.add = null; ui.msg = ''; return draw(); }
      if (a === 'addType') { ui.addType = b.dataset.v; return draw(); }
      if (a === 'renOk') {
        const v = clean(modal.querySelector('#tpRen').value);
        if (!v) { ui.msg = '!กรุณากรอกชื่อ'; return draw(); }
        if (nameTaken(n, v)) { ui.msg = '!มีชื่อนี้อยู่แล้ว'; return draw(); }
        if (n.startsWith('dept:')) { const d = draft.depts.find(x => 'dept:' + x.id === n); d.name = d.short = v; draft.deptsChanged = true; }
        else { const id = n.slice(5), c = draft.roles.custom.find(x => x.id === id); if (c) c.name = v; else draft.roles.names[id] = v; }
        ui.edit = null; ui.msg = 'แก้ชื่อเป็น ' + v + ' แล้ว — กดบันทึกเพื่อใช้งาน'; return draw();
      }
      if (a === 'addOk') {
        const v = clean(modal.querySelector('#tpNew').value), type = ui.addType;
        if (!v) { ui.msg = '!กรุณากรอกชื่อ'; return draw(); }
        if (nameTaken(type + ':', v)) { ui.msg = '!มีชื่อนี้อยู่แล้ว'; return draw(); }
        let id;
        if (type === 'dept') { do id = 'd' + Math.random().toString(36).slice(2, 8); while (draft.depts.some(d => d.id === id)); draft.depts.push({ id, name: v, short: v }); draft.deptsChanged = true; }
        else { const base = modal.querySelector('#tpBase').value; id = 'r_' + Math.random().toString(36).slice(2, 10); draft.roles.custom.push({ id, name: v, base, perms: (D.ROLE_PERMS[base] || []).slice() }); }
        draft.parent[type + ':' + id] = n;
        ui.add = null; ui.msg = 'เพิ่ม' + (type === 'dept' ? 'แผนก ' : 'บทบาท ') + v + ' ใต้ ' + nodeName(n) + ' แล้ว — กดบันทึกเพื่อใช้งาน'; return draw();
      }
      if (a === 'delOk') {
        if (!canDel(n) || usedBy(n)) return;
        const name = nodeName(n), up = parentIn(draft.parent, n);
        // what sat under it moves up to its parent
        for (const x of dNodes()) if (parentIn(draft.parent, x) === n) draft.parent[x] = up;
        delete draft.parent[n];
        if (n.startsWith('dept:')) { draft.depts = draft.depts.filter(d => 'dept:' + d.id !== n); draft.deptsChanged = true; for (const r of Object.values(draft.rules)) r.depts = (r.depts || []).filter(x => 'dept:' + x !== n); }
        else { const id = n.slice(5); draft.roles.custom = draft.roles.custom.filter(c => c.id !== id); delete draft.rules[id]; for (const r of Object.values(draft.rules)) r.roles = (r.roles || []).filter(x => x !== id); }
        ui.del = null; ui.msg = 'ลบ ' + name + ' แล้ว — กดบันทึกเพื่อใช้งาน'; return draw();
      }
    }
    modal.addEventListener('change', e => { const s = e.target.closest('[data-tp-par]'); if (s) { draft.parent[s.dataset.tpPar] = s.value; draw(); } });
    modal.addEventListener('keydown', e => { if (e.key !== 'Enter') return; const t = e.target.id; if (t === 'tpRen' || t === 'tpNew') { e.preventDefault(); const b = modal.querySelector(t === 'tpRen' ? '[data-tp-act="renOk"]' : '[data-tp-act="addOk"]'); if (b) b.click(); } });
    // zoom: buttons, Ctrl/⌘ + wheel, and pinching on a trackpad
    modal.addEventListener('wheel', e => { if (!(e.ctrlKey || e.metaKey) || !e.target.closest('#tpWrap')) return; e.preventDefault(); setZoom(ui.zoom * (e.deltaY < 0 ? 1.1 : 0.9)); }, { passive: false });
    modal.addEventListener('click', e => {
      const z = e.target.closest('[data-tp-zoom]');
      if (z) {
        const v = z.dataset.tpZoom;
        if (v === '+') setZoom(ui.zoom + 0.1); else if (v === '-') setZoom(ui.zoom - 0.1); else if (v === '1') setZoom(1);
        else { const w = modal.querySelector('#tpWrap'), c = modal.querySelector('#tpChart'); if (w && c) { c.style.zoom = 1; const nat = Math.max(c.scrollWidth, c.getBoundingClientRect().width); setZoom(Math.min(1, Math.floor((w.clientWidth - 28) / (nat || 1) * 100) / 100)); for (let k = 0; k < 40 && w.scrollWidth > w.clientWidth + 1 && ui.zoom > 0.3; k++) setZoom(ui.zoom - 0.02); w.scrollLeft = 0; } }
        return;
      }
      const t = e.target.closest('[data-tp-act]'); if (t) { act(t.dataset.tpAct, t.dataset.n, t); return; }
      const b = e.target.closest('[data-tp-mode],[data-tp-role],[data-tp-dept],[data-tp-save],[data-tp-reset]'); if (!b) return;
      if (b.dataset.tpMode) { draft.mode = b.dataset.tpMode; ui.edit = ui.add = ui.del = null; ui.msg = ''; draw(); return; }
      const tog = (r, k, v) => { const x = draft.rules[r] || (draft.rules[r] = { roles: [], depts: [] }), l = x[k] || (x[k] = []), i = l.indexOf(v); if (i >= 0) l.splice(i, 1); else l.push(v); draw(); };
      if (b.dataset.tpRole) return tog(b.dataset.tpRole, 'roles', b.dataset.v);
      if (b.dataset.tpDept) return tog(b.dataset.tpDept, 'depts', b.dataset.v);
      if (b.hasAttribute('data-tp-reset')) { if (draft.mode === 'org') draft.parent = {}; else draft.rules = {}; ui.msg = ''; draw(); return; }
      if (b.hasAttribute('data-tp-save')) {
        const st = Store.db.settings || (Store.db.settings = {});
        // roles: names and added roles (keeping each one's visible sections); departments: only when changed here
        const views = Object.assign({}, (st.roles || {}).views); for (const k of Object.keys(views)) if (k.startsWith('r_') && !draft.roles.custom.some(c => c.id === k)) delete views[k];
        st.roles = Object.assign({}, st.roles, { names: draft.roles.names, custom: draft.roles.custom, views });
        if (draft.deptsChanged) { st.depts = draft.depts.map(d => ({ id: d.id, name: d.name, short: d.short })); D.applyDepts(st.depts); }
        const live = new Set(dNodes());
        const parent = Object.fromEntries(Object.entries(draft.parent).filter(([n, p]) => live.has(n) && live.has(p)));
        const roleIds = new Set(dRoles().map(r => r.id)), deptIds = new Set(draft.depts.map(d => d.id));
        const rules = Object.fromEntries(Object.entries(draft.rules).filter(([k]) => roleIds.has(k)).map(([k, x]) => [k, { roles: (x.roles || []).filter(r => roleIds.has(r)), depts: (x.depts || []).filter(d => deptIds.has(d)) }]).filter(([, x]) => x.roles.length || x.depts.length));
        st.taskPerms = { mode: draft.mode, parent, rules };
        Store.save(); App.closeModal(); App.render(); App.toast('บันทึกสิทธิ์การสั่งงานแล้ว (' + (draft.mode === 'org' ? 'แบบผังองค์กร' : 'แบบฟอร์มสำหรับเลือก') + ')');
      }
    });
    draw();
  };
})(window);
