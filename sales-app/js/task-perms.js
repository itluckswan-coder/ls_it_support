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

  const nodeName = n => { if (n.startsWith('dept:')) { const d = D.DEPTS.find(x => x.id === n.slice(5)); return 'แผนก' + (d ? d.short : n.slice(5)); } const r = roleById(n.slice(5)); return r ? r.name : n.slice(5); };
  App.ACT.tpOpen = function () {
    if (!App.me() || App.me().role !== 'admin') return;
    const cur = TP(), draft = { mode: cur.mode === 'org' ? 'org' : 'form', parent: Object.assign({}, cur.parent), rules: JSON.parse(JSON.stringify(cur.rules || {})) };
    const others = () => roleList().filter(r => r.id !== 'admin');
    const bg = App.modal('<div id="tpBody"></div>'), modal = bg.querySelector('.modal');
    modal.style.maxWidth = '980px';

    function chart() {
      const map = draft.parent, all = nodes(), kids = n => all.filter(x => parentIn(map, x) === n);
      const card = n => {
        const opts = roleList().filter(r => 'role:' + r.id !== n && !below(map, 'role:' + r.id, n));
        const isRole = n.startsWith('role:');
        return `<div class="oc-card ${isRole ? 'role' : 'dept'}${n === 'role:admin' ? ' root' : ''}"><b>${H(nodeName(n))}</b>
          ${n === 'role:admin' ? '<span class="xs muted">สั่งงานได้ทุกคน</span>' : `<label class="xs muted">อยู่ใต้<select class="input sm" data-tp-par="${H(n)}">${opts.map(r => `<option value="role:${H(r.id)}" ${parentIn(map, n) === 'role:' + r.id ? 'selected' : ''}>${H(r.name)}</option>`).join('')}</select></label>`}</div>`;
      };
      const tree = n => { const k = kids(n); return `<li>${card(n)}${k.length ? `<ul>${k.map(tree).join('')}</ul>` : ''}</li>`; };
      return `<p class="small muted" style="margin:0 0 10px">บทบาทที่อยู่ด้านบนสั่งงานได้ทุกบทบาทและทุกแผนกที่อยู่ใต้สาย · เปลี่ยน “อยู่ใต้” เพื่อย้ายตำแหน่งในผัง</p>
        <div class="oc-wrap"><ul class="oc">${tree('role:admin')}</ul></div>`;
    }
    function form() {
      const chip = (on, attrs, label) => `<button type="button" class="chip ${on ? 'on' : ''}" ${attrs}>${H(label)}</button>`;
      return `<p class="small muted" style="margin:0 0 10px">เลือกว่าแต่ละบทบาทสั่งงานให้บทบาทไหน และแผนกไหนได้ (ผู้ดูแลระบบสั่งงานได้ทุกคนเสมอ)</p>
        <div class="tp-rows">${others().map(r => { const x = draft.rules[r.id] || {}, rs = x.roles || [], ds = x.depts || []; return `<div class="tp-row">
          <div class="tp-who"><b>${H(r.name)}</b><span class="xs muted">สั่งงานให้ได้</span></div>
          <div><div class="xs muted" style="margin-bottom:4px">ผู้ใช้งานบทบาท</div><div class="row wrap" style="gap:6px">${others().map(t => chip(rs.includes(t.id), `data-tp-role="${H(r.id)}" data-v="${H(t.id)}"`, t.name)).join('')}</div>
            <div class="xs muted" style="margin:8px 0 4px">แผนก</div><div class="row wrap" style="gap:6px">${D.DEPTS.map(d => chip(ds.includes(d.id), `data-tp-dept="${H(r.id)}" data-v="${H(d.id)}"`, d.short)).join('')}</div></div>
        </div>`; }).join('')}</div>`;
    }
    function draw() {
      const y = modal.scrollTop;
      modal.querySelector('#tpBody').innerHTML = `<h2>ตั้งค่าสิทธิ์การสั่งงาน</h2>
        <div class="seg white" role="tablist" style="margin-bottom:6px"><button type="button" class="${draft.mode === 'org' ? 'on' : ''}" data-tp-mode="org">แบบผังองค์กร</button><button type="button" class="${draft.mode === 'form' ? 'on' : ''}" data-tp-mode="form">แบบฟอร์มสำหรับเลือก</button></div>
        <div class="xs muted" style="margin-bottom:14px">กดบันทึกแล้วระบบจะใช้แบบที่เลือกอยู่</div>
        ${draft.mode === 'org' ? chart() : form()}
        <div class="acts" style="flex-wrap:wrap"><button type="button" class="btn ghost" data-tp-reset style="margin-right:auto">คืนค่าเริ่มต้น</button><button type="button" class="btn ghost" data-act="closeModal">ยกเลิก</button><button type="button" class="btn red" data-tp-save>บันทึก</button></div>`;
      modal.scrollTop = y;
    }
    modal.addEventListener('change', e => { const s = e.target.closest('[data-tp-par]'); if (s) { draft.parent[s.dataset.tpPar] = s.value; draw(); } });
    modal.addEventListener('click', e => {
      const b = e.target.closest('[data-tp-mode],[data-tp-role],[data-tp-dept],[data-tp-save],[data-tp-reset]'); if (!b) return;
      if (b.dataset.tpMode) { draft.mode = b.dataset.tpMode; draw(); return; }
      const tog = (r, k, v) => { const x = draft.rules[r] || (draft.rules[r] = { roles: [], depts: [] }), l = x[k] || (x[k] = []), i = l.indexOf(v); if (i >= 0) l.splice(i, 1); else l.push(v); draw(); };
      if (b.dataset.tpRole) return tog(b.dataset.tpRole, 'roles', b.dataset.v);
      if (b.dataset.tpDept) return tog(b.dataset.tpDept, 'depts', b.dataset.v);
      if (b.hasAttribute('data-tp-reset')) { if (draft.mode === 'org') draft.parent = {}; else draft.rules = {}; draw(); return; }
      if (b.hasAttribute('data-tp-save')) {
        const st = Store.db.settings || (Store.db.settings = {});
        const parent = Object.fromEntries(Object.entries(draft.parent).filter(([n, p]) => p && p !== defParent(n)));
        const rules = Object.fromEntries(Object.entries(draft.rules).map(([k, x]) => [k, { roles: (x.roles || []).slice(), depts: (x.depts || []).slice() }]).filter(([, x]) => x.roles.length || x.depts.length));
        st.taskPerms = { mode: draft.mode, parent, rules };
        Store.save(); App.closeModal(); App.render(); App.toast('บันทึกสิทธิ์การสั่งงานแล้ว (' + (draft.mode === 'org' ? 'แบบผังองค์กร' : 'แบบฟอร์มสำหรับเลือก') + ')');
      }
    });
    draw();
  };
})(window);
