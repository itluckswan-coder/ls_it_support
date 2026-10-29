/* Luck Swan Sales — ตารางการทำงาน: the monthly work roster of a property (26th to the 25th), one row per person,
 * a shift code or a mark (DAY OFF, EXTRA, CDO, WOP, VC) per day, the day's room forecast (OCC%) and function guests,
 * and each person's totals. Heads of department fill their department's rows, HR (leave_hr) and admins any row; it is
 * signed in turn by a head (หัวหน้า) → HR (ทรัพยากรบุคคล) → an admin (ผู้จัดการ), and prints on one A4 landscape page.
 * One roster per department (same form for every department), each signed on its own.
 * Stored as doc kind 'roster', id "<site>:<dept>:<YYYY-MM>" (the month the period ends in); the worker applies the same rules.
 * (Rosters made before the split, "<site>:<YYYY-MM>", hold every department; a department's new roster copies its rows from one.) */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  // shift codes until the admin edits them (settings.shiftCodes)
  const SHIFTS_DEFAULT = [{ code: 'ช6', time: '' }, { code: 'ช7', time: '' }, { code: 'ช8', time: '' }, { code: 'บ2', time: '' }, { code: 'ค4', time: '' }];
  const MARKS = {
    OFF: { name: 'DAY OFF', note: 'วันหยุดประจำสัปดาห์', cls: 'rs-off' },
    EX: { name: 'EXTRA', note: 'ยึดตามตารางเวร ลงตารางถือว่าใช้สิทธิ หากมาทำงานให้เปลี่ยนเป็น QT', cls: 'rs-ex' },
    CDO: { name: 'CDO', note: 'เปลี่ยนวันหยุด **เขียนใบลาให้ HR ทุกครั้ง**', cls: 'rs-cdo' },
    WOP: { name: 'WOP', note: 'ลากิจไม่รับค่าจ้าง', cls: 'rs-wop' },
    VC: { name: 'VC', note: 'พักร้อน ยึดตามตารางเวร ลงตารางถือว่าใช้สิทธิ หากมาทำงานให้แจ้ง', cls: 'rs-vc' }
  };
  const STEPS = [['head', 'หัวหน้า'], ['hr', 'ทรัพยากรบุคคล'], ['gm', 'ผู้จัดการ']];

  const shifts = () => { const s = Store.db.settings && Store.db.settings.shiftCodes; return Array.isArray(s) && s.length ? s : SHIFTS_DEFAULT; };
  const rosters = () => Store.db.rosters || (Store.db.rosters = []);
  const me = () => App.me();
  const isHr = u => !!u && (u.role === 'admin' || !!(u.perms && u.perms.leave_hr));
  const isHead = u => !!u && (u.role === 'head' || u.role === 'manager');
  const myDept = () => App.userDept ? App.userDept(me()) : null;
  // heads work on their own department's roster; HR and admins on any
  const canEdit = () => isHr(me()) || (isHead(me()) && myDept() === UI().dept);
  const canRow = () => canEdit();
  const mayStep = (k, u) => k === 'head' ? u.role === 'admin' || (isHead(u) && myDept() === UI().dept) : k === 'hr' ? isHr(u) : u.role === 'admin';

  /* ---- periods: 26th of one month to the 25th of the next, keyed by the month it ends in ---- */
  const pad = n => String(n).padStart(2, '0');
  const keyOf = iso => { const d = U.d(iso); let y = d.getFullYear(), m = d.getMonth() + 1; if (d.getDate() >= 26) { m++; if (m > 12) { m = 1; y++; } } return y + '-' + pad(m); };
  const shiftKey = (key, n) => { let [y, m] = key.split('-').map(Number); m += n; while (m > 12) { m -= 12; y++; } while (m < 1) { m += 12; y--; } return y + '-' + pad(m); };
  const daysOf = key => { const [y, m] = key.split('-').map(Number); const out = []; let d = new Date(y, m - 2, 26); const end = new Date(y, m - 1, 25); while (d <= end) { out.push(U.iso(d)); d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1); } return out; };
  const label = key => { const ds = daysOf(key), a = U.d(ds[0]), b = U.d(ds[ds.length - 1]); return `${a.getDate()} ${U.ML[a.getMonth()]} - ${b.getDate()} ${U.ML[b.getMonth()]} ${b.getFullYear() + 543}`; };

  const deptName = id => { const d = D.DEPTS.find(x => x.id === id); return d ? d.short || d.name : id || 'อื่น ๆ'; };
  const isShift = v => !!v && !MARKS[v] && !/^CDO/.test(v);
  const tally = (row, days) => {
    const vals = days.map(d => (row.cells || {})[d] || '');
    const work = vals.filter(isShift).length, off = vals.filter(v => v === 'OFF' || /^CDO/.test(v)).length, ex = vals.filter(v => v === 'EX').length, vc = vals.filter(v => v === 'VC').length;
    return { work, off, ex, vc, total: days.length, left: (Number(row.vcStart) || 0) - vc };
  };
  const pax = (site, d) => App.siteQuotes().filter(q => (q.status === 'tentative' || q.status === 'confirmed') && q.start && q.start <= d && d <= (q.end || q.start)).reduce((a, q) => a + (Number(q.pax) || 0), 0);
  const site = () => App.site() || 'resort';
  const rid = (dept, key) => site() + ':' + dept + ':' + key;
  const legacy = key => rosters().find(r => r.id === site() + ':' + key); // a roster from before the split by department
  const cur = () => { const ui = UI(); return rosters().find(r => r.id === rid(ui.dept, ui.key)); };
  const UI = () => {
    const ui = App.state.ui.roster || (App.state.ui.roster = { key: keyOf(U.today()), brush: null });
    if (!ui.dept || !D.DEPTS.some(d => d.id === ui.dept)) { const m = myDept(); ui.dept = D.DEPTS.some(d => d.id === m) ? m : (D.DEPTS[0] || {}).id; }
    return ui;
  };
  const deptEmps = dept => Store.db.employees.filter(e => (App.SITES ? App.SITES.of(e) : 'resort') === site() && e.dept === dept);

  // any change to the table after someone signed sends it back for signatures
  const touched = r => { if (r.sign && Object.keys(r.sign).length) { r.sign = {}; App.toast('แก้ไขตารางแล้ว — ต้องลงนามใหม่ตามลำดับ'); } r.updatedAt = U.nowTH(); Store.save(); App.render(); };

  const cellHtml = (v, d) => {
    if (!v) return '';
    if (/^CDO/.test(v)) return `<span class="rs-cdo-t">CDO</span><span class="rs-cdo-d">${H(v.slice(4))}</span>`;
    if (v === 'OFF') return '';
    return H(v === 'EX' ? 'EX' : MARKS[v] ? MARKS[v].name : v);
  };
  const cellCls = v => !v ? '' : /^CDO/.test(v) ? 'rs-cdo' : MARKS[v] ? MARKS[v].cls : '';

  App.V.staff.renderRoster = function () {
    const ui = UI(), site = App.site() || 'resort', r = cur(), days = daysOf(ui.key), u = me(), edit = canEdit();
    const prop = App.prop ? App.prop(site) : { th: site };
    const dtabs = `<div class="row wrap no-print" style="gap:6px;margin-bottom:14px" role="tablist" aria-label="แผนก">${D.DEPTS.map(d => { const has = rosters().some(x => x.id === rid(d.id, ui.key)); return `<button class="chip ${ui.dept === d.id ? 'on' : ''}" role="tab" data-act="rsDept" data-v="${d.id}"><i class="dot" style="background:${d.color};border-radius:3px"></i> ${H(d.name)} <span class="n">${deptEmps(d.id).length}</span>${has ? ' ✓' : ''}</button>`; }).join('')}</div>`;
    const nav = `<div class="row" style="gap:8px"><button class="icon-btn" data-act="rsNav" data-v="-1" aria-label="รอบก่อนหน้า">‹</button><b style="min-width:220px;text-align:center">${label(ui.key)}</b><button class="icon-btn" data-act="rsNav" data-v="1" aria-label="รอบถัดไป">›</button></div>`;
    const head = `<div class="page-hd no-print"><div><h1>ตารางการทำงาน</h1><div class="small muted">${H(prop.th)} · ${H(deptName(ui.dept))} · ตัดรอบวันที่ 26 – 25 ของทุกเดือน</div></div>
      <div class="row wrap">${nav}${r ? '<button class="btn ghost" data-act="rsPrint">พิมพ์ / PDF</button>' : ''}${isHr(u) && u.role === 'admin' ? '<button class="btn ghost" data-act="rsCodes">รหัสกะ</button>' : ''}</div></div>`;
    if (!r) return head + dtabs + `<div class="card pad" style="text-align:center"><p class="muted">ยังไม่มีตารางการทำงานของ${H(deptName(ui.dept))} รอบ ${label(ui.key)}</p>
      ${edit ? `<button class="btn red" data-act="rsNew">สร้างตารางรอบนี้</button><div class="xs muted" style="margin-top:8px">${legacy(ui.key) && legacy(ui.key).rows.some(x => x.dept === ui.dept) ? 'ใช้ข้อมูลของแผนกนี้จากตารางรวมรอบนี้' : rosters().some(x => x.id === rid(ui.dept, shiftKey(ui.key, -1))) || legacy(shiftKey(ui.key, -1)) ? 'ใช้รายชื่อและยอดพักร้อนคงเหลือจากรอบก่อน' : 'ใช้รายชื่อพนักงานของแผนกนี้'}</div>` : `<p class="small muted">หัวหน้า${H(deptName(ui.dept))} ฝ่ายบุคคล หรือผู้ดูแลระบบเป็นผู้สร้างตาราง</p>`}</div>`;

    const groups = []; for (const row of r.rows) { let gp = groups.find(x => x.dept === row.dept); if (!gp) groups.push(gp = { dept: row.dept, rows: [] }); gp.rows.push(row); }
    const today = U.today();
    const tools = edit ? `<div class="rs-tools no-print"><span class="small muted">เลือกแล้วแตะช่องในตาราง:</span>
      ${shifts().map(s => `<button class="chip ${ui.brush === s.code ? 'on' : ''}" data-act="rsBrush" data-v="${H(s.code)}" title="${H(s.time || s.code)}">${H(s.code)}</button>`).join('')}
      ${Object.entries(MARKS).map(([k, m]) => `<button class="chip ${m.cls} ${ui.brush === k ? 'on' : ''}" data-act="rsBrush" data-v="${k}">${m.name}</button>`).join('')}
      <button class="chip ${ui.brush === '' ? 'on' : ''}" data-act="rsBrush" data-v="">ล้างช่อง</button>
      <span style="flex:1"></span><button class="link" data-act="rsAddRow">+ เพิ่มแถว</button>${edit ? '<button class="link" data-act="rsSync">+ ดึงพนักงานที่ยังไม่มีในตาราง</button>' : ''}</div>` : '';
    const dayHead = days.map(d => { const x = U.d(d), we = x.getDay() === 0 || x.getDay() === 6; return `<th class="rs-dh${we ? ' we' : ''}${d === today ? ' td' : ''}">${U.DS[x.getDay()]}</th>`; }).join('');
    const dateHead = days.map(d => `<th class="rs-dh${d === today ? ' td' : ''}">${U.d(d).getDate()}</th>`).join('');
    const occ = days.map(d => { const v = (r.occ || {})[d]; return `<td class="rs-fc">${isHr(u) ? `<input class="rs-in" inputmode="numeric" value="${v == null ? '' : H(v)}" data-change="rsOcc" data-d="${d}" aria-label="OCC% ${d}">` : H(v == null ? '' : v)}</td>`; }).join('');
    const fn = days.map(d => { const n = pax(site, d); return `<td class="rs-fc">${n || 0}</td>`; }).join('');
    const body = groups.map(gp => gp.rows.map((row, i) => {
      const t = tally(row, days), ok = canRow(row);
      return `<tr>${i ? '' : `<td class="rs-dept" rowspan="${gp.rows.length}">${H(deptName(gp.dept))}</td>`}
        <td class="rs-code">${ok ? `<button class="rs-txt" data-act="rsCode" data-r="${H(row.id)}">${H(row.code || '—')}</button>` : H(row.code || '')}</td>
        <td class="rs-name">${ok ? `<button class="rs-txt" data-act="rsName" data-r="${H(row.id)}">${H(row.name || '—')}</button>` : H(row.name)}${row.note ? `<div class="xs muted">${H(row.note)}</div>` : ''}</td>
        ${days.map(d => { const v = (row.cells || {})[d] || ''; return `<td class="rs-c ${cellCls(v)}"${ok ? ` data-act="rsCell" data-r="${H(row.id)}" data-d="${d}" tabindex="0"` : ''}>${cellHtml(v, d)}</td>`; }).join('')}
        <td class="rs-v">${t.work}</td><td class="rs-v">${t.off}</td><td class="rs-v">${t.ex}</td><td class="rs-v b">${t.total}</td>
        <td class="rs-v${isHr(u) ? ' rs-click' : ''}"${isHr(u) ? ` data-act="rsVc" data-r="${H(row.id)}" title="ยอดพักร้อนยกมา ${Number(row.vcStart) || 0} วัน"` : ''}>${t.left}</td><td class="rs-v">${t.vc || '-'}</td></tr>`;
    }).join('')).join('');
    const S = r.sign || {}, next = STEPS.find(([k]) => !S[k]);
    const signs = STEPS.map(([k, name]) => { const s = S[k]; return `<div class="rs-sig"><div class="rs-sig-line">${s ? `<b>${H(s.byName)}</b>` : '&nbsp;'}</div><div>${name}</div><div class="xs muted">${s ? U.fmtD(s.at.slice(0, 10)) + ' ' + s.at.slice(11, 16) : ''}</div>
      ${!s && next && next[0] === k && u && mayStep(k, u) ? `<button class="btn teal sm no-print" data-act="rsSign" data-k="${k}" style="margin-top:6px">ลงนาม (${name})</button>` : ''}</div>`; }).join('');
    const status = S.gm ? '<span class="badge st-confirmed">ลงนามครบแล้ว</span>' : `<span class="badge st-sent">รอ${next ? next[1] : ''}ลงนาม</span>`;
    return head + dtabs + tools + `<section class="card rs-card"><div class="rs-scroll"><div class="rs-paper">
  <div class="rs-title"><b>ตารางการทำงาน ${label(ui.key)}</b><span class="rs-site">${H(prop.th)} · ${H(deptName(ui.dept))}</span></div>
  <table class="rs"><thead>
    <tr><th rowspan="2">ฝ่าย</th><th rowspan="2">รหัสพนักงาน</th><th rowspan="2">รายชื่อพนักงาน</th>${dayHead}<th rowspan="4" class="rs-vh">วันทำงาน</th><th rowspan="4" class="rs-vh">DAY OFF</th><th rowspan="4" class="rs-vh">EXTRA</th><th rowspan="4" class="rs-vh">รวม</th><th rowspan="4" class="rs-vh">พักร้อนคงเหลือ</th><th rowspan="4" class="rs-vh">พักร้อนใช้ไป</th></tr>
    <tr>${dateHead}</tr>
    <tr class="rs-fr"><td colspan="3">Room Forecast OCC%</td>${occ}</tr>
    <tr class="rs-fr"><td colspan="3">Function Forecast/pax</td>${fn}</tr>
  </thead><tbody>${body || `<tr><td colspan="${days.length + 9}" class="muted" style="text-align:center;padding:14px">ยังไม่มีรายชื่อ</td></tr>`}</tbody></table>
  <div class="rs-foot">
    <div class="rs-note"><b class="red">หมายเหตุ</b> ${edit ? `<textarea class="input rs-notein no-print" rows="2" data-change="rsNote" placeholder="หมายเหตุ (ไม่บังคับ)">${H(r.note || '')}</textarea><div class="print-only">${H(r.note || '')}</div>` : H(r.note || '')}
      <div class="small">ตารางการทำงานอาจมีการเปลี่ยนแปลง หากมีการเปลี่ยนแปลงจะแจ้งให้ทราบล่วงหน้า</div>
      <div class="rs-legend">${Object.values(MARKS).map(m => `<div><i class="${m.cls}"></i><b>${m.name}</b> ${H(m.note)}</div>`).join('')}
        ${shifts().some(s => s.time) ? `<div class="xs">${shifts().filter(s => s.time).map(s => `${H(s.code)} = ${H(s.time)}`).join(' · ')}</div>` : ''}</div></div>
    <div class="rs-sigs">${signs}</div>
  </div></div></div></section>
  <div class="row no-print" style="margin-top:10px;gap:10px">${status}${u && u.role === 'admin' && Object.keys(S).length ? '<button class="link red small" data-act="rsUnsign">ยกเลิกการลงนามทั้งหมด</button>' : ''}${u && u.role === 'admin' ? '<span style="flex:1"></span><button class="link red small" data-act="rsDel">ลบตารางรอบนี้</button>' : ''}</div>`;
  };

  const rowOf = id => cur().rows.find(x => x.id === id);
  const ask = (title, fields, done) => App.modal(`<h2>${H(title)}</h2>${fields.map(f => `<div class="field" style="margin-top:10px"><label class="f">${H(f.label)}</label>${f.opts ? `<select class="input" name="${f.k}">${f.opts.map(o => `<option value="${H(o[0])}" ${o[0] === f.v ? 'selected' : ''}>${H(o[1])}</option>`).join('')}</select>` : `<input class="input" name="${f.k}" value="${H(f.v == null ? '' : f.v)}" ${f.num ? 'inputmode="decimal"' : ''}>`}</div>`).join('')}
    <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="rsOk">บันทึก</button></div>`, bg => { bg.querySelector('#rsOk').onclick = () => { const f = App.formData(bg); App.closeModal(); done(f); }; const i = bg.querySelector('input,select'); if (i) i.focus(); });
  const setCell = (row, d, v) => { row.cells = row.cells || {}; if (v) row.cells[d] = v; else delete row.cells[d]; touched(cur()); };
  const dm = d => { const x = U.d(d); return pad(x.getDate()) + '/' + pad(x.getMonth() + 1); };

  Object.assign(App.ACT, {
    rsNav(el) { const ui = UI(); ui.key = shiftKey(ui.key, Number(el.dataset.v)); App.render(); },
    rsPrint() { window.print(); },
    rsDept(el) { const ui = UI(); ui.dept = el.dataset.v; ui.brush = null; App.render(); },
    rsNew() {
      if (!canEdit()) return;
      const ui = UI(), dept = ui.dept, pk = shiftKey(ui.key, -1), same = legacy(ui.key);
      const mine = rs => rs.filter(x => x.dept === dept);
      const lg = legacy(pk), prev = rosters().find(x => x.id === rid(dept, pk)) || (lg && mine(lg.rows).length ? lg : null);
      // this period's combined roster (before the split) → its rows for this department, cells included;
      // else last period's → names and leave left; else the department's staff
      const rows = same && mine(same.rows).length ? mine(same.rows).map(x => JSON.parse(JSON.stringify(x)))
        : prev ? mine(prev.rows).map(x => ({ id: x.id, emp: x.emp || null, name: x.name, code: x.code || '', dept, note: x.note || '', vcStart: tally(x, daysOf(prev.period)).left, cells: {} }))
        : deptEmps(dept).map(e => ({ id: 'r_' + e.id, emp: e.id, name: e.name, code: e.code || '', dept, vcStart: 0, cells: {} }));
      rosters().push({ id: rid(dept, ui.key), site: site(), dept, period: ui.key, rows, occ: same ? Object.assign({}, same.occ) : {}, note: '', sign: {}, createdAt: U.nowTH(), by: me().id });
      Store.save(); App.render(); App.toast('สร้างตารางการทำงานแล้ว');
    },
    rsSync() {
      const r = cur(), have = new Set(r.rows.map(x => x.emp).filter(Boolean));
      const add = deptEmps(r.dept).filter(e => !have.has(e.id));
      if (!add.length) { App.toast('มีพนักงานครบทุกคนแล้ว'); return; }
      for (const e of add) r.rows.push({ id: 'r_' + e.id, emp: e.id, name: e.name, code: e.code || '', dept: r.dept, vcStart: 0, cells: {} });
      touched(r); App.toast('เพิ่ม ' + add.length + ' คน');
    },
    rsAddRow() {
      if (!canEdit()) return;
      ask('เพิ่มแถว · ' + deptName(UI().dept), [{ k: 'name', label: 'ชื่อ (เช่น ต้อนรับกะบ่าย, รีสอร์ท แทน)', v: '' }, { k: 'code', label: 'รหัสพนักงาน (ไม่บังคับ)', v: '' }], f => {
        if (!String(f.name || '').trim()) { App.toast('กรุณากรอกชื่อ'); return; }
        const r = cur(); r.rows.push({ id: 'r_' + Math.random().toString(36).slice(2, 9), emp: null, name: f.name.trim(), code: (f.code || '').trim(), dept: r.dept, vcStart: 0, cells: {} }); touched(r);
      });
    },
    rsName(el) {
      const row = rowOf(el.dataset.r);
      App.modal(`<h2>${H(row.name)}</h2><div class="field"><label class="f">ชื่อ</label><input class="input" name="name" value="${H(row.name)}"></div><div class="field" style="margin-top:10px"><label class="f">หมายเหตุใต้ชื่อ (ไม่บังคับ)</label><input class="input" name="note" value="${H(row.note || '')}" placeholder="เช่น แทน 7 วัน"></div>
        <div class="acts"><button class="btn ghost red" id="rsDelRow">ลบแถวนี้</button><span style="flex:1"></span><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="rsOk">บันทึก</button></div>`, bg => {
        bg.querySelector('#rsOk').onclick = () => { const f = App.formData(bg); if (!f.name.trim()) return; row.name = f.name.trim(); row.note = f.note.trim(); App.closeModal(); touched(cur()); };
        bg.querySelector('#rsDelRow').onclick = () => { App.closeModal(); App.confirm('ลบแถว ' + row.name + ' ออกจากตารางรอบนี้?', () => { const r = cur(); r.rows = r.rows.filter(x => x !== row); touched(r); }, 'ลบแถว'); };
      });
    },
    rsCode(el) { const row = rowOf(el.dataset.r); ask('รหัสพนักงาน · ' + row.name, [{ k: 'code', label: 'รหัสพนักงาน', v: row.code }], f => { row.code = (f.code || '').trim(); touched(cur()); }); },
    rsVc(el) { if (!isHr(me())) return; const row = rowOf(el.dataset.r); ask('พักร้อนยกมา · ' + row.name, [{ k: 'n', label: 'จำนวนวันพักร้อนที่มีต้นรอบ', v: Number(row.vcStart) || 0, num: true }], f => { row.vcStart = Math.max(0, Number(f.n) || 0); touched(cur()); }); },
    rsBrush(el) { const ui = UI(), v = el.dataset.v; ui.brush = ui.brush === v ? null : v; App.render(); },
    rsCell(el) {
      const row = rowOf(el.dataset.r), d = el.dataset.d, ui = UI(); if (!canRow(row)) return;
      if (ui.brush === 'CDO') { ask('CDO เปลี่ยนวันหยุด', [{ k: 'n', label: 'หยุดแทนวันที่ (วว/ดด)', v: dm(d) }], f => setCell(row, d, 'CDO ' + (f.n || '').trim())); return; }
      if (ui.brush != null) { setCell(row, d, ui.brush); return; }
      const v = (row.cells || {})[d] || '';
      App.modal(`<h2>${H(row.name)} · ${U.fmtD(d, { dow: true })}</h2><div class="row wrap" style="gap:8px">${shifts().map(s => `<button class="chip ${v === s.code ? 'on' : ''}" data-v="${H(s.code)}">${H(s.code)}${s.time ? ' · ' + H(s.time) : ''}</button>`).join('')}
        ${Object.entries(MARKS).map(([k, m]) => `<button class="chip ${m.cls} ${v === k || (k === 'CDO' && /^CDO/.test(v)) ? 'on' : ''}" data-v="${k}">${m.name}</button>`).join('')}<button class="chip" data-v="">ล้างช่อง</button></div>
        <p class="xs muted" style="margin-top:12px">เคล็ดลับ: เลือกรหัสจากแถบด้านบนตาราง แล้วแตะหลายช่องต่อกันได้เลย</p><div class="acts"><button class="btn ghost" data-act="closeModal">ปิด</button></div>`, bg => {
        bg.querySelectorAll('[data-v]').forEach(b => b.onclick = () => { App.closeModal(); const nv = b.dataset.v; if (nv === 'CDO') ask('CDO เปลี่ยนวันหยุด', [{ k: 'n', label: 'หยุดแทนวันที่ (วว/ดด)', v: dm(d) }], f => setCell(row, d, 'CDO ' + (f.n || '').trim())); else setCell(row, d, nv); });
      });
    },
    rsSign(el) {
      const r = cur(), k = el.dataset.k, u = me(); if (!mayStep(k, u)) return;
      const name = STEPS.find(s => s[0] === k)[1];
      App.confirm(`ลงนามช่อง “${name}” ในตารางการทำงานรอบ ${label(r.period)}?`, () => { r.sign = r.sign || {}; r.sign[k] = { by: u.id, byName: u.name, at: U.nowTH() }; Store.save(); App.render(); App.toast('ลงนามแล้ว'); }, 'ลงนาม');
    },
    rsUnsign() { const r = cur(); if (me().role !== 'admin') return; App.confirm('ยกเลิกการลงนามทั้งหมดของตารางรอบนี้?', () => { r.sign = {}; Store.save(); App.render(); }, 'ยกเลิกการลงนาม'); },
    rsDel() { const r = cur(); if (me().role !== 'admin') return; App.confirm('ลบตารางการทำงานรอบ ' + label(r.period) + '?', () => { Store.db.rosters = rosters().filter(x => x !== r); Store.save(); App.render(); App.toast('ลบตารางแล้ว'); }, 'ลบตาราง'); },
    rsCodes() {
      if (me().role !== 'admin') return;
      const list = shifts().map(s => ({ code: s.code, time: s.time || '' }));
      const draw = bg => {
        bg.querySelector('#rcRows').innerHTML = list.map((s, i) => `<div class="row" style="gap:8px;margin-bottom:8px;flex-wrap:nowrap"><input class="input" data-i="${i}" data-f="code" value="${H(s.code)}" placeholder="รหัส" style="width:90px"><input class="input" data-i="${i}" data-f="time" value="${H(s.time)}" placeholder="เวลา เช่น 07:00–16:00" style="flex:1;min-width:0"><button type="button" class="icon-btn sm" data-del="${i}" aria-label="ลบรหัส" style="color:var(--red)">🗑</button></div>`).join('');
        bg.querySelectorAll('#rcRows input').forEach(inp => inp.oninput = () => { list[inp.dataset.i][inp.dataset.f] = inp.value; });
        bg.querySelectorAll('[data-del]').forEach(b => b.onclick = () => { list.splice(b.dataset.del, 1); draw(bg); });
      };
      App.modal(`<h2>รหัสกะ</h2><p class="muted" style="margin-top:-8px">รหัสที่ใช้ลงในตารางการทำงาน (DAY OFF, EXTRA, CDO, WOP, VC มีให้อยู่แล้ว)</p><div id="rcRows"></div><button type="button" class="link" id="rcAdd">+ เพิ่มรหัสกะ</button>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="rcOk">บันทึก</button></div>`, bg => {
        draw(bg); bg.querySelector('#rcAdd').onclick = () => { list.push({ code: '', time: '' }); draw(bg); };
        bg.querySelector('#rcOk').onclick = () => { const out = list.map(s => ({ code: s.code.trim(), time: s.time.trim() })).filter(s => s.code && !MARKS[s.code]); (Store.db.settings || (Store.db.settings = {})).shiftCodes = out; Store.save(); App.closeModal(); App.render(); App.toast('บันทึกรหัสกะแล้ว'); };
      });
    }
  });
  Object.assign(App.CHG, {
    rsOcc(el) { const r = cur(); if (!r || !isHr(me())) return; r.occ = r.occ || {}; const v = String(el.value).trim(); if (v) r.occ[el.dataset.d] = v; else delete r.occ[el.dataset.d]; touched(r); },
    rsNote(el) { const r = cur(); if (!r || !canEdit()) return; r.note = el.value; touched(r); }
  });
})(window);
