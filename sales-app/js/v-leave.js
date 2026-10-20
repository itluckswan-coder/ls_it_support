/* Luck Swan Sales — การแจ้งใบลา: the HR form HRSD08.1 (ใบขอใช้สิทธิการลา QT, การเปลี่ยนแปลงเวลาทำงาน และแจ้งเวลาเข้า-ออกงาน)
 * as three requests — แจ้งใบลา QT, เปลี่ยนวันหยุด / เวลาทำงาน, ลืมสแกนนิ้ว — each signed in turn by
 * หัวหน้าแผนก → ผู้จัดการทั่วไป (อนุมัติ / ไม่อนุมัติ) → รองผู้จัดการทั่วไป → ฝ่ายบุคคล, and printable as the paper form. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  const TYPES = {
    qt: { name: 'แจ้งใบลา QT', sub: 'ทำ QT / ลา QT', color: 'var(--teal)' },
    change: { name: 'เปลี่ยนวันหยุด', sub: 'เปลี่ยนวันหยุด / เปลี่ยนเวลาทำงาน', color: 'var(--blue, #2B3F86)' },
    scan: { name: 'ลืมสแกนนิ้ว', sub: 'สแกนนิ้วไม่ติด / ลืมสแกน / ยังไม่มีลายนิ้วมือ', color: 'var(--yellow)' }
  };
  const CAUSES = { fail: 'สแกนนิ้วไม่ติด', forgot: 'ลืมสแกนนิ้ว', nofp: 'เริ่มงานโดยยังไม่มีลายนิ้วมือ' };
  // signature lines on the form, in signing order; dgm (รองผู้จัดการทั่วไป) signs after the decision and is optional
  const STEPS = [['head', 'หัวหน้าแผนก'], ['gm', 'ผู้จัดการทั่วไป'], ['dgm', 'รองผู้จัดการทั่วไป'], ['hr', 'ฝ่ายบุคคล']];
  const STATUS = { pending: ['รอหัวหน้าแผนก', 'st-sent'], gm: ['รอผู้จัดการทั่วไป', 'st-sent'], approved: ['อนุมัติแล้ว', 'st-confirmed'], done: ['ฝ่ายบุคคลรับทราบแล้ว', 'st-confirmed'], rejected: ['ไม่อนุมัติ', 'st-tentative'], cancelled: ['ยกเลิก', 'st-draft'] };

  const leaves = () => Store.db.leaves || (Store.db.leaves = []);
  const isAdmin = me => !!me && me.role === 'admin';
  const sameDept = (a, b) => !!a && !!b && String(a).trim() === String(b).trim();
  // only admins sign and approve (every line of the form); HR (leave_hr) reads everything; a head of department sees their department's requests
  const maySign = me => isAdmin(me);
  const mayHr = me => isAdmin(me) || !!(me.perms && me.perms.leave_hr);
  const headOf = (l, me) => (me.role === 'head' || me.role === 'manager') && sameDept(l.dept, me.dept);
  const visible = (l, me) => l.by === me.id || mayHr(me) || headOf(l, me);
  const waitingOn = (l, me) => maySign(me) && ['pending', 'gm', 'approved'].includes(l.status);
  App.leaveWaiting = () => { const me = App.me(); return me ? leaves().filter(l => waitingOn(l, me)) : []; };

  const be = iso => { if (!iso) return ['', '', '']; const d = U.d(iso); return [d.getDate(), U.ML[d.getMonth()], d.getFullYear() + 543]; };
  const dTH = iso => iso ? U.fmtD(iso) : '—';
  const mins = t => { const m = /^(\d{1,2}):(\d{2})$/.exec(t || ''); return m ? +m[1] * 60 + +m[2] : null; };
  const hoursOf = (a, b) => { const x = mins(a), y = mins(b); if (x == null || y == null) return 0; return Math.round(((y - x + 1440) % 1440) / 30) / 2; };
  // HRSD08.1/YYMMDDNNN — Buddhist year, month, day and a running number for the day
  function nextNo() {
    const t = U.today(), d = U.d(t), p = 'HRSD08.1/' + String((d.getFullYear() + 543) % 100).padStart(2, '0') + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
    const n = leaves().filter(l => (l.no || '').startsWith(p)).map(l => +l.no.slice(p.length) || 0);
    return p + String(Math.max(0, ...n) + 1).padStart(3, '0');
  }
  const summary = l => {
    if (l.type === 'qt') return `${l.qt.mode === 'do' ? 'ทำ QT' : 'ลา QT'} ${l.qt.hours} ชม. · ${H(l.qt.from)}–${H(l.qt.to)} น. · ${dTH(l.qt.date)}`;
    if (l.type === 'change') return l.ch.mode === 'day' ? `เปลี่ยนวันหยุด ${dTH(l.ch.fromDate)} → ${dTH(l.ch.toDate)}` : `เปลี่ยนเวลาทำงาน ${H(l.ch.fromTime)} → ${H(l.ch.toTime)} น. · ${dTH(l.ch.date)}`;
    return `${H(CAUSES[l.sc.cause] || '')} · เข้า ${H(l.sc.inTime || '—')} ออก ${H(l.sc.outTime || '—')} · ${dTH(l.sc.date)}`;
  };
  const who = l => l.fullName || l.byName; // the name written on the form (the account name when none was given)
  const badge = l => { const s = STATUS[l.status] || STATUS.pending; return `<span class="badge ${s[1]}">${s[0]}</span>`; };

  /* ---------------- page ---------------- */
  App.V.leave = {
    nav: 'leave', title: 'การแจ้งใบลา',
    render() {
      const me = App.me(), ui = App.state.ui.leave || (App.state.ui.leave = { tab: 'mine' });
      const all = leaves().filter(l => visible(l, me)).sort((a, b) => ((b.at || '') + b.no).localeCompare((a.at || '') + a.no));
      const wait = all.filter(l => waitingOn(l, me)), others = all.some(l => l.by !== me.id) || mayHr(me);
      const tabs = [['mine', 'ใบของฉัน', all.filter(l => l.by === me.id).length]].concat(maySign(me) ? [['wait', 'รออนุมัติ', wait.length]] : []).concat(others ? [['all', 'ทั้งหมด', all.length]] : [])
        .concat(mayHr(me) ? [['info', 'ข้อมูลรายละเอียด', 0]] : []);
      if (!tabs.some(t => t[0] === ui.tab)) ui.tab = 'mine';
      const list = ui.tab === 'mine' ? all.filter(l => l.by === me.id) : ui.tab === 'wait' ? wait : all;
      return `
<div class="page-hd"><div><h1>การแจ้งใบลา</h1><div class="small muted">ใบขอใช้สิทธิการลา QT การเปลี่ยนแปลงเวลาทำงาน และแจ้งเวลาเข้า-ออกงาน (HRSD08.1)</div></div></div>
<div class="grid g3" style="margin-bottom:18px">
  ${Object.entries(TYPES).map(([k, t]) => `<button class="card pad lv-new" data-act="lvNew" data-v="${k}" style="text-align:left;border-top:4px solid ${t.color};cursor:pointer">
    <div class="between"><b style="font-size:17px">${t.name}</b><span class="btn red sm" aria-hidden="true">+ แจ้ง</span></div><div class="small muted" style="margin-top:6px">${t.sub}</div></button>`).join('')}
</div>
<div class="seg white" style="margin-bottom:14px" role="tablist">${tabs.map(t => `<button role="tab" class="${ui.tab === t[0] ? 'on' : ''}" data-act="lvTab" data-v="${t[0]}">${t[1]}${t[2] ? ` <span class="n">${t[2]}</span>` : ''}</button>`).join('')}</div>
${ui.tab === 'info' ? info(all) : `<section class="card" style="overflow:hidden">
  <table class="tbl resp"><thead><tr><th>เลขที่ / ประเภท</th><th>ผู้ขอ</th><th>รายละเอียด</th><th>สถานะ</th></tr></thead><tbody>
  ${list.map(l => `<tr class="clickable" data-act="lvOpen" data-id="${H(l.id)}" style="cursor:pointer">
    <td><div style="font-weight:500">${H(TYPES[l.type].name)}</div><div class="sub">${H(l.no)} · ${U.fmtD((l.at || '').slice(0, 10), { year: false })}</div></td>
    <td class="hide-m">${H(who(l))}<div class="sub">${H(l.position || '')}${l.dept ? ' · ' + H(l.dept) : ''}</div></td>
    <td class="hide-m small">${summary(l)}</td><td>${badge(l)}</td></tr>`).join('') || `<tr><td colspan="4"><div class="empty">${ui.tab === 'wait' ? 'ไม่มีใบที่รอคุณลงนาม' : 'ยังไม่มีใบแจ้ง — เลือกประเภทด้านบนเพื่อแจ้ง'}</div></td></tr>`}
  </tbody></table>
</section>`}
<div class="note" style="margin-top:14px">หมายเหตุ: อนุญาตให้ทำหรือใช้ QT จะต้องทำหรือใช้ 2 ชั่วโมงขึ้นไป · ไม่อนุญาตให้ทำหรือใช้ QT ก่อนเวลาเข้างาน จะถือว่าเปลี่ยนเวลาทำงาน · การเปลี่ยนวันหยุด/เวลาทำงาน ตรวจสอบจากตารางเวรที่มีการประกาศทุกครั้ง</div>`;
    }
  };

  /* ---------------- ข้อมูลรายละเอียด: every request in detail, for HR (ฝ่ายทรัพยากรบุคคล) and admins ---------------- */
  // the day a request takes effect (QT day, the new day off / day of the new hours, the day of the missed scan)
  const dayOf = l => l.type === 'qt' ? l.qt.date : l.type === 'change' ? (l.ch.mode === 'day' ? l.ch.toDate : l.ch.date) : l.sc.date;
  const ok = l => l.status === 'approved' || l.status === 'done';
  const signer = (l, k) => { const x = (l.sign || {})[k]; return x ? x.byName + (x.ok === false ? ' (ไม่อนุมัติ)' : '') + ' ' + U.fmtD(x.at.slice(0, 10), { year: false }) : ''; };
  const detailText = l => summary(l).replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
  function infoRows(all) {
    const f = App.state.ui.leaveInfo || (App.state.ui.leaveInfo = { type: 'all', status: 'all', month: U.today().slice(0, 7), dept: 'all', q: '' });
    const q = f.q.trim().toLowerCase();
    return { f, rows: all.filter(l => (f.type === 'all' || l.type === f.type) && (f.status === 'all' || (f.status === 'ok' ? ok(l) : f.status === 'wait' ? ['pending', 'gm'].includes(l.status) : l.status === f.status))
      && (!f.month || (dayOf(l) || '').startsWith(f.month)) && (f.dept === 'all' || (l.dept || '') === f.dept)
      && (!q || [who(l), l.nick, l.position, l.dept, l.no].join(' ').toLowerCase().includes(q))).sort((a, b) => (dayOf(b) || '').localeCompare(dayOf(a) || '') || b.no.localeCompare(a.no)) };
  }
  function info(all) {
    const { f, rows } = infoRows(all);
    const months = [...new Set(all.map(l => (dayOf(l) || '').slice(0, 7)).filter(Boolean).concat(U.today().slice(0, 7)))].sort().reverse();
    const depts = [...new Set(all.map(l => l.dept).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));
    const qtH = m => rows.filter(l => l.type === 'qt' && l.qt.mode === m && ok(l)).reduce((a, l) => a + (Number(l.qt.hours) || 0), 0);
    const sel = (k, opts) => `<select class="input" style="width:auto;min-width:0" data-change="lvInfo" data-k="${k}">${opts.map(([v, t]) => `<option value="${H(v)}" ${f[k] === v ? 'selected' : ''}>${H(t)}</option>`).join('')}</select>`;
    const kpi = (t, v, sub) => `<div class="card kpi"><div class="small muted">${t}</div><div class="v">${v}</div><div class="small muted">${sub}</div></div>`;
    // per person: QT made and used, changes of day off and missed scans (approved requests only)
    const people = {};
    for (const l of rows) {
      const k = who(l) + '|' + (l.dept || ''), p = people[k] || (people[k] = { name: who(l), nick: l.nick, dept: l.dept, pos: l.position, doH: 0, useH: 0, change: 0, scan: 0, wait: 0 });
      if (!ok(l)) { if (['pending', 'gm'].includes(l.status)) p.wait++; continue; }
      if (l.type === 'qt') p[l.qt.mode === 'do' ? 'doH' : 'useH'] += Number(l.qt.hours) || 0; else p[l.type]++;
    }
    const plist = Object.values(people).sort((a, b) => a.name.localeCompare(b.name, 'th'));
    return `
<div class="row wrap" style="margin-bottom:14px;gap:8px">
  <input class="input" style="max-width:240px;border-radius:999px" placeholder="ค้นหาชื่อ ชื่อเล่น แผนก เลขที่…" value="${H(f.q)}" data-change="lvInfo" data-k="q" aria-label="ค้นหา">
  ${sel('month', [['', 'ทุกเดือน']].concat(months.map(m => [m, U.ML[+m.slice(5) - 1] + ' ' + (+m.slice(0, 4) + 543)])))}
  ${sel('type', [['all', 'ทุกประเภท']].concat(Object.entries(TYPES).map(([k, t]) => [k, t.name])))}
  ${sel('status', [['all', 'ทุกสถานะ'], ['wait', 'รออนุมัติ'], ['ok', 'อนุมัติแล้ว'], ['done', 'ฝ่ายบุคคลรับทราบแล้ว'], ['rejected', 'ไม่อนุมัติ'], ['cancelled', 'ยกเลิก']])}
  ${sel('dept', [['all', 'ทุกแผนก']].concat(depts.map(d => [d, d])))}
  <button class="btn ghost" data-act="lvExport" style="margin-left:auto">ส่งออก Excel</button>
</div>
<div class="grid g4" style="margin-bottom:16px">
  ${kpi('ใบแจ้งทั้งหมด', rows.length + ' <small>ใบ</small>', `รออนุมัติ ${rows.filter(l => ['pending', 'gm'].includes(l.status)).length} · รอฝ่ายบุคคล ${rows.filter(l => l.status === 'approved').length}`)}
  ${kpi('QT (อนุมัติแล้ว)', `${qtH('do')} / ${qtH('use')} <small>ชม.</small>`, 'ทำ QT / ลา QT')}
  ${kpi('เปลี่ยนวันหยุด / เวลา', rows.filter(l => l.type === 'change').length + ' <small>ใบ</small>', `อนุมัติแล้ว ${rows.filter(l => l.type === 'change' && ok(l)).length}`)}
  ${kpi('ลืมสแกนนิ้ว', rows.filter(l => l.type === 'scan').length + ' <small>ใบ</small>', `อนุมัติแล้ว ${rows.filter(l => l.type === 'scan' && ok(l)).length}`)}
</div>
<section class="card" style="overflow:hidden;margin-bottom:16px">
  <div style="padding:14px 18px 4px"><h3>สรุปรายบุคคล</h3><div class="xs muted">นับเฉพาะใบที่อนุมัติแล้ว · คงเหลือ QT = ทำ QT − ลา QT</div></div>
  <div style="overflow-x:auto"><table class="tbl"><thead><tr><th>ชื่อ-สกุล</th><th>แผนก</th><th class="r">ทำ QT (ชม.)</th><th class="r">ลา QT (ชม.)</th><th class="r">คงเหลือ QT</th><th class="r">เปลี่ยนวันหยุด/เวลา</th><th class="r">ลืมสแกนนิ้ว</th><th class="r">รออนุมัติ</th></tr></thead><tbody>
  ${plist.map(p => `<tr><td>${H(p.name)}${p.nick ? ` <span class="muted small">(${H(p.nick)})</span>` : ''}<div class="sub">${H(p.pos || '')}</div></td><td>${H(p.dept || '—')}</td><td class="r">${p.doH}</td><td class="r">${p.useH}</td>
    <td class="r b ${p.doH - p.useH < 0 ? 'red' : ''}">${p.doH - p.useH}</td><td class="r">${p.change}</td><td class="r">${p.scan}</td><td class="r">${p.wait || ''}</td></tr>`).join('') || '<tr><td colspan="8"><div class="empty">ไม่มีข้อมูลตามตัวกรอง</div></td></tr>'}
  </tbody></table></div>
</section>
<section class="card" style="overflow:hidden">
  <div style="padding:14px 18px 4px"><h3>รายละเอียดใบแจ้ง</h3><div class="xs muted">${rows.length} ใบ · กดที่แถวเพื่อดูใบและลงนาม</div></div>
  <div style="overflow-x:auto"><table class="tbl" style="font-size:13px;min-width:1100px"><thead><tr><th>เลขที่</th><th>วันที่มีผล</th><th>ชื่อ-สกุล</th><th>ตำแหน่ง / แผนก</th><th>ประเภท</th><th>รายละเอียด</th><th>เหตุผล / พยาน</th><th>ลงนาม</th><th>สถานะ</th></tr></thead><tbody>
  ${rows.map(l => `<tr data-act="lvOpen" data-id="${H(l.id)}" style="cursor:pointer">
    <td style="white-space:nowrap">${H(l.no)}<div class="sub">เขียน ${U.fmtD((l.at || '').slice(0, 10), { year: false })}</div></td><td style="white-space:nowrap">${dTH(dayOf(l))}</td>
    <td>${H(who(l))}${l.nick ? `<div class="sub">${H(l.nick)}</div>` : ''}</td><td>${H(l.position || '—')}<div class="sub">${H(l.dept || '')}${l.branch ? ' · ' + H(l.branch) : ''}</div></td>
    <td style="white-space:nowrap">${H(TYPES[l.type].name)}</td><td>${summary(l)}</td>
    <td>${H(l.reason || '')}${l.type === 'scan' && l.sc.witness ? `<div class="sub">พยาน ${H(l.sc.witness)}${l.sc.witnessPos ? ' · ' + H(l.sc.witnessPos) : ''}</div>` : ''}</td>
    <td class="xs">${STEPS.map(([k, t]) => signer(l, k) ? `<div><span class="muted">${t}:</span> ${H(signer(l, k))}</div>` : '').join('') || '<span class="muted">—</span>'}</td><td>${badge(l)}</td></tr>`).join('') || '<tr><td colspan="9"><div class="empty">ไม่มีใบแจ้งตามตัวกรอง</div></td></tr>'}
  </tbody></table></div>
</section>`;
  }
  App.CHG.lvInfo = el => { const f = App.state.ui.leaveInfo; f[el.dataset.k] = el.value; App.render(); };
  App.ACT.lvExport = () => {
    const me = App.me(); if (!mayHr(me)) return;
    const { f, rows } = infoRows(leaves().filter(l => visible(l, me)));
    const out = [['เลขที่', 'วันที่เขียน', 'วันที่มีผล', 'ชื่อ-สกุล', 'ชื่อเล่น', 'ตำแหน่ง', 'แผนก', 'สาขา', 'ประเภท', 'รายละเอียด', 'ชั่วโมง QT', 'เหตุผล', 'พยาน', 'หัวหน้าแผนก', 'ผู้จัดการทั่วไป', 'รองผู้จัดการทั่วไป', 'ฝ่ายบุคคล', 'สถานะ']]
      .concat(rows.map(l => [l.no, (l.at || '').slice(0, 10), dayOf(l) || '', who(l), l.nick || '', l.position || '', l.dept || '', l.branch || '', TYPES[l.type].name, detailText(l),
        l.type === 'qt' ? (l.qt.mode === 'use' ? -1 : 1) * (Number(l.qt.hours) || 0) : '', l.reason || '', l.type === 'scan' ? [l.sc.witness, l.sc.witnessPos].filter(Boolean).join(' · ') : '',
        signer(l, 'head'), signer(l, 'gm'), signer(l, 'dgm'), signer(l, 'hr'), (STATUS[l.status] || STATUS.pending)[0]]));
    const blob = g.Sheet.writeXlsx(out, { sheetName: 'ใบแจ้ง', widths: [20, 12, 12, 24, 10, 18, 18, 16, 16, 44, 10, 30, 22, 22, 22, 22, 22, 18] });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'ข้อมูลใบลา-' + (f.month || 'ทั้งหมด') + '.xlsx'; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  };

  /* ---------------- request form ---------------- */
  function form(type) {
    const me = App.me(), emp = Store.db.employees.find(e => e.name.replace(/\s+/g, ' ') === (me.name || '').replace(/\s+/g, ' '));
    const last = leaves().filter(l => l.by === me.id).sort((a, b) => (b.at || '').localeCompare(a.at || ''))[0] || {};
    const today = U.today(), t = TYPES[type];
    const fld = (label, html, span) => `<div style="min-width:0${span ? ';grid-column:1/-1' : ''}"><label class="f">${label}</label>${html}</div>`;
    const inp = (name, val, type2, extra) => `<input class="input" name="${name}" type="${type2 || 'text'}" value="${H(val == null ? '' : val)}" ${extra || ''}>`;
    const radio = (name, opts, cur) => `<div class="row wrap" style="gap:16px">${opts.map(([v, l]) => `<label class="row small"><input type="radio" name="${name}" value="${v}" ${v === cur ? 'checked' : ''} style="accent-color:var(--teal);width:18px;height:18px"> ${l}</label>`).join('')}</div>`;
    const body = type === 'qt' ? `
        ${fld('ประเภท', radio('mode', [['do', 'ทำ QT'], ['use', 'ลา QT (ใช้ QT)']], 'use'), 1)}
        ${fld('เวลา (น.)', inp('from', '', 'time'))}${fld('ถึงเวลา (น.)', inp('to', '', 'time'))}
        ${fld('ณ วันที่', inp('date', today, 'date'))}${fld('จำนวนชั่วโมง', `<input class="input" name="hours" type="number" min="2" step="0.5" readonly style="background:var(--bg2,#f4f3ee)">`)}`
      : type === 'change' ? `
        ${fld('ต้องการ', radio('mode', [['day', 'เปลี่ยนวันหยุด'], ['time', 'เปลี่ยนเวลาทำงาน']], 'day'), 1)}
        <div class="lv-day" style="display:contents">${fld('จากวันหยุดวันที่', inp('fromDate', today, 'date'))}${fld('เป็นวันที่', inp('toDate', '', 'date'))}</div>
        <div class="lv-time" style="display:none">${fld('จากเวลา (น.)', inp('fromTime', '', 'time'))}${fld('เป็นเวลา (น.)', inp('toTime', '', 'time'))}${fld('ณ วันที่', inp('date', today, 'date'), 1)}</div>`
      : `
        ${fld('สาเหตุที่ไม่ได้สแกนนิ้วเข้างาน/ออกงาน', radio('cause', Object.entries(CAUSES), 'forgot'), 1)}
        ${fld('เข้างานตามจริง เวลา (น.)', inp('inTime', '', 'time'))}${fld('ออกงาน เวลา (น.)', inp('outTime', '', 'time'))}
        ${fld('ณ วันที่', inp('date', today, 'date'), 1)}
        ${fld('พยาน ชื่อ', inp('witness', ''))}${fld('ตำแหน่งพยาน', inp('witnessPos', ''))}`;
    const bg = App.modal(`<h2>${t.name}</h2><p class="muted" style="margin-top:-8px">${t.sub}</p>
      <div class="grid g2" style="gap:12px 14px">
        ${fld('ชื่อ-สกุล', inp('fullName', last.fullName || ''))}${fld('ชื่อเล่น', inp('nick', last.nick || ''))}
        ${fld('ตำแหน่ง', inp('position', last.position || (emp && emp.role) || ''))}${fld('แผนก', inp('dept', last.dept || me.dept || ''))}
        ${fld('สาขา', inp('branch', last.branch || 'ลักษวรรณ รีสอร์ท'), 1)}
        <div style="grid-column:1/-1;border-top:1px solid var(--line);margin:4px 0 0"></div>
        ${body}
        ${fld(type === 'scan' ? 'หมายเหตุ / สาเหตุเพิ่มเติม' : 'เหตุผล', `<textarea class="input" name="reason" rows="2"></textarea>`, 1)}
      </div>
      ${type === 'qt' ? '<div class="xs muted" style="margin-top:8px">ทำหรือใช้ QT ได้ตั้งแต่ 2 ชั่วโมงขึ้นไป · ห้ามทำหรือใช้ QT ก่อนเวลาเข้างาน (ถือเป็นการเปลี่ยนเวลาทำงาน)</div>' : ''}
      <div class="small red" id="lvErr" style="min-height:20px;margin-top:6px" role="alert"></div>
      <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="lvOk">ส่งใบแจ้ง</button></div>`, bg => {
      const q = n => bg.querySelector(`[name=${n}]`);
      if (type === 'qt') { const upd = () => { q('hours').value = hoursOf(q('from').value, q('to').value) || ''; }; q('from').oninput = q('to').oninput = upd; }
      if (type === 'change') bg.querySelectorAll('[name=mode]').forEach(r => r.onchange = () => {
        const day = bg.querySelector('[name=mode]:checked').value === 'day';
        bg.querySelector('.lv-day').style.display = day ? 'contents' : 'none'; bg.querySelector('.lv-time').style.display = day ? 'none' : 'contents';
      });
      bg.querySelector('#lvOk').onclick = () => {
        const f = App.formData(bg), err = bg.querySelector('#lvErr'), pick = n => (bg.querySelector(`[name=${n}]:checked`) || {}).value;
        const l = { id: U.uid('lv'), no: nextNo(), type, by: me.id, byName: me.name, fullName: f.fullName.trim(), nick: f.nick.trim(), position: f.position.trim(), dept: f.dept.trim(), branch: f.branch.trim(),
          reason: f.reason.trim(), at: U.nowTH(), status: 'pending', sign: {} };
        if (!l.fullName) { err.textContent = 'กรุณากรอกชื่อ-สกุล'; return; }
        if (type === 'qt') {
          const hours = hoursOf(f.from, f.to);
          if (!f.from || !f.to || !f.date) { err.textContent = 'กรุณากรอกเวลาและวันที่'; return; }
          if (hours < 2) { err.textContent = 'ทำหรือใช้ QT ต้อง 2 ชั่วโมงขึ้นไป'; return; }
          l.qt = { mode: pick('mode'), from: f.from, to: f.to, date: f.date, hours };
        } else if (type === 'change') {
          const mode = pick('mode');
          if (mode === 'day' && (!f.fromDate || !f.toDate)) { err.textContent = 'กรุณาเลือกวันหยุดเดิมและวันหยุดใหม่'; return; }
          if (mode === 'day' && f.fromDate === f.toDate) { err.textContent = 'วันหยุดใหม่ต้องไม่ใช่วันเดิม'; return; }
          if (mode === 'time' && (!f.fromTime || !f.toTime || !f.date)) { err.textContent = 'กรุณากรอกเวลาเดิม เวลาใหม่ และวันที่'; return; }
          l.ch = mode === 'day' ? { mode, fromDate: f.fromDate, toDate: f.toDate } : { mode, fromTime: f.fromTime, toTime: f.toTime, date: f.date };
        } else {
          if (!f.date || (!f.inTime && !f.outTime)) { err.textContent = 'กรุณากรอกวันที่ และเวลาเข้างานหรือออกงานตามจริง'; return; }
          l.sc = { cause: pick('cause'), inTime: f.inTime, outTime: f.outTime, date: f.date, witness: f.witness.trim(), witnessPos: f.witnessPos.trim() };
        }
        if (!l.reason && type !== 'scan') { err.textContent = 'กรุณากรอกเหตุผล'; return; }
        leaves().push(l); Store.save(); App.closeModal();
        App.state.ui.leave = { tab: 'mine' }; if (App.state.route !== 'leave') App.go('#/leave'); else App.render();
        App.toast('ส่ง' + TYPES[type].name + ' ' + l.no + ' แล้ว');
      };
    });
    bg.querySelector('.modal').style.maxWidth = '680px';
  }

  /* ---------------- one request ---------------- */
  function view(l) {
    const me = App.me(), s = l.sign || {};
    const line = ([k, label]) => {
      const x = s[k];
      return `<div class="lv-step ${x ? 'on' : ''}" style="display:flex;gap:10px;align-items:flex-start;padding:8px 0;border-top:1px solid var(--line)">
        <i class="dot" style="margin-top:6px;background:${x ? (x.ok === false ? 'var(--red)' : 'var(--teal)') : '#cfccc2'}"></i>
        <div style="flex:1"><b>${label}</b><div class="small muted">${x ? `${x.ok === false ? 'ไม่อนุมัติ' : k === 'gm' ? 'อนุมัติ' : 'ลงนามแล้ว'} · ${H(x.byName)} · ${H((x.at || '').slice(0, 16).replace('T', ' '))}${x.note ? ' · ' + H(x.note) : ''}` : k === 'dgm' ? 'ไม่บังคับ' : 'รอลงนาม'}</div></div></div>`;
    };
    const can = {
      head: l.status === 'pending' && maySign(me),
      gm: l.status === 'gm' && maySign(me),
      dgm: (l.status === 'approved' || l.status === 'done') && !s.dgm && maySign(me),
      hr: l.status === 'approved' && maySign(me)
    };
    const acts = [];
    if (can.head) acts.push(`<button class="btn ghost red" data-act="lvSign" data-id="${H(l.id)}" data-step="head" data-ok="0">ไม่อนุมัติ</button><button class="btn teal" data-act="lvSign" data-id="${H(l.id)}" data-step="head" data-ok="1">ลงนามหัวหน้าแผนก</button>`);
    if (can.gm) acts.push(`<button class="btn ghost red" data-act="lvSign" data-id="${H(l.id)}" data-step="gm" data-ok="0">ไม่อนุมัติ</button><button class="btn teal" data-act="lvSign" data-id="${H(l.id)}" data-step="gm" data-ok="1">อนุมัติ (ผู้จัดการทั่วไป)</button>`);
    if (can.dgm) acts.push(`<button class="btn ghost" data-act="lvSign" data-id="${H(l.id)}" data-step="dgm" data-ok="1">ลงนามรองผู้จัดการทั่วไป</button>`);
    if (can.hr) acts.push(`<button class="btn teal" data-act="lvSign" data-id="${H(l.id)}" data-step="hr" data-ok="1">ฝ่ายบุคคลรับทราบ</button>`);
    if (l.by === me.id && (l.status === 'pending' || l.status === 'gm')) acts.unshift(`<button class="btn ghost red" data-act="lvCancel" data-id="${H(l.id)}">ยกเลิกใบนี้</button>`);
    const kv = (k, v) => `<div class="tk-kv"><span class="muted">${k}</span><span>${v}</span></div>`;
    App.modal(`<div class="between" style="align-items:flex-start;gap:10px"><div><h2 style="margin:0">${H(TYPES[l.type].name)}</h2><div class="small muted">${H(l.no)}</div></div>${badge(l)}</div>
      <div class="tk-kvs" style="margin-top:12px">
        ${kv('ผู้ขอ', `${H(who(l))}${l.nick ? ' (' + H(l.nick) + ')' : ''}`)}${kv('ตำแหน่ง/แผนก', H([l.position, l.dept].filter(Boolean).join(' · ') || '—'))}
        ${kv('รายละเอียด', summary(l))}${l.type === 'scan' && l.sc.witness ? kv('พยาน', H(l.sc.witness + (l.sc.witnessPos ? ' · ' + l.sc.witnessPos : ''))) : ''}
        ${l.reason ? kv(l.type === 'scan' ? 'หมายเหตุ' : 'เหตุผล', H(l.reason)) : ''}${kv('วันที่เขียน', H((l.at || '').slice(0, 16).replace('T', ' ')))}
      </div>
      <div style="margin-top:12px">${STEPS.map(line).join('')}</div>
      <div class="acts" style="flex-wrap:wrap"><a class="btn ghost" href="#/leaveprint/${encodeURIComponent(l.id)}" data-act="closeModal">พิมพ์ฟอร์ม</a>${acts.join('')}<button class="btn ghost" data-act="closeModal">ปิด</button></div>`);
  }

  Object.assign(App.ACT, {
    lvTab(el) { App.state.ui.leave.tab = el.dataset.v; App.render(); },
    lvNew(el) { form(el.dataset.v); },
    lvOpen(el) { const l = leaves().find(x => x.id === el.dataset.id); if (l) view(l); },
    lvCancel(el) {
      const l = leaves().find(x => x.id === el.dataset.id); if (!l) return;
      App.confirm('ยกเลิก' + TYPES[l.type].name + ' ' + l.no + '?', () => { l.status = 'cancelled'; Store.save(); App.render(); App.toast('ยกเลิกแล้ว'); }, 'ยกเลิกใบนี้');
    },
    lvSign(el) {
      const l = leaves().find(x => x.id === el.dataset.id), me = App.me(), step = el.dataset.step, ok = el.dataset.ok === '1'; if (!l) return;
      const allowed = maySign(me) && { head: l.status === 'pending', gm: l.status === 'gm', dgm: !(l.sign || {}).dgm, hr: l.status === 'approved' }[step];
      if (!allowed) { App.toast('คุณลงนามขั้นนี้ไม่ได้'); return; }
      const done = note => {
        l.sign = Object.assign({}, l.sign, { [step]: { by: me.id, byName: me.name, at: U.nowTH(), ok, note: note || '' } });
        if (!ok) l.status = 'rejected';
        else if (step === 'head') l.status = 'gm';
        else if (step === 'gm') l.status = 'approved';
        else if (step === 'hr') l.status = 'done';
        Store.save(); App.closeModal(); App.render(); App.toast(ok ? 'ลงนามแล้ว' : 'ไม่อนุมัติแล้ว');
      };
      if (ok) return done();
      App.modal(`<h2>ไม่อนุมัติ ${H(l.no)}</h2><div class="field"><label class="f" for="lvNote">เหตุผล</label><textarea class="input" id="lvNote" rows="3"></textarea></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="lvNo">ไม่อนุมัติ</button></div>`, bg => bg.querySelector('#lvNo').onclick = () => done(bg.querySelector('#lvNote').value.trim()));
    }
  });

  /* ---------------- the paper form ---------------- */
  App.V.leaveprint = {
    nav: 'leave', title: 'ใบขอใช้สิทธิการลา',
    render(p) {
      const l = leaves().find(x => x.id === p[0]), me = App.me();
      if (!l || !visible(l, me)) return '<div class="empty">ไม่พบใบแจ้ง — <a href="#/leave">กลับ</a></div>';
      const s = l.sign || {}, box = on => `<span style="display:inline-block;width:13px;height:13px;border:1px solid #333;margin-right:5px;vertical-align:-2px;text-align:center;line-height:11px;font-size:11px">${on ? '✓' : ''}</span>`;
      const u = v => `<span style="display:inline-block;min-width:60px;border-bottom:1px dotted #555;padding:0 6px;text-align:center">${H(v == null || v === '' ? '' : v)}</span>`;
      const dt = iso => { const [d, m, y] = be(iso); return `วันที่ ${u(d)} เดือน ${u(m)} พ.ศ. ${u(y)}`; };
      const qt = l.type === 'qt' ? l.qt : {}, ch = l.type === 'change' ? l.ch : {}, sc = l.type === 'scan' ? l.sc : {};
      const why = t => l.type === t ? u(l.reason) : u('');
      const sig = (label, x, decide) => `<td style="vertical-align:top;padding:8px 10px">${decide ? `<div style="margin-bottom:6px">${box(x && x.ok !== false)}อนุมัติ &nbsp; ${box(x && x.ok === false)}ไม่อนุมัติ</div>` : ''}
        <div style="height:26px;border-bottom:1px dotted #555;text-align:center;font-family:inherit">${x ? H(x.byName) : ''}</div>
        <div style="text-align:center">${label}<br>(${x ? H(x.byName) : '&nbsp;'.repeat(30)})<br>วันที่เขียน ${x ? H(U.fmtD(x.at.slice(0, 10))) : '..... / ..... / .....'}</div></td>`;
      const row = html => `<tr><td colspan="3" style="padding:7px 10px;line-height:2">${html}</td></tr>`;
      return `<div class="page-hd no-print"><div><div class="small"><a class="link" href="#/leave">← การแจ้งใบลา</a></div><h1>${H(TYPES[l.type].name)} ${H(l.no)}</h1></div>
        <div class="row">${badge(l)}<button class="btn red" onclick="window.print()">พิมพ์ / PDF</button></div></div>
      <article class="doc lv-doc" style="font-size:12.5px">
        <table style="width:100%;border-collapse:collapse" border="1" cellpadding="0">
          <tr><td colspan="2" style="padding:8px 10px;text-align:center"><b style="font-size:14px">เอกสารใบขอใช้สิทธิการลา QT การเปลี่ยนแปลงเวลาทำงาน และ แจ้งเวลาเข้า-ออกงาน</b></td><td style="padding:8px 10px;width:30%">เลขที่เอกสาร <b>${H(l.no)}</b></td></tr>
          ${row(`ชื่อ-สกุล ${u(who(l))} ชื่อเล่น ${u(l.nick)} ตำแหน่ง/แผนก ${u([l.position, l.dept].filter(Boolean).join(' / '))} สาขา ${u(l.branch)}`)}
          ${row('<b>เปลี่ยนแปลงวันหยุด / เปลี่ยนแปลงเวลาทำงาน</b> <span style="color:#555">**ตรวจสอบจากตารางเวรที่มีการประกาศทุกครั้ง</span>')}
          ${row(`${box(l.type === 'qt' && qt.mode === 'do')}ทำ QT ${u(qt.mode === 'do' ? qt.hours : '')} ชม. เวลา ${u(qt.mode === 'do' ? qt.from : '')} น. ถึง เวลา ${u(qt.mode === 'do' ? qt.to : '')} น. ณ ${dt(qt.mode === 'do' ? qt.date : '')} เหตุผล ${qt.mode === 'do' ? why('qt') : u('')}`)}
          ${row(`${box(l.type === 'qt' && qt.mode === 'use')}ลา QT ${u(qt.mode === 'use' ? qt.hours : '')} ชม. เวลา ${u(qt.mode === 'use' ? qt.from : '')} น. ถึง เวลา ${u(qt.mode === 'use' ? qt.to : '')} น. ณ ${dt(qt.mode === 'use' ? qt.date : '')} เหตุผล ${qt.mode === 'use' ? why('qt') : u('')}`)}
          ${row(`${box(ch.mode === 'time')}เปลี่ยนเวลาทำงาน จาก เวลา ${u(ch.fromTime)} น. เป็น เวลา ${u(ch.toTime)} น. ณ ${dt(ch.mode === 'time' ? ch.date : '')} เหตุผล ${ch.mode === 'time' ? why('change') : u('')}`)}
          ${row(`${box(ch.mode === 'day')}เปลี่ยนวันหยุด จาก ${dt(ch.fromDate)} เป็น ${dt(ch.toDate)} เหตุผล ${ch.mode === 'day' ? why('change') : u('')}`)}
          ${row(`<b>สแกนนิ้วไม่ติด/ลืมสแกนนิ้ว/เริ่มงานโดยยังไม่มีลายนิ้วมือ</b>${l.type === 'scan' ? ' — ' + Object.entries(CAUSES).map(([k, v]) => box(sc.cause === k) + v).join(' &nbsp; ') : ''}`)}
          ${row(`ไม่ได้สแกนนิ้วเข้างาน/ออกงาน สาเหตุเพราะ ${u(l.type === 'scan' ? [CAUSES[sc.cause], l.reason].filter(Boolean).join(' · ') : '')} เข้างานตามจริง เข้างานเวลา ${u(sc.inTime)} น. ออกงาน เวลา ${u(sc.outTime)} น. ณ ${dt(sc.date)} พยาน ชื่อ ${u(sc.witness)} ลายเซ็น ${u('')} ตำแหน่ง ${u(sc.witnessPos)}`)}
          <tr>${sig('พนักงาน', { byName: who(l), at: l.at })}${sig('หัวหน้าแผนก', s.head)}${sig('ผู้จัดการทั่วไป', s.gm, true)}</tr>
          <tr>${sig('รองผู้จัดการทั่วไป', s.dgm)}${sig('ฝ่ายบุคคล', s.hr)}<td></td></tr>
          ${row('<b>หมายเหตุ:</b> อนุญาตให้ทำหรือใช้ QT จะต้องทำหรือใช้ 2 ชั่วโมงขึ้นไป: ไม่อนุญาตให้ทำหรือใช้ QT ก่อนเวลาเข้างาน จะถือว่าเปลี่ยนเวลาทำงาน')}
        </table>
        <div style="text-align:right;color:#555;font-size:11px;margin-top:6px">เอกสารใบขอใช้สิทธิการลาQT การเปลี่ยนแปลงเวลาทำงาน และ แจ้งเวลาเข้า-ออกงาน · บริษัท ลักษวรรณ จำกัด · แก้ไข ณ วันที่ 22 มิถุนายน 2569</div>
      </article>`;
    }
  };
})(window);
