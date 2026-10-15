/* Luck Swan Sales — ใบสรุปรายละเอียดสำหรับฝ่ายขาย (function income summary).
 * Admins, sales managers and sales staff only. Figures come from the quote at prices that already include
 * the discount, service charge and VAT, so the rows add up to the quote's net total; numbers can be edited,
 * and OV, tax invoices and notes are typed in. Kept on the function sheet (b.ss), outside the approval flow. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;

  const ROLES = ['admin', 'manager', 'sales'];
  const allowed = () => { const me = App.me(); return !!me && ROLES.includes(me.role); };
  App.canSalesSum = allowed;
  const curB = () => Store.db.beos.find(b => b.id === App.state.params[0]);
  const num = v => Number(String(v == null ? '' : v).replace(/,/g, '')) || 0;
  const money = v => U.n(Math.round(num(v) * 100) / 100, Math.round(num(v)) === num(v) ? 0 : 2);

  // sections of the paper form: [section, category, Thai, English, column headings]
  const SEC = {
    food: { name: 'ค่าอาหาร', cols: ['ราคาต่อมื้อ', 'จำนวนมื้อ', 'จำนวนคน'], cats: [['breakfast', 'เช้า', 'Breakfast'], ['lunch', 'กลางวัน', 'Lunch'], ['dinner', 'เย็น', 'Dinner']] },
    snack: { name: 'ค่าอาหารว่าง', cols: ['ราคาต่อมื้อ', 'จำนวนมื้อ', 'จำนวนคน'], cats: [['coffee', 'อาหารว่าง', 'Coffee Break'], ['comp', 'อภินันทนาการ', 'Complimentary'], ['other', 'อื่นๆ', 'Other']] },
    room: { name: 'ค่าห้องพัก', cols: ['ราคาต่อห้อง', 'จำนวนห้อง', 'จำนวนคืน'], cats: [['room', 'ห้องพัก', 'Room Charge'], ['extrabed', 'เตียงเสริม', 'Extra Bed'], ['comp', 'อภินันทนาการ', 'Complimentary'], ['other', 'อื่นๆ', 'Other']] },
    other: { name: 'รายรับอื่นๆ', cols: ['ราคา', 'จำนวน', 'ครั้ง/วัน'], cats: [['other', 'อื่นๆ', 'Other']] }
  };
  const total = r => Math.round(num(r.price) * num(r.a) * (r.b === '' || r.b == null ? 1 : num(r.b)) * 100) / 100;

  /* ---------- rows from the quote (prices including discount, service charge and VAT) ---------- */
  function autoRows(q) {
    const c = App.calc(q), f = c.sub ? c.net / c.sub : 1, rows = [];
    const inc = v => num(v) * f; // kept unrounded so the rows add up to the quote's net total
    for (const i of q.items || []) {
      const food = i.sec === 'food' ? D.FOOD.find(x => x.id === i.ref) : null, s = (i.sched || [])[0], t = s ? s.start : '';
      const nm = (i.name + ' ' + (i.desc || '')).toLowerCase();
      let sec = 'other', cat = 'other';
      if (i.sec === 'food') {
        if (i.free) { sec = 'snack'; cat = 'comp'; }
        else if ((food && food.cat === 'coffee') || /coffee|เบรก|อาหารว่าง/.test(nm)) { sec = 'snack'; cat = 'coffee'; }
        else { sec = 'food'; cat = /เช้า|breakfast/.test(nm) || (t && t < '10:30') ? 'breakfast' : /เย็น|dinner|gala|ขันโตก|wedding|โต๊ะจีน|party/.test(nm) || (food && /gala|wedding|cocktail/.test(food.cat)) || (t && t >= '16:00') ? 'dinner' : 'lunch'; }
      } else if (i.sec === 'room') { sec = 'room'; cat = i.free ? 'comp' : /เตียงเสริม|extra bed/.test(nm) ? 'extrabed' : 'room'; }
      const two = i.q2 != null && i.q2 !== '';
      // food: price × meals × people; rooms: price × rooms × nights; other: price × qty × days
      const row = i.sec === 'food' ? { price: inc(i.price), a: two ? i.q2 : 1, b: i.q1 } : { price: inc(i.price), a: i.q1, b: two ? i.q2 : '' };
      rows.push(Object.assign({ sec, cat, label: i.name + (i.free ? ' (ฟรี)' : '') }, row, i.free ? { price: 0 } : {}));
    }
    return rows;
  }
  function data(b, q) {
    const s = b.ss || {}, auto = autoRows(q);
    const rows = Array.isArray(s.rows) ? s.rows : auto;
    const sum = rows.reduce((a, r) => a + total(r), 0), ov = num(s.ov);
    const owner = Store.db.users.find(u => u.id === q.owner);
    return { s, rows, custom: Array.isArray(s.rows), sum, ov, received: sum + ov, no: (App.fnNo && App.fnNo(b)) || b.id, sales: owner ? owner.name : '-', inv: s.inv || [], invSum: (s.inv || []).reduce((a, x) => a + num(x.amt), 0) };
  }

  /* ---------- the paper form (A4 portrait) ---------- */
  function paper(d, q) {
    const sec = k => { const S = SEC[k]; return `<table class="ssf-t"><thead><tr><th class="l">${S.name}</th>${S.cols.map(c => `<th>${c}</th>`).join('')}<th>รวม</th></tr></thead><tbody>
      ${S.cats.map(([cat, th, en]) => { const rs = d.rows.filter(r => r.sec === k && r.cat === cat);
        return `<tr class="ssf-cat"><td class="l" rowspan="${Math.max(1, rs.length)}"><b>${th}</b><br><small>(${en})</small></td>${rs.length ? row1(rs[0]) : '<td></td><td></td><td></td><td></td>'}</tr>${rs.slice(1).map(r => `<tr>${row1(r)}</tr>`).join('')}`; }).join('')}
      ${k === 'other' ? `<tr><td class="l"><b>OV</b></td><td colspan="2" class="r">${d.ov ? money(d.ov) + ' บาท' : 'บาท'}</td><td class="r">${H(d.s.ovPct || '')} %</td><td></td></tr>` : ''}
      </tbody></table>`; };
    const row1 = r => `<td class="r"><div class="ssf-lbl">${H(r.label || '')}</div>${r.price === '' ? '' : money(r.price)}</td><td class="r">${H(String(r.a == null ? '' : r.a))}</td><td class="r">${H(String(r.b == null ? '' : r.b))}</td><td class="r b">${money(total(r))}</td>`;
    return `<div class="ssf-paper" id="ssPaper">
  <h2>ใบสรุปรายละเอียดงานฟังก์ชั่น</h2>
  <div class="ssf-kv"><b>ฟังก์ชั่น เลขที่ :</b> <span>${H(d.no)}</span></div>
  <div class="ssf-kv"><b>พนักงานขาย :</b> <span>${H(d.sales)}</span></div>
  <div class="ssf-kv"><b>ยอดรวมรายรับจากฟังก์ชั่น :</b> <span>${money(d.received)}</span> <b>บาท</b></div>
  <div class="ssf-sub">${H(q.client || '')} · ${H(q.title || '')} · ${U.fmtRange(q.start, q.end)} · ${H(q.id)}</div>
  <h3>รายละเอียดการรับเงิน</h3>
  ${['food', 'snack', 'room', 'other'].map(sec).join('')}
  <div class="ssf-sum"><b>รวมจำนวนเงินที่รับเงิน</b> <span>${money(d.received)}</span> <b>บาท</b></div>
  <h3>รายละเอียดใบกำกับภาษี</h3>
  ${(d.inv.length ? d.inv : [{}, {}, {}, {}]).map(x => `<div class="ssf-inv"><span>เล่มที่/เลขที่</span><span class="dot">${H(x.no || '')}</span><span>จำนวนเงินรวม</span><span class="dot r">${x.amt ? money(x.amt) : ''}</span><span>บาท</span></div>`).join('')}
  <div class="ssf-sum"><b>รวมยอดออกใบกำกับภาษี</b> <span>${d.invSum ? money(d.invSum) : ''}</span> <b>บาท</b></div>
  <h3>หมายเหตุอื่นๆ :</h3><div class="ssf-note">${H(d.s.note || '').replace(/\n/g, '<br>') || '<br><br>'}</div>
</div>`;
  }

  /* ---------- the page ---------- */
  App.V.ss = {
    title: 'ใบสรุปสำหรับฝ่ายขาย', nav: 'beo',
    render() {
      if (!allowed()) return '<div class="card pad" style="max-width:520px;margin:40px auto;text-align:center"><h2>เฉพาะฝ่ายขาย</h2><p class="muted">ใบสรุปสำหรับฝ่ายขายเปิดได้เฉพาะผู้ดูแลระบบ ผู้จัดการฝ่ายขาย และพนักงานขาย</p><a class="btn ghost" href="#/beo">กลับ</a></div>';
      const b = curB(), q = b && Store.quote(b.quoteId);
      if (!b || !q) return '<div class="empty">ไม่พบใบฟังก์ชัน — <a href="#/beo">กลับ</a></div>';
      const d = data(b, q), ui = App.state.ui.ss || (App.state.ui.ss = { tab: 'preview' });
      const head = `<div class="page-hd no-print"><div><div class="small"><a class="link" href="#/beo/${encodeURIComponent(b.id)}">← ใบฟังก์ชัน ${H(b.id)}</a></div><h1>ใบสรุปสำหรับฝ่ายขาย</h1>
        <div class="small muted">${H(q.title)} · ${H(q.client)} · ฟังก์ชั่น ${H(d.no)} · รายรับ ฿${money(d.received)}</div></div>
        <div class="row wrap"><div class="seg white"><button class="${ui.tab === 'preview' ? 'on' : ''}" data-act="ssTab" data-v="preview">ตัวอย่างพิมพ์</button><button class="${ui.tab === 'form' ? 'on' : ''}" data-act="ssTab" data-v="form">แก้ไขข้อมูล</button></div>
        <button class="btn red" data-act="ssPrint">พิมพ์ / PDF</button></div></div>`;
      if (ui.tab === 'preview') return head + `<div class="ssf-wrap">${paper(d, q)}</div>`;
      const c = App.calc(q);
      const rowEd = (r, i) => `<tr><td><input class="input" value="${H(r.label || '')}" data-change="ssRow" data-i="${i}" data-k="label" aria-label="รายการ"></td>
        <td><input class="input r" inputmode="decimal" value="${r.price === '' ? '' : Math.round(num(r.price) * 100) / 100}" data-change="ssRow" data-i="${i}" data-k="price" aria-label="ราคา"></td>
        <td><input class="input r" inputmode="decimal" value="${H(String(r.a == null ? '' : r.a))}" data-change="ssRow" data-i="${i}" data-k="a" aria-label="จำนวน"></td>
        <td><input class="input r" inputmode="decimal" value="${H(String(r.b == null ? '' : r.b))}" data-change="ssRow" data-i="${i}" data-k="b" aria-label="จำนวน"></td>
        <td class="r b num">${money(total(r))}</td><td><button class="icon-btn us-edit emp-del" data-act="ssDel" data-i="${i}" aria-label="ลบแถว">×</button></td></tr>`;
      const secEd = k => { const S = SEC[k]; return `<section class="card" style="overflow:hidden;margin-bottom:14px"><div class="card-hd"><h2>${S.name}</h2></div><div style="overflow-x:auto"><table class="tbl ss-ed">
        <thead><tr><th>รายการ</th><th class="r">${S.cols[0]}</th><th class="r">${S.cols[1]}</th><th class="r">${S.cols[2]}</th><th class="r">รวม</th><th></th></tr></thead><tbody>
        ${S.cats.map(([cat, th, en]) => `<tr class="ss-cat"><td colspan="6"><b>${th}</b> <span class="xs muted">${en}</span> <button class="link xs" data-act="ssAdd" data-s="${k}" data-c="${cat}">+ เพิ่มแถว</button></td></tr>` +
          d.rows.map((r, i) => r.sec === k && r.cat === cat ? rowEd(r, i) : '').join('')).join('')}
        </tbody></table></div></section>`; };
      return head + `<div class="note" style="margin-bottom:14px">ตัวเลขดึงจากใบเสนอราคา ${H(q.id)} โดยใช้ราคาที่รวมส่วนลด ค่าบริการ 10% และ VAT 7% แล้ว (ยอดรวมจากระบบ ฿${money(c.net)}) — แก้ตัวเลข/เพิ่มแถวได้ ${d.custom ? '· <button class="link" data-act="ssReset">ใช้ตัวเลขจากระบบ</button>' : ''}</div>
        <div class="ss-grid"><div>${['food', 'snack', 'room', 'other'].map(secEd).join('')}</div>
        <div><section class="card pad" style="margin-bottom:14px"><h2>OV</h2>
          <div class="grid g2"><div><label class="f" for="ssOv">จำนวนเงิน (บาท)</label><input class="input" id="ssOv" inputmode="decimal" value="${H(d.s.ov || '')}" data-change="ssSet" data-k="ov"></div>
          <div><label class="f" for="ssOvPct">%</label><input class="input" id="ssOvPct" inputmode="decimal" value="${H(d.s.ovPct || '')}" data-change="ssSet" data-k="ovPct"></div></div>
          <div class="xs muted" style="margin-top:6px">ตอนนี้ OV รวมเข้าในยอดรับเงิน</div></section>
        <section class="card pad" style="margin-bottom:14px"><div class="between"><h2>รวมจำนวนเงินที่รับเงิน</h2><b class="num" style="font:700 20px var(--fh)">฿${money(d.received)}</b></div>
          <div class="small muted">รายการ ฿${money(d.sum)}${d.ov ? ' + OV ฿' + money(d.ov) : ''}</div></section>
        <section class="card pad" style="margin-bottom:14px"><div class="between"><h2>ใบกำกับภาษี</h2><button class="link" data-act="ssInvAdd">+ เพิ่ม</button></div>
          ${d.inv.map((x, i) => `<div class="row" style="gap:6px;margin-top:8px"><input class="input" placeholder="เล่มที่/เลขที่" value="${H(x.no || '')}" data-change="ssInv" data-i="${i}" data-k="no" aria-label="เล่มที่/เลขที่"><input class="input r" style="max-width:140px" inputmode="decimal" placeholder="จำนวนเงิน" value="${H(x.amt || '')}" data-change="ssInv" data-i="${i}" data-k="amt" aria-label="จำนวนเงินรวม"><button class="icon-btn us-edit emp-del" data-act="ssInvDel" data-i="${i}" aria-label="ลบ">×</button></div>`).join('') || '<div class="small muted" style="margin-top:6px">ยังไม่มีใบกำกับภาษี</div>'}
          <div class="between small" style="margin-top:10px"><span class="muted">รวมยอดออกใบกำกับภาษี</span><b>฿${money(d.invSum)}</b></div>
          ${d.inv.length && Math.round(d.invSum) !== Math.round(d.received) ? `<div class="xs red" style="margin-top:4px">ต่างจากยอดรับเงิน ฿${money(d.received - d.invSum)}</div>` : ''}</section>
        <section class="card pad"><h2>หมายเหตุอื่นๆ</h2><textarea class="input" rows="4" style="height:auto;padding-top:10px" data-change="ssSet" data-k="note" aria-label="หมายเหตุอื่นๆ">${H(d.s.note || '')}</textarea></section></div></div>`;
    }
  };

  function setSs(fn) {
    const b = curB(); if (!b || !allowed()) return;
    const q = Store.quote(b.quoteId);
    b.ss = Object.assign({}, b.ss); fn(b.ss, () => (b.ss.rows = Array.isArray(b.ss.rows) ? b.ss.rows.map(r => Object.assign({}, r)) : autoRows(q)));
    Store.save(); App.render();
  }
  const clean = (k, v) => k === 'label' ? String(v).slice(0, 120) : String(v).replace(/[^\d.,-]/g, '').slice(0, 14);
  Object.assign(App.ACT, {
    ssTab(el) { App.state.ui.ss.tab = el.dataset.v; App.render(); },
    ssPrint() { const ui = App.state.ui.ss || (App.state.ui.ss = {}); if (ui.tab !== 'preview') { ui.tab = 'preview'; App.render(); } setTimeout(() => window.print(), 50); },
    ssReset() { App.confirm('ใช้ตัวเลขจากใบเสนอราคาแทนตัวเลขที่แก้ไว้?', () => setSs(s => { delete s.rows; }), 'ใช้ค่าจากระบบ'); },
    ssAdd(el) { setSs((s, rows) => { rows().push({ sec: el.dataset.s, cat: el.dataset.c, label: '', price: '', a: 1, b: el.dataset.s === 'other' ? '' : 1 }); }); },
    ssDel(el) { setSs((s, rows) => { rows().splice(Number(el.dataset.i), 1); }); },
    ssInvAdd() { setSs(s => { s.inv = (s.inv || []).concat({ no: '', amt: '' }); }); },
    ssInvDel(el) { setSs(s => { s.inv = (s.inv || []).filter((x, i) => i !== Number(el.dataset.i)); }); }
  });
  Object.assign(App.CHG, {
    ssRow(el) { setSs((s, rows) => { const r = rows()[Number(el.dataset.i)]; if (r) r[el.dataset.k] = clean(el.dataset.k, el.value); }); },
    ssSet(el) { setSs(s => { const k = el.dataset.k; s[k] = k === 'note' ? el.value.slice(0, 2000) : clean(k, el.value); }); },
    ssInv(el) { setSs(s => { s.inv = (s.inv || []).map(x => Object.assign({}, x)); const x = s.inv[Number(el.dataset.i)]; if (x) x[el.dataset.k] = el.dataset.k === 'no' ? el.value.slice(0, 60) : clean('amt', el.value); }); }
  });
})(window);
