/* Luck Swan Sales — ออกแบบผังงาน: drag-and-drop floor plans for a function sheet (one plan per event day).
 * Plans live on the function sheet (b.plans[day]) so they sync, print and go through the same approval as the sheet. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;
  const S = 40; // px per metre

  // room sizes in metres (width × depth)
  const ROOM_DIM = { kinnaree: [24, 18], thikhampon: [13.9, 12.1], ruenhong: [20, 7], prayamengrai: [10, 7], pool: [30, 20] };
  const NAMES = { round: 'โต๊ะกลม', rect: 'โต๊ะยาว', chairrow: 'แถวเก้าอี้', class: 'โต๊ะห้องเรียน', stage: 'เวที', screen: 'จอ', podium: 'โพเดียม', lect: 'โต๊ะบรรยาย', sound: 'คุมเสียง', buffet: 'บุฟเฟ่ต์', reg: 'ลงทะเบียน', sofa: 'โซฟา', carpet: 'พรมแดง', door: 'ประตู' };
  const SEATED = ['round', 'rect', 'chairrow', 'class'];
  const FILL = { stage: '#c9d6e8', buffet: '#e8d3b9', screen: '#555', door: '#fff', sound: '#ddd', podium: '#e6d5ef', sofa: '#f0c9c0', reg: '#d6e8d0', lect: '#f1e2b8', carpet: '#c0392b' };
  const rlen = i => Math.max(1.5, Math.ceil(i.seats / 2) * 0.75);
  function dims(i) {
    switch (i.type) {
      case 'round': return [2.9, 2.9]; case 'rect': return [rlen(i) + 0.2, 2.1]; case 'chairrow': return [i.seats * 0.55, 0.5];
      case 'stage': return [7, 2.4]; case 'buffet': return [3, 0.8]; case 'screen': return [3, 0.25]; case 'sound': return [1.5, 0.8];
      case 'podium': return [0.6, 0.6]; case 'sofa': return [2, 0.8]; case 'reg': return [2.4, 0.7]; case 'class': return [i.seats * 0.6, 1.1];
      case 'lect': return [2.4, 0.9]; case 'carpet': return [2.4, 14.4]; default: return [1, 0.3];
    }
  }
  const seatsDefault = t => ({ round: 10, rect: 8, chairrow: 10, class: 3 }[t] || 0);

  /* ---------- plan data ---------- */
  const curBeo = () => Store.db.beos.find(b => b.id === App.state.params[0]);
  const dayOf = () => Math.max(1, Number(App.state.params[1]) || 1);
  const canEdit = () => App.can('create_quote') || App.can('staff');
  function planOf(b, day) { return b && b.plans && b.plans[day]; }
  function mk(p, type, x, y, rot = 0, seats) {
    const id = (p.items.reduce((m, i) => Math.max(m, i.id), 0) || 0) + 1;
    const it = { id, type, x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10, rot, seats: seats == null ? seatsDefault(type) : seats };
    p.items.push(it); return it;
  }
  const seatsOf = p => p.items.reduce((a, i) => a + (SEATED.includes(i.type) ? i.seats : 0), 0);
  // the meeting room and layout used on a given day of the quote
  function roomFor(q, day) {
    const meets = q.items.filter(i => i.sec === 'meet' && D.MEETING_ROOMS.some(r => r.id === i.ref));
    const it = meets.find(i => (i.sched || []).some(s => s.day === day)) || meets[0];
    const id = (it && it.ref) || q.room;
    return { id, room: D.MEETING_ROOMS.find(r => r.id === id), layout: (it && it.layout) || 'banquet' };
  }

  /* ---------- generators ---------- */
  function addFront(p, w, withScreen = true) { mk(p, 'stage', w / 2, 1.5); if (withScreen) mk(p, 'screen', w / 2, 0.3); }
  function genBanquet(p, w, h, pax, per = 10) {
    addFront(p, w); mk(p, 'sound', Math.max(1, w - 2), 1.4);
    const sx = 3.2, sy = 2.9, nc = Math.max(1, Math.floor((w - 0.4) / sx)); let n = Math.ceil(pax / per);
    for (let r = 0; n > 0; r++) {
      const y = 4.6 + r * sy; if (y + 1.45 > h - 1) break;
      const c = r % 2 ? Math.max(1, nc - 1) : nc;
      for (let k = 0; k < c && n > 0; k++, n--) mk(p, 'round', w / 2 + (k - (c - 1) / 2) * sx, y, 0, per);
    }
    mk(p, 'buffet', 2, h - 0.8); mk(p, 'door', w - 1.5, h);
    return n; // tables that did not fit
  }
  function genTheatre(p, w, h, pax) {
    addFront(p, w); mk(p, 'podium', w / 2 + 4, 1.4);
    const per = Math.max(4, Math.min(30, Math.floor((w - 2.4) / 0.55)));
    let left = pax, r = 0;
    for (; left > 0; r++) { const y = 4.2 + r * 0.9; if (y > h - 1.2) break; const n = Math.min(per, left); mk(p, 'chairrow', w / 2, y, 0, n); left -= n; }
    mk(p, 'door', 1.5, h);
    return Math.ceil(left / per);
  }
  function genClass(p, w, h, pax, vip = true) {
    const rot = w > 2 * h, lw = rot ? h : w, lh = rot ? w : h, wide = lw >= 10;
    const P = (t, x, y, s) => mk(p, t, rot ? y : x, rot ? x : y, rot ? 270 : 0, s);
    const tprMax = Math.max(1, Math.floor((lw - 2.4) / 1.8)), rows = Math.max(1, Math.ceil(pax / (tprMax * 3)));
    const pr = Math.ceil(pax / rows), tpr = Math.ceil(pr / 3), ais = tpr >= 6 ? 1.2 : 0, x0 = (lw - tpr * 1.8 - ais) / 2, y0 = wide ? 4.8 : 5.2;
    P('stage', lw / 2, 1.5); P('podium', wide ? lw / 2 + 4.5 : lw - 0.8, wide ? 1.4 : 3.5); P('lect', wide ? lw / 2 - 4.8 : lw / 2, wide ? 1.4 : 3.5);
    if (vip) P('sofa', lw >= 20 ? 1.6 : 1.3, lw >= 20 ? 1.4 : 3.5);
    let left = pax, r = 0;
    for (; r < rows && left > 0; r++) {
      const y = y0 + r * 1.3; if (y > lh - 0.8) break;
      for (let t = 0; t < tpr && left > 0; t++) { const s = Math.min(3, left); P('class', x0 + t * 1.8 + 0.9 + (ais && t >= tpr / 2 ? ais : 0), y, s); left -= s; }
    }
    if (!rot) mk(p, 'door', 1.5, h);
    return Math.ceil(left / 3);
  }
  function genUshape(p, w, h, pax) {
    // two arms run away from the screen, joined by a head table; turn it sideways in a shallow room
    const segs = n => { const out = []; while (n > 0) { const s2 = Math.min(8, n); out.push(s2); n -= s2; } return out; };
    const head = Math.min(12, Math.max(0, pax - 2 * Math.ceil(pax * 0.4))), side = Math.ceil((pax - head) / 2);
    const armLen = segs(side).reduce((a, s2) => a + rlen({ seats: s2 }) + 0.5, 0);
    const rot = armLen + 4 > h && w > h, lw = rot ? h : w, lh = rot ? w : h;
    const P = (t, x, y, r, s2) => mk(p, t, rot ? y : x, rot ? x : y, ((r || 0) + (rot ? 270 : 0)) % 360, s2);
    P('screen', lw / 2, 0.3);
    const gap = Math.min(3.2, Math.max(1.6, lw / 2 - 1.6));
    for (const sx of [-gap, gap]) { let y = 2.4; for (const s2 of segs(side)) { const L = rlen({ seats: s2 }); P('rect', lw / 2 + sx, y + L / 2, 90, s2); y += L + 0.5; } }
    if (head) P('rect', lw / 2, Math.min(lh - 1.3, 2.4 + armLen + 0.8), 0, head);
    if (!rot) mk(p, 'door', 1.5, h);
    return 0;
  }
  function autoPlan(q, day, layout) {
    const rf = roomFor(q, day), [w, h] = ROOM_DIM[rf.id] || [20, 14], pax = Number(q.pax) || 0;
    const p = { title: (q.title || '') + ' · ' + ((rf.room && (rf.room.th || rf.room.name)) || 'ห้องจัดงาน') + ' · ' + pax + ' ท่าน', W: w, H: h, items: [] };
    const lay = layout || rf.layout, gen = { theatre: genTheatre, classroom: genClass, ushape: genUshape, banquet: genBanquet, cocktail: (p2, a, b2, n) => genBanquet(p2, a, b2, n, 8) }[lay] || genBanquet;
    p.missing = gen(p, w, h, pax) || 0;
    return p;
  }

  /* ---------- past plans (from the hotel's own floor-plan tool) ---------- */
  const HIST_ROOM = { 'กินรี': 'kinnaree', 'ทิฆัมพร': 'thikhampon', 'เรือนหงส์': 'ruenhong' };
  const HIST = [
    { r: 'เรือนหงส์', d: '8-10', l: 'ประชุมกลุ่ม', p: 0, note: 'โต๊ะประชุม 12, โต๊ะเบรกใหญ่ 2 (ผ้าโลโก้โรงแรม), โต๊ะเจ้าหน้าที่, โต๊ะลงทะเบียน, โซฟา VIP, ฟลิปชาร์ต 4' },
    { r: 'กินรี', d: '12', l: 'Classroom+Theatre', p: 250, note: 'โต๊ะประชุม 54 ตัว + แถวเก้าอี้ Theatre' },
    { r: 'กินรี', d: '-', l: 'งานบูท', p: 0, note: 'โต๊ะลงทะเบียน 2, โต๊ะออกบูท 5, ใช้คู่กับห้องเม็งราย' },
    { r: 'ไม่ระบุ', d: '-', l: 'เลี้ยง/ปาร์ตี้', p: 0, note: '7 โต๊ะ + VIP, บาร์น้ำ+น้ำแข็ง, เว้นที่เต้น, ชุดน้ำต่อโต๊ะ: ส้ม1 โค้ก1 โซดา2 เปล่า4, เก็บจาน 22:00, เลิก 23:00 (ต่อเวลาแจ้งเซลล์)' },
    { r: 'กินรี', d: '20-21', l: 'Classroom', p: 50, g: 'class', note: 'โพเดียม, โต๊ะบรรยาย, VIP 2, ป้ายข้อความ' },
    { r: 'ทิฆัมพร', d: '20-21', l: 'แบ่งกลุ่ม', p: 30, note: '14 โต๊ะ, ฟลิปชาร์ต 5, กระดาษ+ปากกาสองหัว' },
    { r: 'ทิฆัมพร', d: '23-25', l: 'Classroom', p: 50, g: 'class', note: 'ฟลิปชาร์ต 4, ไวนิลหลังเวที' },
    { r: 'กินรี', d: '24', l: 'Classroom', p: 291, g: 'class', note: 'ใช้โต๊ะ 88 ตัว 10 แถว' },
    { r: 'กินรี', d: '26', l: 'Classroom', p: 70, g: 'class', note: 'โพเดียม 2, โต๊ะบรรยายย้ายขึ้นหลังพิธีเปิด' },
    { r: 'กินรี', d: '11', l: 'Classroom', p: 140, g: 'class', note: '10 แถว แถวละ 14 ที่, เซตน้ำดื่ม+แก้ว 5 ที่นั่ง' },
    { r: 'กินรี', d: '16-17', l: 'Classroom', p: 108, g: 'class', note: '9 แถว แถวละ 12 ที่ (เผื่อ 8)' },
    { r: 'ทิฆัมพร', d: '18-19', l: 'Classroom', p: 60, g: 'class', note: '6 แถว แถวละ 10 ที่, โต๊ะบรรยายย้ายหลังพิธีเปิด' },
    { r: 'เรือนหงส์', d: '19', l: 'Classroom', p: 30, g: 'class', note: '5 แถว แถวละ 6 ที่, โต๊ะเบรก 2, โต๊ะลงทะเบียน, โต๊ะสังเกตการณ์เจ้าหน้าที่' },
    { r: 'เรือนหงส์', d: '20-22', l: 'Classroom', p: 34, g: 'class', note: '5 แถว แถวละ 6 ที่ + 4 ที่, โต๊ะเบรก 2' },
    { r: 'ทิฆัมพร', d: '23', l: 'Classroom', p: 60, g: 'class', note: 'เหมือนงานวันที่ 18-19' },
    { r: 'กินรี', d: '24-25', l: 'Classroom', p: 205, g: 'class', note: 'ช่วง 200-210 ที่, 69 โต๊ะ, โต๊ะสังเกตการณ์เจ้าหน้าที่' },
    { r: 'กินรี', d: '24', l: 'Classroom+ไลน์อาหาร', p: 120, g: 'class', note: 'ไลน์อาหาร 3, โซนคุมเสียง, ย้าย 3 แถวหลังไปหน้าเรือนหงส์' },
    { r: 'กินรี', d: '25', l: 'โต๊ะกลม', p: 380, g: 'round', note: '38 โต๊ะ (19+19), VIP 3, โซนคุมเสียง, ชุดน้ำต่อโต๊ะ ส้ม1 โค้ก1 โซดา2 เปล่า4' },
    { r: 'เรือนหงส์', d: '26', l: 'Classroom', p: 30, g: 'class', note: '5 แถว แถวละ 6 ที่' }
  ];
  function histPlan(h) {
    const [w, hh] = ROOM_DIM[HIST_ROOM[h.r]] || [20, 14];
    const p = { title: h.r + ' · ' + h.l + ' ' + h.p + ' ท่าน', W: w, H: hh, items: [] };
    p.missing = (h.g === 'class' ? genClass(p, w, hh, h.p) : genBanquet(p, w, hh, h.p)) || 0;
    return p;
  }
  function template(t, w, h) {
    const p = { W: w, H: h, items: [] };
    if (t === 'banquet') genBanquet(p, w, h, Math.floor((w / 3.2) * ((h - 5.5) / 2.9)) * 10);
    if (t === 'theatre') genTheatre(p, w, h, Math.floor((w - 2.4) / 0.55) * Math.floor((h - 5.4) / 0.9));
    if (t === 'board') { mk(p, 'screen', w / 2, 0.3); mk(p, 'rect', w / 2, h / 2, 0, 14); mk(p, 'door', 1.5, h); }
    if (t === 'kin') {
      p.W = 24; p.H = 18; [7, 6, 6, 6].forEach((n, r) => { for (let c = 0; c < n; c++) mk(p, 'round', 12 + (c - (n - 1) / 2) * 3.3, 5.4 + r * 3.2); });
      mk(p, 'stage', 12, 1.5); mk(p, 'podium', 16.4, 1.5); mk(p, 'sound', 19, 1.4); mk(p, 'sofa', 21.6, 1.4); mk(p, 'buffet', 3, 1.2); mk(p, 'buffet', 6.3, 1.2); mk(p, 'reg', 5, 17.4); mk(p, 'door', 2, 18);
    }
    if (t === 'red') {
      p.W = 24; p.H = 18; mk(p, 'carpet', 12, 10.4);
      [[2.2, 5.4, 8.6], [3.8, 7], [2.2, 5.4, 8.6], [3.8, 7]].forEach((r, i) => r.forEach(x => { mk(p, 'round', x, 5 + i * 2.9); mk(p, 'round', 24 - x, 5 + i * 2.9); }));
      mk(p, 'stage', 12, 1.5); mk(p, 'podium', 16.4, 1.5); mk(p, 'sound', 19, 1.4); mk(p, 'sofa', 21.6, 1.4); mk(p, 'buffet', 3, 1.2); mk(p, 'buffet', 6.3, 1.2); mk(p, 'reg', 5, 17.4); mk(p, 'door', 12, 18);
    }
    return p;
  }

  /* ---------- drawing ---------- */
  const chair = (x, y, r = 0) => `<rect x="${x - 9}" y="${y - 9}" width="18" height="18" rx="4" fill="#d9c7a3" stroke="#8a7650" stroke-width="1" transform="rotate(${r} ${x} ${y})"/>`;
  function shape(i) {
    const T = (t, sz = 12, c = '#333') => `<text text-anchor="middle" dominant-baseline="middle" font-size="${sz}" fill="${c}">${t}</text>`;
    let s = '';
    if (i.type === 'round') {
      for (let k = 0; k < i.seats; k++) { const a = 2 * Math.PI * k / i.seats; s += chair(Math.cos(a) * 1.2 * S, Math.sin(a) * 1.2 * S, a * 180 / Math.PI + 90); }
      return s + `<circle r="${0.9 * S}" fill="#f3ece0" stroke="#8a7650" stroke-width="2"/>` + T(i.seats + ' ที่');
    }
    if (i.type === 'rect') {
      const L = rlen(i) * S, n = Math.ceil(i.seats / 2), m = i.seats - n;
      for (let k = 0; k < n; k++) s += chair((k - (n - 1) / 2) * (L / n), -0.72 * S);
      for (let k = 0; k < m; k++) s += chair((k - (m - 1) / 2) * (L / Math.max(m, 1)), 0.72 * S);
      return s + `<rect x="${-L / 2}" y="${-0.45 * S}" width="${L}" height="${0.9 * S}" rx="4" fill="#f3ece0" stroke="#8a7650" stroke-width="2"/>` + T(i.seats + ' ที่');
    }
    if (i.type === 'class') {
      const L = i.seats * 0.6 * S;
      s = `<rect x="${-L / 2}" y="${-0.55 * S}" width="${L}" height="${0.5 * S}" rx="2" fill="#f3ece0" stroke="#8a7650" stroke-width="1.5"/>`;
      for (let k = 0; k < i.seats; k++) s += chair((k - (i.seats - 1) / 2) * 0.6 * S, 0.32 * S);
      return s;
    }
    if (i.type === 'chairrow') { for (let k = 0; k < i.seats; k++) s += chair((k - (i.seats - 1) / 2) * 0.55 * S, 0); return s; }
    const [w, h] = dims(i);
    s = `<rect x="${-w * S / 2}" y="${-h * S / 2}" width="${w * S}" height="${h * S}" rx="3" fill="${FILL[i.type]}" stroke="#555" stroke-width="${i.type === 'door' ? 2 : 1.5}"/>`;
    if (i.type === 'door') s += `<path d="M${-S / 2} 0 A${S} ${S} 0 0 1 ${S / 2} ${S * 0.9}" fill="none" stroke="#999" stroke-dasharray="4 3"/>`;
    if (!['screen', 'door'].includes(i.type)) s += T(NAMES[i.type], w > 2 ? 13 : 9, i.type === 'carpet' ? '#fff' : '#333');
    return s;
  }
  // full drawing; sel = selected item id (left out of exports)
  function build(p, sel) {
    const w = p.W, h = p.H, vw = (w + 2) * S, vh = (h + 3) * S;
    let gr = '';
    for (let x = 0; x <= w; x++) gr += `<line x1="${x * S}" y1="0" x2="${x * S}" y2="${h * S}" stroke="#eee"/>`;
    for (let y = 0; y <= h; y++) gr += `<line x1="0" y1="${y * S}" x2="${w * S}" y2="${y * S}" stroke="#eee"/>`;
    gr += `<rect width="${w * S}" height="${h * S}" fill="none" stroke="#333" stroke-width="4"/>`;
    for (const i of p.items) {
      const [dw, dh] = dims(i);
      gr += `<g data-id="${i.id}" transform="translate(${i.x * S} ${i.y * S}) rotate(${i.rot})" style="cursor:move">${shape(i)}` +
        (sel === i.id ? `<rect x="${-dw * S / 2 - 4}" y="${-dh * S / 2 - 4}" width="${dw * S + 8}" height="${dh * S + 8}" fill="none" stroke="#A8262D" stroke-width="2" stroke-dasharray="6 4"/>` : '') + '</g>';
    }
    const seats = seatsOf(p);
    return { vw, vh, svg: `<rect width="${vw}" height="${vh}" fill="#fff"/><text x="${S}" y="${S * 0.95}" font-size="26" font-weight="700" fill="#222">${H(p.title || '')}</text>
      <g transform="translate(${S} ${2 * S})">${gr}</g><text x="${S}" y="${vh - S * 0.35}" font-size="14" fill="#666">ห้อง ${w}×${h} ม. · ที่นั่งรวม ${seats} ที่</text>` };
  }
  // read-only picture for the function sheet page and print
  App.planSvg = function (p, cls) { const b = build(p); return `<svg class="${cls || ''}" viewBox="0 0 ${b.vw} ${b.vh}" role="img" aria-label="${H((p.title || 'ผังงาน') + ' · ที่นั่งรวม ' + seatsOf(p) + ' ที่')}" font-family="IBM Plex Sans Thai, Tahoma, sans-serif">${b.svg}</svg>`; };
  App.planSeats = seatsOf;

  /* ---------- the editor page ---------- */
  let sel = null, drag = null;
  App.V.plan = {
    title: 'ออกแบบผังงาน', nav: 'beo',
    render(params) {
      const b = curBeo(), q = b && Store.quote(b.quoteId);
      if (!b || !q) return '<div class="empty">ไม่พบใบฟังก์ชัน — <a href="#/beo">กลับไปหน้าใบฟังก์ชัน</a></div>';
      const days = U.daysOf(q.start, q.end), day = Math.min(dayOf(), days.length);
      const p = planOf(b, day), edit = canEdit();
      const head = `<div class="page-hd"><div><div class="small"><a class="link" href="#/beo/${encodeURIComponent(b.id)}">← ใบฟังก์ชัน ${H(b.id)}</a></div><h1>ออกแบบผังงาน</h1><div class="small muted">${H(q.title)} · ${H(q.client)} · ${U.fmtRange(q.start, q.end)} · ${U.n(q.pax)} ท่าน</div></div>
        <div class="row wrap">${days.length > 1 ? `<div class="seg white">${days.map((d, i) => `<a class="${day === i + 1 ? 'on' : ''}" href="#/plan/${encodeURIComponent(b.id)}/${i + 1}">วันที่ ${i + 1}</a>`).join('')}</div>` : ''}
        ${p ? '<button class="btn ghost" data-act="fpPng">ดาวน์โหลด PNG</button>' : ''}</div></div>`;
      if (!p) {
        const rf = roomFor(q, day), lay = (D.LAYOUTS.find(l => l.id === rf.layout) || {}).name || '';
        return head + `<div class="card pad" style="text-align:center;max-width:620px;margin:20px auto">
          <h2>ยังไม่มีผังงานของวันที่ ${day}</h2><p class="muted">สร้างผังตั้งต้นจากใบเสนอราคา: ${H((rf.room && (rf.room.th || rf.room.name)) || 'ห้องจัดงาน')} · ${H(lay)} · ${U.n(q.pax)} ท่าน แล้วลากปรับได้</p>
          ${edit ? `<div class="row wrap" style="justify-content:center"><button class="btn red" data-act="fpAuto">สร้างผังจากใบเสนอราคา</button><button class="btn ghost" data-act="fpBlank">เริ่มจากห้องว่าง</button></div>` : '<p class="small muted">คุณดูได้อย่างเดียว</p>'}</div>`;
      }
      if (!p.items.some(i => i.id === sel)) sel = null;
      const pax = Number(q.pax) || 0, seats = seatsOf(p);
      return head + `
<div class="fp">
  <aside class="fp-side">
    <section class="card pad"><h3>ชื่อผัง / ขนาดห้อง (ม.)</h3>
      <input class="input" value="${H(p.title || '')}" data-change="fpTitle" aria-label="ชื่อผัง" ${edit ? '' : 'disabled'} style="margin-bottom:8px">
      <div class="row" style="gap:6px"><input class="input" type="number" min="4" max="80" step="0.1" value="${p.W}" data-change="fpW" aria-label="กว้าง (เมตร)" ${edit ? '' : 'disabled'}><span>×</span><input class="input" type="number" min="4" max="80" step="0.1" value="${p.H}" data-change="fpH" aria-label="ลึก (เมตร)" ${edit ? '' : 'disabled'}></div></section>
    ${edit ? `<section class="card pad"><h3>เทมเพลต</h3><div class="fp-grid">
      <button class="btn ghost sm" data-act="fpAuto">ตามใบเสนอราคา</button><button class="btn ghost sm" data-act="fpTpl" data-v="banquet">โต๊ะกลม</button>
      <button class="btn ghost sm" data-act="fpTpl" data-v="theatre">Theatre</button><button class="btn ghost sm" data-act="fpTpl" data-v="board">ห้องประชุม</button>
      <button class="btn ghost sm fp-wide" data-act="fpTpl" data-v="kin">ห้องกินรี 250 ท่าน</button><button class="btn ghost sm fp-wide" data-act="fpTpl" data-v="red">กินรี 200 ท่าน พรมแดงกลาง</button>
      <button class="btn ghost sm fp-wide" data-act="fpTpl" data-v="blank" style="color:var(--red)">ล้างทั้งหมด</button></div></section>
    <section class="card pad"><h3>เพิ่มของ</h3><div class="fp-grid">${Object.keys(NAMES).map(t => `<button class="btn ghost sm" data-act="fpAdd" data-v="${t}">+ ${NAMES[t]}</button>`).join('')}</div></section>
    <section class="card pad" id="fpSel">${selPanel(p)}</section>
    <section class="card pad"><h3>แปลนที่เคยจัด (${HIST.length} งาน)</h3>
      <select class="input" id="fpHist" aria-label="แปลนที่เคยจัด">${HIST.map((h, i) => `<option value="${i}">${i + 1}. ${H(h.r)} · ${H(h.l)}${h.p ? ' ' + h.p : ''} (วันที่ ${H(h.d)})</option>`).join('')}</select>
      <div class="xs muted" id="fpHistInfo" style="margin:6px 0">${histInfo(0)}</div><button class="btn ghost sm" style="width:100%" data-act="fpHist">โหลดเป็นผังตั้งต้น</button></section>` : '<div class="note">คุณดูผังได้อย่างเดียว</div>'}
  </aside>
  <div class="fp-main">
    <div class="fp-paper"><svg id="fpSvg" viewBox="0 0 10 10" font-family="IBM Plex Sans Thai, Tahoma, sans-serif"></svg></div>
    <div class="small muted" style="margin-top:8px">รวมที่นั่ง <b id="fpTot">${seats}</b> ที่ · ใบฟังก์ชัน ${U.n(pax)} ท่าน${edit ? ' · คลิกเพื่อเลือก ลากเพื่อย้าย · Delete ลบ · R หมุน' : ''}</div>
    <div id="fpWarn">${warn(p, pax)}</div>
  </div>
</div>`;
    },
    after() { drawSvg(); bindSvg(); const hs = document.getElementById('fpHist'); if (hs) hs.onchange = () => { document.getElementById('fpHistInfo').textContent = histInfo(Number(hs.value)); }; }
  };
  function histInfo(i) { const h = HIST[i], d = ROOM_DIM[HIST_ROOM[h.r]]; return (h.p && d ? 'พื้นที่ต่อคน ≈ ' + (d[0] * d[1] / h.p).toFixed(2) + ' ตร.ม. · ' : '') + h.note + (h.g ? '' : ' · (ยังไม่มีพิกัดพอจะโหลดเป็นผัง)'); }
  function warn(p, pax) {
    const seats = seatsOf(p), out = [];
    if (pax && seats && seats !== pax) out.push(seats < pax ? `ที่นั่งน้อยกว่าจำนวนท่าน ${pax - seats} ที่` : `ที่นั่งมากกว่าจำนวนท่าน ${seats - pax} ที่`);
    if (p.items.some(i => i.x < 0 || i.y < 0 || i.x > p.W || i.y > p.H)) out.push('มีของวางเกินขอบห้อง');
    if (p.missing) out.push(`วางไม่พอ ${p.missing} โต๊ะในห้องขนาดนี้ — ลองลดขนาดโต๊ะ ขยายห้อง หรือจัดชิดขึ้น`);
    return out.length ? `<div class="note" style="margin-top:8px">${out.map(H).join(' · ')}</div>` : '';
  }
  function selPanel(p) {
    const it = p.items.find(i => i.id === sel);
    if (!it) return '<h3>ที่เลือกอยู่</h3><div class="small muted">แตะที่โต๊ะหรือของในผังเพื่อเลือก</div>';
    return `<h3>ที่เลือกอยู่ · ${H(NAMES[it.type])}</h3>
      ${SEATED.includes(it.type) ? `<div class="row" style="margin-bottom:8px"><span style="flex:1">ที่นั่ง: <b>${it.seats}</b></span><button class="icon-btn" data-act="fpSeat" data-v="-1" aria-label="ลดที่นั่ง">−</button><button class="icon-btn" data-act="fpSeat" data-v="1" aria-label="เพิ่มที่นั่ง">+</button></div>` : ''}
      <div class="fp-grid"><button class="btn ghost sm" data-act="fpRot">⟳ หมุน 15°</button><button class="btn ghost sm" data-act="fpDup">คัดลอก</button><button class="btn ghost sm fp-wide" data-act="fpDel" style="color:var(--red)">ลบ</button></div>`;
  }
  const P = () => planOf(curBeo(), dayOf());
  function drawSvg() {
    const p = P(), s = document.getElementById('fpSvg'); if (!p || !s) return;
    const b = build(p, sel); s.setAttribute('viewBox', `0 0 ${b.vw} ${b.vh}`); s.innerHTML = b.svg;
    const t = document.getElementById('fpTot'); if (t) t.textContent = seatsOf(p);
  }
  function refreshSide() {
    const p = P(), q = Store.quote(curBeo().quoteId), el = document.getElementById('fpSel'), w = document.getElementById('fpWarn');
    if (el) el.innerHTML = selPanel(p); if (w) w.innerHTML = warn(p, Number(q.pax) || 0);
  }
  function save() { const p = P(); if (p) delete p.missing; Store.save(); }
  function bindSvg() {
    const s = document.getElementById('fpSvg'); if (!s || !canEdit()) return;
    const pos = e => { const pt = s.createSVGPoint(); pt.x = e.clientX; pt.y = e.clientY; const q = pt.matrixTransform(s.getScreenCTM().inverse()); return [(q.x - S) / S, (q.y - 2 * S) / S]; };
    s.addEventListener('pointerdown', e => {
      const gEl = e.target.closest('[data-id]'), p = P();
      if (!gEl) { sel = null; drawSvg(); refreshSide(); return; }
      sel = Number(gEl.dataset.id); const it = p.items.find(v => v.id === sel), [x, y] = pos(e);
      drag = { dx: it.x - x, dy: it.y - y, moved: false }; s.setPointerCapture(e.pointerId); drawSvg(); refreshSide();
    });
    s.addEventListener('pointermove', e => {
      if (!drag) return; const it = P().items.find(v => v.id === sel); if (!it) return;
      const [x, y] = pos(e); it.x = Math.round((x + drag.dx) * 10) / 10; it.y = Math.round((y + drag.dy) * 10) / 10; drag.moved = true; drawSvg();
    });
    const end = () => { if (drag && drag.moved) { save(); refreshSide(); } drag = null; };
    s.addEventListener('pointerup', end); s.addEventListener('pointercancel', end);
  }
  // Delete / R while the editor is open (not while typing or in a dialog)
  document.addEventListener('keydown', e => {
    if (App.state.route !== 'plan' || !canEdit() || sel == null || document.getElementById('modalBg')) return;
    if (/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return;
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); App.ACT.fpDel(); }
    if (e.key === 'r' || e.key === 'R') App.ACT.fpRot();
  });

  function setPlan(p) {
    const b = curBeo(); b.plans = b.plans || {}; b.plans[dayOf()] = p; sel = null; Store.save(); App.render();
    if (p.missing) App.toast(`วางไม่พอ ${p.missing} โต๊ะ — ปรับผังเพิ่มเติมได้`);
  }
  const step = fn => { const p = P(), it = p && p.items.find(i => i.id === sel); if (!it) return; fn(it, p); save(); drawSvg(); refreshSide(); };
  Object.assign(App.ACT, {
    fpAuto() { const b = curBeo(), q = Store.quote(b.quoteId); const go = () => setPlan(autoPlan(q, dayOf())); if (P() && P().items.length) App.confirm('แทนที่ผังปัจจุบันด้วยผังจากใบเสนอราคา?', go, 'แทนที่'); else go(); },
    fpBlank() { const b = curBeo(), q = Store.quote(b.quoteId), rf = roomFor(q, dayOf()), [w, h] = ROOM_DIM[rf.id] || [20, 14]; setPlan({ title: q.title || '', W: w, H: h, items: [] }); },
    fpTpl(el) {
      const p = P(), t = el.dataset.v;
      const go = () => { const n = t === 'blank' ? Object.assign({}, p, { items: [] }) : Object.assign(template(t, p.W, p.H), { title: p.title }); if (t === 'kin' || t === 'red') n.title = t === 'kin' ? 'งานเลี้ยง 250 ท่าน · ห้องกินรี' : 'งานเลี้ยง 200 ท่าน · ห้องกินรี · พรมแดงกลางห้อง'; setPlan(n); };
      if (p.items.length) App.confirm(t === 'blank' ? 'ล้างของทั้งหมดในผังนี้?' : 'แทนที่ผังปัจจุบันด้วยเทมเพลตนี้?', go, t === 'blank' ? 'ล้าง' : 'แทนที่'); else go();
    },
    fpHist() {
      const h = HIST[Number(document.getElementById('fpHist').value)];
      if (!h.g) { App.toast('แปลนนี้ยังไม่มีพิกัดพอจะโหลดเป็นผัง'); return; }
      const go = () => setPlan(histPlan(h));
      if (P().items.length) App.confirm('แทนที่ผังปัจจุบันด้วยแปลน “' + h.r + ' · ' + h.l + '”?', go, 'แทนที่'); else go();
    },
    fpAdd(el) { const p = P(), n = p.items.length % 5; sel = mk(p, el.dataset.v, p.W / 2 + n * 0.4, p.H / 2 + n * 0.4).id; save(); drawSvg(); refreshSide(); },
    fpRot() { step(it => { it.rot = (it.rot + 15) % 360; }); },
    fpDup() { step((it, p) => { const n = mk(p, it.type, it.x + 1, it.y + 1, it.rot, it.seats); sel = n.id; }); },
    fpDel() { const p = P(); if (!p) return; p.items = p.items.filter(i => i.id !== sel); sel = null; save(); drawSvg(); refreshSide(); },
    fpSeat(el) { step(it => { it.seats = Math.max(1, Math.min(it.type === 'chairrow' ? 40 : 16, it.seats + Number(el.dataset.v))); }); },
    fpPng() {
      const p = P(); if (!p) return;
      const b = build(p), sc = 2, x = `<svg xmlns="http://www.w3.org/2000/svg" width="${b.vw * sc}" height="${b.vh * sc}" viewBox="0 0 ${b.vw} ${b.vh}" font-family="IBM Plex Sans Thai, Tahoma, sans-serif">${b.svg}</svg>`;
      const img = new Image();
      img.onload = () => { const c = document.createElement('canvas'); c.width = b.vw * sc; c.height = b.vh * sc; c.getContext('2d').drawImage(img, 0, 0); const a = document.createElement('a'); a.download = 'floorplan-' + curBeo().id + '-day' + dayOf() + '.png'; a.href = c.toDataURL('image/png'); a.click(); };
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(x);
    }
  });
  Object.assign(App.CHG, {
    fpTitle(el) { P().title = el.value.slice(0, 120); save(); drawSvg(); },
    fpW(el) { P().W = Math.max(4, Math.min(80, Number(el.value) || 20)); save(); drawSvg(); refreshSide(); },
    fpH(el) { P().H = Math.max(4, Math.min(80, Number(el.value) || 14)); save(); drawSvg(); refreshSide(); }
  });
})(window);
