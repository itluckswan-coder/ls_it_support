/* Luck Swan Sales — ออกแบบผังงาน: drag-and-drop floor plans for a function sheet (one plan per event day).
 * Plans live on the function sheet (b.plans[day]) so they sync, print and go through the same approval as the sheet. */
(function (g) {
  'use strict';
  const { U, App, Store } = g, D = g.DATA, H = U.esc;
  const S = 40; // px per metre

  // room sizes in metres (width × depth)
  const ROOM_DIM = { kinnaree: [24, 18], thikhampon: [13.9, 12.1], ruenhong: [20, 7], prayamengrai: [10, 7], pool: [30, 20] };
  const NAMES = { round: 'โต๊ะกลม', rect: 'โต๊ะยาว', chairrow: 'แถวเก้าอี้', class: 'โต๊ะห้องเรียน', stage: 'เวที', screen: 'จอ', podium: 'โพเดียม', lect: 'โต๊ะบรรยาย', sound: 'คุมเสียง', buffet: 'บุฟเฟ่ต์', reg: 'ลงทะเบียน', sofa: 'โซฟา', carpet: 'พรมแดง', door: 'ประตู', pillar: 'เสา', block: 'บล็อก/ป้าย', custom: 'สิ่งของอื่น' };
  // default sizes in metres (w × d); a round table is its diameter. Tables sized by seat count keep their own length.
  const DEF_SIZE = { round: [2.9, 2.9], stage: [7, 2.4], buffet: [3, 0.8], screen: [3, 0.25], sound: [1.5, 0.8], podium: [0.6, 0.6], sofa: [2, 0.8], reg: [2.4, 0.7], lect: [2.4, 0.9], carpet: [2.4, 14.4], door: [1, 0.3], pillar: [0.6, 0.6], block: [2, 1] };
  const SIZED = Object.keys(DEF_SIZE);
  const floor = () => (Store.db.settings && Store.db.settings.floor) || {};
  // room size and locked fixtures from แก้ไขปรับปรุง, else the built-in size
  const roomDim = id => { const r = (floor().rooms || {})[id]; return r && r.W && r.H ? [r.W, r.H] : ROOM_DIM[id] || [20, 14]; };
  const fixedOf = id => (id && ((floor().rooms || {})[id] || {}).fixed) || [];
  const sizeOf = t => (floor().sizes || {})[t] || DEF_SIZE[t] || [1, 0.3];
  const canLib = () => App.can('edit_prices') || App.can('staff'); // the shared settings the libraries live in
  const libTpl = () => floor().templates || [];
  const libItems = () => floor().custom || [];
  function setFloor(fn) { const st = Store.db.settings, f = st.floor = Object.assign({ rooms: {}, sizes: {} }, st.floor); fn(f); Store.save(); }
  const roomName = id => { const r = D.MEETING_ROOMS.find(x => x.id === id); return r ? r.th || r.name : id; };
  const SEATED = ['round', 'rect', 'chairrow', 'class'];
  const FILL = { stage: '#c9d6e8', buffet: '#e8d3b9', screen: '#555', door: '#fff', sound: '#ddd', podium: '#e6d5ef', sofa: '#f0c9c0', reg: '#d6e8d0', lect: '#f1e2b8', carpet: '#c0392b', pillar: '#9a9a9a', block: '#e4e2da' };
  const rlen = i => Math.max(1.5, Math.ceil(i.seats / 2) * 0.75);
  function dims(i) {
    if (i.type === 'rect') return [rlen(i) + 0.2, 2.1];
    if (i.type === 'chairrow') return [i.seats * 0.55, 0.5];
    if (i.type === 'class') return [i.seats * 0.6, 1.1];
    // a size set on this one object wins over the standard size
    if (i.type === 'custom') return [i.w || 1, i.h || 1];
    const d = i.w && i.h ? [i.w, i.h] : sizeOf(i.type);
    return i.type === 'round' ? [d[0], d[0]] : d;
  }
  const seatsDefault = t => ({ round: 10, rect: 8, chairrow: 10, class: 3 }[t] || 0);

  /* ---------- plan data ---------- */
  const curBeo = () => Store.db.beos.find(b => b.id === App.state.params[0]);
  const dayOf = () => Math.max(1, Number(App.state.params[1]) || 1);
  const canEdit = () => App.canPlan();
  const isAdmin = () => { const me = App.me(); return !!me && me.role === 'admin'; };
  function planOf(b, day) { return b && b.plans && b.plans[day]; }
  // bounding box of an item in metres (rotation rounded to the nearest quarter turn, other angles use the longer side)
  function box(i) {
    const [w, h] = dims(i), q = ((Math.round(i.rot / 90) % 2) + 2) % 2, odd = i.rot % 90 !== 0;
    const bw = odd ? Math.max(w, h) : q ? h : w, bh = odd ? Math.max(w, h) : q ? w : h;
    return { x0: i.x - bw / 2, x1: i.x + bw / 2, y0: i.y - bh / 2, y1: i.y + bh / 2 };
  }
  const overlaps = (a, b) => a.x0 < b.x1 - 0.05 && b.x0 < a.x1 - 0.05 && a.y0 < b.y1 - 0.05 && b.y0 < a.y1 - 0.05;
  function mk(p, type, x, y, rot = 0, seats) {
    const id = (p.items.reduce((m, i) => Math.max(m, i.id), 0) || 0) + 1;
    const it = { id, type, x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10, rot, seats: seats == null ? seatsDefault(type) : seats };
    if (p._auto) {
      // generated plans work around the room's locked fixtures and do not repeat them (a fixed stage means no second stage)
      const fixed = fixedOf(p.room);
      if (!SEATED.includes(type) && fixed.some(f => f.type === type)) return null;
      const b = box(it); if (fixed.some(f => overlaps(box(f), b))) return null;
    }
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
    const d = sizeOf('round')[0], sx = d + 0.3, sy = d, nc = Math.max(1, Math.floor((w - 0.4) / sx)); let n = Math.ceil(pax / per);
    for (let r = 0; n > 0; r++) {
      const y = 3.15 + d / 2 + r * sy; if (y + d / 2 > h - 1) break;
      const c = r % 2 ? Math.max(1, nc - 1) : nc;
      for (let k = 0; k < c && n > 0; k++) if (mk(p, 'round', w / 2 + (k - (c - 1) / 2) * sx, y, 0, per)) n--;
    }
    mk(p, 'buffet', 2, h - 0.8); mk(p, 'door', w - 1.5, h);
    return n; // tables that did not fit
  }
  function genTheatre(p, w, h, pax) {
    addFront(p, w); mk(p, 'podium', w / 2 + 4, 1.4);
    const per = Math.max(4, Math.min(30, Math.floor((w - 2.4) / 0.55)));
    let left = pax, r = 0;
    for (; left > 0; r++) { const y = 4.2 + r * 0.9; if (y > h - 1.2) break; const n = Math.min(per, left); if (mk(p, 'chairrow', w / 2, y, 0, n)) left -= n; }
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
    for (; left > 0; r++) {
      const y = y0 + r * 1.3; if (y > lh - 0.8 || (!p._auto && r >= rows)) break;
      for (let t = 0; t < tpr && left > 0; t++) { const s = Math.min(3, left); if (P('class', x0 + t * 1.8 + 0.9 + (ais && t >= tpr / 2 ? ais : 0), y, s)) left -= s; }
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
    const rf = roomFor(q, day), [w, h] = roomDim(rf.id), pax = Number(q.pax) || 0;
    const p = { title: (q.title || '') + ' · ' + ((rf.room && (rf.room.th || rf.room.name)) || 'ห้องจัดงาน') + ' · ' + pax + ' ท่าน', room: rf.room ? rf.id : undefined, W: w, H: h, items: [] };
    const lay = layout || rf.layout, gen = { theatre: genTheatre, classroom: genClass, ushape: genUshape, banquet: genBanquet, cocktail: (p2, a, b2, n) => genBanquet(p2, a, b2, n, 8) }[lay] || genBanquet;
    p._auto = true; p.missing = gen(p, w, h, pax) || 0; delete p._auto;
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
    const [w, hh] = roomDim(HIST_ROOM[h.r]);
    const p = { title: h.r + ' · ' + h.l + ' ' + h.p + ' ท่าน', room: HIST_ROOM[h.r], W: w, H: hh, items: [] };
    p._auto = true; p.missing = (h.g === 'class' ? genClass(p, w, hh, h.p) : genBanquet(p, w, hh, h.p)) || 0; delete p._auto;
    return p;
  }
  function template(t, w, h) {
    const p = { W: w, H: h, items: [] };
    if (t === 'banquet') genBanquet(p, w, h, Math.floor((w / 3.2) * ((h - 5.5) / 2.9)) * 10);
    if (t === 'theatre') genTheatre(p, w, h, Math.floor((w - 2.4) / 0.55) * Math.floor((h - 5.4) / 0.9));
    if (t === 'board') { mk(p, 'screen', w / 2, 0.3); mk(p, 'rect', w / 2, h / 2, 0, 14); mk(p, 'door', 1.5, h); }
    if (t === 'kin') {
      p.W = 24; p.H = 18; p.room = 'kinnaree'; [7, 6, 6, 6].forEach((n, r) => { for (let c = 0; c < n; c++) mk(p, 'round', 12 + (c - (n - 1) / 2) * 3.3, 5.4 + r * 3.2); });
      mk(p, 'stage', 12, 1.5); mk(p, 'podium', 16.4, 1.5); mk(p, 'sound', 19, 1.4); mk(p, 'sofa', 21.6, 1.4); mk(p, 'buffet', 3, 1.2); mk(p, 'buffet', 6.3, 1.2); mk(p, 'reg', 5, 17.4); mk(p, 'door', 2, 18);
    }
    if (t === 'red') {
      p.W = 24; p.H = 18; p.room = 'kinnaree'; mk(p, 'carpet', 12, 10.4);
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
      const d = dims(i)[0];
      for (let k = 0; k < i.seats; k++) { const a = 2 * Math.PI * k / i.seats; s += chair(Math.cos(a) * d * 0.414 * S, Math.sin(a) * d * 0.414 * S, a * 180 / Math.PI + 90); }
      return s + `<circle r="${d * 0.31 * S}" fill="#f3ece0" stroke="#8a7650" stroke-width="2"/>` + T(i.seats + ' ที่');
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
    if (i.type === 'custom') {
      const fill = /^#[0-9a-f]{6}$/i.test(i.color || '') ? i.color : '#e4e2da';
      return (i.shape === 'round' ? `<ellipse rx="${w * S / 2}" ry="${h * S / 2}" fill="${fill}" stroke="#555" stroke-width="1.5"/>` : `<rect x="${-w * S / 2}" y="${-h * S / 2}" width="${w * S}" height="${h * S}" rx="3" fill="${fill}" stroke="#555" stroke-width="1.5"/>`) + T(H(i.label || NAMES.custom), w > 2 ? 13 : 9);
    }
    s = `<rect x="${-w * S / 2}" y="${-h * S / 2}" width="${w * S}" height="${h * S}" rx="3" fill="${FILL[i.type]}" stroke="#555" stroke-width="${i.type === 'door' ? 2 : 1.5}"/>`;
    if (i.type === 'door') s += `<path d="M${-S / 2} 0 A${S} ${S} 0 0 1 ${S / 2} ${S * 0.9}" fill="none" stroke="#999" stroke-dasharray="4 3"/>`;
    if (!['screen', 'door', 'pillar'].includes(i.type)) s += T(H(i.type === 'block' && i.label ? i.label : NAMES[i.type]), w > 2 ? 13 : 9, i.type === 'carpet' ? '#fff' : '#333');
    return s;
  }
  // full drawing; sel = selected item ids (left out of exports)
  // a plan tied to a room takes the room's size and locked fixtures from แก้ไขปรับปรุง
  const geom = p => p.room ? { W: roomDim(p.room)[0], H: roomDim(p.room)[1], fixed: fixedOf(p.room) } : { W: p.W, H: p.H, fixed: [] };
  function build(p, sel, noFixed) {
    const G = noFixed ? { W: p.W, H: p.H, fixed: [] } : geom(p), w = G.W, h = G.H, vw = (w + 2) * S, vh = (h + 3) * S;
    let gr = '';
    for (let x = 0; x <= w; x++) gr += `<line x1="${x * S}" y1="0" x2="${x * S}" y2="${h * S}" stroke="#eee"/>`;
    for (let y = 0; y <= h; y++) gr += `<line x1="0" y1="${y * S}" x2="${w * S}" y2="${y * S}" stroke="#eee"/>`;
    gr += `<rect width="${w * S}" height="${h * S}" fill="none" stroke="#333" stroke-width="4"/>`;
    for (const i of G.fixed) {
      const [dw, dh] = dims(i);
      gr += `<g transform="translate(${i.x * S} ${i.y * S}) rotate(${i.rot})" opacity=".9">${shape(i)}<rect x="${-dw * S / 2 - 3}" y="${-dh * S / 2 - 3}" width="${dw * S + 6}" height="${dh * S + 6}" fill="none" stroke="#6B716E" stroke-width="1.5" stroke-dasharray="2 3"/></g>`;
    }
    for (const i of p.items) {
      const [dw, dh] = dims(i);
      gr += `<g data-id="${i.id}" transform="translate(${i.x * S} ${i.y * S}) rotate(${i.rot})" style="cursor:move">${shape(i)}` +
        (sel && sel.includes(i.id) ? `<rect x="${-dw * S / 2 - 4}" y="${-dh * S / 2 - 4}" width="${dw * S + 8}" height="${dh * S + 8}" fill="none" stroke="#A8262D" stroke-width="2" stroke-dasharray="6 4"/>` : '') + '</g>';
    }
    const seats = seatsOf(p);
    return { vw, vh, svg: `<rect width="${vw}" height="${vh}" fill="#fff"/><text x="${S}" y="${S * 0.95}" font-size="26" font-weight="700" fill="#222">${H(p.title || '')}</text>
      <g transform="translate(${S} ${2 * S})">${gr}</g><text x="${S}" y="${vh - S * 0.35}" font-size="14" fill="#666">ห้อง ${w}×${h} ม. · ที่นั่งรวม ${seats} ที่</text>` };
  }
  // read-only picture for the function sheet page and print
  App.planSvg = function (p, cls) { const b = build(p); return `<svg class="${cls || ''}" viewBox="0 0 ${b.vw} ${b.vh}" role="img" aria-label="${H((p.title || 'ผังงาน') + ' · ที่นั่งรวม ' + seatsOf(p) + ' ที่')}" font-family="IBM Plex Sans Thai, Tahoma, sans-serif">${b.svg}</svg>`; };
  App.planSeats = seatsOf;

  /* ---------- the editor page ---------- */
  let sel = [], drag = null, marq = null;
  // undo: snapshots of the plan (or the room setup) taken before every change
  const UNDO_MAX = 60;
  let hist = { key: null, stack: [], last: null };
  const histKey = () => setupMode() ? 'setup:' + setupRoom() : 'plan:' + App.state.params[0] + ':' + dayOf();
  const snapNow = () => { const d = P(); return JSON.stringify(d ? (setupMode() ? { W: d.W, H: d.H, items: d.items } : d) : null); };
  function histSync() { if (hist.key !== histKey()) hist = { key: histKey(), stack: [], last: snapNow() }; }
  function record() {
    histSync(); const now = snapNow();
    if (now !== hist.last) { hist.stack.push(hist.last); if (hist.stack.length > UNDO_MAX) hist.stack.shift(); hist.last = now; }
    const btn = document.getElementById('fpUndo'); if (btn) btn.disabled = !hist.stack.length;
  }
  const undoBtn = () => `<button class="btn ghost" id="fpUndo" data-act="fpUndo" ${hist.key === histKey() && hist.stack.length ? '' : 'disabled'} title="ย้อนกลับ (Ctrl+Z)"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/></svg>ย้อนกลับ</button>`;
  App.V.plan = {
    title: 'ออกแบบผังงาน', nav: 'beo',
    render(params) {
      const b = curBeo(), q = b && Store.quote(b.quoteId);
      if (!b || !q) return '<div class="empty">ไม่พบใบฟังก์ชัน — <a href="#/beo">กลับไปหน้าใบฟังก์ชัน</a></div>';
      if (!App.canPlan()) return `<div class="empty">ออกแบบผังงานใช้ได้เฉพาะผู้ดูแลระบบ ฝ่ายขาย แผนก F&B และแผนกจัดเลี้ยง — <a href="#/beo/${encodeURIComponent(b.id)}">กลับไปใบฟังก์ชัน</a></div>`;
      const days = U.daysOf(q.start, q.end), day = Math.min(dayOf(), days.length);
      const p = planOf(b, day), edit = canEdit();
      histSync();
      const head = `<div class="page-hd"><div><div class="small"><a class="link" href="#/beo/${encodeURIComponent(b.id)}">← ใบฟังก์ชัน ${H(b.id)}</a></div><h1>ออกแบบผังงาน</h1><div class="small muted">${H(q.title)} · ${H(q.client)} · ${U.fmtRange(q.start, q.end)} · ${U.n(q.pax)} ท่าน</div></div>
        <div class="row wrap">${days.length > 1 ? `<div class="seg white">${days.map((d, i) => `<a class="${day === i + 1 ? 'on' : ''}" href="#/plan/${encodeURIComponent(b.id)}/${i + 1}">วันที่ ${i + 1}</a>`).join('')}</div>` : ''}
        ${edit && (p || hist.stack.length) ? undoBtn() : ''}${isAdmin() ? `<a class="btn ghost" href="#/plansetup${p && p.room ? '/' + encodeURIComponent(p.room) : ''}">แก้ไขปรับปรุง</a>` : ''}${p ? '<button class="btn ghost" data-act="fpPng">ดาวน์โหลด PNG</button>' : ''}</div></div>`;
      if (!p) {
        const rf = roomFor(q, day), lay = (D.LAYOUTS.find(l => l.id === rf.layout) || {}).name || '';
        return head + `<div class="card pad" style="text-align:center;max-width:620px;margin:20px auto">
          <h2>ยังไม่มีผังงานของวันที่ ${day}</h2><p class="muted">สร้างผังตั้งต้นจากใบเสนอราคา: ${H((rf.room && (rf.room.th || rf.room.name)) || 'ห้องจัดงาน')} · ${H(lay)} · ${U.n(q.pax)} ท่าน แล้วลากปรับได้</p>
          ${edit ? `<div class="row wrap" style="justify-content:center"><button class="btn red" data-act="fpAuto">สร้างผังจากใบเสนอราคา</button><button class="btn ghost" data-act="fpBlank">เริ่มจากห้องว่าง</button></div>` : '<p class="small muted">คุณดูได้อย่างเดียว</p>'}</div>`;
      }
      sel = sel.filter(id => p.items.some(i => i.id === id));
      const pax = Number(q.pax) || 0, seats = seatsOf(p);
      return head + `
<div class="fp">
  <aside class="fp-side">
    <section class="card pad"><h3>ชื่อผัง / ขนาดห้อง (ม.)</h3>
      <input class="input" value="${H(p.title || '')}" data-change="fpTitle" aria-label="ชื่อผัง" ${edit ? '' : 'disabled'} style="margin-bottom:8px">
      ${p.room ? `<div class="small"><b>${H(roomName(p.room))}</b> ${geom(p).W} × ${geom(p).H} ม.${fixedOf(p.room).length ? ' · ของที่ห้ามย้าย ' + fixedOf(p.room).length + ' ชิ้น (เส้นประ)' : ''}</div>
        <div class="xs muted" style="margin-top:4px">ขนาดห้องและของที่ห้ามย้ายแก้ได้ที่ ${isAdmin() ? `<a href="#/plansetup/${encodeURIComponent(p.room)}">แก้ไขปรับปรุง</a>` : 'แก้ไขปรับปรุง (ผู้ดูแลระบบ)'}</div>`
      : `<div class="row" style="gap:6px"><input class="input" type="number" min="4" max="80" step="0.1" value="${p.W}" data-change="fpW" aria-label="กว้าง (เมตร)" ${edit ? '' : 'disabled'}><span>×</span><input class="input" type="number" min="4" max="80" step="0.1" value="${p.H}" data-change="fpH" aria-label="ลึก (เมตร)" ${edit ? '' : 'disabled'}></div>`}</section>
    ${edit ? `<section class="card pad"><div class="fp-h"><h3>เทมเพลต</h3>${canLib() ? '<button class="fp-plus" data-act="fpTplNew" aria-label="เพิ่มเทมเพลต" title="เพิ่มเทมเพลต (บันทึกผังนี้ / อัปโหลดไฟล์)">+</button>' : ''}</div><div class="fp-grid">
      <button class="btn ghost sm" data-act="fpAuto">ตามใบเสนอราคา</button><button class="btn ghost sm" data-act="fpTpl" data-v="banquet">โต๊ะกลม</button>
      <button class="btn ghost sm" data-act="fpTpl" data-v="theatre">Theatre</button><button class="btn ghost sm" data-act="fpTpl" data-v="board">ห้องประชุม</button>
      <button class="btn ghost sm fp-wide" data-act="fpTpl" data-v="kin">ห้องกินรี 250 ท่าน</button><button class="btn ghost sm fp-wide" data-act="fpTpl" data-v="red">กินรี 200 ท่าน พรมแดงกลาง</button>
      ${libTpl().map(t => `<span class="fp-lib fp-wide"><button class="btn ghost sm" data-act="fpTplC" data-id="${H(t.id)}">${H(t.name)}</button>${canLib() ? `<button class="fp-x" data-act="fpTplDel" data-id="${H(t.id)}" aria-label="ลบเทมเพลต ${H(t.name)}">×</button>` : ''}</span>`).join('')}
      <button class="btn ghost sm fp-wide" data-act="fpTpl" data-v="blank" style="color:var(--red)">ล้างทั้งหมด</button></div></section>
    <section class="card pad"><div class="fp-h"><h3>เพิ่มของ</h3><button class="fp-plus" data-act="fpItemNew" aria-label="เพิ่มสิ่งของใหม่" title="เพิ่มสิ่งของใหม่ กำหนดชื่อและขนาดเอง">+</button></div><div class="fp-grid">${Object.keys(NAMES).filter(t => t !== 'custom').map(t => `<button class="btn ghost sm" data-act="fpAdd" data-v="${t}">+ ${NAMES[t]}</button>`).join('')}${libButtons()}</div></section>
    <section class="card pad" id="fpSel">${selPanel(p)}</section>
    <section class="card pad"><h3>แปลนที่เคยจัด (${HIST.length} งาน)</h3>
      <select class="input" id="fpHist" aria-label="แปลนที่เคยจัด">${HIST.map((h, i) => `<option value="${i}">${i + 1}. ${H(h.r)} · ${H(h.l)}${h.p ? ' ' + h.p : ''} (วันที่ ${H(h.d)})</option>`).join('')}</select>
      <div class="xs muted" id="fpHistInfo" style="margin:6px 0">${histInfo(0)}</div><button class="btn ghost sm" style="width:100%" data-act="fpHist">โหลดเป็นผังตั้งต้น</button></section>` : '<div class="note">คุณดูผังได้อย่างเดียว</div>'}
  </aside>
  <div class="fp-main">
    <div class="fp-paper"><svg id="fpSvg" viewBox="0 0 10 10" font-family="IBM Plex Sans Thai, Tahoma, sans-serif"></svg></div>
    <div class="small muted" style="margin-top:8px">รวมที่นั่ง <b id="fpTot">${seats}</b> ที่ · ใบฟังก์ชัน ${U.n(pax)} ท่าน${edit ? ' · คลิกเพื่อเลือก · ลากคลุมบนพื้นที่ว่างเพื่อเลือกหลายชิ้น · ลากเพื่อย้าย · Delete ลบ · R หมุน' : ''}</div>
    <div id="fpWarn">${warn(p, pax)}</div>
  </div>
</div>`;
    },
    after() { drawSvg(); bindSvg(); const hs = document.getElementById('fpHist'); if (hs) hs.onchange = () => { document.getElementById('fpHistInfo').textContent = histInfo(Number(hs.value)); }; }
  };
  /* ---------- แก้ไขปรับปรุง: room sizes, locked fixtures, standard item sizes (admins) ---------- */
  const setupRoom = () => { const id = App.state.params[0]; return D.MEETING_ROOMS.some(r => r.id === id) ? id : D.MEETING_ROOMS[0].id; };
  function setupDoc() {
    const ui = App.state.ui, id = setupRoom();
    if (!ui.fpSetup || ui.fpSetup.room !== id) { const [w, h] = roomDim(id); ui.fpSetup = { room: id, title: 'ห้อง' + roomName(id).replace(/^ห้อง/, '') + ' · ของที่ห้ามย้าย', W: w, H: h, items: JSON.parse(JSON.stringify(fixedOf(id))) }; }
    return ui.fpSetup;
  }
  function saveSetup() {
    const d = App.state.ui.fpSetup, st = Store.db.settings, f = st.floor = Object.assign({ rooms: {}, sizes: {} }, st.floor);
    f.rooms = Object.assign({}, f.rooms, { [d.room]: { W: d.W, H: d.H, fixed: d.items.map(i => Object.assign({}, i)) } });
    Store.save();
  }
  App.V.plansetup = {
    title: 'แก้ไขปรับปรุง', nav: 'beo',
    render() {
      if (!isAdmin()) return '<div class="card pad" style="max-width:520px;margin:40px auto;text-align:center"><h2>เฉพาะผู้ดูแลระบบ</h2><p class="muted">หน้าแก้ไขปรับปรุงผังห้องเปิดได้เฉพาะผู้ดูแลระบบ</p><a class="btn ghost" href="#/beo">กลับ</a></div>';
      App.state.ui.fpSetup = null; // rebuild from the shared settings on every render
      const d = setupDoc(), id = d.room, sizes = floor().sizes || {};
      histSync();
      return `<div class="page-hd"><div><div class="small"><a class="link" href="#/beo">← ใบฟังก์ชัน</a></div><h1>แก้ไขปรับปรุง</h1><div class="small muted">ขนาดห้อง ของในห้องที่ห้ามย้าย และขนาดมาตรฐานของสิ่งของ — ใช้กับทุกผังงานของห้องนั้น</div></div>
        <div class="row wrap">${undoBtn()}<button class="btn ghost" data-act="fpPng">ดาวน์โหลด PNG</button></div></div>
<div class="row wrap" style="margin-bottom:14px">${D.MEETING_ROOMS.map(r => `<a class="chip ${r.id === id ? 'on' : ''}" href="#/plansetup/${r.id}">${H(r.th || r.name)}</a>`).join('')}</div>
<div class="fp">
  <aside class="fp-side">
    <section class="card pad"><h3>ขนาดห้อง${H(roomName(id).replace(/^ห้อง/, ''))} (ม.)</h3>
      <div class="row" style="gap:6px"><input class="input" type="number" min="4" max="80" step="0.1" value="${d.W}" data-change="fsW" aria-label="กว้าง (เมตร)"><span>×</span><input class="input" type="number" min="4" max="80" step="0.1" value="${d.H}" data-change="fsH" aria-label="ลึก (เมตร)"></div>
      <div class="xs muted" style="margin-top:6px">พื้นที่ ${(d.W * d.H).toFixed(1)} ตร.ม.${(D.MEETING_ROOMS.find(r => r.id === id) || {}).area ? ' · ข้อมูลห้อง ' + D.MEETING_ROOMS.find(r => r.id === id).area + ' ตร.ม.' : ''}</div></section>
    <section class="card pad"><div class="fp-h"><h3>เพิ่มของที่ห้ามย้าย</h3><button class="fp-plus" data-act="fpItemNew" aria-label="เพิ่มของที่ห้ามย้ายชิ้นใหม่" title="เพิ่มของที่ห้ามย้ายชิ้นใหม่ กำหนดชื่อ ขนาด และสีเอง">+</button></div><div class="xs muted" style="margin-bottom:8px">เช่น เวทีถาวร เสา ประตู จอ — จะแสดงเป็นเส้นประในทุกผังของห้องนี้ และย้ายในผังงานไม่ได้</div>
      <div class="fp-grid">${['stage', 'screen', 'door', 'pillar', 'block', 'sound', 'buffet', 'carpet', 'podium', 'reg'].map(t => `<button class="btn ghost sm" data-act="fpAdd" data-v="${t}">+ ${NAMES[t]}</button>`).join('')}${libButtons()}</div></section>
    <section class="card pad" id="fpSel">${selPanel(d)}</section>
  </aside>
  <div class="fp-main">
    <div class="fp-paper"><svg id="fpSvg" viewBox="0 0 10 10" font-family="IBM Plex Sans Thai, Tahoma, sans-serif"></svg></div>
    <div class="small muted" style="margin-top:8px">ของที่ห้ามย้าย ${d.items.length} ชิ้น · คลิกเพื่อเลือก · ลากคลุมบนพื้นที่ว่างเพื่อเลือกหลายชิ้น · ลากเพื่อย้าย · Delete ลบ · R หมุน · บันทึกอัตโนมัติ</div>
    <section class="card" style="margin-top:16px;overflow:hidden">
      <div class="card-hd"><div><h2>ขนาดมาตรฐานของสิ่งของ</h2><div class="small muted">ใช้กับของที่ไม่ได้ตั้งขนาดเฉพาะชิ้น ในทุกห้อง (เมตร) · โต๊ะยาว แถวเก้าอี้ และโต๊ะห้องเรียน ยาวตามจำนวนที่นั่ง</div></div></div>
      <div style="overflow-x:auto"><table class="tbl"><thead><tr><th>สิ่งของ</th><th>กว้าง</th><th>ลึก</th><th class="r"></th></tr></thead><tbody>
      ${SIZED.map(t => { const v = sizeOf(t), custom = !!sizes[t]; return `<tr><td>${H(NAMES[t])}${t === 'round' ? ' <span class="xs muted">(เส้นผ่านศูนย์กลาง)</span>' : ''}</td>
        <td><input class="input" style="width:90px;height:36px" type="number" min="0.1" max="40" step="0.1" value="${v[0]}" data-change="fsSize" data-t="${t}" data-k="0" aria-label="${H(NAMES[t])} กว้าง"></td>
        <td>${t === 'round' ? '—' : `<input class="input" style="width:90px;height:36px" type="number" min="0.1" max="40" step="0.1" value="${v[1]}" data-change="fsSize" data-t="${t}" data-k="1" aria-label="${H(NAMES[t])} ลึก">`}</td>
        <td class="r">${custom ? `<button class="link" data-act="fsReset" data-t="${t}">คืนค่าเดิม (${DEF_SIZE[t][0]}${t === 'round' ? '' : '×' + DEF_SIZE[t][1]})</button>` : '<span class="xs muted">ค่าเดิม</span>'}</td></tr>`; }).join('')}
      </tbody></table></div>
    </section>
  </div>
</div>`;
    },
    after() { drawSvg(); bindSvg(); }
  };
  function setSizes(fn) { const st = Store.db.settings, f = st.floor = Object.assign({ rooms: {}, sizes: {} }, st.floor); f.sizes = Object.assign({}, f.sizes); fn(f.sizes); Store.save(); App.render(); }
  Object.assign(App.CHG, {
    fsW(el) { const d = setupDoc(); d.W = Math.max(4, Math.min(80, Number(el.value) || d.W)); saveSetup(); record(); App.render(); },
    fsH(el) { const d = setupDoc(); d.H = Math.max(4, Math.min(80, Number(el.value) || d.H)); saveSetup(); record(); App.render(); },
    fsSize(el) { const t = el.dataset.t, k = Number(el.dataset.k), v = Math.max(0.1, Math.min(40, Number(el.value) || 0.1)); setSizes(z => { const cur = sizeOf(t).slice(); cur[k] = v; if (t === 'round') cur[1] = v; z[t] = cur; }); }
  });
  App.ACT.fsReset = el => setSizes(z => { delete z[el.dataset.t]; });

  function histInfo(i) { const h = HIST[i], d = HIST_ROOM[h.r] && roomDim(HIST_ROOM[h.r]); return (h.p && d ? 'พื้นที่ต่อคน ≈ ' + (d[0] * d[1] / h.p).toFixed(2) + ' ตร.ม. · ' : '') + h.note + (h.g ? '' : ' · (ยังไม่มีพิกัดพอจะโหลดเป็นผัง)'); }
  function warn(p, pax) {
    const seats = seatsOf(p), out = [];
    if (pax && seats && seats !== pax) out.push(seats < pax ? `ที่นั่งน้อยกว่าจำนวนท่าน ${pax - seats} ที่` : `ที่นั่งมากกว่าจำนวนท่าน ${seats - pax} ที่`);
    const G = geom(p);
    if (p.items.some(i => i.x < 0 || i.y < 0 || i.x > G.W || i.y > G.H)) out.push('มีของวางเกินขอบห้อง');
    if (p.missing) out.push(`วางไม่พอ ${p.missing} โต๊ะในห้องขนาดนี้ — ลองลดขนาดโต๊ะ ขยายห้อง หรือจัดชิดขึ้น`);
    return out.length ? `<div class="note" style="margin-top:8px">${out.map(H).join(' · ')}</div>` : '';
  }
  const picked = p => p.items.filter(i => sel.includes(i.id));
  function selPanel(p) {
    const list = picked(p), tip = '<div class="xs muted" style="margin-top:8px">ลากบนพื้นที่ว่างเพื่อคลุมเลือกหลายชิ้น · Shift/Ctrl + คลิก เพื่อเพิ่ม/เอาออกทีละชิ้น · ลูกศรขยับทีละ 10 ซม.</div>';
    if (!list.length) return `<h3>ที่เลือกอยู่</h3><div class="small muted">แตะที่โต๊ะหรือของในผังเพื่อเลือก</div>${p.items.length ? '<button class="btn ghost sm" style="width:100%;margin-top:8px" data-act="fpAll">เลือกทั้งหมด</button>' : ''}${tip}`;
    const seated = list.filter(i => SEATED.includes(i.type)), seatBtns = `<button class="icon-btn" data-act="fpSeat" data-v="-1" aria-label="ลดที่นั่ง">−</button><button class="icon-btn" data-act="fpSeat" data-v="1" aria-label="เพิ่มที่นั่ง">+</button>`;
    const actions = `<div class="fp-grid"><button class="btn ghost sm" data-act="fpRot">⟳ หมุน 15°</button><button class="btn ghost sm" data-act="fpDup">คัดลอก</button><button class="btn ghost sm fp-wide" data-act="fpDel" style="color:var(--red)">ลบ${list.length > 1 ? ' ' + list.length + ' ชิ้น' : ''}</button></div>`;
    if (list.length > 1) {
      const kinds = Object.entries(list.reduce((m, i) => (m[NAMES[i.type]] = (m[NAMES[i.type]] || 0) + 1, m), {})).map(([k, n]) => k + ' ' + n).join(' · ');
      return `<h3>เลือกอยู่ ${list.length} ชิ้น</h3><div class="small muted" style="margin-bottom:8px">${H(kinds)}</div>
        ${seated.length ? `<div class="row" style="margin-bottom:8px"><span style="flex:1">ที่นั่งต่อโต๊ะ (${seated.length} โต๊ะ) รวม <b>${seated.reduce((a, i) => a + i.seats, 0)}</b></span>${seatBtns}</div>` : ''}
        ${actions}<button class="link" style="margin-top:8px" data-act="fpNone">ยกเลิกการเลือก</button>${tip}`;
    }
    const it = list[0];
    return `<h3>ที่เลือกอยู่ · ${H(it.type === 'custom' && it.label ? it.label : NAMES[it.type])}</h3>
      ${SEATED.includes(it.type) ? `<div class="row" style="margin-bottom:8px"><span style="flex:1">ที่นั่ง: <b>${it.seats}</b></span>${seatBtns}</div>` : ''}
      ${SIZED.includes(it.type) || it.type === 'custom' ? `<div class="small muted" style="margin-bottom:4px">${it.type === 'round' ? 'เส้นผ่านศูนย์กลาง (ม.)' : 'ขนาด กว้าง × ลึก (ม.)'}</div>
        <div class="row" style="gap:6px;margin-bottom:8px"><input class="input" type="number" min="0.1" max="40" step="0.1" value="${dims(it)[0]}" data-change="fpSize" data-k="w" aria-label="กว้าง">${it.type === 'round' ? '' : `<span>×</span><input class="input" type="number" min="0.1" max="40" step="0.1" value="${dims(it)[1]}" data-change="fpSize" data-k="h" aria-label="ลึก">`}</div>` : ''}
      ${it.type === 'block' || it.type === 'custom' ? `<input class="input" value="${H(it.label || '')}" placeholder="ข้อความบนบล็อก เช่น ห้องควบคุม" data-change="fpLabel" aria-label="ข้อความบนบล็อก" style="margin-bottom:8px">` : ''}
      ${actions}${tip}`;
  }
  const setupMode = () => App.state.route === 'plansetup';
  const P = () => setupMode() ? setupDoc() : planOf(curBeo(), dayOf());
  function drawSvg() {
    const p = P(), s = document.getElementById('fpSvg'); if (!p || !s) return;
    const b = build(p, sel, setupMode()); s.setAttribute('viewBox', `0 0 ${b.vw} ${b.vh}`);
    s.innerHTML = b.svg + (marq ? `<rect x="${(Math.min(marq.x0, marq.x1) + 1) * S}" y="${(Math.min(marq.y0, marq.y1) + 2) * S}" width="${Math.abs(marq.x1 - marq.x0) * S}" height="${Math.abs(marq.y1 - marq.y0) * S}" fill="rgba(168,38,45,.08)" stroke="#A8262D" stroke-width="1.5" stroke-dasharray="5 4"/>` : '');
    const t = document.getElementById('fpTot'); if (t) t.textContent = seatsOf(p);
  }
  function refreshSide() {
    const p = P(), el = document.getElementById('fpSel'), w = document.getElementById('fpWarn');
    if (el) el.innerHTML = selPanel(p);
    if (w && !setupMode()) w.innerHTML = warn(p, Number(Store.quote(curBeo().quoteId).pax) || 0);
  }
  function save() { const p = P(); if (p) delete p.missing; if (setupMode()) saveSetup(); else Store.save(); record(); }
  function bindSvg() {
    const s = document.getElementById('fpSvg'); if (!s || !(setupMode() ? isAdmin() : canEdit())) return;
    const pos = e => { const pt = s.createSVGPoint(); pt.x = e.clientX; pt.y = e.clientY; const q = pt.matrixTransform(s.getScreenCTM().inverse()); return [(q.x - S) / S, (q.y - 2 * S) / S]; };
    s.addEventListener('pointerdown', e => {
      const gEl = e.target.closest('[data-id]'), p = P(), [x, y] = pos(e), add = e.shiftKey || e.ctrlKey || e.metaKey;
      e.preventDefault(); s.setPointerCapture(e.pointerId);
      if (!gEl) {
        // empty floor: start a selection box (Shift/Ctrl keeps what is already selected)
        marq = { x0: x, y0: y, x1: x, y1: y, base: add ? sel.slice() : [] };
        if (!add) sel = [];
        drawSvg(); refreshSide(); return;
      }
      const id = Number(gEl.dataset.id);
      if (add) { sel = sel.includes(id) ? sel.filter(v => v !== id) : sel.concat(id); drawSvg(); refreshSide(); return; }
      if (!sel.includes(id)) sel = [id];
      // drag every selected item together
      drag = { x, y, moved: false, orig: picked(p).map(i => [i, i.x, i.y]) };
      drawSvg(); refreshSide();
    });
    s.addEventListener('pointermove', e => {
      const [x, y] = pos(e);
      if (marq) {
        marq.x1 = x; marq.y1 = y;
        const r = { x0: Math.min(marq.x0, x), x1: Math.max(marq.x0, x), y0: Math.min(marq.y0, y), y1: Math.max(marq.y0, y) };
        // an item is picked when its centre is inside the box, so grazing a neighbour does not grab it
        const hit = P().items.filter(i => i.x >= r.x0 && i.x <= r.x1 && i.y >= r.y0 && i.y <= r.y1).map(i => i.id);
        sel = marq.base.concat(hit.filter(id => !marq.base.includes(id)));
        drawSvg(); return;
      }
      if (!drag) return;
      const dx = x - drag.x, dy = y - drag.y;
      for (const [i, ox, oy] of drag.orig) { i.x = Math.round((ox + dx) * 10) / 10; i.y = Math.round((oy + dy) * 10) / 10; }
      drag.moved = true; drawSvg();
    });
    const end = () => {
      if (marq) { marq = null; drawSvg(); refreshSide(); }
      if (drag && drag.moved) { save(); refreshSide(); }
      drag = null;
    };
    s.addEventListener('pointerup', end); s.addEventListener('pointercancel', end);
  }
  // Delete / R while the editor is open (not while typing or in a dialog)
  document.addEventListener('keydown', e => {
    if (!(App.state.route === 'plan' ? canEdit() : setupMode() && isAdmin()) || document.getElementById('modalBg')) return;
    if (/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) || !P()) return;
    if ((e.ctrlKey || e.metaKey) && (e.key === 'a' || e.key === 'A')) { e.preventDefault(); App.ACT.fpAll(); return; }
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && (e.key === 'z' || e.key === 'Z')) { e.preventDefault(); App.ACT.fpUndo(); return; }
    if (!sel.length) return;
    if (e.key === 'Escape') App.ACT.fpNone();
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); App.ACT.fpDel(); }
    if (e.key === 'r' || e.key === 'R') App.ACT.fpRot();
    const mv = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (mv) { e.preventDefault(); const d = e.shiftKey ? 0.5 : 0.1; step(it => { it.x = Math.round((it.x + mv[0] * d) * 10) / 10; it.y = Math.round((it.y + mv[1] * d) * 10) / 10; }); }
  });

  // a template keeps the items and room size of a plan
  function tplOf(p, name) { const G = geom(p); return { id: 't' + Date.now().toString(36), name: String(name).slice(0, 60), room: p.room, W: G.W, H: G.H, items: p.items.map(({ id, ...rest }) => rest) }; }
  function cleanTpl(t) {
    if (!t || !Array.isArray(t.items) || !(Number(t.W) > 0 && Number(t.H) > 0)) return null;
    const items = t.items.filter(i => i && NAMES[i.type] && isFinite(i.x) && isFinite(i.y)).slice(0, 800).map(i => {
      const o = { type: i.type, x: Number(i.x), y: Number(i.y), rot: Number(i.rot) || 0, seats: Math.max(0, Math.min(40, Number(i.seats) || 0)) };
      for (const k of ['w', 'h']) if (Number(i[k]) > 0) o[k] = Math.min(40, Number(i[k]));
      if (i.label) o.label = String(i.label).slice(0, 40);
      if (i.type === 'custom') { o.shape = i.shape === 'round' ? 'round' : 'rect'; if (/^#[0-9a-f]{6}$/i.test(i.color || '')) o.color = i.color; }
      return o;
    });
    if (!items.length) return null;
    return { id: 't' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), name: String(t.name || 'เทมเพลตจากไฟล์').slice(0, 60), room: D.MEETING_ROOMS.some(r => r.id === t.room) ? t.room : undefined, W: Math.min(80, Number(t.W)), H: Math.min(80, Number(t.H)), items };
  }
  function addCustom(def) {
    const p = P(); if (!p) return; const G = setupMode() ? p : geom(p), n = p.items.length % 5;
    const it = mk(p, 'custom', G.W / 2 + n * 0.4, G.H / 2 + n * 0.4);
    Object.assign(it, { label: def.name, w: def.w, h: def.h, shape: def.shape, color: def.color });
    sel = [it.id]; save(); drawSvg(); refreshSide();
  }
  function libButtons() {
    return libItems().map(c => `<span class="fp-lib"><button class="btn ghost sm" data-act="fpAddC" data-id="${H(c.id)}">+ ${H(c.name)}</button>${canLib() ? `<button class="fp-x" data-act="fpItemDel" data-id="${H(c.id)}" aria-label="เอา ${H(c.name)} ออกจากรายการ">×</button>` : ''}</span>`).join('');
  }
  function setPlan(p) {
    const b = curBeo(); b.plans = b.plans || {}; b.plans[dayOf()] = p; sel = []; Store.save(); record(); App.render();
    if (p.missing) App.toast(`วางไม่พอ ${p.missing} โต๊ะ — ปรับผังเพิ่มเติมได้`);
  }
  // apply a change to every selected item
  const step = fn => { const p = P(), list = p ? picked(p) : []; if (!list.length) return; list.forEach(it => fn(it, p, list)); save(); drawSvg(); refreshSide(); };
  Object.assign(App.ACT, {
    fpAuto() { const b = curBeo(), q = Store.quote(b.quoteId); const go = () => setPlan(autoPlan(q, dayOf())); if (P() && P().items.length) App.confirm('แทนที่ผังปัจจุบันด้วยผังจากใบเสนอราคา?', go, 'แทนที่'); else go(); },
    fpBlank() { const b = curBeo(), q = Store.quote(b.quoteId), rf = roomFor(q, dayOf()), [w, h] = roomDim(rf.id); setPlan({ title: q.title || '', room: rf.room ? rf.id : undefined, W: w, H: h, items: [] }); },
    fpTpl(el) {
      const p = P(), t = el.dataset.v;
      const G = geom(p), go = () => { const n = t === 'blank' ? Object.assign({}, p, { items: [] }) : Object.assign({ room: p.room }, template(t, G.W, G.H), { title: p.title }); if (t === 'kin' || t === 'red') n.title = t === 'kin' ? 'งานเลี้ยง 250 ท่าน · ห้องกินรี' : 'งานเลี้ยง 200 ท่าน · ห้องกินรี · พรมแดงกลางห้อง'; setPlan(n); };
      if (p.items.length) App.confirm(t === 'blank' ? 'ล้างของทั้งหมดในผังนี้?' : 'แทนที่ผังปัจจุบันด้วยเทมเพลตนี้?', go, t === 'blank' ? 'ล้าง' : 'แทนที่'); else go();
    },
    fpHist() {
      const h = HIST[Number(document.getElementById('fpHist').value)];
      if (!h.g) { App.toast('แปลนนี้ยังไม่มีพิกัดพอจะโหลดเป็นผัง'); return; }
      const go = () => setPlan(histPlan(h));
      if (P().items.length) App.confirm('แทนที่ผังปัจจุบันด้วยแปลน “' + h.r + ' · ' + h.l + '”?', go, 'แทนที่'); else go();
    },
    fpAdd(el) { const p = P(), G = setupMode() ? p : geom(p), n = p.items.length % 5; sel = [mk(p, el.dataset.v, G.W / 2 + n * 0.4, G.H / 2 + n * 0.4).id]; save(); drawSvg(); refreshSide(); },
    fpUndo() {
      histSync(); if (!hist.stack.length) return;
      const prev = hist.stack.pop(), data = JSON.parse(prev);
      if (setupMode()) { const d = setupDoc(); Object.assign(d, { W: data.W, H: data.H, items: data.items }); saveSetup(); }
      else { const b = curBeo(); b.plans = b.plans || {}; if (data) b.plans[dayOf()] = data; else delete b.plans[dayOf()]; Store.save(); }
      hist.last = prev; sel = []; App.render(); App.toast('ย้อนกลับแล้ว');
    },
    fpTplNew() {
      const p = setupMode() ? null : P();
      App.modal(`<h2>เพิ่มเทมเพลต</h2>
        ${p && p.items.length ? `<div class="field"><label class="f" for="tpName">บันทึกผังนี้เป็นเทมเพลต</label><div class="row" style="gap:8px"><input class="input" id="tpName" value="${H(p.title || '')}" placeholder="ชื่อเทมเพลต เช่น กินรี Classroom 120"><button class="btn red" id="tpSave" style="flex:none">บันทึก</button></div></div>` : ''}
        <div class="field"><label class="f" for="tpFile">อัปโหลดไฟล์เทมเพลต (.json)</label><input class="input" id="tpFile" type="file" accept=".json,application/json" style="padding-top:10px"></div>
        <div class="xs muted">ไฟล์เทมเพลตได้จากปุ่ม “ดาวน์โหลดเป็นไฟล์เทมเพลต” ด้านล่าง หรือจากเครื่องอื่นของโรงแรม · เทมเพลตใช้ร่วมกันทุกคน</div>
        <div class="small red" id="tpErr" style="min-height:20px;margin-top:6px" role="alert"></div>
        <div class="acts">${p && p.items.length ? '<button class="btn ghost" id="tpDl">ดาวน์โหลดเป็นไฟล์เทมเพลต</button>' : ''}<button class="btn ghost" data-act="closeModal">ปิด</button></div>`, bg => {
        const err = bg.querySelector('#tpErr');
        const add = list => { setFloor(f => { f.templates = (f.templates || []).concat(list); }); App.closeModal(); App.render(); App.toast('เพิ่มเทมเพลต ' + list.map(t => t.name).join(', ') + ' แล้ว'); };
        const save = bg.querySelector('#tpSave');
        if (save) save.onclick = () => { const name = bg.querySelector('#tpName').value.trim(); if (!name) { err.textContent = 'กรุณาตั้งชื่อเทมเพลต'; return; } add([tplOf(p, name)]); };
        const dl = bg.querySelector('#tpDl');
        if (dl) dl.onclick = () => U.download('template-' + curBeo().id + '-day' + dayOf() + '.json', JSON.stringify(tplOf(p, p.title || 'ผังงาน'), null, 1));
        bg.querySelector('#tpFile').onchange = e => {
          const file = e.target.files[0]; if (!file) return;
          if (file.size > 500000) { err.textContent = 'ไฟล์ใหญ่เกินไป'; return; }
          file.text().then(text => {
            let j; try { j = JSON.parse(text); } catch (x) { err.textContent = 'ไฟล์ไม่ใช่ JSON ที่อ่านได้'; return; }
            const list = (Array.isArray(j) ? j : Array.isArray(j.templates) ? j.templates : [j]).map(cleanTpl).filter(Boolean);
            if (!list.length) { err.textContent = 'ไม่พบผังในไฟล์นี้ (ต้องมี items และขนาดห้อง W, H)'; return; }
            add(list);
          });
        };
      });
    },
    fpTplC(el) {
      const t = libTpl().find(x => x.id === el.dataset.id), p = P(); if (!t || !p) return;
      const go = () => setPlan(Object.assign({}, p, { items: t.items.map((it, n) => Object.assign({}, it, { id: n + 1 })) }, p.room ? {} : { W: t.W, H: t.H }));
      if (p.items.length) App.confirm('แทนที่ผังปัจจุบันด้วยเทมเพลต “' + t.name + '”?', go, 'แทนที่'); else go();
    },
    fpTplDel(el) { const t = libTpl().find(x => x.id === el.dataset.id); if (t) App.confirm('ลบเทมเพลต “' + t.name + '”?', () => { setFloor(f => { f.templates = (f.templates || []).filter(x => x.id !== t.id); }); App.render(); }, 'ลบ'); },
    fpItemNew() {
      const sw = ['#e4e2da', '#c9d6e8', '#d6e8d0', '#f1e2b8', '#f0c9c0', '#e6d5ef', '#cfe8e5'];
      App.modal(`<h2>${setupMode() ? 'เพิ่มของที่ห้ามย้ายชิ้นใหม่' : 'เพิ่มสิ่งของใหม่'}</h2>
        <div class="field"><label class="f" for="ciName">ชื่อ</label><input class="input" id="ciName" placeholder="เช่น โต๊ะค็อกเทล, ซุ้มดอกไม้, จอ LED"></div>
        <div class="grid g2" style="margin-top:14px"><div><label class="f" for="ciShape">รูปทรง</label><select class="input" id="ciShape"><option value="rect">สี่เหลี่ยม</option><option value="round">วงกลม / วงรี</option></select></div>
          <div><label class="f">สี</label><div class="row wrap" style="gap:6px;margin-top:6px" id="ciColors">${sw.map((c, i) => `<button type="button" class="fp-sw ${i ? '' : 'on'}" data-c="${c}" style="background:${c}" aria-label="สี ${c}"></button>`).join('')}</div></div></div>
        <div class="grid g2" style="margin-top:14px"><div><label class="f" for="ciW">กว้าง (ม.)</label><input class="input" id="ciW" type="number" min="0.1" max="40" step="0.1" value="1.2"></div><div><label class="f" for="ciH">ลึก (ม.)</label><input class="input" id="ciH" type="number" min="0.1" max="40" step="0.1" value="0.8"></div></div>
        ${canLib() ? '<label class="row small" style="gap:8px;margin-top:14px;cursor:pointer"><input type="checkbox" id="ciKeep" checked> บันทึกไว้ในรายการ “เพิ่มของ” ให้ทุกคนใช้ครั้งหน้า</label>' : ''}
        <div class="small red" id="ciErr" style="min-height:20px;margin-top:6px" role="alert"></div>
        <div class="acts"><button class="btn ghost" data-act="closeModal">ยกเลิก</button><button class="btn red" id="ciOk">${setupMode() ? 'เพิ่มเป็นของที่ห้ามย้าย' : 'เพิ่มลงผัง'}</button></div>`, bg => {
        let color = sw[0];
        bg.querySelectorAll('.fp-sw').forEach(b => { b.onclick = () => { color = b.dataset.c; bg.querySelectorAll('.fp-sw').forEach(x => x.classList.toggle('on', x === b)); }; });
        bg.querySelector('#ciOk').onclick = () => {
          const name = bg.querySelector('#ciName').value.trim().slice(0, 40), w = Number(bg.querySelector('#ciW').value), h = Number(bg.querySelector('#ciH').value);
          const err = bg.querySelector('#ciErr');
          if (!name) { err.textContent = 'กรุณาตั้งชื่อสิ่งของ'; return; }
          if (!(w >= 0.1 && w <= 40 && h >= 0.1 && h <= 40)) { err.textContent = 'ขนาดต้องอยู่ระหว่าง 0.1–40 เมตร'; return; }
          const def = { id: 'c' + Date.now().toString(36), name, w: Math.round(w * 10) / 10, h: Math.round(h * 10) / 10, shape: bg.querySelector('#ciShape').value, color };
          const keep = bg.querySelector('#ciKeep');
          if (keep && keep.checked) setFloor(f => { f.custom = (f.custom || []).concat(def); });
          App.closeModal(); addCustom(def); if (keep && keep.checked) App.render();
        };
      });
    },
    fpAddC(el) { const def = libItems().find(x => x.id === el.dataset.id); if (def) addCustom(def); },
    fpItemDel(el) { const c = libItems().find(x => x.id === el.dataset.id); if (c) App.confirm('เอา “' + c.name + '” ออกจากรายการเพิ่มของ? (ของที่วางในผังแล้วยังอยู่)', () => { setFloor(f => { f.custom = (f.custom || []).filter(x => x.id !== c.id); }); App.render(); }, 'เอาออก'); },
    fpAll() { const p = P(); if (!p) return; sel = p.items.map(i => i.id); drawSvg(); refreshSide(); },
    fpNone() { sel = []; drawSvg(); refreshSide(); },
    fpRot() {
      // several items turn together around their common centre
      const p = P(), list = p ? picked(p) : []; if (!list.length) return;
      const cx = list.reduce((a, i) => a + i.x, 0) / list.length, cy = list.reduce((a, i) => a + i.y, 0) / list.length, r = Math.PI / 12, c = Math.cos(r), sn = Math.sin(r);
      for (const i of list) {
        if (list.length > 1) { const dx = i.x - cx, dy = i.y - cy; i.x = Math.round((cx + dx * c - dy * sn) * 10) / 10; i.y = Math.round((cy + dx * sn + dy * c) * 10) / 10; }
        i.rot = (i.rot + 15) % 360;
      }
      save(); drawSvg(); refreshSide();
    },
    fpDup() {
      const p = P(), list = p ? picked(p) : []; if (!list.length) return;
      sel = list.map(it => { const n = mk(p, it.type, it.x + 1, it.y + 1, it.rot, it.seats); ['w', 'h', 'label'].forEach(k => { if (it[k] != null) n[k] = it[k]; }); return n.id; });
      save(); drawSvg(); refreshSide();
    },
    fpDel() { const p = P(); if (!p || !sel.length) return; p.items = p.items.filter(i => !sel.includes(i.id)); sel = []; save(); drawSvg(); refreshSide(); },
    fpSeat(el) { step(it => { if (SEATED.includes(it.type)) it.seats = Math.max(1, Math.min(it.type === 'chairrow' ? 40 : 16, it.seats + Number(el.dataset.v))); }); },
    fpPng() {
      const p = P(); if (!p) return;
      const b = build(p, null, setupMode()), sc = 2, x = `<svg xmlns="http://www.w3.org/2000/svg" width="${b.vw * sc}" height="${b.vh * sc}" viewBox="0 0 ${b.vw} ${b.vh}" font-family="IBM Plex Sans Thai, Tahoma, sans-serif">${b.svg}</svg>`;
      const img = new Image();
      img.onload = () => { const c = document.createElement('canvas'); c.width = b.vw * sc; c.height = b.vh * sc; c.getContext('2d').drawImage(img, 0, 0); const a = document.createElement('a'); a.download = setupMode() ? 'room-' + setupRoom() + '.png' : 'floorplan-' + curBeo().id + '-day' + dayOf() + '.png'; a.href = c.toDataURL('image/png'); a.click(); };
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(x);
    }
  });
  Object.assign(App.CHG, {
    fpSize(el) {
      step(it => { const d = dims(it), v = Math.max(0.1, Math.min(40, Number(el.value) || 0.1)); if (it.type === 'round') { it.w = it.h = v; } else { it.w = el.dataset.k === 'w' ? v : d[0]; it.h = el.dataset.k === 'h' ? v : d[1]; } });
    },
    fpLabel(el) { step(it => { it.label = el.value.slice(0, 40); }); },
    fpTitle(el) { P().title = el.value.slice(0, 120); save(); drawSvg(); },
    fpW(el) { P().W = Math.max(4, Math.min(80, Number(el.value) || 20)); save(); drawSvg(); refreshSide(); },
    fpH(el) { P().H = Math.max(4, Math.min(80, Number(el.value) || 14)); save(); drawSvg(); refreshSide(); }
  });
})(window);
