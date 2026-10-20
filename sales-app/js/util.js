/* Luck Swan Sales — shared helpers */
(function (g) {
  'use strict';
  const MS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const ML = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  const DS = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
  const DL = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

  const U = {
    MS, ML, DS, DL,
    esc(s) {
      return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    },
    // date <-> ISO yyyy-mm-dd (local)
    d(iso) { const [y, m, dd] = String(iso).split('-').map(Number); return new Date(y, m - 1, dd); },
    iso(dt) { const p = n => String(n).padStart(2, '0'); return dt.getFullYear() + '-' + p(dt.getMonth() + 1) + '-' + p(dt.getDate()); },
    // all timestamps are Thailand time (UTC+7), whatever the device's own time zone is
    nowTH() { return new Date(Date.now() + 7 * 3600e3).toISOString().slice(0, 19); },
    today() { return U.nowTH().slice(0, 10); },
    addDays(iso, n) { const x = U.d(iso); x.setDate(x.getDate() + n); return U.iso(x); },
    diffDays(a, b) { return Math.round((U.d(b) - U.d(a)) / 864e5); },
    be(y) { return y + 543; },
    fmtD(iso, o = {}) { // 12 พ.ย. 2569
      if (!iso) return '—';
      const x = U.d(iso);
      let s = x.getDate() + ' ' + (o.long ? ML : MS)[x.getMonth()];
      if (o.year !== false) s += ' ' + U.be(x.getFullYear());
      if (o.dow) s = (o.long ? DL : DS)[x.getDay()] + ' ' + s;
      return s;
    },
    fmtRange(a, b, o = {}) { // 12–13 พ.ย. 2569
      if (!b || a === b) return U.fmtD(a, o);
      const x = U.d(a), y = U.d(b), M = o.long ? ML : MS;
      const yr = o.year === false ? '' : ' ' + U.be(y.getFullYear());
      if (x.getMonth() === y.getMonth() && x.getFullYear() === y.getFullYear()) return x.getDate() + '–' + y.getDate() + ' ' + M[y.getMonth()] + yr;
      return x.getDate() + ' ' + M[x.getMonth()] + ' – ' + y.getDate() + ' ' + M[y.getMonth()] + yr;
    },
    daysOf(a, b) { const out = []; let c = a; while (c <= (b || a)) { out.push(c); c = U.addDays(c, 1); } return out; },
    n(v, dp = 0) {
      const x = Number(v) || 0;
      return x.toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp });
    },
    n2(v) { return U.n(v, 2); },
    short(v) { // 1.92M
      const x = Number(v) || 0;
      if (Math.abs(x) >= 1e6) return (x / 1e6).toFixed(2).replace(/\.?0+$/, '') + 'M';
      if (Math.abs(x) >= 1e3) return (x / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
      return String(Math.round(x));
    },
    initials(name) {
      const parts = String(name || '').replace(/^(คุณ|นาย|นาง|นางสาว|น\.ส\.)\s*/, '').trim().split(/\s+/);
      const cs = (parts[0] || '').match(/[ก-ฮA-Za-z]/g) || [];
      if (/^[A-Za-z]/.test(parts[0] || '') && parts[1]) return (cs[0] + parts[1].charAt(0)).toUpperCase();
      return (cs.slice(0, 2).join('') || (parts[0] || '?').charAt(0)).toUpperCase();
    },
    uid(p = '') { return p + Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-3); },
    clamp(v, a, b) { return Math.max(a, Math.min(b, v)); },
    hm(t) { const [h, m] = String(t || '0:0').split(':').map(Number); return h + (m || 0) / 60; },
    fromHm(x) { const h = Math.floor(x), m = Math.round((x - h) * 60); return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0'); },

    // Thai baht text
    bahtText(amount) {
      const NUM = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
      const POS = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
      const readInt = s => {
        s = s.replace(/^0+/, '');
        if (!s) return '';
        if (s.length > 6) return readInt(s.slice(0, -6)) + 'ล้าน' + readInt(s.slice(-6));
        let out = '';
        const L = s.length;
        for (let i = 0; i < L; i++) {
          const dgt = +s[i], p = L - i - 1;
          if (!dgt) continue;
          if (p === 1 && dgt === 1) out += 'สิบ';
          else if (p === 1 && dgt === 2) out += 'ยี่สิบ';
          else if (p === 0 && dgt === 1 && L > 1) out += 'เอ็ด';
          else out += NUM[dgt] + POS[p];
        }
        return out;
      };
      const v = Math.round((Number(amount) || 0) * 100);
      const baht = String(Math.floor(v / 100)), sat = v % 100;
      let t = readInt(baht);
      if (!t && !sat) return 'ศูนย์บาทถ้วน';
      t = t ? t + 'บาท' : '';
      t += sat ? readInt(String(sat)) + 'สตางค์' : 'ถ้วน';
      return t;
    },

    // Quote math. items: {price,q1,q2}
    lineTotal(it) { return (Number(it.price) || 0) * (Number(it.q1) || 0) * (it.q2 == null || it.q2 === '' ? 1 : Number(it.q2) || 0) * (it.free ? 0 : 1); },
    calc(q, settings) {
      const S = settings || { service: 0.10, vat: 0.07 };
      const D = g.DATA, items = q.items || [];
      /* meeting-room rules: a meeting room's rent is paid down by the lunch and coffee-break spend of the same quote
         (rent − food; free once the food reaches the rent, the room still listed). Kinnaree's LED screen is 3,000
         when the rest of the quote is 50,000 or more (before service and VAT), otherwise 5,000. */
      const isRoom = it => it.sec === 'meet' && D && D.MEETING_ROOMS.some(r => r.id === it.ref);
      const isMeeting = it => isRoom(it) && !['banquet', 'cocktail'].includes(it.layout);
      const isMeal = it => it.sec === 'food' && D && ['lunch', 'coffee'].includes((D.FOOD.find(f => f.id === it.ref) || {}).cat);
      const isLed = it => it.ref === 'svc:ledk';
      const roomFee = items.filter(isMeeting).reduce((a, it) => a + U.lineTotal(it), 0);
      const meals = items.filter(isMeal).reduce((a, it) => a + U.lineTotal(it), 0);
      const credit = roomFee ? Math.min(roomFee, meals) : 0;
      const rest = items.filter(it => !isLed(it)).reduce((a, it) => a + U.lineTotal(it), 0) - credit;
      for (const it of items) if (isLed(it)) it.price = rest >= 50000 ? 3000 : 5000;
      const by = { meet: 0, food: 0, room: 0, other: 0 };
      for (const it of items) by[it.sec] = (by[it.sec] || 0) + U.lineTotal(it);
      const gross = by.meet + by.food + by.room + by.other, sub = gross - credit;
      const pct = q.discount && q.discount.type === 'amt' ? (sub ? q.discount.value / sub : 0) : ((q.discount && q.discount.value) || 0) / 100;
      const disc = q.discount && q.discount.type === 'amt' ? Math.min(sub, Number(q.discount.value) || 0) : Math.round(sub * pct * 100) / 100;
      const after = sub - disc;
      const svc = Math.round(after * S.service * 100) / 100;
      const vat = Math.round((after + svc) * S.vat * 100) / 100;
      const net = Math.round((after + svc + vat) * 100) / 100;
      return { by, gross, credit, roomFee, roomFree: !!roomFee && credit >= roomFee, sub, disc, pct: pct * 100, svc, vat, net, perPax: q.pax ? net / q.pax : 0 };
    },
    debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; },
    download(name, text, type = 'application/json') {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([text], { type }));
      a.download = name; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    },
    parseCSV(text) {
      const rows = []; let row = [], cur = '', q = false;
      for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (q) { if (c === '"' && text[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
        else if (c === '"') q = true;
        else if (c === ',') { row.push(cur); cur = ''; }
        else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
        else cur += c;
      }
      if (cur || row.length) { row.push(cur); rows.push(row); }
      return rows.filter(r => r.some(x => x.trim()));
    }
  };
  g.U = U;
})(window);
