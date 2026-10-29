/* Luck Swan Sales — catalog, seed data and local store */
(function (g) {
  'use strict';
  const U = g.U;

  /* ---------------- static catalog ---------------- */
  const LAYOUTS = [
    { id: 'theatre', name: 'Theatre' },
    { id: 'classroom', name: 'Classroom' },
    { id: 'ushape', name: 'U-shape' },
    { id: 'banquet', name: 'Banquet' },
    { id: 'cocktail', name: 'Cocktail' }
  ];

  const MEETING_ROOMS = [
    { id: 'kinnaree', minBanquet: 80, name: 'Kinnaree', th: 'ห้องกินรี', area: 430, desc: 'โดมโค้ง ตกแต่งล้านนา', img: 'img/meet-kinnaree.jpg',
      cap: { theatre: [200, 330], classroom: [150, 300], ushape: [60, 130], banquet: [150, 320], cocktail: [250, 450] }, full: 30000, half: 15000,
      amen: ['Wi-Fi 1 Gbps', 'จอ LED 6×3 ม.', 'ไมค์ 4 ตัว', 'เวที', 'น้ำดื่ม'] },
    { id: 'thikhampon', minBanquet: 50, name: 'Thikhampon', th: 'ห้องทิฆัมพร', area: 170, desc: 'อาคารล้านนาส่วนตัว · มีห้องน้ำในตัว', img: 'img/meet-thikhampon.jpg',
      cap: { theatre: [100, 200], classroom: [50, 120], ushape: [30, 50], banquet: [50, 120], cocktail: [80, 150] }, full: 12000, half: 7000,
      amen: ['Wi-Fi 1 Gbps', 'จอ LED 4×2.5 ม.', 'ไมค์ 2 ตัว', 'น้ำดื่ม'] },
    { id: 'ruenhong', minBanquet: 50, name: 'Ruen Hong', th: 'เรือนหงส์', area: 140, desc: 'ห้องกระจก วิวสวน', img: 'img/meet-ruenhong.jpg',
      cap: { theatre: [50, 100], classroom: [30, 50], ushape: [30, 50], banquet: [50, 70], cocktail: [60, 100] }, full: 10000, half: 6000,
      amen: ['Wi-Fi', 'ทีวี 75"', 'ไมค์ 2 ตัว', 'น้ำดื่ม'] },
    { id: 'prayamengrai', minBanquet: 5, name: 'Praya Meng Rai', th: 'ห้องพญาเม็งราย', area: 70, desc: 'ห้องประชุมเล็กสีทอง', img: 'img/meet-prayamengrai.jpg',
      cap: { theatre: [10, 20], classroom: [10, 15], ushape: [10, 15], banquet: [10, 20], cocktail: [15, 25] }, full: 3000, half: 2000,
      amen: ['Wi-Fi', 'ทีวี 65"', 'ไวท์บอร์ด', 'น้ำดื่ม'] },
    { id: 'pool', name: 'ลานสระว่ายน้ำ', th: 'ลานสระว่ายน้ำ', area: 600, desc: 'พื้นที่กลางแจ้ง · งานเลี้ยง/ขันโตก', img: '',
      cap: { theatre: [0, 0], classroom: [0, 0], ushape: [0, 0], banquet: [100, 300], cocktail: [150, 350] }, full: 15000, half: 8000,
      amen: ['ไฟประดับ', 'เวทีกลางแจ้ง', 'เครื่องเสียงพื้นฐาน'] }
  ];
  // every room's photos (the first is its cover); the admin edits photos, prices, capacity, size and description
  // on the meeting-room page (settings.meetRooms), laid over these defaults — pages hold a reference to MEETING_ROOMS
  MEETING_ROOMS.forEach(r => { r.imgs = r.img ? [r.img] : []; });
  const MEETING_ROOMS_DEFAULT = JSON.parse(JSON.stringify(MEETING_ROOMS));
  const MEET_EDITABLE = ['imgs', 'full', 'half', 'cap', 'area', 'desc', 'minBanquet', 'amen'];
  function applyMeetRooms(over) {
    MEETING_ROOMS.forEach((r, i) => {
      const base = JSON.parse(JSON.stringify(MEETING_ROOMS_DEFAULT[i])), o = (over && over[r.id]) || {};
      for (const k of MEET_EDITABLE) if (o[k] != null) base[k] = k === 'cap' ? Object.assign(base.cap, o.cap) : o[k];
      base.img = (base.imgs || [])[0] || '';
      for (const k of Object.keys(r)) delete r[k];
      Object.assign(r, base);
    });
  }

  /* ---------- rooms: types from the hotel rate card, priced by season ----------
   * Every room type (a room family + bed) has a count and a full (rack) price; each season has a "with breakfast"
   * (ABF) and a "room only" price for 2 guests (Triple: 3). Prices are net (service charge and VAT included), as every quote line is (U.calc).
   * Admins change counts, prices and season dates on the room-rates page (settings.roomRates). */
  const ROOM_DEF = [
    // id, family, bed, zone, rooms, rack, pax, [season1 ABF, RO], [season2 ABF, RO], img
    ['sup', 'Superior ชั้น 3', 'Twin', 'Resident', 20, 1240, 2, [800, 550], [850, 600], 'img/room-superior.jpg'],
    ['dlx', 'Deluxe ชั้น 1-2', 'Twin', 'Resident', 19, 1340, 2, [850, 600], [900, 650], 'img/room-deluxe.jpg'],
    ['dlx_db', 'Deluxe ชั้น 1-2', 'Double', 'Resident', 9, 1440, 2, [900, 650], [950, 700], 'img/room-deluxe.jpg'],
    ['dlx_tr', 'Deluxe ชั้น 1-2', 'Triple (3 ท่าน)', 'Resident', 6, 1600, 3, [1150, 950], [1250, 1050], 'img/room-deluxe.jpg'],
    ['gpl', 'Grand Pool', 'Twin (queen)', 'Resident', 4, 1350, 2, [1100, 750], [1150, 800], ''],
    ['gch', 'Grand Chiang', 'Twin (queen)', 'Resident', 4, 1350, 2, [1100, 750], [1150, 800], ''],
    ['gdx', 'Grand Deluxe ชั้น 2-3', 'Twin', 'Resort', 16, 1600, 2, [1050, 850], [1200, 900], 'img/room-grand-deluxe.jpg'],
    ['gdx_db', 'Grand Deluxe ชั้น 2-3', 'Double', 'Resort', 4, 1650, 2, [1150, 900], [1250, 950], 'img/room-grand-deluxe.jpg'],
    ['gdw', 'Grand Deluxe (Wheel Chair Access) ชั้น 1', 'Twin', 'Resort', 8, 1700, 2, [1200, 950], [1300, 1000], 'img/room-grand-deluxe.jpg'],
    ['gdw_db', 'Grand Deluxe (Wheel Chair Access) ชั้น 1', 'Double', 'Resort', 2, 1750, 2, [1250, 1050], [1350, 1050], 'img/room-grand-deluxe.jpg'],
    ['jrs', 'Junior Suite', 'Double', 'Resort', 5, 2150, 2, [1500, 1250], [1650, 1300], 'img/room-junior-suite.jpg'],
    ['hms', 'Honeymoon Suite', 'Double', 'Resort', 1, 2150, 2, [1600, 1400], [1700, 1450], '']
  ];
  // seasons repeat every year; from/to are month-day, a season may wrap past the new year
  const SEASONS_DEFAULT = [{ id: 's1', name: '1 มี.ค. – 31 ก.ค.', from: '03-01', to: '07-31' }, { id: 's2', name: '1 ส.ค. – 28 ก.พ.', from: '08-01', to: '02-29' }];
  const ROOM_RATES_DEFAULT = {
    seasons: SEASONS_DEFAULT,
    types: Object.fromEntries(ROOM_DEF.map(r => [r[0], { total: r[4], rack: r[5], rates: { s1: r[7], s2: r[8] } }]))
  };
  const ROOM_TYPES = []; // filled by applyRoomRates (pages hold a reference)
  /* the boutique (ลักษวรรณ บูติก): rooms per room per night, each including congee, coffee and crackers for 2; breakfast and
   * a dinner party bought on top, per person, at three price levels. Admins change them on the room-rates page (roomRates.boutique). */
  const BOUTIQUE_DEFAULT = {
    rooms: [{ id: 'bq_std', name: 'Standard', price: 400 }, { id: 'bq_sup', name: 'Superior', price: 500 }, { id: 'bq_dlx', name: 'Deluxe', price: 600 }],
    incl: 'ข้าวต้ม กาแฟ แครกเกอร์ สำหรับ 2 ท่าน', breakfast: [50, 100, 150], dinner: [350, 399, 450]
  };
  const BOUTIQUE = JSON.parse(JSON.stringify(BOUTIQUE_DEFAULT));
  function applyBoutique(b) {
    b = b && typeof b === 'object' ? b : {};
    const num = (v, d) => (Number(v) >= 0 && v !== '' && v != null ? Math.round(Number(v)) : d);
    BOUTIQUE.rooms = BOUTIQUE_DEFAULT.rooms.map(r => Object.assign({}, r, { price: num(((b.rooms || []).find(x => x && x.id === r.id) || {}).price, r.price) }));
    BOUTIQUE.incl = String(b.incl || BOUTIQUE_DEFAULT.incl).slice(0, 120);
    for (const k of ['breakfast', 'dinner']) BOUTIQUE[k] = BOUTIQUE_DEFAULT[k].map((d, i) => num((b[k] || [])[i], d));
  }
  let SEASONS = SEASONS_DEFAULT;
  const md = s => /^\d{2}-\d{2}$/.test(s || '') ? s : null;
  function applyRoomRates(rr) {
    rr = rr && typeof rr === 'object' ? rr : {};
    const ss = (Array.isArray(rr.seasons) && rr.seasons.length ? rr.seasons : SEASONS_DEFAULT).filter(x => x && md(x.from) && md(x.to)).slice(0, 6);
    SEASONS = ss.length ? ss.map(x => ({ id: String(x.id), name: String(x.name || (x.from + ' – ' + x.to)).slice(0, 40), from: x.from, to: x.to })) : SEASONS_DEFAULT;
    const types = rr.types || {};
    applyBoutique(rr.boutique);
    ROOM_TYPES.length = 0;
    for (const [id, fam, bed, zone, total, rack, pax, s1, s2, img] of ROOM_DEF) {
      const o = types[id] || {}, num = (v, d) => (Number(v) >= 0 && v !== '' && v != null ? Number(v) : d);
      const rates = {};
      for (const se of SEASONS) { const def = se.id === 's2' ? s2 : s1, r = (o.rates || {})[se.id] || def; rates[se.id] = [num(r[0], def[0]), num(r[1], def[1])]; }
      ROOM_TYPES.push({ id, fam, bed, name: fam.replace(/ ชั้น .*$/, '') + ' · ' + bed, full: fam + ' · ' + bed, zone: 'Zone : ' + zone, zoneId: zone, total: Math.round(num(o.total, total)), rack: num(o.rack, rack), pax, img, rates,
        beds: bed + ' · ' + pax + ' ท่าน/ห้อง', size: '', get group() { return roomPrice(this, U.today(), true); } });
    }
  }
  // which season a date falls in (by month-day, repeating every year)
  function seasonOf(date) {
    const m = String(date || U.today()).slice(5, 10);
    const hit = SEASONS.find(se => se.from <= se.to ? m >= se.from && m <= se.to : m >= se.from || m <= se.to);
    return hit || SEASONS[0];
  }
  const roomPrice = (t, date, abf) => { const r = t.rates[seasonOf(date).id] || Object.values(t.rates)[0] || [0, 0]; return abf ? r[0] : r[1]; };
  // a stay split by season: [{ season, from, nights, price }], so a stay across the change of season is priced per night
  function roomStay(t, ci, nights, abf) {
    const out = [];
    for (let k = 0; k < nights; k++) {
      const d = U.addDays(ci, k), se = seasonOf(d), last = out[out.length - 1];
      if (last && last.season.id === se.id) last.nights++; else out.push({ season: se, from: d, nights: 1, price: roomPrice(t, d, abf) });
    }
    return out;
  }
  applyRoomRates(null);

  // catering categories (ประเภทจัดเลี้ยง): editable by price editors, kept in settings.foodCats; 'all' is always first and
  // 'package' (seminar packages) can be renamed but not removed, since the packages pages rely on it
  const FOOD_CATS_DEFAULT = [
    { id: 'all', name: 'ทั้งหมด' },
    { id: 'package', name: 'แพ็กเกจสัมมนา' },
    { id: 'coffee', name: 'Coffee break' },
    { id: 'lunch', name: 'อาหารกลางวัน' },
    { id: 'gala', name: 'Gala dinner' },
    { id: 'cocktail', name: 'Cocktail' },
    { id: 'wedding', name: 'งานแต่งงาน' }
  ];
  const FOOD_CATS = FOOD_CATS_DEFAULT.map(c => Object.assign({}, c));
  function applyFoodCats(list) {
    const src = Array.isArray(list) && list.length ? list.filter(c => c && c.id && c.id !== 'all') : FOOD_CATS_DEFAULT.slice(1);
    if (!src.some(c => c.id === 'package')) src.unshift(FOOD_CATS_DEFAULT[1]);
    FOOD_CATS.length = 0; FOOD_CATS.push(FOOD_CATS_DEFAULT[0]);
    src.forEach(c => FOOD_CATS.push({ id: String(c.id), name: String(c.name || c.id).slice(0, 40) }));
  }

  // photos used by the package item lists (img: null shows a "รอรูปเมนู" placeholder in the carousel)
  const PH = { coffee: 'img/food1-0.jpg', lunch: 'img/food2-0.jpg', dinner: 'img/food2-1.jpg', room: 'img/room-deluxe.jpg', meeting: 'img/meet-thikhampon.jpg' };
  const it = (name, qty, unit, detail, img) => ({ name, qty: qty || '', unit: unit || '', detail: detail || '', img: img || null });

  const FOOD = [
    { id: 'pkg-basic', cat: 'package', name: 'Basic Package', sub: '1 วัน 1 คืน', price: 790, unit: 'ท่าน', min: 30, rec: 2, tags: ['halal'], photos: [],
      items: [it('Coffee break', 2, 'มื้อ', '(ชา กาแฟ พร้อมขนม)', PH.coffee), it('Thai Lunch', 1, 'มื้อ', 'เมนู 7 รายการ', PH.lunch), it('ห้องพัก Deluxe เตียงคู่', '', '', 'พร้อมอาหารเช้า', PH.room), it('ห้องจัดประชุม', 8, 'ชั่วโมง/วัน', '', PH.meeting)] },
    { id: 'pkg-standard', cat: 'package', name: 'Standard Package', sub: '2 วัน 1 คืน', price: 1390, unit: 'ท่าน', min: 30, rec: 1, best: true, tags: ['halal'], photos: [],
      items: [it('Coffee break', 4, 'มื้อ', '(ชา กาแฟ พร้อมขนม)', PH.coffee), it('Thai Lunch', 2, 'มื้อ', 'เมนู 7 รายการ', PH.lunch), it('Thai Dinner', 1, 'มื้อ', 'เมนู 7 รายการ', PH.dinner), it('ห้องพัก Deluxe เตียงคู่', '', '', 'พร้อมอาหารเช้า', PH.room), it('ห้องจัดประชุม', 8, 'ชั่วโมง/วัน', '', PH.meeting)] },
    { id: 'pkg-premium', cat: 'package', name: 'Premium Package', sub: '2 วัน 1 คืน', price: 1520, unit: 'ท่าน', min: 30, rec: 3, tags: ['halal'], photos: [],
      items: [it('Coffee break', 4, 'มื้อ', '(ชา กาแฟ พร้อมขนม)', PH.coffee), it('Thai Lunch', 2, 'มื้อ', 'เมนู 7 รายการ', PH.lunch), it('Thai Dinner', 1, 'มื้อ', '', PH.dinner), it('Soft drink และสิทธิพิเศษเพิ่มเติม'), it('ห้องพัก Deluxe เตียงคู่', '', '', 'พร้อมอาหารเช้า', PH.room)] },

    { id: 'cb-morning', cat: 'coffee', name: 'Morning Bright', price: 180, unit: 'ท่าน/รอบ', tags: ['halal', 'veg'], feat: ['ชา กาแฟ น้ำผลไม้', 'ขนมไทย 2 อย่าง', 'เบเกอรี่ 1 อย่าง'] },
    { id: 'cb-thaisweet', cat: 'coffee', name: 'Thai Sweet Break', price: 220, unit: 'ท่าน/รอบ', tags: ['halal', 'veg'], feat: ['ชา กาแฟ ชาสมุนไพร', 'ขนมไทยล้านนา 3 อย่าง'] },
    { id: 'cb-healthy', cat: 'coffee', name: 'Healthy Break', price: 200, unit: 'ท่าน/รอบ', tags: ['veg', 'gf'], feat: ['น้ำผลไม้สด', 'ผลไม้ตามฤดูกาล', 'โยเกิร์ต'] },
    { id: 'cb-bakery', cat: 'coffee', name: 'Bakery Corner', price: 240, unit: 'ท่าน/รอบ', tags: ['veg'], feat: ['ครัวซองต์ / มัฟฟิน', 'ชา กาแฟสด'] },
    { id: 'cb-lanna', cat: 'coffee', name: 'Lanna Snack', price: 190, unit: 'ท่าน/รอบ', tags: ['halal'], feat: ['ข้าวแต๋น ขนมจ็อก', 'ชาใบหม่อน'] },
    { id: 'cb-fruit', cat: 'coffee', name: 'Fresh Fruit Break', price: 160, unit: 'ท่าน/รอบ', tags: ['halal', 'veg', 'gf'], feat: ['ผลไม้ 4 ชนิด', 'น้ำสมุนไพร'] },
    { id: 'cb-premium', cat: 'coffee', name: 'Premium Coffee Bar', price: 290, unit: 'ท่าน/รอบ', tags: ['veg'], feat: ['บาริสต้าประจำจุด', 'ขนม 3 อย่าง'] },
    { id: 'cb-basic', cat: 'coffee', name: 'Basic Break', price: 120, unit: 'ท่าน/รอบ', tags: ['halal', 'veg'], feat: ['ชา กาแฟ', 'คุกกี้'] },

    { id: 'ln-buffet', cat: 'lunch', name: 'Thai–International Buffet', price: 650, unit: 'ท่าน/มื้อ', place: 'Swan Café', tags: ['halal'], feat: ['อาหารไทย-นานาชาติ 12 รายการ', 'สลัดบาร์ ของหวาน ผลไม้'] },
    { id: 'ln-bento', cat: 'lunch', name: 'Healthy Bento Set', price: 420, unit: 'ท่าน/มื้อ', tags: ['gf'], feat: ['เซตเบนโตะ 5 ช่อง', 'น้ำผลไม้'] },
    { id: 'lb-250', cat: 'lunch', name: 'บุฟเฟ่ต์อาหารกลางวัน 250', price: 250, unit: 'ท่าน/มื้อ', rc: true, tags: [], feat: ['บุฟเฟ่ต์อาหารกลางวัน'] },
    { id: 'lb-300', cat: 'lunch', name: 'บุฟเฟ่ต์อาหารกลางวัน 300', price: 300, unit: 'ท่าน/มื้อ', rc: true, tags: [], feat: ['บุฟเฟ่ต์อาหารกลางวัน'] },
    { id: 'lb-350', cat: 'lunch', name: 'บุฟเฟ่ต์อาหารกลางวัน 350', price: 350, unit: 'ท่าน/มื้อ', rc: true, tags: [], feat: ['บุฟเฟ่ต์อาหารกลางวัน'] },
    { id: 'lb-400', cat: 'lunch', name: 'บุฟเฟ่ต์อาหารกลางวัน 400', price: 400, unit: 'ท่าน/มื้อ', rc: true, tags: [], feat: ['บุฟเฟ่ต์อาหารกลางวัน'] },
    { id: 'lb-450', cat: 'lunch', name: 'บุฟเฟ่ต์อาหารกลางวัน 450', price: 450, unit: 'ท่าน/มื้อ', rc: true, tags: [], feat: ['บุฟเฟ่ต์อาหารกลางวัน'] },
    { id: 'cbk-50', cat: 'coffee', name: 'เบรก 50', price: 50, unit: 'ท่าน/รอบ', rc: true, tags: [], feat: ['ของว่าง / เบรกระหว่างวัน'] },
    { id: 'cbk-60', cat: 'coffee', name: 'เบรก 60', price: 60, unit: 'ท่าน/รอบ', rc: true, tags: [], feat: ['ของว่าง / เบรกระหว่างวัน'] },
    { id: 'cbk-70', cat: 'coffee', name: 'เบรก 70', price: 70, unit: 'ท่าน/รอบ', rc: true, tags: [], feat: ['ของว่าง / เบรกระหว่างวัน'] },
    { id: 'ln-thai7', cat: 'lunch', name: 'Thai Lunch 7 รายการ', price: 450, unit: 'ท่าน/มื้อ', tags: ['halal'], feat: ['เมนูไทย 7 รายการ (โต๊ะจีน/บุฟเฟต์)'] },
    { id: 'ln-northern', cat: 'lunch', name: 'Northern Thai Set', price: 520, unit: 'ท่าน/มื้อ', tags: [], feat: ['ข้าวซอย แกงฮังเล น้ำพริกหนุ่ม'] },
    { id: 'ln-veg', cat: 'lunch', name: 'Vegetarian Buffet', price: 480, unit: 'ท่าน/มื้อ', tags: ['veg', 'halal'], feat: ['อาหารมังสวิรัติ 9 รายการ'] },
    { id: 'ln-intl', cat: 'lunch', name: 'International Buffet', price: 750, unit: 'ท่าน/มื้อ', tags: [], feat: ['อาหารนานาชาติ 14 รายการ', 'สถานีพาสต้า'] },
    { id: 'ln-box', cat: 'lunch', name: 'Lunch Box', price: 250, unit: 'ท่าน/มื้อ', tags: ['halal'], feat: ['ข้าวกล่อง + ผลไม้ + น้ำ'] },
    { id: 'ln-chinese', cat: 'lunch', name: 'Chinese Table (10 ท่าน)', price: 5500, unit: 'โต๊ะ', tags: [], feat: ['โต๊ะจีน 8 รายการ'] },
    { id: 'ln-halal', cat: 'lunch', name: 'Halal Buffet', price: 620, unit: 'ท่าน/มื้อ', tags: ['halal'], feat: ['อาหารฮาลาล 10 รายการ'] },

    { id: 'gd-khantoke', cat: 'gala', name: 'ขันโตกล้านนา (Northern Khantoke)', price: 1200, unit: 'ท่าน/มื้อ', place: 'ลานสระว่ายน้ำ', tags: [], feat: ['สำรับขันโตก 7 อย่าง', 'การแสดงฟ้อนเล็บ'] },
    { id: 'gd-bbq', cat: 'gala', name: 'Poolside BBQ', price: 1100, unit: 'ท่าน/มื้อ', place: 'ลานสระว่ายน้ำ', tags: [], feat: ['บาร์บีคิว ซีฟู้ด', 'สลัดบาร์'] },
    { id: 'gd-intl', cat: 'gala', name: 'International Gala Buffet', price: 990, unit: 'ท่าน/มื้อ', tags: ['halal'], feat: ['บุฟเฟต์นานาชาติ 16 รายการ'] },
    { id: 'gd-thai', cat: 'gala', name: 'Royal Thai Set', price: 1350, unit: 'ท่าน/มื้อ', tags: [], feat: ['อาหารไทยชาววัง 6 คอร์ส'] },
    { id: 'gd-chinese', cat: 'gala', name: 'Chinese Banquet', price: 8800, unit: 'โต๊ะ', tags: [], feat: ['โต๊ะจีน 9 รายการ'] },
    { id: 'gd-western', cat: 'gala', name: 'Western Set Dinner', price: 1450, unit: 'ท่าน/มื้อ', tags: [], feat: ['4 คอร์ส'] },

    { id: 'ck-sunset', cat: 'cocktail', name: 'Sunset Cocktail', price: 890, unit: 'ท่าน', sub: '2 ชม.', tags: [], feat: ['คานาเป้ 8 อย่าง', 'ม็อกเทล/ค็อกเทล 2 ชม.'] },
    { id: 'ck-welcome', cat: 'cocktail', name: 'Welcome Drink', price: 150, unit: 'ท่าน', tags: ['halal', 'veg'], feat: ['น้ำผลไม้ต้อนรับ'] },
    { id: 'ck-canape', cat: 'cocktail', name: 'Canapé Party', price: 650, unit: 'ท่าน', tags: [], feat: ['คานาเป้ 10 อย่าง'] },
    { id: 'ck-freeflow', cat: 'cocktail', name: 'Free-flow Beer 3 ชม.', price: 590, unit: 'ท่าน', tags: [], feat: ['เบียร์สด 3 ชั่วโมง'] },

    { id: 'wd-silver', cat: 'wedding', name: 'Wedding Silver', price: 990, unit: 'ท่าน', tags: [], feat: ['บุฟเฟต์ 10 รายการ', 'ซุ้มถ่ายรูป'] },
    { id: 'wd-gold', cat: 'wedding', name: 'Wedding Gold', price: 1290, unit: 'ท่าน', tags: [], feat: ['บุฟเฟต์ 14 รายการ', 'ดอกไม้เวที'] },
    { id: 'wd-chinese', cat: 'wedding', name: 'Wedding Chinese Table', price: 9900, unit: 'โต๊ะ', tags: [], feat: ['โต๊ะจีน 9 รายการ'] },
    { id: 'wd-morning', cat: 'wedding', name: 'พิธีเช้า (ไทย)', price: 35000, unit: 'งาน', tags: [], feat: ['จัดสถานที่พิธีหมั้น/รดน้ำ'] },
    { id: 'wd-cocktail', cat: 'wedding', name: 'Wedding Cocktail', price: 790, unit: 'ท่าน', tags: [], feat: ['คานาเป้ + เครื่องดื่ม'] }
  ];

  const SERVICES = [
    { id: 'av', name: 'AV Package', desc: 'โปรเจกเตอร์คู่ ไมค์ไร้สาย 4 ตัว พร้อมช่างเทคนิค', price: 8000, unit: 'วัน' },
    { id: 'led', name: 'จอ LED เพิ่ม', desc: 'จอ LED 3×2 ม. พร้อมติดตั้ง', price: 12000, unit: 'วัน' },
    { id: 'ledk', name: 'จอ LED ห้องกินรี', desc: 'ใช้ได้ตลอดการใช้งานห้อง · ยอดรวม 50,000 บาทขึ้นไป 3,000 บาท / ต่ำกว่านั้น 5,000 บาท', price: 5000, unit: 'งาน' },
    { id: 'photo', name: 'ช่างภาพ', desc: 'ช่างภาพ 1 ท่าน 8 ชม.', price: 6500, unit: 'วัน' },
    { id: 'stage', name: 'ตกแต่งเวที/ป้าย', desc: 'ป้ายไวนิล + ดอกไม้หน้าเวที', price: 4500, unit: 'งาน' },
    { id: 'shuttle', name: 'รถรับส่ง', desc: 'รถตู้ VIP พร้อมคนขับ', price: 2500, unit: 'เที่ยว' }
  ];

  const DEPTS_DEFAULT = [
    { id: 'kitchen', name: 'ครัว / อาหารและเครื่องดื่ม', short: 'ครัว/F&B', color: '#A8262D', bg: '#F8E1E1', fg: '#8E1F25' },
    { id: 'banquet', name: 'จัดเลี้ยง (Banquet)', short: 'จัดเลี้ยง', color: '#0E9F8E', bg: '#DDF3EF', fg: '#0B7C6F' },
    { id: 'av', name: 'ช่าง / ระบบเสียงและภาพ', short: 'ช่าง/AV', color: '#2B3F86', bg: '#E3E7F5', fg: '#2B3F86' },
    { id: 'front', name: 'ส่วนหน้า (Front Office)', short: 'ส่วนหน้า', color: '#D4A21A', bg: '#FFF1C2', fg: '#7a5c00' },
    { id: 'hk', name: 'แม่บ้าน (Housekeeping)', short: 'แม่บ้าน', color: '#6B716E', bg: '#EDECE6', fg: '#4A524F' },
    { id: 'acc', name: 'บัญชี (Accounting)', short: 'บัญชี', color: '#17211F', bg: '#EDECE6', fg: '#17211F' }
  ];

  // departments: the built-in six, renamed or added to on the staff form (settings.depts); pages hold a reference to DEPTS
  const DEPTS = [];
  const DEPT_TINTS = [['#7B4FA0', '#EFE6F6', '#5E3A7D'], ['#2F7D3A', '#E2F2E4', '#25632E'], ['#C2571A', '#FBE7DA', '#93410f'], ['#1F7FA8', '#DDF0F8', '#155f7f'], ['#9C2F6B', '#F7E1EC', '#7a2453']];
  function applyDepts(list) {
    const src = Array.isArray(list) && list.length ? list : DEPTS_DEFAULT;
    DEPTS.length = 0;
    src.forEach((d, i) => {
      const base = DEPTS_DEFAULT.find(x => x.id === d.id), t = DEPT_TINTS[i % DEPT_TINTS.length];
      DEPTS.push(Object.assign({ color: t[0], bg: t[1], fg: t[2] }, base, { id: d.id, name: d.name || (base && base.name) || d.id, short: d.short || d.name || (base && base.short) || d.id }));
    });
  }
  applyDepts(null);

  const ROLES = [
    { id: 'sales', name: 'พนักงานขาย' },
    { id: 'manager', name: 'ผู้จัดการฝ่ายขาย' },
    { id: 'head', name: 'หัวหน้าแผนก' },
    { id: 'acc', name: 'บัญชี' },
    { id: 'admin', name: 'ผู้ดูแลระบบ' }
  ];
  const PERMS = [
    { id: 'create_quote', name: 'สร้างและส่งใบเสนอราคา' },
    { id: 'discount10', name: 'ให้ส่วนลดได้ถึง 10%' },
    { id: 'discount_over', name: 'อนุมัติส่วนลดเกิน 10%', note: 'เฉพาะผู้จัดการฝ่ายขายขึ้นไป' },
    { id: 'view_beo', name: 'ดูใบฟังก์ชันและปฏิทินการจอง' },
    { id: 'customers', name: 'จัดการข้อมูลลูกค้า' },
    { id: 'staff', name: 'จัดรอบพนักงาน / อนุมัติต่อรอบ' },
    { id: 'users', name: 'จัดการผู้ใช้งาน' },
    { id: 'edit_prices', name: 'แก้ไขราคาและแพ็กเกจ', note: 'เฉพาะผู้ดูแลระบบ' },
    { id: 'leave_hr', name: 'ดูข้อมูลรายละเอียดใบลา (ฝ่ายทรัพยากรบุคคล)', note: 'การอนุมัติใบลาเป็นของผู้ดูแลระบบเท่านั้น' }
  ];
  const ROLE_PERMS = {
    sales: ['create_quote', 'discount10', 'view_beo', 'customers'],
    manager: ['create_quote', 'discount10', 'discount_over', 'view_beo', 'customers', 'staff', 'users'],
    head: ['view_beo', 'staff'],
    acc: ['view_beo', 'customers'],
    admin: PERMS.map(p => p.id)
  };
  const STATUS = {
    draft: 'ร่าง', sent: 'ส่งแล้ว', tentative: 'จองชั่วคราว', confirmed: 'ยืนยันแล้ว', cancelled: 'ยกเลิก'
  };

  /* ---------------- catalog helpers (menus can be edited in the app) ---------------- */
  const itemText = x => [x.name, x.qty !== '' && x.qty != null ? x.qty + (x.unit ? ' ' + x.unit : '') : '', x.detail].filter(v => v !== '' && v != null).join(' ').trim();
  function normFood(f) {
    if (!Array.isArray(f.items)) f.items = (f.feat || []).map(t => it(t));
    if (!Array.isArray(f.photos)) f.photos = f.img ? [f.img] : [];
    f.tags = f.tags || [];
    delete f.feat; delete f.img; delete f.gallery;
    return f;
  }
  FOOD.forEach(normFood);
  const FOOD_DEFAULT = JSON.parse(JSON.stringify(FOOD));
  const foodFeat = f => (f.items || []).map(itemText).filter(Boolean);
  // carousel slides: main photos first, then a slide per item (packages also show placeholders for items without a photo)
  const foodSlides = f => {
    const out = (f.photos || []).map(img => ({ img, cap: f.name }));
    for (const x of f.items || []) if (x.img || f.cat === 'package') out.push({ img: x.img || null, cap: itemText(x) });
    return out.length ? out : [{ img: null, cap: f.name }];
  };
  // replace the menu list in place (every page holds a reference to FOOD)
  function applyCatalog(list) {
    let src = Array.isArray(list) && list.length ? list : FOOD_DEFAULT;
    // a saved menu from before the buffet/break sets existed gets them added once (after that, the saved menu decides)
    if (src !== FOOD_DEFAULT && !src.some(f => f && f.rc)) src = src.concat(FOOD_DEFAULT.filter(f => f.rc && !src.some(x => x && x.id === f.id)));
    FOOD.length = 0;
    for (const f of JSON.parse(JSON.stringify(src))) FOOD.push(normFood(f));
  }


  const meetItem = (roomId, days, layout, extra = {}) => {
    const r = MEETING_ROOMS.find(x => x.id === roomId);
    return Object.assign({ id: U.uid('i'), sec: 'meet', ref: roomId, name: r.name + ' — ' + (LAYOUTS.find(l => l.id === layout) || {}).name + (extra.pax ? ' ' + extra.pax + ' ท่าน' : ''),
      desc: extra.desc || 'เต็มวัน 08:00–17:00 · พักกลางวัน 12:00–13:00', price: extra.half ? r.half : r.full, q1: days, u1: 'วัน', q2: null, u2: '', layout,
      sched: extra.sched || [] }, extra.more || {});
  };
  const svcItem = (sid, days, sched) => {
    const s = SERVICES.find(x => x.id === sid);
    return { id: U.uid('i'), sec: 'meet', ref: 'svc:' + sid, name: s.name, desc: s.desc, price: s.price, q1: days, u1: s.unit, q2: null, u2: '', sched: sched || [] };
  };
  const foodItem = (fid, pax, times, desc, sched) => {
    const f = FOOD.find(x => x.id === fid) || FOOD_DEFAULT.find(x => x.id === fid) || { id: fid, cat: 'lunch', name: fid, price: 0, unit: 'ท่าน' };
    const per = /โต๊ะ|งาน/.test(f.unit);
    return { id: U.uid('i'), sec: 'food', ref: fid, name: (f.cat === 'coffee' ? 'Coffee break · ' : '') + f.name, desc: desc || f.place || '', price: f.price,
      q1: pax, u1: per ? f.unit : 'ท่าน', q2: per || times == null ? null : times, u2: per || times == null ? '' : (f.cat === 'coffee' ? 'รอบ' : 'มื้อ'), sched: sched || [] };
  };
  const roomItem = (name, rooms, nights, price, desc, ref, extra = {}) =>
    Object.assign({ id: U.uid('i'), sec: 'room', ref: ref || '', name, desc: desc || '', price, q1: rooms, u1: 'ห้อง', q2: nights, u2: 'คืน', sched: [] }, extra);

  function seedQuotes() {
    const mk = (o) => Object.assign({ discount: { type: 'pct', value: 0 }, history: [], views: 0, cc: [], deposit: '30% ภายใน 7 วันหลังยืนยัน', owner: 'u1', dietary: [] }, o);
    const sch = (day, start, end, place) => ({ day, start, end, place: place || '' });
    return [
      mk({ id: 'QT-2611-042', title: 'สัมมนาประจำปี 2026', client: 'บจก. สยามเทค', clientFull: 'บริษัท สยามเทค จำกัด', contact: 'คุณวรวุฒิ ศรีสุข', phone: '081-234-5678',
        email: 'worawut@siamtech.co.th', cc: ['hr@siamtech.co.th'], address: '', taxId: '', start: '2026-11-12', end: '2026-11-13', pax: 80, status: 'sent',
        validUntil: '2026-10-20', holdUntil: '2026-10-15', createdAt: '2026-10-02T09:12:00', updatedAt: '2026-10-05T10:42:00', views: 3, room: 'thikhampon',
        discount: { type: 'pct', value: 5 }, dietary: [{ k: 'ฮาลาล', n: 6 }, { k: 'มังสวิรัติ', n: 4 }, { k: 'แพ้อาหารทะเล', n: 2 }],
        history: [{ at: '2026-10-05T10:42:00', t: 'เวอร์ชัน 2 · ปรับส่วนลดเป็น 5%' }, { at: '2026-10-02T15:00:00', t: 'เวอร์ชัน 1 · ลูกค้าเปิดอ่าน 3 ครั้ง' }],
        items: [
          meetItem('thikhampon', 2, 'classroom', { pax: 80, desc: 'เต็มวัน 08:00–17:00 · ชั้น 2', sched: [sch(1, '08:30', '17:00', 'Thikhampon'), sch(2, '08:30', '16:00', 'Thikhampon')] }),
          svcItem('av', 2, []),
          foodItem('cb-morning', 80, 4, 'เช้า–บ่าย วันละ 2 รอบ', [sch(1, '10:00', '10:30'), sch(1, '14:30', '15:00'), sch(2, '10:00', '10:30'), sch(2, '14:30', '15:00')]),
          foodItem('ln-buffet', 80, 2, 'อาหารกลางวัน · Swan Café', [sch(1, '12:00', '13:00'), sch(2, '12:00', '13:00')]),
          foodItem('gd-khantoke', 80, 1, 'คืนวันที่ 1 · ลานสระว่ายน้ำ', [sch(1, '18:30', '21:30', 'ลานสระว่ายน้ำ')]),
          roomItem('Grand Deluxe + Superior (เตียงแฝด) รวมอาหารเช้า', 40, 1, 1290, 'เช็กอิน 12 พ.ย. 14:00 → เช็กเอาต์ 13 พ.ย. 12:00', 'block', { block: { gdx: 25, sup: 15 } }),
          roomItem('ห้องเจ้าหน้าที่ พักห้องละ 2 ท่าน รวมอาหารเช้า', 2, 1, 1290, 'ห้องฟรีสำหรับผู้จัด ทุก 20 ห้อง ฟรี 1 ห้อง', 'comp', { free: true })
        ] }),
      mk({ id: 'QT-2611-044', title: 'ประชุมวิชาการ', client: 'สมาคมแพทย์ภาคใต้', clientFull: 'สมาคมแพทย์ภาคใต้', contact: 'นพ.ธนกร วิเศษ', phone: '089-111-2233', email: 'secretary@southmed.or.th',
        start: '2026-11-20', end: '2026-11-22', pax: 220, status: 'draft', validUntil: '2026-10-30', createdAt: '2026-10-02T11:00:00', updatedAt: '2026-10-02T16:20:00', room: 'kinnaree',
        items: [
          meetItem('kinnaree', 3, 'theatre', { pax: 220, sched: [sch(1, '08:30', '17:00', 'Kinnaree'), sch(2, '08:30', '17:00', 'Kinnaree'), sch(3, '08:30', '12:00', 'Kinnaree')] }),
          svcItem('av', 3), svcItem('led', 3),
          foodItem('cb-thaisweet', 220, 5, 'วันละ 2 รอบ (วันสุดท้าย 1 รอบ)'),
          foodItem('ln-buffet', 220, 3, 'อาหารกลางวัน · Swan Café'),
          foodItem('gd-intl', 220, 1, 'Welcome dinner คืนแรก'),
          roomItem('Deluxe เตียงแฝด รวมอาหารเช้า', 34, 2, 1390, 'เช็กอิน 20 พ.ย. → เช็กเอาต์ 22 พ.ย.', 'dlx'),
          roomItem('Grand Deluxe เตียงแฝด รวมอาหารเช้า', 30, 2, 1290, '', 'gdx'),
          roomItem('Superior เตียงแฝด รวมอาหารเช้า', 20, 2, 1290, '', 'sup')
        ] }),
      mk({ id: 'QT-2611-039', title: 'อบรมผู้จัดการสาขา', client: 'ธนาคารกรุงไทย ภาค 4', clientFull: 'ธนาคารกรุงไทย จำกัด (มหาชน) สำนักงานภาค 4', contact: 'คุณศิริพร ใจงาม', phone: '086-555-1200', email: 'siriporn.j@ktb.co.th',
        start: '2026-11-03', end: '2026-11-03', pax: 45, status: 'tentative', validUntil: '2026-10-18', holdUntil: '2026-10-16', createdAt: '2026-09-28T10:00:00', updatedAt: '2026-10-04T14:10:00', room: 'ruenhong',
        followUp: 'ขอปรับเมนูกลางวัน ต้องการเมนูฮาลาลเพิ่ม 10 ที่',
        items: [meetItem('ruenhong', 1, 'classroom', { pax: 45, sched: [sch(1, '08:30', '16:30', 'Ruen Hong')] }), svcItem('av', 1),
          foodItem('cb-morning', 45, 2, 'เช้า–บ่าย'), foodItem('ln-thai7', 45, 1, 'อาหารกลางวัน'),
          roomItem('Deluxe เตียงแฝด รวมอาหารเช้า', 23, 1, 1390, 'คืนก่อนวันอบรม', 'dlx')] }),
      mk({ id: 'QT-2610-031', title: 'งานเลี้ยงฉลองหมั้น', client: 'คุณพิมพ์ชนก & คุณธีร์', clientFull: 'คุณพิมพ์ชนก แก้วมณี', contact: 'คุณพิมพ์ชนก แก้วมณี', phone: '092-345-6677', email: 'pimchanok.k@gmail.com',
        start: '2026-10-24', end: '2026-10-24', pax: 150, status: 'confirmed', validUntil: '2026-10-01', createdAt: '2026-09-10T10:00:00', updatedAt: '2026-09-20T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'banquet', { pax: 150, desc: 'ครึ่งวันเย็น 17:00–22:00', half: true, sched: [sch(1, '17:00', '22:00', 'Kinnaree')] }),
          foodItem('wd-gold', 150, null, 'บุฟเฟต์ 14 รายการ'), svcItem('stage', 1), svcItem('photo', 1),
          roomItem('Junior Suite (ห้องเจ้าสาว)', 1, 1, 2490, '', 'jrs'), roomItem('Grand Deluxe รวมอาหารเช้า', 20, 1, 1290, '', 'gdx')] }),
      mk({ id: 'QT-2610-028', title: 'Town hall Q4', client: 'บมจ. ไทยออยล์', clientFull: 'บริษัท ไทยออยล์ จำกัด (มหาชน)', contact: 'คุณภาณุ ทองดี', phone: '081-900-1122', email: 'panu.t@thaioil.co.th',
        start: '2026-10-18', end: '2026-10-18', pax: 120, status: 'confirmed', validUntil: '2026-09-30', createdAt: '2026-09-05T10:00:00', updatedAt: '2026-09-25T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'theatre', { pax: 120, sched: [sch(1, '08:30', '17:00', 'Kinnaree')] }), svcItem('av', 1), svcItem('led', 1),
          foodItem('cb-premium', 120, 2), foodItem('ln-intl', 120, 1), foodItem('ck-sunset', 120, null, 'ค็อกเทลปิดงาน'),
          roomItem('Grand Deluxe รวมอาหารเช้า', 30, 1, 1290, '', 'gdx'), roomItem('Deluxe รวมอาหารเช้า', 15, 1, 1390, '', 'dlx')] }),
      // calendar fillers
      mk({ id: 'QT-2610-024', title: 'งานแต่ง คุณเมย์', client: 'คุณเมย์ & คุณบอส', contact: 'คุณเมย์ สุดสวย', phone: '080-111-2222', email: 'may.wedding@gmail.com', start: '2026-11-06', end: '2026-11-06', pax: 250, status: 'confirmed',
        validUntil: '2026-09-30', createdAt: '2026-08-20T10:00:00', updatedAt: '2026-09-15T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'banquet', { pax: 250, sched: [sch(1, '17:00', '22:00', 'Kinnaree')] }), foodItem('wd-chinese', 25, null, 'โต๊ะจีน 25 โต๊ะ'), svcItem('stage', 1), roomItem('Grand Deluxe', 15, 1, 1290, '', 'gdx')] }),
      mk({ id: 'QT-2611-033', title: 'สัมมนา SCG', client: 'SCG ภาคเหนือ', contact: 'คุณอรุณ ศักดิ์ดี', phone: '081-777-8899', email: 'arun@scg.com', start: '2026-11-10', end: '2026-11-10', pax: 180, status: 'sent',
        validUntil: '2026-10-25', createdAt: '2026-10-01T10:00:00', updatedAt: '2026-10-03T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'classroom', { pax: 180, sched: [sch(1, '08:30', '17:00', 'Kinnaree')] }), svcItem('av', 1), foodItem('cb-morning', 180, 2), foodItem('ln-buffet', 180, 1)] }),
      mk({ id: 'QT-2611-040', title: 'Gala dinner', client: 'บจก. นอร์ทเทิร์น ทราเวล', contact: 'คุณปวีณา', phone: '083-222-1100', email: 'paweena@northtravel.co.th', start: '2026-11-12', end: '2026-11-12', pax: 120, status: 'sent',
        validUntil: '2026-10-25', createdAt: '2026-10-03T10:00:00', updatedAt: '2026-10-03T10:00:00', room: 'pool',
        items: [meetItem('pool', 1, 'banquet', { pax: 120, desc: 'Terrace ริมสระ 18:00–22:00', half: true, sched: [sch(1, '18:00', '22:00', 'ลานสระว่ายน้ำ')] }), foodItem('gd-bbq', 120, 1)] }),
      mk({ id: 'QT-2611-041', title: 'สัมมนาตัวแทนจำหน่าย', client: 'บจก. ล้านนาฟาร์ม', contact: 'คุณสมศักดิ์', phone: '084-555-6677', email: 'somsak@lannafarm.co.th', start: '2026-11-13', end: '2026-11-13', pax: 200, status: 'tentative',
        validUntil: '2026-10-22', holdUntil: '2026-10-20', createdAt: '2026-10-04T10:00:00', updatedAt: '2026-10-04T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'theatre', { pax: 200, sched: [sch(1, '08:30', '17:00', 'Kinnaree')] }), foodItem('cb-basic', 200, 2), foodItem('ln-thai7', 200, 1)] }),
      mk({ id: 'QT-2610-027', title: 'งานแต่ง คุณปาล์ม', client: 'คุณปาล์ม & คุณแนน', contact: 'คุณปาล์ม', phone: '085-121-3434', email: 'palm.nan@gmail.com', start: '2026-11-14', end: '2026-11-14', pax: 180, status: 'confirmed',
        validUntil: '2026-09-30', createdAt: '2026-08-28T10:00:00', updatedAt: '2026-09-28T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'banquet', { pax: 180, sched: [sch(1, '17:00', '22:00', 'Kinnaree')] }), foodItem('wd-silver', 180, null), svcItem('stage', 1)] }),
      mk({ id: 'QT-2611-045', title: 'Workshop', client: 'มหาวิทยาลัยแม่ฟ้าหลวง', contact: 'ผศ.ดร.วิชัย', phone: '053-916-000', email: 'wichai@mfu.ac.th', start: '2026-11-17', end: '2026-11-17', pax: 18, status: 'draft',
        validUntil: '2026-11-01', createdAt: '2026-10-05T10:00:00', updatedAt: '2026-10-05T10:00:00', room: 'prayamengrai',
        items: [meetItem('prayamengrai', 1, 'ushape', { pax: 18, sched: [sch(1, '09:00', '16:00', 'Praya Meng Rai')] }), foodItem('cb-thaisweet', 18, 2), foodItem('ln-bento', 18, 1)] }),
      mk({ id: 'QT-2611-030', title: 'ประชุม อบจ.', client: 'องค์การบริหารส่วนจังหวัดเชียงราย', contact: 'คุณสุรเชษฐ์', phone: '053-711-111', email: 'admin@chiangraipao.go.th', start: '2026-11-19', end: '2026-11-19', pax: 40, status: 'confirmed',
        validUntil: '2026-10-10', createdAt: '2026-09-20T10:00:00', updatedAt: '2026-10-01T10:00:00', room: 'ruenhong',
        items: [meetItem('ruenhong', 1, 'ushape', { pax: 40, sched: [sch(1, '09:00', '16:00', 'Ruen Hong')] }), foodItem('cb-morning', 40, 2), foodItem('ln-thai7', 40, 1)] }),
      mk({ id: 'QT-2612-002', title: 'Year-end party', client: 'ธนาคารกสิกรไทย', contact: 'คุณชลธิชา', phone: '086-999-0011', email: 'chonticha@kbank.com', start: '2026-11-25', end: '2026-11-25', pax: 300, status: 'tentative',
        validUntil: '2026-10-28', holdUntil: '2026-10-24', createdAt: '2026-10-05T10:00:00', updatedAt: '2026-10-05T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'banquet', { pax: 300, sched: [sch(1, '18:00', '22:00', 'Kinnaree')] }), foodItem('gd-intl', 300, 1), svcItem('led', 1), svcItem('stage', 1)] }),
      mk({ id: 'QT-2611-036', title: 'Gala ไทยออยล์', client: 'บมจ. ไทยออยล์', contact: 'คุณภาณุ ทองดี', phone: '081-900-1122', email: 'panu.t@thaioil.co.th', start: '2026-11-27', end: '2026-11-27', pax: 150, status: 'confirmed',
        validUntil: '2026-10-15', createdAt: '2026-09-30T10:00:00', updatedAt: '2026-10-04T10:00:00', room: 'pool',
        items: [meetItem('pool', 1, 'banquet', { pax: 150, sched: [sch(1, '18:30', '22:00', 'ลานสระว่ายน้ำ')] }), foodItem('gd-khantoke', 150, 1)] }),
      mk({ id: 'QT-2611-037', title: 'เลี้ยงเกษียณ กฟภ.', client: 'การไฟฟ้าส่วนภูมิภาค เขต 1', contact: 'คุณมานพ', phone: '053-222-333', email: 'manop@pea.co.th', start: '2026-11-27', end: '2026-11-27', pax: 90, status: 'confirmed',
        validUntil: '2026-10-15', createdAt: '2026-09-30T10:00:00', updatedAt: '2026-10-04T10:00:00', room: 'thikhampon',
        items: [meetItem('thikhampon', 1, 'banquet', { pax: 90, half: true, sched: [sch(1, '17:30', '21:30', 'Thikhampon')] }), foodItem('gd-thai', 90, 1)] }),
      mk({ id: 'QT-2610-029', title: 'งานแต่ง คุณฝน', client: 'คุณฝน & คุณต้น', contact: 'คุณฝน', phone: '087-121-9090', email: 'fon.ton@gmail.com', start: '2026-11-28', end: '2026-11-28', pax: 200, status: 'confirmed',
        validUntil: '2026-09-30', createdAt: '2026-09-01T10:00:00', updatedAt: '2026-09-29T10:00:00', room: 'kinnaree',
        items: [meetItem('kinnaree', 1, 'banquet', { pax: 200, sched: [sch(1, '17:00', '22:00', 'Kinnaree')] }), foodItem('wd-gold', 200, null), svcItem('stage', 1)] }),
      mk({ id: 'QT-2611-046', title: 'ประชุมวิชาการ (ต่อเนื่อง)', client: 'สมาคมแพทย์ภาคใต้', contact: 'นพ.ธนกร วิเศษ', phone: '089-111-2233', email: 'secretary@southmed.or.th', start: '2026-10-08', end: '2026-10-08', pax: 60, status: 'confirmed',
        validUntil: '2026-09-30', createdAt: '2026-09-10T10:00:00', updatedAt: '2026-09-29T10:00:00', room: 'thikhampon',
        items: [meetItem('thikhampon', 1, 'classroom', { pax: 60, sched: [sch(1, '08:30', '17:00', 'Thikhampon')] }), foodItem('cb-morning', 60, 2), foodItem('ln-buffet', 60, 1), roomItem('Deluxe', 30, 2, 1390, '', 'dlx')] })
    ];
  }

  const FIRST = ['สมชาย', 'วิไล', 'ประยุทธ์', 'นงลักษณ์', 'อนุชา', 'กมลชนก', 'ธนพล', 'มาลี', 'อรุณี', 'วิภา', 'สุนทร', 'จันทร์เพ็ญ', 'บุญมี', 'ศิริพร', 'อำนาจ', 'พรทิพย์', 'ชาตรี', 'ดวงใจ', 'ณัฐวุฒิ', 'สุดารัตน์', 'เกียรติ', 'รัตนา', 'วีระ', 'ปราณี', 'สมศรี', 'ยุทธนา', 'อัญชลี', 'สมพร', 'ไพโรจน์', 'กาญจนา', 'ทองใบ', 'จิราพร', 'ศักดิ์ชัย', 'นิตยา', 'ประเสริฐ', 'อุไร', 'ธีระ', 'บุษบา', 'สุรชัย', 'ลำดวน'];
  const LAST = ['กิตติ', 'สุขใจ', 'ทองดี', 'พรหมมา', 'ภาศสุวรรณ', 'แก้วกาญจน์', 'ใจกล้า', 'ศรีวงศ์', 'บุญมา', 'อินทร์แก้ว', 'คำแสน', 'ปัญญา', 'วงศ์ไชย', 'ชัยมงคล', 'สายทอง', 'มณีวงศ์', 'ดวงดี', 'ศรีสุข', 'แสงคำ', 'ยอดยิ่ง'];

  function seedEmployees() {
    const plan = { kitchen: 14, banquet: 10, av: 4, front: 8, hk: 9 };
    const roleOf = { kitchen: ['เชฟ', 'ผู้ช่วยเชฟ', 'พนักงานครัว'], banquet: ['หัวหน้าจัดเลี้ยง', 'พนักงานจัดเลี้ยง'], av: ['หัวหน้าช่าง', 'ช่างเทคนิค'], front: ['หัวหน้าส่วนหน้า', 'พนักงานต้อนรับ'], hk: ['หัวหน้าแม่บ้าน', 'แม่บ้าน'] };
    const out = []; let i = 0;
    const fixed = [['สมชาย', 'กิตติ', 'kitchen', 'เชฟ', 38], ['วิไล', 'สุขใจ', 'kitchen', 'ผู้ช่วยเชฟ', 34], ['ประยุทธ์', 'ทองดี', 'kitchen', 'พนักงานครัว', 40], ['นงลักษณ์', 'พรหมมา', 'kitchen', 'พนักงานครัว', 44], ['อนุชา', 'ภาศสุวรรณ', 'kitchen', 'พนักงานครัว', 48],
      ['ธนพล', 'ใจกล้า', 'av', 'หัวหน้าช่าง', 32], ['กมลชนก', 'แก้วกาญจน์', 'front', 'หัวหน้าส่วนหน้า', 40], ['มาลี', 'ศรีวงศ์', 'hk', 'หัวหน้าแม่บ้าน', 42], ['อรุณี', 'บุญมา', 'banquet', 'หัวหน้าจัดเลี้ยง', 40]];
    for (const f of fixed) { out.push({ id: 'e' + (++i), name: f[0] + ' ' + f[1], dept: f[2], role: f[3], hours: f[4] }); plan[f[2]]--; }
    let k = 9;
    for (const [dept, n] of Object.entries(plan)) for (let j = 0; j < n; j++) {
      const fn = FIRST[k % FIRST.length], ln = LAST[(k * 7) % LAST.length]; k++;
      out.push({ id: 'e' + (++i), name: fn + ' ' + ln, dept, role: roleOf[dept][roleOf[dept].length - 1], hours: 24 + ((k * 5) % 21) });
    }
    return out;
  }

  function seedStaffPlan(emps) {
    const by = d => emps.filter(e => e.dept === d).map(e => e.id);
    const K = by('kitchen'), B = by('banquet'), A = by('av'), F = by('front'), H = by('hk');
    const day = (sh) => sh;
    const p12 = {
      kitchen: [{ id: 's1', type: 'morning', label: 'รอบเช้า', start: '06:00', end: '14:00', need: 4, staff: K.slice(5, 9) },
        { id: 's2', type: 'evening', label: 'รอบบ่าย', start: '14:00', end: '22:00', need: 6, staff: K.slice(0, 4) }],
      banquet: [{ id: 's3', type: 'morning', label: 'รอบเช้า', start: '06:00', end: '14:00', need: 4, staff: B.slice(0, 4) },
        { id: 's4', type: 'evening', label: 'รอบบ่าย', start: '14:00', end: '22:00', need: 4, staff: B.slice(4, 8) }],
      av: [{ id: 's5', type: 'morning', label: 'รอบเต็มวัน', start: '07:00', end: '17:00', need: 2, staff: A.slice(0, 2) },
        { id: 's6', type: 'evening', label: 'รอบค่ำ', start: '17:00', end: '22:00', need: 1, staff: [] }],
      front: [{ id: 's7', type: 'morning', label: 'รอบเช้า', start: '06:00', end: '14:00', need: 3, staff: F.slice(0, 3) },
        { id: 's8', type: 'evening', label: 'รอบบ่าย', start: '14:00', end: '22:00', need: 3, staff: F.slice(3, 6) }],
      hk: [{ id: 's9', type: 'morning', label: 'รอบกลางวัน', start: '08:00', end: '16:00', need: 7, staff: H.slice(0, 7) }]
    };
    const p13 = JSON.parse(JSON.stringify(p12));
    p13.kitchen[1].need = 4; p13.kitchen[1].end = '18:00';
    p13.av[1].need = 0;
    p13.banquet[1].end = '18:00';
    return {
      'QT-2611-042': {
        days: { '2026-11-12': day(p12), '2026-11-13': day(p13) },
        ot: [
          { id: 'ot1', date: '2026-11-12', dept: 'kitchen', shiftId: 's2', hours: 2, reason: 'เก็บงานเลี้ยงขันโตก + ล้างครัว', staff: K.slice(0, 4), status: 'draft' },
          { id: 'ot2', date: '2026-11-12', dept: 'banquet', shiftId: 's4', hours: 2, reason: 'เก็บสถานที่งานเลี้ยง', staff: B.slice(4, 5), status: 'approved' }
        ]
      }
    };
  }

  function seedBEO() {
    const t = (time, task, place, dept) => ({ time, task, place, dept });
    return [
      { id: 'FS-2611-018', quoteId: 'QT-2611-042', rev: 2, revNote: 'เปลี่ยนจำนวนเป็น 80', status: 'sent', confirmPax: '80 ท่าน (+5% สำรอง)', payment: 'มัดจำ 30% แล้ว · วางบิล',
        schedule: {
          1: [t('06:00', 'จัดห้อง Classroom 81 ที่ + โต๊ะลงทะเบียนหน้าห้อง', 'Thikhampon', 'banquet'), t('07:00', 'ทดสอบเสียง ไมค์ไร้สาย 4 ตัว โปรเจกเตอร์คู่', 'Thikhampon', 'av'),
            t('07:30', 'เปิดลงทะเบียน · ป้ายบอกทางจากล็อบบี้', 'หน้าห้อง', 'banquet'), t('10:00', 'Coffee break · ชา กาแฟ ขนมไทย 80 ที่', 'หน้าห้อง', 'kitchen'),
            t('12:00', 'อาหารกลางวัน Thai Lunch 7 รายการ (กันโต๊ะ 10 ตัว)', 'ห้องอาหารราชพฤกษ์', 'kitchen'), t('14:00', 'เช็กอินกรุ๊ป 40 ห้อง · แจกกุญแจที่โต๊ะลงทะเบียน', 'ล็อบบี้', 'front'),
            t('14:30', 'Coffee break รอบบ่าย', 'หน้าห้อง', 'kitchen'), t('18:30', 'ขันโตกล้านนา + การแสดงฟ้อนเล็บ', 'ลานสระว่ายน้ำ', 'kitchen')],
          2: [t('07:00', 'จัดห้องคืนสภาพ Classroom', 'Thikhampon', 'banquet'), t('08:00', 'ทดสอบเสียงและภาพ', 'Thikhampon', 'av'), t('10:00', 'Coffee break', 'หน้าห้อง', 'kitchen'),
            t('12:00', 'อาหารกลางวัน Thai–International Buffet', 'Swan Café', 'kitchen'), t('12:00', 'เช็กเอาต์กรุ๊ป · เก็บกุญแจ', 'ล็อบบี้', 'front'), t('14:30', 'Coffee break รอบบ่าย', 'หน้าห้อง', 'kitchen'), t('16:00', 'เก็บห้องประชุม', 'Thikhampon', 'hk')]
        },
        depts: {
          kitchen: { items: ['Coffee break 4 รอบ × 80 ที่', 'Thai Lunch 2 มื้อ · ขันโตก 1 มื้อ', 'ฮาลาล 6 · มังสวิรัติ 4 · แพ้อาหารทะเล 2'], owner: 'เชฟสมชาย', ack: true },
          banquet: { items: ['Classroom 81 ที่ ผ้าคลุมโต๊ะสีครีม', 'โต๊ะลงทะเบียน 2 ตัว + ป้ายต้อนรับ', 'น้ำดื่ม กระดาษ ปากกา ทุกที่นั่ง'], owner: 'คุณอรุณี', ack: true },
          av: { items: ['โปรเจกเตอร์คู่ + จอ 2 ด้าน', 'ไมค์ไร้สาย 4 · ไมค์โพเดียม 1', 'ช่างประจำห้อง 08:00–17:00'], owner: 'คุณธนพล', ack: false },
          front: { items: ['Room block Grand Deluxe 25 + Superior 15', 'รายชื่อผู้เข้าพักภายใน 29 ต.ค.', 'ที่จอดรถ VIP 3 คัน'], owner: 'คุณกมลชนก', ack: true },
          hk: { items: ['เตรียม 40 ห้องก่อน 13:00', 'ของต้อนรับในห้อง: น้ำ + ผลไม้', 'ทำความสะอาดห้องประชุมหลังเลิก'], owner: 'คุณมาลี', ack: false },
          acc: { items: ['มัดจำ 30% รับแล้ว 117,137 ฿', 'ยอดคงเหลือ วางบิล 30 วัน', 'ใบกำกับภาษีในนาม บจก. สยามเทค'], owner: 'คุณวิภา', ack: true }
        },
        note: 'ประธานบริษัทมาถึง 08:00 ต้องการที่จอดรถ VIP 3 คัน · ฮาลาล 6 / มังสวิรัติ 4 / แพ้อาหารทะเล 2 (ติดป้ายชื่อที่โต๊ะ) · ยืนยันจำนวนสุดท้ายภายใน 9 พ.ย. 12:00' },
      { id: 'FS-2611-015', quoteId: 'QT-2610-024', rev: 1, status: 'sent', confirmPax: '250 ท่าน', payment: 'ชำระครบแล้ว', schedule: { 1: [t('13:00', 'จัดโต๊ะจีน 25 โต๊ะ + เวที', 'Kinnaree', 'banquet'), t('16:00', 'ทดสอบเสียง', 'Kinnaree', 'av'), t('18:00', 'เสิร์ฟโต๊ะจีน', 'Kinnaree', 'kitchen')] },
        depts: { kitchen: { items: ['โต๊ะจีน 9 รายการ × 25 โต๊ะ'], owner: 'เชฟสมชาย', ack: true }, banquet: { items: ['โต๊ะจีน 25 โต๊ะ ผ้าสีทอง'], owner: 'คุณอรุณี', ack: true }, av: { items: ['ไมค์ 2 ตัว + จอ LED'], owner: 'คุณธนพล', ack: true }, front: { items: ['ห้องพัก 15 ห้อง'], owner: 'คุณกมลชนก', ack: true }, hk: { items: ['ห้องเจ้าสาว'], owner: 'คุณมาลี', ack: true }, acc: { items: ['ชำระครบ'], owner: 'คุณวิภา', ack: true } }, note: '' },
      { id: 'FS-2611-012', quoteId: 'QT-2611-039', rev: 1, status: 'sent', confirmPax: '45 ท่าน', payment: 'มัดจำ 30%', schedule: { 1: [t('07:30', 'จัดห้อง Classroom 45 ที่', 'Ruen Hong', 'banquet'), t('10:00', 'Coffee break', 'หน้าห้อง', 'kitchen'), t('12:00', 'Thai Lunch', 'ห้องอาหาร', 'kitchen')] },
        depts: { kitchen: { items: ['Coffee break 2 รอบ', 'Thai Lunch 45 ที่'], owner: 'เชฟสมชาย', ack: true }, banquet: { items: ['Classroom 45 ที่'], owner: 'คุณอรุณี', ack: true }, av: { items: ['AV Package'], owner: 'คุณธนพล', ack: true }, front: { items: ['Room 23 ห้อง'], owner: 'คุณกมลชนก', ack: true }, hk: { items: ['เตรียมห้องพัก'], owner: 'คุณมาลี', ack: true }, acc: { items: ['มัดจำ 30%'], owner: 'คุณวิภา', ack: true } }, note: '' },
      { id: 'FS-2611-020', quoteId: 'QT-2610-027', rev: 0, status: 'draft', confirmPax: '180 ท่าน', payment: 'มัดจำ 50%', schedule: { 1: [] }, depts: {}, note: '' },
      { id: 'FS-2611-021', quoteId: 'QT-2611-030', rev: 0, status: 'draft', confirmPax: '40 ท่าน', payment: 'วางบิล', schedule: { 1: [] }, depts: {}, note: '' }
    ];
  }

  function seedUsers() {
    const P = r => Object.fromEntries(PERMS.map(p => [p.id, ROLE_PERMS[r].includes(p.id)]));
    const base = [
      ['u1', 'นภัสสร ใจดี', 'napassorn', 'sales', 'ฝ่ายขายและการตลาด', 'active', '085-029-9944'],
      ['u2', 'ปิยะพงษ์ ศรีวงศ์', 'piyapong', 'manager', 'ฝ่ายขายและการตลาด', 'active', '081-555-0101'],
      ['u3', 'ลักษิกา เตชะธีราวัฒน์', 'laksika', 'admin', 'ผู้บริหาร', 'active', '081-222-3333'],
      ['u4', 'สมชาย กิตติ', 'somchai', 'head', 'ครัว / F&B', 'active', '089-100-2000'],
      ['u5', 'วิภา รัตนากร', 'wipa', 'acc', 'บัญชี', 'active', '088-200-3000'],
      ['u6', 'ชนิดา มั่นคง', 'chanida', 'sales', 'ฝ่ายขายและการตลาด', 'invited', '087-300-4000'],
      ['u7', 'ธีรพงษ์ แสงทอง', 'teerapong', 'sales', 'ฝ่ายขายและการตลาด', 'disabled', '086-400-5000'],
      ['u8', 'กิตติพัฒน์ บุญเรือง', 'kittipat', 'sales', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u9', 'อรวรรณ ศรีสวัสดิ์', 'orawan', 'sales', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u10', 'ณัฐพล ทองคำ', 'nattapon', 'sales', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u11', 'สุภาวดี แก้วใส', 'supawadee', 'sales', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u12', 'ภัทรา จันทร์หอม', 'pattra', 'sales', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u13', 'ศุภชัย วงศ์ใหญ่', 'supachai', 'sales', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u14', 'จิรายุ พรหมสุข', 'jirayu', 'sales', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u15', 'พิชญา มหาวงศ์', 'pitchaya', 'sales', 'ฝ่ายขายและการตลาด', 'invited', ''],
      ['u16', 'อนันต์ ศรีเมือง', 'anan', 'manager', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u17', 'วรรณา ดีงาม', 'wanna', 'manager', 'ฝ่ายขายและการตลาด', 'active', ''],
      ['u18', 'อรุณี บุญมา', 'arunee', 'head', 'จัดเลี้ยง', 'active', ''],
      ['u19', 'ธนพล ใจกล้า', 'tanapol', 'head', 'ช่าง / AV', 'active', ''],
      ['u20', 'กมลชนก แก้วกาญจน์', 'kamonchanok', 'head', 'ส่วนหน้า', 'active', ''],
      ['u21', 'มาลี ศรีวงศ์', 'malee', 'head', 'แม่บ้าน', 'active', ''],
      ['u22', 'สุนทร คำแสน', 'soontorn', 'head', 'ครัว / F&B', 'active', ''],
      ['u23', 'ปราณี อินทร์แก้ว', 'pranee', 'acc', 'บัญชี', 'active', ''],
      ['u24', 'เกียรติศักดิ์ ปัญญา', 'kiattisak', 'sales', 'ฝ่ายขายและการตลาด', 'active', '']
    ];
    const last = ['2026-10-05T10:42:00', '2026-10-05T09:15:00', '2026-10-04T17:30:00', '2026-10-04T16:20:00', '2026-10-04T08:00:00', null, '2026-09-12T10:00:00'];
    return base.map((b, i) => ({ id: b[0], name: b[1], email: b[2] + '@luckswangroup.com', role: b[3], dept: b[4], status: b[5], phone: b[6],
      lastLogin: i < 7 ? last[i] : '2026-10-0' + (1 + (i % 4)) + 'T09:00:00', perms: P(b[3]), twoFA: i % 3 === 0, pw: null }));
  }

  function seed() {
    const employees = seedEmployees();
    return {
      v: 1,
      settings: { service: 0.10, vat: 0.07, target: 3000000, company: 'บริษัท ลักษวรรณ จำกัด', address: '499 หมู่ 4 ต.ริมกก อ.เมือง จ.เชียงราย 57000',
        phone: '(+66)53-750-855', hq: '(+66)86-756-9933', bank: 'บัญชี บจก.ลักษวรรณ หมายเลขบัญชี 045-8-38475-4 ธนาคารกสิกรไทย', email: 'Sale.marketing@luckswangroup.com', line: '@luckswanresort',
        gm: 'น.ส. ลักษิกา เตชะธีราวัฒน์' },
      users: seedUsers(),
      quotes: seedQuotes(),
      beos: seedBEO(),
      employees,
      staff: seedStaffPlan(employees),
      tasks: [
        { id: 'TK-1', title: 'โทรยืนยันจำนวนผู้เข้าร่วม SCG ภาคเหนือ', detail: 'ลูกค้าแจ้งว่าอาจเพิ่มเป็น 200 ท่าน ให้สอบถามและอัปเดตใบเสนอราคา', quoteId: 'QT-2611-033', due: '2026-10-20', priority: 'high',
          to: { type: 'user', id: 'u1', name: 'นภัสสร ใจดี' }, by: 'u3', byName: 'ลักษิกา เตชะธีราวัฒน์', at: '2026-10-14T09:10:00', status: 'open' },
        { id: 'TK-2', title: 'เตรียมเมนูชิมอาหารงานแต่งคุณเมย์ & คุณบอส', detail: 'ชิมอาหาร 6 ท่าน วันเสาร์ 24 ต.ค. 11:00', quoteId: 'QT-2610-024', due: '2026-10-24', priority: 'normal',
          to: { type: 'dept', dept: 'kitchen', name: 'ครัว / อาหารและเครื่องดื่ม', emps: ['e1', 'e2'] }, by: 'u3', byName: 'ลักษิกา เตชะธีราวัฒน์', at: '2026-10-13T15:30:00', status: 'done',
          doneBy: 'u4', doneByName: 'สมชาย กิตติ', doneAt: '2026-10-14T08:20:00', doneNote: 'จัดเตรียมแล้ว' }
      ],
      seq: { QT: 47, FS: 22 }
    };
  }

  /* ---------------- store ---------------- */
  const KEY = 'lsSales.v1';
  const Store = {
    db: null,
    KEY,
    load() {
      try { const raw = localStorage.getItem(this.KEY); if (raw) { const d = JSON.parse(raw); if (d && d.v === 1) { this.db = d; this.loaded(); return this.db; } } } catch (e) { /* ignore */ }
      this.db = seed(); this.write(); this.loaded(); return this.db;
    },
    loaded() { if (this.db && !Array.isArray(this.db.tasks)) this.db.tasks = []; if (this.db && !Array.isArray(this.db.leaves)) this.db.leaves = []; if (this.db && !Array.isArray(this.db.rosters)) this.db.rosters = []; applyCatalog(this.db && this.db.catalog && this.db.catalog.food); applyRoomRates(this.db && this.db.settings && this.db.settings.roomRates); applyMeetRooms(this.db && this.db.settings && this.db.settings.meetRooms); applyDepts(this.db && this.db.settings && this.db.settings.depts); applyFoodCats(this.db && this.db.settings && this.db.settings.foodCats); if (g.API) g.API.rememberBeos(); },
    write() { try { localStorage.setItem(this.KEY, JSON.stringify(this.db)); } catch (e) { console.warn('save failed', e); } },
    save() { this.write(); },
    reset() { this.db = seed(); this.write(); this.loaded(); },
    seedData: seed,
    exportJSON() { return JSON.stringify(this.db, null, 2); },
    importJSON(text) { const d = JSON.parse(text); if (!d || d.v !== 1 || !Array.isArray(d.quotes)) throw new Error('ไฟล์ไม่ถูกต้อง'); this.db = d; this.write(); this.loaded(); },
    quote(id) { return this.db.quotes.find(q => q.id === id); },
    nextId(prefix) {
      const now = new Date(); const yy = String(now.getFullYear()).slice(-2); const mm = String(now.getMonth() + 1).padStart(2, '0');
      // continue after the highest number already used (also by other people when data is shared)
      const list = prefix === 'QT' ? this.db.quotes : prefix === 'FS' ? this.db.beos : [];
      const used = list.map(x => parseInt(String(x.id).split('-')[2], 10) || 0);
      this.db.seq[prefix] = Math.max((this.db.seq[prefix] || 0), ...used, 0) + 1;
      return prefix + '-' + yy + mm + '-' + String(this.db.seq[prefix]).padStart(3, '0');
    }
  };

  async function hash(pw) {
    try {
      const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('ls::' + pw));
      return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
    } catch (e) { return 'plain:' + pw; }
  }

  g.DATA = { FOOD_CATS_DEFAULT, applyFoodCats, LAYOUTS, MEETING_ROOMS, MEETING_ROOMS_DEFAULT, applyMeetRooms, ROOM_TYPES, ROOM_DEF, ROOM_RATES_DEFAULT, BOUTIQUE, BOUTIQUE_DEFAULT, applyRoomRates, seasonOf, roomPrice, roomStay, seasons: () => SEASONS, FOOD_CATS, FOOD, SERVICES, DEPTS, DEPTS_DEFAULT, applyDepts, ROLES, PERMS, ROLE_PERMS, STATUS, meetItem, svcItem, foodItem, roomItem, DEFAULT_PW: 'luckswan',
    itemText, foodFeat, foodSlides, normFood, applyCatalog, FOOD_DEFAULT };
  g.Store = Store;
  g.hashPw = hash;
})(window);
