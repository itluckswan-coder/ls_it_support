/* Luck Swan Sales — catering menu / package editor (add, edit, delete, prices, item list, photos) */
(function (g) {
  'use strict';
  const U = g.U, D = g.DATA, Store = g.Store, App = g.App, ICON = App.ICON;
  const H = U.esc;
  const UNITS = ['ท่าน', 'ท่าน/รอบ', 'ท่าน/มื้อ', 'โต๊ะ', 'งาน'];
  const ITEM_UNITS = ['มื้อ', 'รอบ', 'คืน', 'ห้อง', 'ชั่วโมง/วัน', 'ชั่วโมง', 'อย่าง', 'รายการ', 'ที่'];
  const TAGS = [['halal', 'ฮาลาล'], ['veg', 'มังสวิรัติ'], ['gf', 'ไม่มีกลูเตน']];

  /* ---------- images: shrink in the browser, then store (cloud: D1 via API · local: inside the menu data) ---------- */
  async function shrink(file, max = 1000, quality = 0.82) {
    if (!/^image\//.test(file.type)) throw new Error('ไฟล์ ' + file.name + ' ไม่ใช่รูปภาพ');
    const url = URL.createObjectURL(file);
    try {
      const img = await new Promise((ok, bad) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => bad(new Error('เปิดรูป ' + file.name + ' ไม่ได้')); i.src = url; });
      const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement('canvas');
      c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
      const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(img, 0, 0, c.width, c.height);
      return c.toDataURL('image/jpeg', quality);
    } finally { URL.revokeObjectURL(url); }
  }
  async function storeImage(file) {
    const cloud = g.API && g.API.online;
    const dataUrl = await shrink(file, cloud ? 1200 : 800, cloud ? 0.82 : 0.72);
    if (!cloud) return dataUrl;
    const j = await g.API.req('POST', 'images', { mime: 'image/jpeg', data: dataUrl.slice(dataUrl.indexOf(',') + 1) });
    return j.url;
  }
  const pickFiles = multiple => new Promise(res => {
    const inp = document.createElement('input'); inp.type = 'file'; inp.accept = 'image/*'; inp.multiple = !!multiple;
    inp.onchange = () => res(Array.from(inp.files || [])); inp.click();
  });

  /* ---------- save the whole menu list (synced to the shared database when online) ---------- */
  function saveCatalog() {
    Store.db.catalog = { food: JSON.parse(JSON.stringify(D.FOOD)) };
    Store.save();
  }

  /* ---------- editor ---------- */
  function openEditor(id) {
    const orig = id ? D.FOOD.find(f => f.id === id) : null;
    const draft = orig ? JSON.parse(JSON.stringify(orig)) : D.normFood({ id: '', cat: (App.state.ui.food && App.state.ui.food.cat !== 'all' ? App.state.ui.food.cat : 'package'), name: '', sub: '', price: 0, unit: 'ท่าน', min: '', place: '', tags: [], photos: [], items: [{ name: '', qty: '', unit: '', detail: '', img: null }] });
    let busy = 0, err = '';
    const bg = App.modal('<div id="ceBody"></div>');
    const modal = bg.querySelector('.modal');
    modal.style.maxWidth = '760px';

    const thumb = (src, act, i, label) => `<div class="ce-ph">${src ? `<img src="${H(src)}" alt="">` : '<span>📷</span>'}<button type="button" class="ce-ph-btn" data-ce="${act}" data-i="${i}" aria-label="${H(label)}"></button>${src ? `<button type="button" class="ce-x" data-ce="${act}Del" data-i="${i}" aria-label="ลบรูป">${ICON.x}</button>` : ''}</div>`;
    function draw() {
      const y = modal.scrollTop, pkg = draft.cat === 'package';
      modal.querySelector('#ceBody').innerHTML = `
        <h2>${orig ? 'แก้ไข ' + H(orig.name) : 'เพิ่มเมนู / แพ็กเกจ'}</h2>
        <div class="grid g2">
          <div><label class="f">ประเภท</label><select class="input" data-top="cat">${D.FOOD_CATS.filter(c => c.id !== 'all').map(c => `<option value="${c.id}" ${draft.cat === c.id ? 'selected' : ''}>${H(c.name)}</option>`).join('')}</select></div>
          <div><label class="f">ชื่อ${pkg ? 'แพ็กเกจ' : 'เมนู'} *</label><input class="input" data-top="name" value="${H(draft.name)}" placeholder="${pkg ? 'เช่น Standard Package' : 'เช่น Morning Bright'}"></div>
          <div><label class="f">ราคา (บาท) *</label><input class="input" data-top="price" inputmode="decimal" value="${H(draft.price)}"></div>
          <div><label class="f">ต่อ</label><input class="input" data-top="unit" list="ceUnits" value="${H(draft.unit)}"><datalist id="ceUnits">${UNITS.map(u => `<option>${u}</option>`).join('')}</datalist></div>
          <div><label class="f">คำอธิบายสั้น</label><input class="input" data-top="sub" value="${H(draft.sub || '')}" placeholder="${pkg ? 'เช่น 2 วัน 1 คืน' : 'เช่น 2 ชม.'}"></div>
          <div><label class="f">ขั้นต่ำ (ท่าน/วัน)</label><input class="input" data-top="min" inputmode="numeric" value="${H(draft.min || '')}" placeholder="ไม่บังคับ"></div>
          <div><label class="f">สถานที่ (ถ้ามี)</label><input class="input" data-top="place" value="${H(draft.place || '')}"></div>
          <div><label class="f">ป้ายและความเหมาะสม</label><div class="row wrap" style="gap:6px">
            <button type="button" class="chip ${draft.best ? 'on' : ''}" data-ce="best">ขายดี</button>
            ${['lunch', 'coffee'].includes(draft.cat) ? `<button type="button" class="chip ${draft.rc ? 'on' : ''}" data-ce="rc" title="ใช้หักค่าห้องประชุมในงานประชุม">หักค่าห้องประชุมได้</button>` : ''}
            ${TAGS.map(([k, l]) => `<button type="button" class="chip ${draft.tags.includes(k) ? 'on-teal' : ''}" data-ce="tag" data-v="${k}">${l}</button>`).join('')}</div></div>
        </div>

        <h3 style="margin:20px 0 8px">รูปหลัก</h3>
        <div class="ce-photos">${draft.photos.map((p, i) => thumb(p, 'photo', i, 'เปลี่ยนรูป')).join('')}<button type="button" class="ce-add-ph" data-ce="photoAdd">+ อัปโหลดรูป<small>เลือกได้หลายรูป</small></button></div>

        <div class="between" style="margin:20px 0 8px"><h3>${pkg ? 'รายการในแพ็กเกจ' : 'รายละเอียดเมนู'}</h3><span class="xs muted">กดช่องรูปเพื่ออัปโหลดรูปของรายการ</span></div>
        <div class="ce-items">${draft.items.map((x, i) => `<div class="ce-row">
          ${thumb(x.img, 'itemImg', i, 'อัปโหลดรูปรายการ')}
          <input class="input sm ce-name" data-item="${i}" data-f="name" value="${H(x.name)}" placeholder="${pkg ? 'เช่น Coffee break' : 'เช่น ชา กาแฟ น้ำผลไม้'}" aria-label="ชื่อรายการ">
          <span class="stepper ce-qty"><button type="button" class="minus" data-ce="qty" data-i="${i}" data-d="-1" aria-label="ลด">−</button><input data-item="${i}" data-f="qty" inputmode="numeric" value="${H(x.qty)}" placeholder="–" aria-label="จำนวน"><button type="button" class="plus" data-ce="qty" data-i="${i}" data-d="1" aria-label="เพิ่ม">+</button></span>
          <input class="input sm ce-unit" data-item="${i}" data-f="unit" list="ceItemUnits" value="${H(x.unit)}" placeholder="หน่วย" aria-label="หน่วย">
          <input class="input sm ce-detail" data-item="${i}" data-f="detail" value="${H(x.detail)}" placeholder="รายละเอียดเพิ่มเติม" aria-label="รายละเอียด">
          <span class="ce-ord"><button type="button" class="icon-btn" data-ce="up" data-i="${i}" aria-label="เลื่อนขึ้น" ${i ? '' : 'disabled'}>↑</button><button type="button" class="icon-btn" data-ce="down" data-i="${i}" aria-label="เลื่อนลง" ${i < draft.items.length - 1 ? '' : 'disabled'}>↓</button><button type="button" class="icon-btn" data-ce="del" data-i="${i}" aria-label="ลบรายการ" style="color:var(--red)">${ICON.trash}</button></span>
          <div class="ce-preview xs muted">${H(D.itemText(x)) || '&nbsp;'}</div>
        </div>`).join('')}</div>
        <datalist id="ceItemUnits">${ITEM_UNITS.map(u => `<option>${u}</option>`).join('')}</datalist>
        <button type="button" class="add-row" style="padding:12px 0" data-ce="add">+ เพิ่มรายการ</button>

        <div class="small red" style="min-height:20px;margin-top:6px" role="alert">${H(err)}</div>
        ${busy ? `<div class="small teal">กำลังอัปโหลดรูป ${busy} รูป…</div>` : ''}
        <div class="acts" style="flex-wrap:wrap">${orig ? '<button type="button" class="btn ghost" data-ce="remove" style="color:var(--red);margin-right:auto">ลบเมนูนี้</button>' : ''}<button type="button" class="btn ghost" data-act="closeModal">ยกเลิก</button><button type="button" class="btn red" data-ce="save" ${busy ? 'disabled' : ''}>บันทึก</button></div>`;
      modal.scrollTop = y;
    }

    async function upload(files, put) {
      if (!files.length) return;
      busy += files.length; err = ''; draw();
      for (const f of files) {
        try { put(await storeImage(f)); }
        catch (e) { err = e.message || 'อัปโหลดรูปไม่สำเร็จ'; }
        busy--; draw();
      }
    }

    modal.addEventListener('input', e => {
      const t = e.target;
      if (t.dataset.top) {
        draft[t.dataset.top] = t.value;
        if (t.dataset.top === 'cat') draw();
      } else if (t.dataset.item != null) {
        const x = draft.items[Number(t.dataset.item)];
        x[t.dataset.f] = t.dataset.f === 'qty' ? (t.value.trim() === '' ? '' : Math.max(0, parseInt(t.value, 10) || 0)) : t.value;
        const pv = t.closest('.ce-row').querySelector('.ce-preview'); if (pv) pv.textContent = D.itemText(x) || ' ';
      }
    });
    modal.addEventListener('click', async e => {
      const b = e.target.closest('[data-ce]'); if (!b) return;
      const i = Number(b.dataset.i), act = b.dataset.ce;
      if (act === 'best') { draft.best = !draft.best; draw(); }
      if (act === 'rc') { draft.rc = !draft.rc; draw(); }
      if (act === 'tag') { const k = b.dataset.v, j = draft.tags.indexOf(k); if (j >= 0) draft.tags.splice(j, 1); else draft.tags.push(k); draw(); }
      if (act === 'add') { draft.items.push({ name: '', qty: '', unit: '', detail: '', img: null }); draw(); const ins = modal.querySelectorAll('.ce-name'); ins[ins.length - 1].focus(); }
      if (act === 'del') { draft.items.splice(i, 1); draw(); }
      if (act === 'up' && i > 0) { [draft.items[i - 1], draft.items[i]] = [draft.items[i], draft.items[i - 1]]; draw(); }
      if (act === 'down' && i < draft.items.length - 1) { [draft.items[i + 1], draft.items[i]] = [draft.items[i], draft.items[i + 1]]; draw(); }
      if (act === 'qty') { const x = draft.items[i], n = (x.qty === '' ? 0 : Number(x.qty)) + Number(b.dataset.d); x.qty = n <= 0 ? '' : n; draw(); }
      if (act === 'photoAdd') upload(await pickFiles(true), url => draft.photos.push(url));
      if (act === 'photo') upload(await pickFiles(false), url => { draft.photos[i] = url; });
      if (act === 'photoDel') { draft.photos.splice(i, 1); draw(); }
      if (act === 'itemImg') upload(await pickFiles(false), url => { draft.items[i].img = url; });
      if (act === 'itemImgDel') { draft.items[i].img = null; draw(); }
      if (act === 'remove') {
        App.confirm(`ลบ “${orig.name}” ออกจากรายการจัดเลี้ยง? (ใบเสนอราคาเดิมที่มีเมนูนี้จะไม่เปลี่ยน)`, () => {
          const j = D.FOOD.findIndex(f => f.id === orig.id); if (j >= 0) D.FOOD.splice(j, 1);
          saveCatalog(); App.render(); App.toast('ลบ ' + orig.name + ' แล้ว');
        }, 'ลบเมนู');
      }
      if (act === 'save') {
        draft.name = String(draft.name || '').trim();
        const price = parseFloat(String(draft.price).replace(/,/g, ''));
        if (!draft.name) { err = 'กรุณาใส่ชื่อ'; draw(); return; }
        if (!(price >= 0)) { err = 'ราคาไม่ถูกต้อง'; draw(); return; }
        draft.price = price;
        if (!['lunch', 'coffee'].includes(draft.cat)) delete draft.rc; else if (!draft.rc) delete draft.rc;
        draft.unit = String(draft.unit || '').trim() || 'ท่าน';
        draft.min = draft.min === '' || draft.min == null ? '' : Math.max(0, parseInt(draft.min, 10) || 0) || '';
        draft.items = draft.items.map(x => Object.assign(x, { name: String(x.name || '').trim(), unit: String(x.unit || '').trim(), detail: String(x.detail || '').trim() })).filter(x => x.name || x.detail || x.img);
        if (orig) Object.keys(orig).forEach(k => delete orig[k]);
        const target = orig || {};
        Object.assign(target, draft);
        if (!orig) { target.id = 'c-' + U.uid(); D.FOOD.push(target); }
        D.normFood(target);
        saveCatalog(); App.closeModal(); App.render();
        App.toast((orig ? 'บันทึก ' : 'เพิ่ม ') + target.name + ' แล้ว' + (orig ? ' — ราคาใหม่ใช้กับรายการที่เพิ่มใหม่' : ''));
      }
    });
    draw();
  }

  App.ACT.foodNew = () => { if (App.can('edit_prices')) openEditor(null); };
  App.ACT.foodEdit = el => { if (App.can('edit_prices')) openEditor(el.dataset.v); };
})(window);
