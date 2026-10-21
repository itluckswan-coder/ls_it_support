/* Luck Swan Sales — colour theme, chosen with the rainbow button beside the language button.
 * The standard theme is the app's own colours and is what everyone starts with; a choice (a ready-made theme,
 * any colour, or a mix of a main colour and a background) is kept on this device until it is changed back.
 * A theme only swaps the colour tokens on :root — printed documents keep the standard colours. */
(function (g) {
  'use strict';
  const KEY = 'lsSales.theme';
  const STD = { accent: '#A8262D', bg: '#F5F4EF', ink: '#17211F' };
  // ready-made themes: [id, Thai name, main colour, background, text]
  const PRESETS = [
    ['std', 'มาตรฐาน', STD.accent, STD.bg, STD.ink],
    ['ocean', 'ทะเล', '#1F6FB2', '#F1F5F9', '#13263A'],
    ['emerald', 'มรกต', '#0E8A6A', '#F0F6F2', '#12302A'],
    ['royal', 'ม่วงราชา', '#6D3FB0', '#F5F2FA', '#231A33'],
    ['sunset', 'ส้มอาทิตย์', '#D2591A', '#FBF4EE', '#2D1D12'],
    ['sakura', 'ชมพูซากุระ', '#C2416F', '#FBF1F4', '#2E1820'],
    ['gold', 'ทองล้านนา', '#9A6E00', '#F8F5EA', '#2A2410'],
    ['graphite', 'เทากราไฟต์', '#3D4A57', '#F2F3F5', '#1A1F24']
  ];
  // every colour: a quick row of hues around the colour wheel
  const HUES = ['#D32F2F', '#E8590C', '#F2A900', '#7CB342', '#2E9D4F', '#00897B', '#0097A7', '#1E88E5', '#3949AB', '#6A3FB5', '#AD1F8F', '#D81B60'];

  /* ---------- colour arithmetic ---------- */
  const rgb = h => { h = h.replace('#', ''); if (h.length === 3) h = h.replace(/./g, c => c + c); const n = parseInt(h, 16); return [n >> 16 & 255, n >> 8 & 255, n & 255]; };
  const hex = c => '#' + c.map(v => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('').toUpperCase();
  const mix = (a, b, t) => { const x = rgb(a), y = rgb(b); return hex(x.map((v, i) => v + (y[i] - v) * t)); };
  const lum = h => { const [r, gg, b] = rgb(h).map(v => { v /= 255; return v <= .03928 ? v / 12.92 : Math.pow((v + .055) / 1.055, 2.4); }); return .2126 * r + .7152 * gg + .0722 * b; };
  const contrast = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + .05) / (Math.min(x, y) + .05); };
  const valid = h => typeof h === 'string' && /^#[0-9a-f]{6}$/i.test(h);
  // white text has to stay readable on the main colour (buttons, badges), so a light colour is deepened
  const readable = h => { let c = h; for (let i = 0; i < 20 && contrast(c, '#FFFFFF') < 4; i++) c = mix(c, '#000000', .08); return c; };
  // the page background stays light so the white cards and dark text still read
  const light = h => { let c = h; for (let i = 0; i < 20 && lum(c) < .72; i++) c = mix(c, '#FFFFFF', .12); return c; };

  let cur = { id: 'std' };
  try { const v = JSON.parse(localStorage.getItem(KEY) || 'null'); if (v && (PRESETS.some(p => p[0] === v.id) || (v.id === 'custom' && valid(v.accent) && valid(v.bg)))) cur = v; } catch (e) { /* ignore */ }

  const colours = t => {
    if (t.id === 'custom') { const accent = readable(t.accent), bg = light(t.bg); return { accent, bg, ink: mix('#141816', accent, .14) }; }
    const p = PRESETS.find(x => x[0] === t.id) || PRESETS[0]; return { accent: p[2], bg: p[3], ink: p[4] };
  };
  function apply() {
    let st = document.getElementById('themeVars');
    if (cur.id === 'std') { if (st) st.remove(); setMeta(STD.accent); return; }
    const { accent, bg, ink } = colours(cur);
    const vars = `--red:${accent};--red-d:${mix(accent, '#000000', .16)};--red-soft:${mix(accent, '#FFFFFF', .84)};--red-tint:${mix(accent, '#FFFFFF', .94)};`
      + `--bg:${bg};--line:${mix(bg, ink, .09)};--line2:${mix(bg, ink, .05)};--gray-soft:${mix(bg, ink, .07)};--ink:${ink};--ink2:${mix(ink, bg, .25)};--dark:${mix(ink, accent, .1)};`;
    const keep = `--red:${STD.accent};--red-d:#8E1F25;--red-soft:#F8E1E1;--red-tint:#FDF2F2;--bg:${STD.bg};--line:#E7E5DD;--line2:#EFEDE6;--gray-soft:#EDECE6;--ink:${STD.ink};--ink2:#4A524F;--dark:#1B2421;`;
    if (!st) { st = document.createElement('style'); st.id = 'themeVars'; document.head.appendChild(st); }
    st.textContent = `:root{${vars}}.doc,.fnf-wrap,.ssf-wrap{${keep}}`;
    setMeta(accent);
  }
  function setMeta(c) { const m = document.querySelector('meta[name=theme-color]'); if (m) m.setAttribute('content', c); }
  function save() { try { if (cur.id === 'std') localStorage.removeItem(KEY); else localStorage.setItem(KEY, JSON.stringify(cur)); } catch (e) { /* ignore */ } }
  apply();

  const THEME = g.THEME = {
    PRESETS, get current() { return cur; },
    set(t) { cur = t; save(); apply(); draw(); },
    button: () => `<button class="bell-btn theme-btn" data-act="themeMenu" aria-haspopup="true" aria-expanded="false" aria-label="ธีมสี" title="ธีมสี"><span class="theme-dot" aria-hidden="true"></span></button>`
  };

  /* ---------- the menu ---------- */
  const H = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  function body() {
    const c = colours(cur), mine = cur.id === 'custom' ? cur : { accent: c.accent, bg: c.bg };
    return `<div class="tp-hd"><b>ธีมสี</b><span class="xs muted">จำไว้ในเครื่องนี้</span></div>
      <div class="tp-sec">ธีมของระบบ</div>
      <div class="tp-grid">${PRESETS.map(([id, name, a, b]) => `<button type="button" class="tp-pre ${cur.id === id ? 'on' : ''}" data-theme="${id}" aria-pressed="${cur.id === id}"><span class="tp-sw" style="background:linear-gradient(135deg,${a} 0 50%,${b} 50% 100%)"></span><span class="tp-name">${H(name)}${id === 'std' ? '<small>ค่าเริ่มต้น</small>' : ''}</span></button>`).join('')}</div>
      <div class="tp-sec">ทุกสี</div>
      <div class="tp-hues">${HUES.map(h => `<button type="button" class="tp-hue ${cur.id === 'custom' && cur.accent.toUpperCase() === h ? 'on' : ''}" data-hue="${h}" style="background:${h}" aria-label="สี ${h}" title="${h}"></button>`).join('')}</div>
      <div class="tp-sec">ผสมสีเอง</div>
      <div class="tp-mix"><label><input type="color" id="tpAccent" value="${H(mine.accent)}"><span>สีหลัก</span></label><label><input type="color" id="tpBg" value="${H(mine.bg)}"><span>สีพื้นหลัง</span></label></div>
      <div class="xs muted" style="margin-top:6px">สีที่อ่อนหรือเข้มเกินไป ระบบจะปรับให้อ่านง่ายเอง · เอกสารที่พิมพ์ใช้สีมาตรฐานเสมอ</div>
      <button type="button" class="btn ghost sm block tp-reset" data-theme="std" ${cur.id === 'std' ? 'disabled' : ''}>กลับเป็นสีมาตรฐาน</button>`;
  }
  function draw() { const pop = document.querySelector('.theme-pop'); if (pop) { pop.innerHTML = body(); if (g.I18N) g.I18N.apply(pop); } }
  const close = () => { document.querySelectorAll('.theme-pop').forEach(x => x.remove()); document.querySelectorAll('[data-act=themeMenu]').forEach(b => b.setAttribute('aria-expanded', 'false')); };
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-act=themeMenu]'), pop = e.target.closest('.theme-pop');
    if (pop) {
      const t = e.target.closest('[data-theme]'), h = e.target.closest('[data-hue]');
      if (t) THEME.set({ id: t.dataset.theme });
      else if (h) THEME.set({ id: 'custom', accent: h.dataset.hue, bg: colours(cur).bg });
      return;
    }
    if (!b) { close(); return; }
    e.preventDefault(); e.stopPropagation();
    if (document.querySelector('.theme-pop')) { close(); return; }
    const r = b.getBoundingClientRect(), el = document.createElement('div');
    el.className = 'nav-pop theme-pop'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-label', 'ธีมสี');
    el.style.top = (r.bottom + 6) + 'px'; el.style.left = Math.max(8, Math.min(r.right - 300, innerWidth - 308)) + 'px';
    el.style.maxHeight = (innerHeight - r.bottom - 20) + 'px';
    el.innerHTML = body(); document.body.appendChild(el); b.setAttribute('aria-expanded', 'true');
  }, true);
  // the colour pickers recolour as they move
  document.addEventListener('input', e => {
    if (e.target.id !== 'tpAccent' && e.target.id !== 'tpBg') return;
    const pop = e.target.closest('.theme-pop');
    cur = { id: 'custom', accent: pop.querySelector('#tpAccent').value.toUpperCase(), bg: pop.querySelector('#tpBg').value.toUpperCase() };
    save(); apply();
    pop.querySelectorAll('.tp-pre.on,.tp-hue.on').forEach(x => x.classList.remove('on'));
    const rs = pop.querySelector('.tp-reset'); if (rs) rs.disabled = false;
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
  window.addEventListener('resize', close);
  window.addEventListener('hashchange', close);
})(window);
