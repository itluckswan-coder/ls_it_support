/* Luck Swan Sales — interface language: ไทย (default), English (US), မြန်မာ.
 * The screens are written in Thai; after anything is drawn, Thai text and the placeholder / title / aria-label
 * attributes are swapped for the chosen language from the dictionaries (exact phrases first, then the few
 * patterns with numbers or dates in them, then dates and known words — but only when nothing Thai is left).
 * What people type (names, reasons, task titles) is left exactly as written, and
 * printed documents (.doc) and editable boxes are never touched. Words missing from a dictionary stay Thai. */
(function (g) {
  'use strict';
  const KEY = 'lsSales.lang';
  const LANGS = [['th', 'ไทย', 'TH'], ['en', 'English (US)', 'EN'], ['my', 'မြန်မာ', 'MY']];
  let lang = 'th';
  try { const v = localStorage.getItem(KEY); if (LANGS.some(l => l[0] === v)) lang = v; } catch (e) { /* ignore */ }

  // dictionaries: Thai phrase → [English, Burmese]; filled per section by I18N.add()
  const DICT = new Map();
  // patterns: [regexp on the whole text, (match, lang) => text]; and pieces swapped inside longer text
  const PATTERNS = [], PIECES = [];
  const missing = new Set();

  const MS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const ML = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  const EN_MS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const EN_ML = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const MY_ML = ['ဇန်နဝါရီ', 'ဖေဖော်ဝါရီ', 'မတ်', 'ဧပြီ', 'မေ', 'ဇွန်', 'ဇူလိုင်', 'ဩဂုတ်', 'စက်တင်ဘာ', 'အောက်တိုဘာ', 'နိုဝင်ဘာ', 'ဒီဇင်ဘာ'];
  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const RE_DATE = new RegExp('(\\d{1,2}) (' + MS.map(esc).join('|') + ')(?: (25\\d\\d))?', 'g');
  const RE_RANGE = new RegExp('(\\d{1,2})–(\\d{1,2}) (' + MS.map(esc).join('|') + ')(?: (25\\d\\d))?', 'g');
  const RE_MONTH = new RegExp('(' + ML.map(esc).join('|') + ') (25\\d\\d)', 'g');
  const year = y => y ? String(+y - 543) : '';
  // Thai dates (20 ต.ค. 2569, ตุลาคม 2569) in the chosen language, with Christian-era years
  function dates(s, L) {
    return s.replace(RE_RANGE, (m, a, b, mo, y) => { const i = MS.indexOf(mo);
      return L === 'en' ? EN_MS[i] + ' ' + a + '–' + b + (y ? ', ' + year(y) : '') : a + '–' + b + ' ' + MY_ML[i] + (y ? ' ' + year(y) : ''); })
      .replace(RE_DATE, (m, d, mo, y) => { const i = MS.indexOf(mo);
      return L === 'en' ? EN_MS[i] + ' ' + d + (y ? ', ' + year(y) : '') : d + ' ' + MY_ML[i] + (y ? ' ' + year(y) : ''); })
      .replace(RE_MONTH, (m, mo, y) => { const i = ML.indexOf(mo); return (L === 'en' ? EN_ML[i] : MY_ML[i]) + ' ' + year(y); });
  }
  const THAI = /[\u0E01-\u0E3A\u0E40-\u0E5B]/; // Thai letters, not the ฿ sign
  const idx = L => L === 'en' ? 0 : 1;

  function tr(text, L) {
    const t = text.trim(); if (!t || !THAI.test(t)) return null;
    const hit = DICT.get(t);
    if (hit && hit[idx(L)]) return text.replace(t, hit[idx(L)]);
    for (const [re, fn] of PATTERNS) { const m = re.exec(t); if (m) { const out = fn(m, L, x => tr(x, L) || x); if (out != null) return text.replace(t, out); } }
    // a line made of parts ("name · role · department", "QT-… · customer · date"): each part on its own, so
    // the system's parts are translated and a typed part (a name, a customer) stays as written
    if (t.includes(' · ')) {
      const parts = t.split(' · '), outs = parts.map(x => tr(x, L));
      if (outs.some(Boolean)) { if (outs.some(x => !x)) missing.add(t); return text.replace(t, outs.map((x, i) => x == null ? parts[i] : x).join(' · ')); }
    }
    // dates and known words inside longer text — only when that translates all of it: anything left in Thai
    // is something people typed (a task title, a reason, a name), and that stays exactly as written
    let out = dates(t, L);
    for (const [re, fn] of PIECES) out = out.replace(re, (...m) => fn(m, L));
    if (THAI.test(out)) { missing.add(t); return null; }
    return out !== t ? text.replace(t, out) : null;
  }
  const SKIP = 'script,style,textarea,.doc,[contenteditable],[data-noi18n],.lang-pop';
  function apply(root) {
    if (lang === 'th' || !root) return;
    if (root.nodeType === 3) { if (!root.parentElement || !root.parentElement.closest(SKIP)) { const v = tr(root.nodeValue, lang); if (v != null) root.nodeValue = v; } return; }
    if (root.nodeType !== 1 || root.closest(SKIP)) return;
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, { acceptNode: n => n.nodeType === 1 && n.matches(SKIP) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT });
    const els = [root], texts = [];
    for (let n = w.nextNode(); n; n = w.nextNode()) (n.nodeType === 3 ? texts : els).push(n);
    for (const n of texts) { const v = tr(n.nodeValue, lang); if (v != null) n.nodeValue = v; }
    for (const el of els) for (const a of ['placeholder', 'title', 'aria-label']) { const v = el.getAttribute(a); if (v && THAI.test(v)) { const o = tr(v, lang); if (o != null) el.setAttribute(a, o); } }
  }
  const obs = new MutationObserver(list => { if (lang === 'th') return; for (const m of list) for (const n of m.addedNodes) apply(n); });

  const I18N = g.I18N = {
    LANGS, missing,
    get lang() { return lang; },
    // add a section's words: { 'ไทย': ['English', 'မြန်မာ'] }
    add(dict) { for (const [k, v] of Object.entries(dict)) DICT.set(k, v); },
    // a pattern for text with values in it: I18N.pattern(/^รวม (\d+) วัน$/, (m, L) => L === 'en' ? m[1] + ' days' : ...)
    pattern(re, fn) { PATTERNS.push([re, fn]); },
    piece(re, fn, first) { PIECES.push([re, fn, first ? 2e9 : 1e9]); PIECES.sort((x, y) => y[2] - x[2]); },
    // words swapped inside longer text (department and shift names, units): across all sections the longest
    // words go first, so "หัวหน้าแผนก" is replaced whole before "แผนก" on its own
    pieces(dict) {
      for (const k of Object.keys(dict)) PIECES.push([new RegExp(esc(k), 'g'), (m, L) => dict[k][L === 'en' ? 0 : 1], k.length]);
      PIECES.sort((x, y) => y[2] - x[2]);
    },
    t(text) { return lang === 'th' ? text : (tr(text, lang) || text); },
    apply,
    set(next) {
      if (!LANGS.some(l => l[0] === next) || next === lang) return;
      lang = next; try { localStorage.setItem(KEY, next); } catch (e) { /* ignore */ }
      document.documentElement.lang = next;
      if (g.App) { g.App.closeModal && g.App.closeModal(); g.App.render(); }
      apply(document.body);
    }
  };
  document.documentElement.lang = lang;
  const start = () => { obs.observe(document.body, { childList: true, subtree: true }); apply(document.body); };
  if (document.body) start(); else document.addEventListener('DOMContentLoaded', start);

  /* ---------- the language button (beside the bell) and its menu ---------- */
  const ICON = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m5 8 6 6"/><path d="m4 14 6-6 2-3"/><path d="M2 5h12"/><path d="M7 2h1"/><path d="m22 22-5-10-5 10"/><path d="M14 18h6"/></svg>';
  I18N.button = () => `<button class="bell-btn lang-btn" data-act="langMenu" aria-haspopup="true" aria-expanded="false" aria-label="Language / ภาษา" title="Language / ภาษา">${ICON}<span class="lang-code" data-noi18n>${LANGS.find(l => l[0] === lang)[2]}</span></button>`;
  const close = () => { document.querySelectorAll('.lang-pop').forEach(x => x.remove()); document.querySelectorAll('[data-act=langMenu]').forEach(b => b.setAttribute('aria-expanded', 'false')); };
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-act=langMenu]'), opt = e.target.closest('[data-lang]');
    if (opt) { e.preventDefault(); close(); I18N.set(opt.dataset.lang); return; }
    if (!b) { if (!e.target.closest('.lang-pop')) close(); return; }
    e.preventDefault(); e.stopPropagation();
    if (document.querySelector('.lang-pop')) { close(); return; }
    const r = b.getBoundingClientRect(), pop = document.createElement('div');
    pop.className = 'nav-pop lang-pop'; pop.setAttribute('role', 'menu');
    pop.style.top = (r.bottom + 6) + 'px'; pop.style.left = Math.max(8, Math.min(r.right - 210, innerWidth - 218)) + 'px';
    pop.innerHTML = LANGS.map(([id, name, code]) => `<a role="menuitemradio" aria-checked="${id === lang}" href="#" data-lang="${id}" class="${id === lang ? 'on' : ''}" lang="${id}"><span class="lang-code">${code}</span>${name}${id === 'th' ? ' <span class="xs" style="opacity:.7">(ค่าเริ่มต้น)</span>' : ''}</a>`).join('');
    document.body.appendChild(pop); b.setAttribute('aria-expanded', 'true');
  }, true);
  window.addEventListener('resize', close);
  window.addEventListener('hashchange', close);
})(window);
