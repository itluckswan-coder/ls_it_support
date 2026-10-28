/* Luck Swan Sales — properties (สถานที่): the picker after sign-in, the current property, and who may open which.
 * Quotations carry q.site (missing = resort); function sheets and staff plans follow their quotation.
 * user.sites lists the properties an account may open (missing or empty = all); admins open all. */
(function (g) {
  const H = U.esc;
  const PROPS = [
    { id: 'resort', word: 'Luck Swan', sub: 'RESORT', name: 'Luckswan Resort', th: 'ลักษวรรณ รีสอร์ต', addr: '499 Moo.4 T.RIMKOK A.MUANG CHIANGRAI CHIANGRAI 57100', logo: 'img/logo-wide.png' },
    { id: 'boutique', word: 'Luck Swan', sub: 'BOUTIQUE', name: 'Luck Swan Boutique Hotel', th: 'ลักษวรรณ บูติก', addr: '225 Moo. 4 T.RIMKOK A.MUANG CHIANGRAI', logo: '' },
    { id: 'maneein', word: 'Manee Inn', sub: 'HOTEL', name: 'Manee Inn Hotel', th: 'มณีอินทร์', addr: '25 Moo. 14 T.ROBWIANG A.MUANG CHIANGRAI', logo: '' }
  ];
  const KEY = 'lsSales.site';
  const prop = id => PROPS.find(p => p.id === id) || PROPS[0];
  // a property without its logo file yet shows its name in the brand colours
  const logo = (p, cls) => p.logo ? `<img class="${cls}" src="${p.logo}" alt="${H(p.name)}">` : `<span class="${cls} site-word"><b>${H(p.word)}</b><small>${p.sub}</small></span>`;

  Object.assign(App, {
    PROPS, prop, propLogo: logo,
    sitesOf(u) { return !u ? [] : u.role === 'admin' || !Array.isArray(u.sites) || !u.sites.length ? PROPS.map(p => p.id) : PROPS.map(p => p.id).filter(id => u.sites.includes(id)); },
    mySites() { return App.sitesOf(App.me()); },
    site() {
      let id = null; try { id = sessionStorage.getItem(KEY); } catch (e) { /* ignore */ }
      id = id || App.state.site;
      return id && App.mySites().includes(id) ? id : null;
    },
    setSite(id) { App.state.site = id; try { sessionStorage.setItem(KEY, id); } catch (e) { /* ignore */ } },
    clearSite() { App.state.site = null; try { sessionStorage.removeItem(KEY); } catch (e) { /* ignore */ } },
    siteOfQ: q => (q && q.site) || 'resort',
    inSite: q => App.siteOfQ(q) === (App.site() || 'resort'),
    siteQuotes() { return Store.db.quotes.filter(App.inSite); },
    siteBeos() { return Store.db.beos.filter(b => { const q = Store.quote(b.quoteId); return !q || App.inSite(q); }); }
  });

  App.V.site = {
    bare: true, title: 'เลือกสถานที่',
    render() {
      const me = App.me(), mine = App.mySites(), cur = App.site();
      return `<div class="site-page">
  <div class="between site-top"><img src="img/logo-wide-sm.png" alt="Luck Swan" style="width:120px">
    <div class="row" style="gap:8px">${g.I18N ? g.I18N.button() : ''}${g.THEME ? g.THEME.button() : ''}<button class="btn ghost sm" data-act="logout">ออกจากระบบ</button></div></div>
  <h1>เลือกสถานที่</h1><p class="muted" style="margin:-4px 0 22px">สวัสดีคุณ${H((me && me.name || '').split(' ')[0])} · เลือกสถานที่ที่จะทำงาน เปลี่ยนได้ภายหลังจากแถบด้านบน</p>
  <div class="site-cards">${PROPS.filter(p => mine.includes(p.id)).map(p => `<button class="site-card ${p.id === cur ? 'on' : ''}" data-act="sitePick" data-v="${p.id}">
    <span class="sc-hero">${logo(p, 'sc-logo')}</span>
    <span class="sc-foot"><span class="sc-thumb">${logo(p, 'sc-mini')}</span><span class="sc-txt"><b>${H(p.name)}</b><span class="small muted">${H(p.th)}</span><span class="small">${H(p.addr)}</span></span></span>
  </button>`).join('')}</div>
  ${mine.length ? '' : '<div class="empty">บัญชีนี้ยังไม่ได้รับสิทธิ์เข้าสถานที่ใด — กรุณาติดต่อผู้ดูแลระบบ</div>'}
</div>`;
    }
  };
  Object.assign(App.ACT, {
    sitePick(el) { App.setSite(el.dataset.v); App.state.ctx = null; try { sessionStorage.removeItem('lsSales.ctx'); } catch (e) { /* ignore */ } App.go('#/home'); App.toast('สถานที่: ' + prop(el.dataset.v).name); },
    siteSwitch() { App.go('#/site'); }
  });
})(window);
