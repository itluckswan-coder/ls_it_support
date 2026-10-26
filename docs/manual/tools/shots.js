// every screen for the user manual, in th / en / my
const { chromium } = require('playwright'); const BASE = 'http://localhost:8765/'; const OUT = process.argv[2]; // folder for the PNGs, e.g. /tmp/manual-raw
const ONLY = process.argv[3] ? process.argv[3].split(',') : null;
const W = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const b = await chromium.launch(); const errs = [];
  for (const L of ['th', 'en', 'my']) {
    const p = await (await b.newContext({ viewport: { width: 1280, height: 820 } })).newPage(); p.on('pageerror', e => errs.push(L + ':' + e.message));
    const shot = async (n, full) => { if (ONLY && !ONLY.includes(n)) return; await W(350); await p.screenshot({ path: `${OUT}/${n}-${L}.png`, fullPage: !!full }); };
    const go = async h => { await p.evaluate(h => { App.closeModal && App.closeModal(); App.closeNavPop && App.closeNavPop(); location.hash = h; }, h); await W(500); };
    const click = async (s, opt) => { try { await p.click(s, Object.assign({ timeout: 4000 }, opt)); await W(350); } catch (e) { errs.push(L + ': click ' + s); } };
    const close = async () => { await p.evaluate(() => { App.closeModal && App.closeModal(); App.closeNavPop && App.closeNavPop(); document.querySelectorAll('.lang-pop,.theme-pop,.pop').forEach(x => x.remove()); }); await W(150); };
    await p.goto(BASE); await p.evaluate(L => { localStorage.clear(); localStorage.setItem('lsSales.lang', L); }, L); await p.goto(BASE + '#/login'); await p.reload(); await p.waitForSelector('#lgU'); await W(500);
    await shot('login');
    await p.fill('#lgU', 'laksika@luckswangroup.com'); await p.fill('#lgP', 'luckswan'); await p.click('button[type=submit]'); await W(1200);
    await shot('home');
    await click('[data-act=langMenu]'); await shot('lang'); await close();
    await click('[data-act=themeMenu]'); await shot('theme'); await close();
    await click('[data-act=notifs]'); await shot('notifs'); await close();
    await click('[data-act=userMenu]'); await shot('usermenu'); await close();
    await go('#/meet'); await shot('meet'); await click('[data-act=meetSel]'); await shot('meet-sel'); await click('[data-act=meetEdit]'); await shot('meet-edit'); await close();
    await go('#/rooms'); await shot('rooms'); await click('[data-act=rmSite][data-v=boutique]'); await click('[data-act=bqStep][data-k=bq_sup][data-v="1"]'); await click('[data-act=bqTier][data-k=bf][data-v="1"]'); await click('[data-act=bqPax][data-k=bf][data-v="1"]'); await p.evaluate(() => window.scrollTo(0, 0)); await shot('rooms-bq'); await p.evaluate(() => { App.state.ui.rooms = null; });
    await go('#/roomrates'); await shot('roomrates'); await p.evaluate(() => document.querySelector('[data-change=rrBq]').scrollIntoView({ block: 'center' })); await shot('roomrates-bq');
    await go('#/food'); await shot('food'); await click('[data-act=foodEdit]'); await shot('food-edit'); await close();
    await go('#/quotes'); await shot('quotes');
    await go('#/quote/QT-2611-042'); await shot('quote-edit');
    await go('#/doc/QT-2611-042'); await shot('quote-doc'); await click('[data-act=docVer][data-v="1"]'); await shot('quote-doc-full'); await click('[data-act=qdTab][data-v=edit]'); await shot('quote-doc-edit'); await click('[data-act=qdCancel]'); await click('[data-act=docVer][data-v="0"]');
    await p.evaluate(() => { App.state.wiz = null; }); await go('#/new');
    for (const [f, v] of [['title', 'สัมมนาประจำปี 2027'], ['client', 'บจก. ตัวอย่าง จำกัด']]) { try { await p.fill(`[data-f=${f}]`, v); await p.dispatchEvent(`[data-f=${f}]`, 'input'); await p.dispatchEvent(`[data-f=${f}]`, 'change'); } catch (e) { errs.push('fill ' + f); } }
    await shot('wiz1');
    for (let i = 2; i <= 5; i++) { await click('[data-act=wizNext]'); await W(300); const e = await p.$('[data-act=wizSel]'); if (e && await e.isVisible()) { await e.click().catch(() => {}); await W(250); await close(); } await shot('wiz' + i); }
    await go('#/beo'); await shot('beo-list');
    await go('#/beo/FS-2611-018'); await shot('beo'); await click('[data-act=btFold][data-k=guests]'); await click('[data-act=btFold][data-k=kinds]'); await click('[data-act=btPick][data-k=guests][data-v=g4]'); await click('[data-act=btPick][data-k=kinds][data-v=k6]'); await p.evaluate(() => document.querySelector('.bt-grid').scrollIntoView({ block: 'center' })); await shot('beo-tags');
    await go('#/plan/FS-2611-018/1'); await click('[data-act=fpAuto]'); await shot('plan');
    await go('#/plansetup'); await shot('plansetup');
    await go('#/fn/FS-2611-018'); await shot('fn');
    await go('#/ss/FS-2611-018'); await click('[data-act=ssTab][data-v=form]'); await shot('ss');
    await go('#/cal'); await click('[data-act=calNav][data-v="1"]'); await shot('cal-month'); await click('[data-act=calView][data-v=week]'); await shot('cal-week'); await click('[data-act=calView][data-v=pipe]'); await shot('cal-pipe');
    await go('#/staff'); await click('[data-act=stTab][data-v=shifts]'); await shot('staff-shifts'); await click('[data-act=stAssign]'); await shot('staff-assign'); await close();
    await click('[data-act=stTab][data-v=emps]'); await shot('staff-emps'); await click('[data-act=empEdit]'); await shot('staff-empform');
    await click('#emSiteBtn'); await click('#emDeptBtn'); await click('#emPosBtn'); await shot('staff-plus', true); await close();
    await click('[data-act=empImport]'); await shot('staff-import'); await close();
    await click('[data-act=stTab][data-v=rules]'); await shot('staff-rules');
    await go('#/tasks'); await shot('tasks'); await click('[data-act=tkNew]'); await shot('task-new'); await click('[data-to=dept]'); await shot('task-site'); await close();
    await click('[data-act=tpOpen]'); await click('[data-tp-mode=org]'); await click('[data-tp-zoom=fit]'); await shot('task-perms-org'); await click('[data-tp-mode=form]'); await shot('task-perms-form'); await close();
    await go('#/home'); await click('[data-act=tkView]'); await shot('task-view'); await close();
    await go('#/tasks/approve'); await shot('tasks-approve');
    // leave: sample requests
    await p.evaluate(() => { const me = App.me(); const mk = (type, x, st) => Store.db.leaves.push(Object.assign({ id: U.uid('lv'), no: 'HRSD08.1/69102100' + (Store.db.leaves.length + 1), type, by: me.id, byName: me.name, fullName: 'วิศรุต เสนตา', nick: 'โอ๊ค', position: 'IT', dept: 'IT', branch: 'ลักษวรรณ รีสอร์ท', reason: 'ติดตั้งและทดสอบโปรแกรม', at: U.nowTH(), status: st, sign: st === 'approved' ? { head: { byName: me.name, at: U.nowTH(), ok: true }, gm: { byName: me.name, at: U.nowTH(), ok: true } } : {} }, x));
      mk('qt', { qt: { mode: 'do', from: '08:00', to: '17:00', date: U.today(), hours: 9 } }, 'approved'); mk('change', { ch: { mode: 'day', fromDate: U.today(), toDate: U.addDays(U.today(), 3) } }, 'pending');
      mk('scan', { sc: { cause: 'fail', inTime: '08:01', outTime: '', date: U.today(), witness: 'ชนิดา มั่นคง', witnessPos: 'แม่บ้าน' } }, 'pending'); Store.save(); });
    await go('#/leave'); await shot('leave');
    await click('[data-act=lvNew][data-v=qt]'); await shot('leave-qt'); await close();
    await click('[data-act=lvNew][data-v=change]'); await shot('leave-change'); await close();
    await click('[data-act=lvNew][data-v=scan]'); await shot('leave-scan'); await close();
    await click('[data-act=lvTab][data-v=all]'); await click('tbody tr:nth-child(2)'); await shot('leave-view'); await close();
    await click('[data-act=lvTab][data-v=info]'); await shot('leave-info');
    await go('#/users'); await shot('users'); await click('[data-act=usEdit]'); await shot('user-edit'); await close(); await click('[data-act=roleAdd]'); await shot('role'); await close();
    await go('#/dash'); await shot('dash');
    await p.close();
  }
  console.log('errors:', errs.join(' | ') || 'none'); await b.close();
})();
