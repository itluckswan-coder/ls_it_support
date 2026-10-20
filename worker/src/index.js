/* Luck Swan Sales — API on Cloudflare Workers + D1
 * Serves /api/*; every other path is the static web app (../sales-app). */

const RANK = { sales: 1, acc: 1, head: 2, manager: 3, admin: 4 };
const ALL_PERMS = ['create_quote', 'discount10', 'discount_over', 'view_beo', 'customers', 'staff', 'users', 'edit_prices', 'leave_hr'];
const PBKDF2_ITER = 10000; // kept modest so login fits the Workers free-plan CPU budget
const SESSION_DAYS = { remember: 30, normal: 1 };
const KINDS = ['quote', 'beo', 'staff', 'employees', 'settings', 'catalog', 'task', 'leave'];
const MAX_IMAGE_BYTES = 1500000;

/* ---------------- helpers ---------------- */
const now = () => new Date(Date.now() + 7 * 3600e3).toISOString().slice(0, 19); // Thailand time (UTC+7)
const json = (data, status = 200) => new Response(JSON.stringify(data), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', 'x-content-type-options': 'nosniff' }
});
class HttpError extends Error { constructor(status, msg) { super(msg); this.status = status; } }
const fail = (status, msg) => { throw new HttpError(status, msg); };
const hex = buf => Array.from(new Uint8Array(buf), b => b.toString(16).padStart(2, '0')).join('');
const randomHex = n => hex(crypto.getRandomValues(new Uint8Array(n)));
const sha256 = async s => hex(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)));

async function hashPassword(pw, saltHex) {
  const salt = saltHex || randomHex(16);
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: new TextEncoder().encode(salt), iterations: PBKDF2_ITER }, key, 256);
  return { hash: hex(bits), salt };
}
function safeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let r = 0; for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
async function readJson(req) {
  try { return await req.json(); } catch (e) { fail(400, 'ข้อมูลไม่ถูกต้อง'); }
}
const can = (u, p) => u.role === 'admin' || !!(u.perms && u.perms[p]);
const publicUser = row => Object.assign(JSON.parse(row.data), { id: row.id, email: row.email });

/* ---------------- auth ---------------- */
async function userCount(env) { return (await env.DB.prepare('SELECT COUNT(*) AS n FROM users').first()).n; }
async function createSession(env, userId, remember) {
  const token = randomHex(32);
  const days = remember ? SESSION_DAYS.remember : SESSION_DAYS.normal;
  await env.DB.prepare('INSERT INTO sessions (token_hash, user_id, expires_at, created_at) VALUES (?, ?, ?, ?)')
    .bind(await sha256(token), userId, Date.now() + days * 864e5, now()).run();
  return token;
}
async function authUser(req, env) {
  const m = /^Bearer ([0-9a-f]{64})$/.exec(req.headers.get('authorization') || '');
  if (!m) fail(401, 'กรุณาเข้าสู่ระบบ');
  const s = await env.DB.prepare('SELECT user_id, expires_at FROM sessions WHERE token_hash = ?').bind(await sha256(m[1])).first();
  if (!s || s.expires_at < Date.now()) fail(401, 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const row = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(s.user_id).first();
  if (!row) fail(401, 'ไม่พบผู้ใช้');
  const u = publicUser(row);
  if (u.status === 'disabled') fail(403, 'บัญชีนี้ถูกปิดใช้งาน');
  u._tokenHash = await sha256(m[1]);
  return u;
}

async function login(req, env) {
  const { email, password, remember } = await readJson(req);
  const ident = String(email || '').trim().toLowerCase();
  if (!ident || !password) fail(400, 'กรุณากรอกอีเมลและรหัสผ่าน');
  const lockKey = ident.includes('@') ? ident : ident + '@';
  const lf = await env.DB.prepare('SELECT * FROM login_fail WHERE email = ?').bind(lockKey).first();
  if (lf && lf.count >= 5 && Date.now() - lf.first_at < 15 * 60e3) fail(429, 'ใส่รหัสผิดหลายครั้ง กรุณารอ 15 นาทีแล้วลองใหม่');
  const row = await env.DB.prepare("SELECT * FROM users WHERE email = ? OR (instr(email, '@') > 0 AND lower(substr(email, 1, instr(email, '@') - 1)) = ?) OR lower(id) = ? LIMIT 1")
    .bind(ident, ident, ident).first();
  const check = row && row.pw_hash ? await hashPassword(String(password), row.pw_salt) : null;
  if (!row || !check || !safeEqual(check.hash, row.pw_hash)) {
    const fresh = !lf || Date.now() - lf.first_at >= 15 * 60e3;
    await env.DB.prepare('INSERT INTO login_fail (email, count, first_at) VALUES (?, 1, ?) ON CONFLICT(email) DO UPDATE SET count = ?, first_at = ?')
      .bind(lockKey, Date.now(), fresh ? 1 : lf.count + 1, fresh ? Date.now() : lf.first_at).run();
    fail(401, row ? 'รหัสผ่านไม่ถูกต้อง' : 'ไม่พบผู้ใช้นี้ในระบบ');
  }
  const u = publicUser(row);
  if (u.status === 'disabled') fail(403, 'บัญชีนี้ถูกปิดใช้งาน กรุณาติดต่อผู้ดูแลระบบ');
  await env.DB.prepare('DELETE FROM login_fail WHERE email = ?').bind(lockKey).run();
  if (u.status === 'invited') u.status = 'active';
  u.lastLogin = now();
  let dev = String(req.headers.get('x-device') || '');
  try { dev = decodeURIComponent(dev); } catch (e) { /* keep as is */ }
  u.lastDevice = dev.slice(0, 40);
  await saveUserRow(env, u, row.version + 1);
  await env.DB.prepare('DELETE FROM sessions WHERE expires_at < ?').bind(Date.now()).run();
  return json({ token: await createSession(env, u.id, !!remember), user: u });
}

async function saveUserRow(env, u, version) {
  const { id, email, pw, _tokenHash, ...rest } = u;
  await env.DB.prepare('UPDATE users SET email = ?, data = ?, version = ?, updated_at = ? WHERE id = ?')
    .bind(email, JSON.stringify(rest), version, now(), id).run();
}

/* ---------------- state ---------------- */
/* leave requests (การแจ้งใบลา, HRSD08.1) are personal: the requester, the head of their department and the
   approvers (ผู้จัดการทั่วไป / ฝ่ายบุคคล) see them; nobody else gets them from the server */
const sameDept = (a, b) => !!a && !!b && String(a).trim() === String(b).trim();
const leaveHead = (l, u) => (u.role === 'head' || u.role === 'manager') && sameDept(l.dept, u.dept);
const leaveSeen = (l, u) => !!l && (l.by === u.id || can(u, 'leave_hr') || leaveHead(l, u));
// every signature and the approval on a leave request is the admin's
const leaveSigner = u => u.role === 'admin';
async function state(env, since, user) {
  const docs = since
    ? (await env.DB.prepare('SELECT kind, id, data, version, deleted, updated_at FROM docs WHERE updated_at >= ?').bind(since).all()).results
    : (await env.DB.prepare('SELECT kind, id, data, version, deleted, updated_at FROM docs WHERE deleted = 0').all()).results;
  const users = since
    ? (await env.DB.prepare('SELECT * FROM users WHERE updated_at >= ?').bind(since).all()).results
    : (await env.DB.prepare('SELECT * FROM users').all()).results;
  // every user id, so clients can drop accounts that were deleted since their last sync
  const userIds = since ? (await env.DB.prepare('SELECT id FROM users').all()).results.map(r => r.id) : users.map(r => r.id);
  return {
    serverTime: now(),
    userIds,
    docs: docs.filter(d => d.kind !== 'leave' || !user || d.deleted || leaveSeen(JSON.parse(d.data), user)).map(d => ({ kind: d.kind, id: d.id, version: d.version, deleted: !!d.deleted, data: d.deleted ? null : JSON.parse(d.data) })),
    users: users.map(r => Object.assign(publicUser(r), { _v: r.version }))
  };
}

/* ---------------- business rules ---------------- */
const beoContent = b => {
  const { approval, status, rev, revNote, signoff, ss, ...c } = b || {};
  c.depts = Object.fromEntries(Object.entries((b && b.depts) || {}).map(([k, v]) => [k, { items: v.items || [], owner: v.owner || '' }]));
  return c;
};
// department of a user, from the free-text dept on their account (mirrors DEPT_WORDS in js/v-tasks.js)
const DEPT_WORDS = { kitchen: ['ครัว', 'f&b'], banquet: ['จัดเลี้ยง', 'banquet'], av: ['ช่าง', 'av'], front: ['ส่วนหน้า', 'front'], hk: ['แม่บ้าน', 'housekeeping'], acc: ['บัญชี', 'account'] };
const deptOf = u => { const d = String((u && u.dept) || '').toLowerCase(); return Object.keys(DEPT_WORDS).find(k => DEPT_WORDS[k].some(w => d.includes(w))) || null; };
const isAssignee = (t, u) => !!t && !!t.to && (t.to.type === 'user' ? t.to.id === u.id : t.to.type === 'dept' && t.to.dept === deptOf(u));
const lineTotal = it => (Number(it.price) || 0) * (Number(it.q1) || 0) * (it.q2 == null || it.q2 === '' ? 1 : Number(it.q2) || 0) * (it.free ? 0 : 1);

async function applyRules(env, user, kind, id, data, prev) {
  if (kind === 'quote') {
    if (!can(user, 'create_quote')) fail(403, 'คุณไม่มีสิทธิ์แก้ไขใบเสนอราคา');
    const sub = (data.items || []).reduce((a, it) => a + lineTotal(it), 0);
    const d = data.discount || {};
    const pct = d.type === 'amt' ? (sub ? (Number(d.value) || 0) / sub * 100 : 0) : Number(d.value) || 0;
    if (pct > 10.0001 && !can(user, 'discount_over') && !(prev && prev.approval === 'approved' && JSON.stringify(prev.discount) === JSON.stringify(data.discount))) data.approval = 'pending';
    // approval to book a meeting room over another booking's times: only admins give or withdraw it
    if (user.role !== 'admin') { if (prev && prev.overlapOk) data.overlapOk = prev.overlapOk; else delete data.overlapOk; }
  }
  if (kind === 'beo') {
    if (!can(user, 'create_quote') && !can(user, 'staff')) fail(403, 'คุณมีสิทธิ์ดูใบฟังก์ชันอย่างเดียว');
    const rank = RANK[user.role] || 1;
    const changed = !prev || JSON.stringify(beoContent(prev)) !== JSON.stringify(beoContent(data));
    const prevAppr = prev ? (prev.approval || { state: 'approved' }) : null;
    if (changed) {
      if (rank >= RANK.manager) data.approval = { state: 'approved', by: user.id, byName: user.name, at: now(), auto: true };
      else data.approval = {
        state: 'pending', requestedBy: user.id, requestedByName: user.name, requestedAt: now(),
        base: prevAppr && prevAppr.state === 'approved' ? beoContent(prev) : (prevAppr ? prevAppr.base || null : null)
      };
    } else data.approval = prevAppr || data.approval || { state: 'approved' };
    // the sales summary (ใบสรุปสำหรับฝ่ายขาย) belongs to admins, sales managers and sales staff
    if (!['admin', 'manager', 'sales'].includes(user.role)) { if (prev && prev.ss) data.ss = prev.ss; else delete data.ss; }
    // admin sign-off: only admins set it, and any change to the content clears it
    if (user.role !== 'admin') { if (prev && prev.signoff) data.signoff = prev.signoff; else delete data.signoff; }
    if (changed) delete data.signoff;
    const sending = data.status === 'sent' && (!prev || prev.status !== 'sent' || (data.rev || 0) > (prev.rev || 0));
    if (sending && data.approval.state !== 'approved') fail(409, 'ใบฟังก์ชันนี้ต้องได้รับอนุมัติก่อนส่งให้แผนก');
  }
  if (kind === 'staff') {
    if (!can(user, 'staff') && !can(user, 'create_quote')) fail(403, 'คุณไม่มีสิทธิ์แก้ไขรอบพนักงาน');
    const prevOt = Object.fromEntries(((prev && prev.ot) || []).map(o => [o.id, o]));
    for (const o of data.ot || []) {
      const was = prevOt[o.id] && prevOt[o.id].status;
      if ((o.status === 'approved' || o.status === 'rejected') && was !== o.status && !can(user, 'staff')) o.status = 'pending';
    }
  }
  if (kind === 'task') {
    if (user.role !== 'admin') {
      // assignees may only confirm their own task; everything else stays as the admin wrote it
      if (!prev || !isAssignee(prev, user)) fail(403, 'เฉพาะผู้ดูแลระบบที่มอบหมายงานได้');
      const keep = ({ status, doneBy, doneByName, doneAt, doneNote, ...rest }) => rest;
      if (JSON.stringify(keep(prev)) !== JSON.stringify(keep(data))) fail(403, 'แก้ไขได้เฉพาะการยืนยันงานที่ได้รับมอบหมาย');
      if (prev.status === 'done') fail(409, 'งานนี้ยืนยันแล้ว');
      if (data.status !== 'done') fail(400, 'สถานะไม่ถูกต้อง');
      Object.assign(data, { doneBy: user.id, doneByName: user.name, doneAt: now(), doneNote: String(data.doneNote || '').slice(0, 500) });
    }
    if (!data.to || !['user', 'dept'].includes(data.to.type)) fail(400, 'ระบุผู้รับมอบหมาย');
    if (JSON.stringify(data).length > 20000) fail(413, 'รายละเอียดงานยาวเกินไป');
  }
  if (kind === 'leave') {
    if (!prev) {
      // a new request: always the caller's own, unsigned and waiting for the head of department
      Object.assign(data, { by: user.id, byName: user.name, at: now(), status: 'pending', sign: {} });
    } else {
      if (!leaveSeen(prev, user)) fail(403, 'คุณไม่มีสิทธิ์แก้ไขใบลานี้');
      const keep = ({ status, sign, ...rest }) => rest;
      if (JSON.stringify(keep(prev)) !== JSON.stringify(keep(data))) fail(403, 'แก้ไขรายละเอียดใบลาที่ส่งแล้วไม่ได้');
      const ps = prev.sign || {}, ns = data.sign || {}, steps = ['head', 'gm', 'dgm', 'hr'];
      const added = steps.filter(k => JSON.stringify(ps[k] || null) !== JSON.stringify(ns[k] || null));
      if (added.length > 1 || added.some(k => ps[k])) fail(409, 'ใบลานี้ถูกลงนามไปแล้ว — โหลดใหม่แล้วลองอีกครั้ง');
      const k = added[0], ok = k && ns[k].ok !== false;
      const allowed = !k ? (data.status === prev.status || (data.status === 'cancelled' && prev.by === user.id && ['pending', 'gm'].includes(prev.status)))
        : !leaveSigner(user) ? false
        : k === 'head' ? prev.status === 'pending' && data.status === (ok ? 'gm' : 'rejected')
        : k === 'gm' ? prev.status === 'gm' && data.status === (ok ? 'approved' : 'rejected')
        : k === 'dgm' ? ['approved', 'done'].includes(prev.status) && ok && data.status === prev.status
        : prev.status === 'approved' && ok && data.status === 'done';
      if (!allowed) fail(403, 'คุณลงนามหรือเปลี่ยนสถานะใบลานี้ไม่ได้');
      if (k) Object.assign(ns[k], { by: user.id, byName: user.name, at: now(), note: String(ns[k].note || '').slice(0, 500) });
    }
    if (!['qt', 'change', 'scan'].includes(data.type)) fail(400, 'ประเภทใบลาไม่ถูกต้อง');
    if (JSON.stringify(data).length > 20000) fail(413, 'ข้อมูลใบลายาวเกินไป');
  }
  if (kind === 'catalog') {
    if (!can(user, 'edit_prices')) fail(403, 'คุณไม่มีสิทธิ์แก้ไขเมนูและแพ็กเกจ');
    if (id !== 'food' || !Array.isArray(data.list)) fail(400, 'ข้อมูลเมนูไม่ถูกต้อง');
    if (JSON.stringify(data).length > 800000) fail(413, 'ข้อมูลเมนูใหญ่เกินไป');
  }
  if (kind === 'employees' || kind === 'settings') {
    if (!can(user, 'edit_prices') && !can(user, 'staff')) fail(403, 'คุณไม่มีสิทธิ์แก้ไขข้อมูลนี้');
    // role names and added roles (settings.roles) are the admin's to change
    if (kind === 'settings' && user.role !== 'admin') { if (prev && prev.roles) data.roles = prev.roles; else delete data.roles; }
  }
  return data;
}

async function putDoc(req, env, user, kind, id) {
  if (!KINDS.includes(kind)) fail(404, 'ไม่รู้จักประเภทข้อมูล');
  if (!id || id.length > 80) fail(400, 'รหัสไม่ถูกต้อง');
  const { data, version } = await readJson(req);
  if (!data || typeof data !== 'object') fail(400, 'ข้อมูลไม่ถูกต้อง');
  const row = await env.DB.prepare('SELECT * FROM docs WHERE kind = ? AND id = ?').bind(kind, id).first();
  const live = row && !row.deleted;
  if ((live ? row.version : 0) !== (Number(version) || 0)) {
    return json({ error: 'conflict', message: 'ข้อมูลนี้ถูกแก้ไขโดยผู้อื่น — โหลดเวอร์ชันล่าสุดแล้ว', doc: live ? { kind, id, version: row.version, data: JSON.parse(row.data) } : null }, 409);
  }
  const prev = live ? JSON.parse(row.data) : null;
  const out = await applyRules(env, user, kind, id, data, prev);
  const v = (row ? row.version : 0) + 1, at = now();
  await env.DB.prepare('INSERT INTO docs (kind, id, data, version, updated_at, updated_by, deleted) VALUES (?, ?, ?, ?, ?, ?, 0) ON CONFLICT(kind, id) DO UPDATE SET data = excluded.data, version = excluded.version, updated_at = excluded.updated_at, updated_by = excluded.updated_by, deleted = 0')
    .bind(kind, id, JSON.stringify(out), v, at, user.id).run();
  if (kind === 'beo' || kind === 'quote' || kind === 'task' || kind === 'leave') await audit(env, user, kind, id, prev ? 'update' : 'create', kind === 'beo' ? out.approval && out.approval.state : kind === 'task' || kind === 'leave' ? out.status : '');
  return json({ doc: { kind, id, version: v, data: out } });
}

async function deleteDoc(env, user, kind, id) {
  if (user.role !== 'admin' && !(RANK[user.role] >= RANK.manager)) fail(403, 'เฉพาะผู้จัดการขึ้นไปที่ลบข้อมูลได้');
  if ((kind === 'task' || kind === 'leave') && user.role !== 'admin') fail(403, 'เฉพาะผู้ดูแลระบบที่ลบข้อมูลนี้ได้');
  await env.DB.prepare('UPDATE docs SET deleted = 1, version = version + 1, updated_at = ?, updated_by = ? WHERE kind = ? AND id = ?').bind(now(), user.id, kind, id).run();
  await audit(env, user, kind, id, 'delete');
  return json({ ok: true });
}

async function decideBeo(req, env, user, id, decision) {
  const { comment } = await readJson(req);
  const row = await env.DB.prepare("SELECT * FROM docs WHERE kind = 'beo' AND id = ? AND deleted = 0").bind(id).first();
  if (!row) fail(404, 'ไม่พบใบฟังก์ชัน');
  const b = JSON.parse(row.data), a = b.approval || {};
  if (a.state !== 'pending') fail(409, 'ใบฟังก์ชันนี้ไม่ได้รออนุมัติ');
  const req2 = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(a.requestedBy).first();
  const reqRank = req2 ? RANK[publicUser(req2).role] || 1 : 1;
  if ((RANK[user.role] || 1) <= reqRank) fail(403, 'ต้องเป็นผู้มีตำแหน่งสูงกว่าผู้ขอจึงจะอนุมัติได้');
  if (decision === 'reject' && !String(comment || '').trim()) fail(400, 'กรุณาระบุเหตุผลที่ตีกลับ');
  b.approval = decision === 'approve'
    ? { state: 'approved', by: user.id, byName: user.name, at: now(), comment: String(comment || '').slice(0, 500), requestedBy: a.requestedBy, requestedByName: a.requestedByName }
    : Object.assign({}, a, { state: 'rejected', by: user.id, byName: user.name, at: now(), comment: String(comment).slice(0, 500) });
  const v = row.version + 1;
  await env.DB.prepare("UPDATE docs SET data = ?, version = ?, updated_at = ?, updated_by = ? WHERE kind = 'beo' AND id = ?").bind(JSON.stringify(b), v, now(), user.id, id).run();
  await audit(env, user, 'beo', id, decision, comment || '');
  return json({ doc: { kind: 'beo', id, version: v, data: b } });
}

async function audit(env, user, kind, id, action, note) {
  await env.DB.prepare('INSERT INTO audit (at, user_id, kind, doc_id, action, note) VALUES (?, ?, ?, ?, ?, ?)').bind(now(), user && user.id, kind, id, action, note || null).run();
}

/* ---------------- uploaded photos (menus / packages) ---------------- */
async function putImage(req, env, user) {
  if (!can(user, 'edit_prices')) fail(403, 'คุณไม่มีสิทธิ์อัปโหลดรูปเมนู');
  const { mime, data } = await readJson(req);
  if (!['image/jpeg', 'image/png', 'image/webp'].includes(mime)) fail(400, 'รองรับเฉพาะรูป JPG, PNG, WebP');
  if (typeof data !== 'string' || !/^[A-Za-z0-9+/=]+$/.test(data)) fail(400, 'ข้อมูลรูปไม่ถูกต้อง');
  if (data.length * 0.75 > MAX_IMAGE_BYTES) fail(413, 'รูปใหญ่เกินไป (สูงสุด 1.5 MB)');
  const id = randomHex(16);
  await env.DB.prepare('INSERT INTO images (id, mime, data, created_at, created_by) VALUES (?, ?, ?, ?, ?)').bind(id, mime, data, now(), user.id).run();
  await audit(env, user, 'image', id, 'upload');
  return json({ id, url: 'api/images/' + id });
}
async function getImage(env, id) {
  const row = await env.DB.prepare('SELECT mime, data FROM images WHERE id = ?').bind(id).first();
  if (!row) return new Response('Not found', { status: 404 });
  const bin = atob(row.data), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Response(bytes, { headers: { 'content-type': row.mime, 'cache-control': 'public, max-age=31536000, immutable', 'x-content-type-options': 'nosniff' } });
}

/* ---------------- users ---------------- */
async function putUser(req, env, user, id) {
  if (!can(user, 'users')) fail(403, 'คุณไม่มีสิทธิ์จัดการผู้ใช้');
  const { data, version, password } = await readJson(req);
  if (!data || !/.+@.+\..+/.test(String(data.email || ''))) fail(400, 'อีเมลไม่ถูกต้อง');
  if (password != null && (String(password).length < 6 || String(password).length > 128)) fail(400, 'รหัสผ่านต้องมี 6–128 ตัวอักษร');
  if (!['sales', 'manager', 'head', 'acc', 'admin'].includes(data.role)) fail(400, 'บทบาทไม่ถูกต้อง');
  if (data.role === 'admin' && user.role !== 'admin') fail(403, 'เฉพาะผู้ดูแลระบบที่ตั้งผู้ดูแลระบบได้');
  const row = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(id).first();
  const dupe = await env.DB.prepare('SELECT id FROM users WHERE email = ? AND id != ?').bind(data.email, id).first();
  if (dupe) fail(409, 'อีเมลนี้มีผู้ใช้อยู่แล้ว');
  if (id === user.id && !(data.perms && data.perms.users) && user.role !== 'admin') fail(400, 'ไม่สามารถถอดสิทธิ์จัดการผู้ใช้ของตัวเองได้');
  const clean = { name: String(data.name || '').slice(0, 100), role: data.role, roleId: /^r_[a-z0-9]{4,20}$/.test(String(data.roleId || '')) ? data.roleId : '', dept: String(data.dept || '').slice(0, 100), phone: String(data.phone || '').slice(0, 40),
    status: ['active', 'invited', 'disabled'].includes(data.status) ? data.status : 'invited', twoFA: !!data.twoFA,
    perms: Object.fromEntries(ALL_PERMS.map(p => [p, !!(data.perms && data.perms[p])])), lastLogin: row ? JSON.parse(row.data).lastLogin || null : null, lastDevice: row ? JSON.parse(row.data).lastDevice || '' : '' };
  if (row) {
    if (row.version !== Number(version)) fail(409, 'ข้อมูลผู้ใช้นี้ถูกแก้ไขโดยผู้อื่น กรุณาโหลดใหม่');
    await env.DB.prepare('UPDATE users SET email = ?, data = ?, version = version + 1, updated_at = ? WHERE id = ?').bind(data.email, JSON.stringify(clean), now(), id).run();
    if (clean.status === 'disabled') await env.DB.prepare('DELETE FROM sessions WHERE user_id = ?').bind(id).run();
  } else {
    const pw = await hashPassword(password != null ? String(password) : (env.DEFAULT_PASSWORD || 'luckswan'));
    await env.DB.prepare('INSERT INTO users (id, email, data, pw_hash, pw_salt, version, updated_at) VALUES (?, ?, ?, ?, ?, 1, ?)').bind(id, data.email, JSON.stringify(clean), pw.hash, pw.salt, now()).run();
  }
  await audit(env, user, 'user', id, row ? 'update' : 'create');
  const fresh = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(id).first();
  return json({ user: Object.assign(publicUser(fresh), { _v: fresh.version }) });
}

async function setPassword(env, id, password) {
  if (String(password || '').length < 6) fail(400, 'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
  if (String(password).length > 128) fail(400, 'รหัสผ่านยาวเกินไป');
  const pw = await hashPassword(String(password));
  await env.DB.prepare('UPDATE users SET pw_hash = ?, pw_salt = ?, updated_at = ? WHERE id = ?').bind(pw.hash, pw.salt, now(), id).run();
}

/* ---------------- setup & import ---------------- */
async function importData(env, user, data, opts = {}) {
  const stmts = [], at = now();
  const put = (kind, id, d) => stmts.push(env.DB.prepare('INSERT INTO docs (kind, id, data, version, updated_at, updated_by, deleted) VALUES (?, ?, ?, 1, ?, ?, 0) ON CONFLICT(kind, id) DO UPDATE SET data = excluded.data, version = docs.version + 1, updated_at = excluded.updated_at, updated_by = excluded.updated_by, deleted = 0')
    .bind(kind, String(id), JSON.stringify(d), at, user.id));
  for (const q of data.quotes || []) put('quote', q.id, q);
  for (const b of data.beos || []) put('beo', b.id, Object.assign({ approval: { state: 'approved', imported: true } }, b));
  for (const [k, v] of Object.entries(data.staff || {})) put('staff', k, v);
  if (Array.isArray(data.employees)) put('employees', 'all', { list: data.employees });
  if (data.settings) put('settings', 'main', data.settings);
  if (data.catalog && Array.isArray(data.catalog.food) && JSON.stringify(data.catalog).length < 800000) put('catalog', 'food', { list: data.catalog.food });
  if (opts.users) for (const u of data.users || []) {
    if (!u.email || !u.id) continue;
    const exists = await env.DB.prepare('SELECT id FROM users WHERE email = ? OR id = ?').bind(u.email, u.id).first();
    if (exists) continue;
    const { id, email, pw, ...rest } = u;
    const h = await hashPassword(env.DEFAULT_PASSWORD || 'luckswan');
    stmts.push(env.DB.prepare('INSERT INTO users (id, email, data, pw_hash, pw_salt, version, updated_at) VALUES (?, ?, ?, ?, ?, 1, ?)').bind(id, email, JSON.stringify(rest), h.hash, h.salt, at));
  }
  for (let i = 0; i < stmts.length; i += 50) await env.DB.batch(stmts.slice(i, i + 50));
  await audit(env, user, 'all', '-', 'import', `${(data.quotes || []).length} quotes, ${(data.beos || []).length} beos`);
  return stmts.length;
}

async function setup(req, env) {
  if (await userCount(env) > 0) fail(409, 'ระบบถูกตั้งค่าแล้ว');
  const { name, email, password, sample } = await readJson(req);
  if (!name || !/.+@.+\..+/.test(String(email || ''))) fail(400, 'กรุณากรอกชื่อและอีเมลให้ถูกต้อง');
  if (String(password || '').length < 8) fail(400, 'รหัสผ่านผู้ดูแลต้องมีอย่างน้อย 8 ตัวอักษร');
  const id = 'u1';
  const pw = await hashPassword(String(password));
  const profile = { name: String(name).slice(0, 100), role: 'admin', dept: 'ผู้บริหาร', phone: '', status: 'active', twoFA: false, perms: Object.fromEntries(ALL_PERMS.map(p => [p, true])), lastLogin: now() };
  const r = await env.DB.prepare('INSERT INTO users (id, email, data, pw_hash, pw_salt, version, updated_at) SELECT ?, ?, ?, ?, ?, 1, ? WHERE NOT EXISTS (SELECT 1 FROM users)')
    .bind(id, String(email).trim(), JSON.stringify(profile), pw.hash, pw.salt, now()).run();
  if (!r.meta || !r.meta.changes) fail(409, 'ระบบถูกตั้งค่าแล้ว');
  const user = Object.assign({ id, email }, profile);
  if (sample && typeof sample === 'object') await importData(env, user, sample, { users: false });
  return json({ token: await createSession(env, id, true), user });
}

/* ---------------- router ---------------- */
async function api(req, env, url) {
  const path = url.pathname.replace(/\/+$/, '');
  const m = req.method;
  if (path === '/api/health') return json({ ok: true, setup: (await userCount(env)) === 0, serverTime: now() });
  if (path === '/api/setup' && m === 'POST') return setup(req, env);
  if (path === '/api/login' && m === 'POST') return login(req, env);
  let ri;
  if ((ri = /^\/api\/images\/([0-9a-f]{32})$/.exec(path)) && m === 'GET') return getImage(env, ri[1]);

  const user = await authUser(req, env);
  if (path === '/api/logout' && m === 'POST') { await env.DB.prepare('DELETE FROM sessions WHERE token_hash = ?').bind(user._tokenHash).run(); return json({ ok: true }); }
  if (path === '/api/state' && m === 'GET') {
    const since = url.searchParams.get('since');
    if (since && !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(since)) fail(400, 'since ไม่ถูกต้อง');
    return json(Object.assign(await state(env, since, user), { me: user.id }));
  }
  let r;
  if ((r = /^\/api\/doc\/([a-z]+)\/(.+)$/.exec(path))) {
    const id = decodeURIComponent(r[2]);
    if (m === 'PUT') return putDoc(req, env, user, r[1], id);
    if (m === 'DELETE') return deleteDoc(env, user, r[1], id);
  }
  if ((r = /^\/api\/beo\/(.+)\/(approve|reject)$/.exec(path)) && m === 'POST') return decideBeo(req, env, user, decodeURIComponent(r[1]), r[2]);
  if ((r = /^\/api\/users\/([^/]+)$/.exec(path)) && m === 'PUT') return putUser(req, env, user, decodeURIComponent(r[1]));
  if ((r = /^\/api\/users\/([^/]+)$/.exec(path)) && m === 'DELETE') {
    if (user.role !== 'admin') fail(403, 'เฉพาะผู้ดูแลระบบที่ลบบัญชีผู้ใช้ได้');
    const id = decodeURIComponent(r[1]);
    if (id === user.id) fail(400, 'ไม่สามารถลบบัญชีของตัวเองได้');
    const target = await env.DB.prepare('SELECT id FROM users WHERE id = ?').bind(id).first();
    if (!target) fail(404, 'ไม่พบผู้ใช้');
    await env.DB.batch([env.DB.prepare('DELETE FROM sessions WHERE user_id = ?').bind(id), env.DB.prepare('DELETE FROM users WHERE id = ?').bind(id)]);
    await audit(env, user, 'user', id, 'delete');
    return json({ ok: true });
  }
  if ((r = /^\/api\/users\/([^/]+)\/password$/.exec(path)) && m === 'POST') {
    if (!can(user, 'users')) fail(403, 'คุณไม่มีสิทธิ์จัดการผู้ใช้');
    const { password } = await readJson(req);
    const id = decodeURIComponent(r[1]);
    const target = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(id).first();
    if (!target) fail(404, 'ไม่พบผู้ใช้');
    if (publicUser(target).role === 'admin' && user.role !== 'admin') fail(403, 'เฉพาะผู้ดูแลระบบที่เปลี่ยนรหัสผ่านของผู้ดูแลระบบได้');
    await setPassword(env, id, password);
    // sign the user out everywhere (keep the current session when admins change their own password)
    await env.DB.prepare('DELETE FROM sessions WHERE user_id = ? AND token_hash != ?').bind(id, user._tokenHash).run();
    await audit(env, user, 'user', id, 'reset-password');
    return json({ ok: true });
  }
  if (path === '/api/me/password' && m === 'POST') {
    const { current, next } = await readJson(req);
    const row = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(user.id).first();
    const chk = await hashPassword(String(current || ''), row.pw_salt);
    if (!safeEqual(chk.hash, row.pw_hash)) fail(400, 'รหัสผ่านปัจจุบันไม่ถูกต้อง');
    await setPassword(env, user.id, next);
    await env.DB.prepare('DELETE FROM sessions WHERE user_id = ? AND token_hash != ?').bind(user.id, user._tokenHash).run();
    return json({ ok: true });
  }
  if (path === '/api/images' && m === 'POST') return putImage(req, env, user);
  if (path === '/api/import' && m === 'POST') {
    if (user.role !== 'admin') fail(403, 'เฉพาะผู้ดูแลระบบที่นำเข้าข้อมูลได้');
    const { data, users } = await readJson(req);
    if (!data || !Array.isArray(data.quotes)) fail(400, 'ไฟล์สำรองไม่ถูกต้อง');
    return json({ ok: true, rows: await importData(env, user, data, { users: !!users }) });
  }
  if (path === '/api/audit' && m === 'GET') {
    const kind = url.searchParams.get('kind'), id = url.searchParams.get('id');
    const rows = (await env.DB.prepare('SELECT a.at, a.action, a.note, a.user_id FROM audit a WHERE a.kind = ? AND a.doc_id = ? ORDER BY a.id DESC LIMIT 50').bind(kind, id).all()).results;
    return json({ rows });
  }
  fail(404, 'ไม่พบ API นี้');
}

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    if (url.pathname.startsWith('/api/')) {
      try { return await api(req, env, url); }
      catch (e) {
        if (e instanceof HttpError) return json({ error: e.message }, e.status);
        console.error(e);
        return json({ error: 'เกิดข้อผิดพลาดของเซิร์ฟเวอร์' }, 500);
      }
    }
    return env.ASSETS.fetch(req);
  }
};
