// Account recovery when no administrator can sign in (run by the "Reset admin account" workflow).
// Makes ADMIN_RESET_EMAIL an active administrator with every permission and ADMIN_RESET_PASSWORD as its password:
// creates the account if it does not exist, otherwise updates it (its other data is kept), and clears its failed-login lock.
// The password is hashed exactly as worker/src/index.js does. Nothing about the users is printed: the repository is public.
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { webcrypto as crypto } from 'node:crypto';

const NAME = 'luckswan-sales';
const WHERE = (process.env.D1_WHERE || '--remote').split(' '); // e.g. "--local --persist-to dir" for testing
const ALL_PERMS = ['create_quote', 'discount10', 'discount_over', 'view_beo', 'customers', 'staff', 'users', 'edit_prices', 'leave_hr'];
const PBKDF2_ITER = 10000;

const email = String(process.env.ADMIN_RESET_EMAIL || '').trim().toLowerCase();
const password = String(process.env.ADMIN_RESET_PASSWORD || '');
if (!/^[^\s@'"]+@[^\s@'"]+\.[^\s@'"]+$/.test(email)) throw new Error('ADMIN_RESET_EMAIL is not an e-mail address');
if (password.length < 8 || password.length > 128) throw new Error('ADMIN_RESET_PASSWORD must be 8–128 characters');

const hex = buf => Array.from(new Uint8Array(buf), b => b.toString(16).padStart(2, '0')).join('');
const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveBits']);
const hash = hex(await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: new TextEncoder().encode(salt), iterations: PBKDF2_ITER }, key, 256));
const now = new Date(Date.now() + 7 * 3600e3).toISOString().slice(0, 19);
const q = s => "'" + String(s).replace(/'/g, "''") + "'";
const perms = JSON.stringify(Object.fromEntries(ALL_PERMS.map(p => [p, true])));
const profile = JSON.stringify({ name: email.split('@')[0], role: 'admin', roleId: '', dept: 'ผู้บริหาร', phone: '', status: 'active', twoFA: false, perms: JSON.parse(perms), lastLogin: null, lastDevice: '', sites: [] });
const id = 'u' + hex(crypto.getRandomValues(new Uint8Array(6)));

const sql = `
INSERT INTO users (id, email, data, pw_hash, pw_salt, version, updated_at)
  SELECT ${q(id)}, ${q(email)}, ${q(profile)}, ${q(hash)}, ${q(salt)}, 1, ${q(now)}
  WHERE NOT EXISTS (SELECT 1 FROM users WHERE email = ${q(email)} COLLATE NOCASE);
UPDATE users SET pw_hash = ${q(hash)}, pw_salt = ${q(salt)}, version = version + 1, updated_at = ${q(now)},
  data = json_set(data, '$.role', 'admin', '$.status', 'active', '$.perms', json(${q(perms)}))
  WHERE email = ${q(email)} COLLATE NOCASE AND id != ${q(id)};
DELETE FROM login_fail WHERE email IN (${q(email)}, ${q(email.split('@')[0] + '@')});
`;

const dir = mkdtempSync(join(tmpdir(), 'reset-'));
const file = join(dir, 'reset.sql');
try {
  writeFileSync(file, sql, { mode: 0o600 });
  // the SQL holds the password hash: keep wrangler's echo of it out of the (public) log
  execFileSync('npx', ['wrangler', 'd1', 'execute', NAME, ...WHERE, '--file', file, '--yes'], { stdio: ['ignore', 'ignore', 'pipe'] });
} finally {
  rmSync(dir, { recursive: true, force: true });
}
console.log('Done: the account is an active administrator with the new password. Sign in with it now.');
