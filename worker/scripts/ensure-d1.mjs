// Creates the D1 database on first deploy (if missing) and writes its id into wrangler.toml.
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';

const NAME = 'luckswan-sales';
const wrangler = args => execFileSync('npx', ['wrangler', ...args], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'inherit'] });
const find = () => {
  const out = wrangler(['d1', 'list', '--json']);
  const list = JSON.parse(out.slice(out.indexOf('[')));
  return list.find(d => d.name === NAME);
};

let db = find();
if (!db) {
  console.log(`Creating D1 database "${NAME}"…`);
  wrangler(['d1', 'create', NAME]);
  db = find();
}
if (!db || !db.uuid) throw new Error('Could not find or create the D1 database');
const toml = readFileSync('wrangler.toml', 'utf8').replace(/database_id = "[^"]*"/, `database_id = "${db.uuid}"`);
writeFileSync('wrangler.toml', toml);
console.log(`Using D1 database ${NAME} (${db.uuid})`);
