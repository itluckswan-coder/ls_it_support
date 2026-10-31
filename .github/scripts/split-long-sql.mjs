// D1 refuses any SQL statement longer than 100 KB (SQLITE_TOOBIG), and a D1 export writes each row as one
// INSERT, so a row holding a photo or a long log does not import. This rewrites every such INSERT as an
// INSERT with the long text values cut short, followed by UPDATEs that append the rest in pieces
// (… SET "col" = "col" || '…' WHERE rowid = last_insert_rowid()). Usage: node split-long-sql.mjs in.sql out.sql
import { readFileSync, writeFileSync } from 'node:fs';

const MAX = 90000, PIECE = 60000;
const sql = readFileSync(process.argv[2], 'utf8');

// split into statements at ; outside quotes ('' inside a string is a quote, "" inside a name likewise)
function statements(s){
  const out = []; let start = 0, q = null;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) { if (c === q) { if (s[i + 1] === q) i++; else q = null; } continue; }
    if (c === "'" || c === '"') q = c;
    else if (c === ';') { out.push(s.slice(start, i + 1)); start = i + 1; }
  }
  if (s.slice(start).trim()) out.push(s.slice(start));
  return out;
}
// the top-level parts of "( … , … )" starting at index i (the opening bracket); returns [parts, end index]
function list(s, i){
  const parts = []; let depth = 0, q = null, from = i + 1;
  for (let j = i; j < s.length; j++) {
    const c = s[j];
    if (q) { if (c === q) { if (s[j + 1] === q) j++; else q = null; } continue; }
    if (c === "'" || c === '"') q = c;
    else if (c === '(') depth++;
    else if (c === ')') { depth--; if (depth === 0) { parts.push(s.slice(from, j)); return [parts, j]; } }
    else if (c === ',' && depth === 1) { parts.push(s.slice(from, j)); from = j + 1; }
  }
  throw new Error('unbalanced statement');
}
// cut a quoted literal 'abc' into pieces of plain text, never between the two quotes of an escaped ''
function pieces(lit){
  const body = lit.slice(1, -1), out = [];
  for (let i = 0; i < body.length;) {
    let j = Math.min(body.length, i + PIECE);
    while (j < body.length && body[j - 1] === "'" && ((body.slice(i, j).match(/'+$/) || [''])[0].length % 2)) j++;
    out.push("'" + body.slice(i, j) + "'"); i = j;
  }
  return out;
}

let changed = 0, longest = 0;
const out = [];
for (const st of statements(sql)) {
  if (st.length <= MAX) { out.push(st); continue; }
  const m = /^\s*INSERT INTO\s+("[^"]+"|\w+)\s*/i.exec(st);
  if (!m) throw new Error('A statement over 90 KB is not an INSERT: ' + st.slice(0, 80));
  const table = m[1];
  let at = m[0].length, cols = null;
  if (st[at] === '(') { const [c, end] = list(st, at); cols = c.map(x => x.trim()); at = end + 1; }
  const v = /\s*VALUES\s*/i.exec(st.slice(at)); if (!v || v.index !== 0) throw new Error('Unexpected INSERT form in ' + table);
  at += v[0].length;
  const [vals, end] = list(st, at);
  if (!cols) throw new Error('INSERT without a column list in ' + table);
  const later = [];
  const shortVals = vals.map((x, i) => {
    const t = x.trim();
    if (t.length > PIECE && /^'.*'$/s.test(t)) { const p = pieces(t); later.push([cols[i], p.slice(1)]); return p[0]; }
    if (t.length > PIECE) throw new Error('A long value in ' + table + '.' + cols[i] + ' is not a plain text value');
    return t;
  });
  out.push('INSERT INTO ' + table + ' (' + cols.join(',') + ') VALUES(' + shortVals.join(',') + ')' + st.slice(end + 1));
  for (const [col, rest] of later) for (const p of rest) out.push('\nUPDATE ' + table + ' SET ' + col + ' = ' + col + ' || ' + p + ' WHERE rowid = last_insert_rowid();');
  changed++;
}
const text = out.join('');
for (const st of statements(text)) longest = Math.max(longest, st.length);
writeFileSync(process.argv[3], text);
console.log('Rows split into pieces: ' + changed + ' · longest statement now: ' + longest + ' characters');
if (longest > 100000) { console.log('::error::A statement is still over 100 KB'); process.exit(1); }
