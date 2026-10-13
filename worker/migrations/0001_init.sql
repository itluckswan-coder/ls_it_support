-- Luck Swan Sales — central database (Cloudflare D1)
CREATE TABLE IF NOT EXISTS users (
  id         TEXT PRIMARY KEY,
  email      TEXT NOT NULL UNIQUE COLLATE NOCASE,
  data       TEXT NOT NULL,            -- profile JSON: name, role, dept, phone, perms, status, twoFA, lastLogin…
  pw_hash    TEXT,
  pw_salt    TEXT,
  version    INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS sessions (
  token_hash TEXT PRIMARY KEY,
  user_id    TEXT NOT NULL,
  expires_at INTEGER NOT NULL,
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS sessions_user ON sessions(user_id);

-- quotes, function sheets (beo), staff plans, employees, settings
CREATE TABLE IF NOT EXISTS docs (
  kind       TEXT NOT NULL,
  id         TEXT NOT NULL,
  data       TEXT NOT NULL,
  version    INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL,
  updated_by TEXT,
  deleted    INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (kind, id)
);
CREATE INDEX IF NOT EXISTS docs_updated ON docs(updated_at);

CREATE TABLE IF NOT EXISTS audit (
  id      INTEGER PRIMARY KEY AUTOINCREMENT,
  at      TEXT NOT NULL,
  user_id TEXT,
  kind    TEXT,
  doc_id  TEXT,
  action  TEXT NOT NULL,
  note    TEXT
);
CREATE INDEX IF NOT EXISTS audit_doc ON audit(kind, doc_id);

CREATE TABLE IF NOT EXISTS login_fail (
  email TEXT PRIMARY KEY COLLATE NOCASE,
  count INTEGER NOT NULL,
  first_at INTEGER NOT NULL
);
