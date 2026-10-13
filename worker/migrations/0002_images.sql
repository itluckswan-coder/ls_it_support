-- photos uploaded for catering menus / packages (base64 JPEG/PNG/WebP, max 1.5 MB each)
CREATE TABLE IF NOT EXISTS images (
  id         TEXT PRIMARY KEY,
  mime       TEXT NOT NULL,
  data       TEXT NOT NULL,
  created_at TEXT NOT NULL,
  created_by TEXT
);
