const express = require('express');
const Database = require('better-sqlite3');
const crypto = require('crypto');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;

// --- Database setup ---

const db = new Database(path.join(__dirname, 'urls.db'));
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
  CREATE TABLE IF NOT EXISTS urls (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    short_code  TEXT    UNIQUE NOT NULL,
    original    TEXT    NOT NULL,
    created_at  TEXT    DEFAULT (datetime('now')),
    clicks      INTEGER DEFAULT 0,
    last_click  TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_short_code ON urls(short_code);
`);

// --- Base62 short code generator ---

const BASE62 = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';

function generateShortCode(length = 7) {
  const bytes = crypto.randomBytes(length);
  let code = '';
  for (let i = 0; i < length; i++) {
    code += BASE62[bytes[i] % 62];
  }
  return code;
}

// --- Prepared statements ---

const insertUrl = db.prepare(
  'INSERT INTO urls (short_code, original) VALUES (?, ?)'
);
const findByCode = db.prepare(
  'SELECT * FROM urls WHERE short_code = ?'
);
const findByOriginal = db.prepare(
  'SELECT * FROM urls WHERE original = ?'
);
const incrementClicks = db.prepare(
  `UPDATE urls SET clicks = clicks + 1, last_click = datetime('now') WHERE short_code = ?`
);
const recentUrls = db.prepare(
  'SELECT * FROM urls ORDER BY created_at DESC LIMIT ?'
);
const totalStats = db.prepare(
  'SELECT COUNT(*) as total_urls, COALESCE(SUM(clicks), 0) as total_clicks FROM urls'
);

// --- Middleware ---

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// --- API Routes ---

app.post('/api/shorten', (req, res) => {
  let { url, customCode } = req.body;

  if (!url) {
    return res.status(400).json({ error: 'URL is required' });
  }

  if (!/^https?:\/\//i.test(url)) {
    url = 'https://' + url;
  }

  try {
    new URL(url);
  } catch {
    return res.status(400).json({ error: 'Invalid URL format' });
  }

  const existing = findByOriginal.get(url);
  if (existing && !customCode) {
    return res.json({
      shortUrl: `${BASE_URL}/${existing.short_code}`,
      shortCode: existing.short_code,
      original: existing.original,
      created: existing.created_at,
      isExisting: true,
    });
  }

  let code = customCode || generateShortCode();

  if (customCode) {
    if (!/^[a-zA-Z0-9_-]{3,30}$/.test(customCode)) {
      return res.status(400).json({
        error: 'Custom code must be 3-30 alphanumeric characters, hyphens, or underscores',
      });
    }
    if (findByCode.get(customCode)) {
      return res.status(409).json({ error: 'Custom code is already taken' });
    }
  }

  let attempts = 0;
  while (attempts < 5) {
    try {
      insertUrl.run(code, url);
      return res.status(201).json({
        shortUrl: `${BASE_URL}/${code}`,
        shortCode: code,
        original: url,
        created: new Date().toISOString(),
      });
    } catch (err) {
      if (err.code === 'SQLITE_CONSTRAINT_UNIQUE' && !customCode) {
        code = generateShortCode();
        attempts++;
      } else {
        throw err;
      }
    }
  }

  res.status(500).json({ error: 'Failed to generate unique code' });
});

app.get('/api/stats/:code', (req, res) => {
  const row = findByCode.get(req.params.code);
  if (!row) {
    return res.status(404).json({ error: 'Short URL not found' });
  }
  res.json({
    shortCode: row.short_code,
    shortUrl: `${BASE_URL}/${row.short_code}`,
    original: row.original,
    clicks: row.clicks,
    created: row.created_at,
    lastClick: row.last_click,
  });
});

app.get('/api/recent', (req, res) => {
  const limit = Math.min(parseInt(req.query.limit) || 20, 100);
  const rows = recentUrls.all(limit);
  const stats = totalStats.get();
  res.json({
    urls: rows.map((r) => ({
      shortCode: r.short_code,
      shortUrl: `${BASE_URL}/${r.short_code}`,
      original: r.original,
      clicks: r.clicks,
      created: r.created_at,
      lastClick: r.last_click,
    })),
    totalUrls: stats.total_urls,
    totalClicks: stats.total_clicks,
  });
});

// --- Redirect ---

app.get('/:code', (req, res) => {
  const code = req.params.code;

  if (code.includes('.') || code === 'api') {
    return res.status(404).send('Not found');
  }

  const row = findByCode.get(code);
  if (!row) {
    return res.status(404).sendFile(path.join(__dirname, 'public', 'index.html'));
  }

  incrementClicks.run(code);
  res.redirect(301, row.original);
});

// --- Start ---

app.listen(PORT, () => {
  console.log(`URL Shortener running at ${BASE_URL}`);
});
