const db = require('../db');

db.exec(`
  CREATE TABLE IF NOT EXISTS sync_meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
  );
`);

function getLastSync(key) {
  const row = db
    .prepare('SELECT value FROM sync_meta WHERE key = ?')
    .get(`last_${key}_sync_at`);
  return row?.value ?? null;
}

function setLastSync(key, isoTimestamp) {
  db.prepare(
    `INSERT INTO sync_meta (key, value) VALUES (?, ?)
     ON CONFLICT(key) DO UPDATE SET value = excluded.value`
  ).run(`last_${key}_sync_at`, isoTimestamp);
}

function getDefaultSince() {
  const since = new Date();
  since.setHours(since.getHours() - 24);
  return since.toISOString();
}

module.exports = {
  getLastSync,
  setLastSync,
  getDefaultSince,
};
