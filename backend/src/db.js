const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

const dataDir = process.env.DATA_DIR || path.join(__dirname, '..', 'data');
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
}

const dbPath = path.join(dataDir, 'reconciliation.db');
const db = new Database(dbPath);

db.exec(`
  CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id TEXT NOT NULL,
    amount REAL NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id TEXT NOT NULL,
    amount REAL NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE INDEX IF NOT EXISTS idx_orders_order_id ON orders(order_id);
  CREATE INDEX IF NOT EXISTS idx_payments_order_id ON payments(order_id);
`);

function ensureColumn(table, column, type) {
  const cols = db.prepare(`PRAGMA table_info(${table})`).all();
  if (!cols.some((c) => c.name === column)) {
    db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`);
  }
}

ensureColumn('orders', 'line_ref', 'TEXT');
ensureColumn('payments', 'line_ref', 'TEXT');
ensureColumn('orders', 'note', 'TEXT');
ensureColumn('orders', 'comment', 'TEXT');
ensureColumn('payments', 'note', 'TEXT');
ensureColumn('payments', 'comment', 'TEXT');

db.exec(`
  CREATE INDEX IF NOT EXISTS idx_orders_line_ref ON orders(line_ref);
  CREATE INDEX IF NOT EXISTS idx_payments_line_ref ON payments(line_ref);

  CREATE TABLE IF NOT EXISTS stripe_payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    description TEXT NOT NULL,
    match_key TEXT NOT NULL,
    amount REAL NOT NULL,
    net REAL,
    fee REAL,
    reporting_category TEXT,
    currency TEXT,
    source_row INTEGER,
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE INDEX IF NOT EXISTS idx_stripe_payments_match_key ON stripe_payments(match_key);
`);

module.exports = db;
