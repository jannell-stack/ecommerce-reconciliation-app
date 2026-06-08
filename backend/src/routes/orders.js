const express = require('express');
const multer = require('multer');
const db = require('../db');
const { insertOrders, clearOrders } = require('../db/repository');
const { parseCsvRows } = require('../utils/csvParser');

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });

function normalizeOrderInput(body) {
  if (Array.isArray(body)) return body;
  if (body.orders && Array.isArray(body.orders)) return body.orders;
  if (body.order_id != null) return [body];
  return null;
}

router.get('/', (_req, res) => {
  const orders = db.prepare('SELECT * FROM orders ORDER BY id DESC').all();
  res.json({ orders, count: orders.length });
});

router.post('/', (req, res) => {
  const records = normalizeOrderInput(req.body);
  if (!records?.length) {
    return res.status(400).json({
      error: 'Provide { order_id, amount } or an array of orders',
    });
  }

  const invalid = records.find(
    (r) => !(r.order_id ?? r.orderId) || Number.isNaN(Number(r.amount))
  );
  if (invalid) {
    return res.status(400).json({ error: 'Each order needs order_id and amount' });
  }

  const count = insertOrders(records);
  res.status(201).json({ message: `${count} order(s) added`, count });
});

router.post('/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'CSV file required (field: file)' });
  }

  try {
    const rows = parseCsvRows(req.file.buffer);
    if (!rows.length) {
      return res.status(400).json({
        error:
          'No valid rows found. Use a UTF-8 CSV with columns order_id, amount (not Excel .xlsx).',
      });
    }
    const count = insertOrders(rows);
    res.status(201).json({ message: `${count} order(s) uploaded`, count });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

router.delete('/', (_req, res) => {
  clearOrders();
  res.json({ message: 'All orders cleared' });
});

module.exports = router;
