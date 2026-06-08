const express = require('express');
const multer = require('multer');
const db = require('../db');
const { insertPayments, clearPayments } = require('../db/repository');
const { parseCsvRows } = require('../utils/csvParser');

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });

function normalizePaymentInput(body) {
  if (Array.isArray(body)) return body;
  if (body.payments && Array.isArray(body.payments)) return body.payments;
  if (body.order_id != null) return [body];
  return null;
}

router.get('/', (_req, res) => {
  const payments = db.prepare('SELECT * FROM payments ORDER BY id DESC').all();
  res.json({ payments, count: payments.length });
});

router.post('/', (req, res) => {
  const records = normalizePaymentInput(req.body);
  if (!records?.length) {
    return res.status(400).json({
      error: 'Provide { order_id, amount } or an array of payments',
    });
  }

  const invalid = records.find(
    (r) => !(r.order_id ?? r.orderId) || Number.isNaN(Number(r.amount))
  );
  if (invalid) {
    return res.status(400).json({ error: 'Each payment needs order_id and amount' });
  }

  const count = insertPayments(records);
  res.status(201).json({ message: `${count} payment(s) added`, count });
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
    const count = insertPayments(rows);
    res.status(201).json({ message: `${count} payment(s) uploaded`, count });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

router.delete('/', (_req, res) => {
  clearPayments();
  res.json({ message: 'All payments cleared' });
});

module.exports = router;
