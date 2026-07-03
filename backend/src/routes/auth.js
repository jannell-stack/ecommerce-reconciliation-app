const express = require('express');
const { APP_PASSWORD } = require('../middleware/appAuth');
const {
  clearOrders,
  clearPayments,
  clearStripePayments,
} = require('../db/repository');

const router = express.Router();

router.post('/login', (req, res) => {
  const password = String(req.body?.password ?? '');
  if (password === APP_PASSWORD) {
    clearOrders();
    clearPayments();
    clearStripePayments();
    return res.json({ ok: true, dataCleared: true });
  }
  return res.status(401).json({ error: 'Invalid password' });
});

module.exports = router;
