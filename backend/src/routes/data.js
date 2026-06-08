const express = require('express');
const {
  clearOrders,
  clearPayments,
  clearStripePayments,
} = require('../db/repository');
const db = require('../db');
const { isValidOrderId } = require('../utils/csvParser');

const router = express.Router();

router.get('/health', (_req, res) => {
  const orders = db.prepare('SELECT order_id FROM orders').all();
  const payments = db.prepare('SELECT order_id FROM payments').all();
  const stripePayments = db.prepare('SELECT COUNT(*) AS c FROM stripe_payments').get().c;

  const badOrders = orders.filter((r) => !isValidOrderId(r.order_id)).length;
  const badPayments = payments.filter((r) => !isValidOrderId(r.order_id)).length;

  res.json({
    orderCount: orders.length,
    paymentCount: payments.length,
    stripePaymentCount: stripePayments,
    corruptedOrders: badOrders,
    corruptedPayments: badPayments,
    dataHealthy: badOrders === 0 && badPayments === 0,
  });
});

router.post('/reset', (_req, res) => {
  clearOrders();
  clearPayments();
  clearStripePayments();
  res.json({ message: 'All orders, payments, and Stripe rows have been cleared' });
});

module.exports = router;
