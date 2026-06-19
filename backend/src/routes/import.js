const express = require('express');
const multer = require('multer');
const { parseOrderXlsx } = require('../services/orderXlsx');
const { parseStripeXlsx } = require('../services/stripeXlsx');
const {
  insertOrders,
  insertPayments,
  clearOrders,
  clearPayments,
  insertStripePayments,
  clearStripePayments,
} = require('../db/repository');

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });

function shouldReplace(req) {
  return req.body?.replace === true || req.query.replace === 'true';
}

router.post('/order-xlsx', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Excel file required (field: file)' });
  }

  try {
    const paymentType = req.query.paymentType || req.body?.paymentType || 'STRIPE';
    const sheetName = req.query.sheet || req.body?.sheet;

    const parsed = parseOrderXlsx(req.file.buffer, { paymentType, sheetName });

    if (shouldReplace(req)) {
      clearOrders();
      clearPayments();
    }

    const ordersSynced = parsed.orders.length
      ? insertOrders(parsed.orders, { dedupe: true })
      : 0;
    const paymentsSynced = parsed.payments.length
      ? insertPayments(parsed.payments, { dedupe: true })
      : 0;

    res.json({
      message: `Imported ${ordersSynced} order(s) and ${paymentsSynced} payment(s) for ${parsed.paymentType}. Click "Run Reconciliation 1" to match.`,
      source: 'order-xlsx',
      sheetName: parsed.sheetName,
      paymentType: parsed.paymentType,
      mapping: {
        orderId: 'Column A (Ord No)',
        lineRef: 'Column C (DOC.NO) — F and H matched per row',
        orderAmount: 'Column F (AMT)',
        note: 'Column G (NOTE)',
        paymentAmount: 'Column H (Final rounded price paid)',
        comment: 'Column K (Comment)',
        filter: `Column E (PAYMENT TYPE) = ${parsed.paymentType}`,
      },
      ordersSynced,
      paymentsSynced,
      scanned: parsed.scanned,
      skipped: parsed.skipped,
      replaced: shouldReplace(req),
    });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

router.delete('/order-xlsx', (_req, res) => {
  clearOrders();
  clearPayments();
  res.json({ message: 'NC import data cleared (orders and payments)' });
});

router.delete('/stripe-xlsx', (_req, res) => {
  const cleared = clearStripePayments();
  res.json({
    message: `Stripe import data cleared (${cleared} row(s) removed)`,
    cleared,
  });
});

router.post('/stripe-xlsx', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Excel file required (field: file)' });
  }

  try {
    const sheetName = req.query.sheet || req.body?.sheet;
    const parsed = parseStripeXlsx(req.file.buffer, { sheetName });

    if (shouldReplace(req)) {
      clearStripePayments();
    }

    const stripeSynced = parsed.payments.length
      ? insertStripePayments(parsed.payments, { dedupe: true })
      : 0;

    res.json({
      message: `Imported ${stripeSynced} Stripe payment row(s) from ${parsed.sheetName}`,
      source: 'stripe-xlsx',
      sheetName: parsed.sheetName,
      mapping: {
        matchKey:
          'First 6 chars of Order column A (e.g. 864859-290426 → 864859) maps to Stripe column D',
        reconciliationAmount: 'Order column F (AMT) vs Stripe column F (gross)',
        net: 'Column H (net, reference only)',
        fee: 'Column G (fee, reference only)',
      },
      stripeSynced,
      scanned: parsed.scanned,
      skipped: parsed.skipped,
      replaced: shouldReplace(req),
    });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

module.exports = router;
