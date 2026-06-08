const express = require('express');
const {
  syncOrdersFromShopify,
  syncPaymentsFromStripe,
} = require('../services/syncRunner');

const router = express.Router();

function shouldReplace(req) {
  return req.body?.replace === true || req.query.replace === 'true';
}

router.post('/orders', async (req, res) => {
  try {
    const result = await syncOrdersFromShopify({
      replace: shouldReplace(req),
    });

    res.json({
      ...result,
      message: `Synced ${result.synced} order(s) from Shopify`,
    });
  } catch (err) {
    const status = err.message?.includes('Missing required') ? 503 : 502;
    res.status(status).json({ error: err.message });
  }
});

router.post('/payments', async (req, res) => {
  try {
    const result = await syncPaymentsFromStripe({
      replace: shouldReplace(req),
    });

    res.json({
      ...result,
      message: `Synced ${result.synced} payment(s) from Stripe`,
    });
  } catch (err) {
    const status = err.message?.includes('Missing required') ? 503 : 502;
    res.status(status).json({ error: err.message });
  }
});

module.exports = router;
