const express = require('express');
const { runReconciliation } = require('../services/reconciliation');
const { runReconciliation2 } = require('../services/reconciliation2');

const router = express.Router();

router.get('/', (_req, res) => {
  try {
    const result = runReconciliation();
    res.json(result);
  } catch (err) {
    console.error('[reconcile] GET failed', err);
    res.status(500).json({ error: err.message || 'Reconciliation failed' });
  }
});

router.post('/', (_req, res) => {
  try {
    const result = runReconciliation();
    const { rows, ...payload } = result;
    res.json({
      message: 'Reconciliation complete',
      ...payload,
    });
  } catch (err) {
    console.error('[reconcile] POST failed', err);
    res.status(500).json({ error: err.message || 'Reconciliation failed' });
  }
});

router.get('/2', (_req, res) => {
  try {
    const result = runReconciliation2();
    res.json(result);
  } catch (err) {
    console.error('[reconcile] GET /2 failed', err);
    res.status(500).json({ error: err.message || 'Reconciliation 2 failed' });
  }
});

router.post('/2', (_req, res) => {
  try {
    const result = runReconciliation2();
    const { rows, ...payload } = result;
    res.json({
      message:
        'Reconciliation 2 complete (Stripe col D vs Order col A prefix; amounts Order col F vs Stripe col F)',
      ...payload,
    });
  } catch (err) {
    console.error('[reconcile] POST /2 failed', err);
    res.status(500).json({ error: err.message || 'Reconciliation 2 failed' });
  }
});

module.exports = router;
