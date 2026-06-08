const express = require('express');
const { runReconciliation } = require('../services/reconciliation');
const { runReconciliation2 } = require('../services/reconciliation2');

const router = express.Router();

router.get('/', (_req, res) => {
  const result = runReconciliation();
  res.json(result);
});

router.post('/', (_req, res) => {
  const result = runReconciliation();
  res.json({
    message: 'Reconciliation complete',
    ...result,
  });
});

router.get('/2', (_req, res) => {
  const result = runReconciliation2();
  res.json(result);
});

router.post('/2', (_req, res) => {
  const result = runReconciliation2();
  res.json({
    message:
      'Reconciliation 2 complete (Stripe col D vs Order col A prefix; amounts Order col F vs Stripe col F)',
    ...result,
  });
});

module.exports = router;
