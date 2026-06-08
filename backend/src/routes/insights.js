const express = require('express');
const { getInsights } = require('../services/insights');

const router = express.Router();

router.get('/', async (_req, res) => {
  try {
    const result = await getInsights();
    res.json(result);
  } catch (err) {
    const status = err.message?.includes('OPENAI_API_KEY') ? 503 : 502;
    res.status(status).json({ error: err.message });
  }
});

module.exports = router;
