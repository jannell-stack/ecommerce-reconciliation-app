const express = require('express');
const { APP_PASSWORD } = require('../middleware/appAuth');

const router = express.Router();

router.post('/login', (req, res) => {
  const password = String(req.body?.password ?? '');
  if (password === APP_PASSWORD) {
    return res.json({ ok: true });
  }
  return res.status(401).json({ error: 'Invalid password' });
});

module.exports = router;
