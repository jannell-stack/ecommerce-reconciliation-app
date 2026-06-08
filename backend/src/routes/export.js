const express = require('express');
const {
  exportReconciliationBuffer,
  exportReconciliation2Buffer,
} = require('../services/exportExcel');

const router = express.Router();

function sendWorkbook(res, { buffer, filename }) {
  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(buffer);
}

router.get('/reconciliation', (_req, res) => {
  try {
    sendWorkbook(res, exportReconciliationBuffer());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/reconciliation-2', (_req, res) => {
  try {
    sendWorkbook(res, exportReconciliation2Buffer());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
