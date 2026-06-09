const express = require('express');
const multer = require('multer');
const {
  exportReconciliationBuffer,
  exportReconciliation2Buffer,
} = require('../services/exportExcel');
const { createEskPaymentWorkbook } = require('../services/eskPayment');

const router = express.Router();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

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

router.post(
  '/esk-payment',
  upload.fields([
    { name: 'skFile', maxCount: 1 },
    { name: 'ncFile', maxCount: 1 },
  ]),
  async (req, res) => {
    const skFile = req.files?.skFile?.[0];
    const ncFile = req.files?.ncFile?.[0];

    if (!skFile) {
      return res.status(400).json({ error: 'SK Combined file required (field: skFile)' });
    }
    if (!ncFile) {
      return res.status(400).json({ error: 'NC Payment Listing file required (field: ncFile)' });
    }

    try {
      const result = await createEskPaymentWorkbook(skFile.buffer, ncFile.buffer);
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="${result.filename}"`,
      );
      res.setHeader('X-Esk-Generated-Rows', String(result.generated));
      res.setHeader('X-Esk-Skipped-Rows', String(result.skipped));
      sendWorkbook(res, {
        buffer: result.buffer,
        filename: result.filename,
      });
    } catch (err) {
      console.error('[esk-payment]', err);
      res.status(500).json({
        error: err.message || 'Failed to create ESK-Payment file',
      });
    }
  },
);

module.exports = router;
