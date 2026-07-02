const express = require('express');
const multer = require('multer');
const {
  exportReconciliationBuffer,
  exportReconciliation2Buffer,
} = require('../services/exportExcel');
const { createEskPaymentWorkbook } = require('../services/eskPayment');
const { createStripeAdviceWorkbook } = require('../services/stripeAdvice');
const { createGpayWorkbook } = require('../services/gpay');
const { createShopbackWorkbook } = require('../services/shopback');

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

    const startedAt = Date.now();
    const memBefore = process.memoryUsage();

    console.info('[esk-payment] start', {
      skMB: Math.round(skFile.buffer.length / 1024 / 1024),
      ncMB: Math.round(ncFile.buffer.length / 1024 / 1024),
      heapLimitMB: Math.round((require('v8').getHeapStatistics().heap_size_limit) / 1024 / 1024),
    });

    try {
      const result = await createEskPaymentWorkbook(skFile.buffer, ncFile.buffer, {
        ncFileName: ncFile.originalname,
      });
      const memAfter = process.memoryUsage();
      console.info('[esk-payment] ok', {
        generated: result.generated,
        durationMs: Date.now() - startedAt,
        heapUsedMB: Math.round(memAfter.heapUsed / 1024 / 1024),
        heapDeltaMB: Math.round((memAfter.heapUsed - memBefore.heapUsed) / 1024 / 1024),
        outputMB: Math.round(result.buffer.length / 1024 / 1024),
      });
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
      console.error('[esk-payment] failed', {
        durationMs: Date.now() - startedAt,
        message: err.message,
        stack: err.stack,
      });
      res.status(500).json({
        error: err.message || 'Failed to create ESK-Payment file',
      });
    }
  },
);

router.post(
  '/stripe-advice',
  upload.fields([
    { name: 'stripeFile', maxCount: 1 },
    { name: 'ncFile', maxCount: 12 },
  ]),
  async (req, res) => {
    const stripeFile = req.files?.stripeFile?.[0];
    const ncFiles = req.files?.ncFile || [];

    if (!stripeFile) {
      return res.status(400).json({
        error: 'Stripe Itemized_payout_reconciliation file required (field: stripeFile)',
      });
    }
    if (ncFiles.length === 0) {
      return res.status(400).json({
        error: 'At least one NC Payment Listing file required (field: ncFile)',
      });
    }

    const startedAt = Date.now();

    try {
      const result = await createStripeAdviceWorkbook(
        stripeFile.buffer,
        ncFiles.map((file) => ({ buffer: file.buffer, fileName: file.originalname })),
        { stripeFileName: stripeFile.originalname },
      );
      console.info('[stripe-advice] ok', {
        generated: result.generated,
        linked: result.linked,
        ncFiles: result.ncFileNames,
        durationMs: Date.now() - startedAt,
        outputMB: Math.round(result.buffer.length / 1024 / 1024),
      });
      res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"`);
      res.setHeader('X-Stripe-Advice-Generated-Rows', String(result.generated));
      res.setHeader('X-Stripe-Advice-Linked-Rows', String(result.linked));
      sendWorkbook(res, {
        buffer: result.buffer,
        filename: result.filename,
      });
    } catch (err) {
      console.error('[stripe-advice] failed', {
        durationMs: Date.now() - startedAt,
        message: err.message,
        stack: err.stack,
      });
      res.status(500).json({
        error: err.message || 'Failed to create Stripe Advice file',
      });
    }
  },
);

router.post(
  '/gpay',
  upload.fields([
    { name: 'grabFile', maxCount: 1 },
    { name: 'eskFile', maxCount: 12 },
  ]),
  async (req, res) => {
    const grabFile = req.files?.grabFile?.[0];
    const eskFiles = req.files?.eskFile || [];

    if (!grabFile) {
      return res.status(400).json({
        error: 'Grab Transaction file required (field: grabFile)',
      });
    }
    if (eskFiles.length === 0) {
      return res.status(400).json({
        error: 'At least one ESK Payment Details file required (field: eskFile)',
      });
    }

    const startedAt = Date.now();

    try {
      const result = await createGpayWorkbook(
        grabFile.buffer,
        eskFiles.map((file) => ({ buffer: file.buffer, fileName: file.originalname })),
        { grabFileName: grabFile.originalname },
      );
      console.info('[gpay] ok', {
        generated: result.generated,
        linked: result.linked,
        eskFiles: result.ncFileNames,
        durationMs: Date.now() - startedAt,
        outputMB: Math.round(result.buffer.length / 1024 / 1024),
      });
      res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"`);
      res.setHeader('X-GPay-Generated-Rows', String(result.generated));
      res.setHeader('X-GPay-Linked-Rows', String(result.linked));
      sendWorkbook(res, {
        buffer: result.buffer,
        filename: result.filename,
      });
    } catch (err) {
      console.error('[gpay] failed', {
        durationMs: Date.now() - startedAt,
        message: err.message,
        stack: err.stack,
      });
      res.status(500).json({
        error: err.message || 'Failed to create GPay file',
      });
    }
  },
);

router.post(
  '/shopback',
  upload.fields([
    { name: 'shopbackFile', maxCount: 1 },
    { name: 'eskFile', maxCount: 12 },
  ]),
  async (req, res) => {
    const shopbackFile = req.files?.shopbackFile?.[0];
    const eskFiles = req.files?.eskFile || [];

    if (!shopbackFile) {
      return res.status(400).json({
        error: 'Shopback Transaction file required (field: shopbackFile)',
      });
    }
    if (eskFiles.length === 0) {
      return res.status(400).json({
        error: 'At least one ESK Payment Details file required (field: eskFile)',
      });
    }

    const startedAt = Date.now();

    try {
      const result = await createShopbackWorkbook(
        shopbackFile.buffer,
        eskFiles.map((file) => ({ buffer: file.buffer, fileName: file.originalname })),
        { shopbackFileName: shopbackFile.originalname },
      );
      console.info('[shopback] ok', {
        generated: result.generated,
        linked: result.linked,
        eskFiles: result.ncFileNames,
        durationMs: Date.now() - startedAt,
        outputMB: Math.round(result.buffer.length / 1024 / 1024),
      });
      res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"`);
      res.setHeader('X-Shopback-Generated-Rows', String(result.generated));
      res.setHeader('X-Shopback-Linked-Rows', String(result.linked));
      sendWorkbook(res, {
        buffer: result.buffer,
        filename: result.filename,
      });
    } catch (err) {
      console.error('[shopback] failed', {
        durationMs: Date.now() - startedAt,
        message: err.message,
        stack: err.stack,
      });
      res.status(500).json({
        error: err.message || 'Failed to create Shopback file',
      });
    }
  },
);

module.exports = router;
