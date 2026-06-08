const XLSX = require('xlsx');
const { isValidOrderId } = require('../utils/csvParser');

const COL = {
  ORDER_ID: 0,
  DOC_NO: 2,
  PAYMENT_TYPE: 4,
  AMT: 5,
  NOTE: 6,
  FINAL_PAID: 7,
  COMMENT: 10,
};

const DATA_START_ROW = 2;

function normalizePaymentType(value) {
  return String(value ?? '')
    .trim()
    .toUpperCase();
}

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function lineRefFromRow(row, excelRowNumber) {
  const docNo = String(row[COL.DOC_NO] ?? '').trim();
  return docNo || `row-${excelRowNumber}`;
}

function parseOrderXlsx(buffer, { paymentType = 'STRIPE', sheetName } = {}) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const targetSheet =
    sheetName && workbook.SheetNames.includes(sheetName)
      ? sheetName
      : workbook.SheetNames.find((n) => /payment listing/i.test(n)) ||
        workbook.SheetNames[0];

  const sheet = workbook.Sheets[targetSheet];
  if (!sheet) {
    throw new Error('No worksheet found in Excel file');
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const filterType = normalizePaymentType(paymentType);

  const orders = [];
  const payments = [];
  let skipped = 0;
  let scanned = 0;

  for (let i = DATA_START_ROW; i < rows.length; i += 1) {
    const row = rows[i];
    if (!row?.length) continue;

    scanned += 1;
    const rowPaymentType = normalizePaymentType(row[COL.PAYMENT_TYPE]);
    if (rowPaymentType !== filterType) continue;

    const order_id = String(row[COL.ORDER_ID] ?? '').trim();
    const orderAmount = parseNumber(row[COL.AMT]);
    const paymentAmount = parseNumber(row[COL.FINAL_PAID]);
    const line_ref = lineRefFromRow(row, i + 1);
    const note = String(row[COL.NOTE] ?? '').trim();
    const comment = String(row[COL.COMMENT] ?? '').trim();

    if (!isValidOrderId(order_id)) {
      skipped += 1;
      continue;
    }

    // Reconcile column F and H on the same Excel row (linked by DOC.NO / line_ref).
    if (!Number.isNaN(orderAmount)) {
      orders.push({
        order_id,
        amount: orderAmount,
        line_ref,
        note: note || null,
        comment: comment || null,
        sourceRow: i + 1,
      });
    }

    if (!Number.isNaN(paymentAmount)) {
      payments.push({
        order_id,
        amount: paymentAmount,
        line_ref,
        note: note || null,
        comment: comment || null,
        sourceRow: i + 1,
      });
    }

    if (Number.isNaN(orderAmount) && Number.isNaN(paymentAmount)) {
      skipped += 1;
    }
  }

  return {
    sheetName: targetSheet,
    paymentType: filterType,
    orders,
    payments,
    scanned,
    skipped,
  };
}

module.exports = {
  parseOrderXlsx,
  COL,
};
