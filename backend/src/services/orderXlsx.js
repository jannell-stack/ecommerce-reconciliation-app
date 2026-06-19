const XLSX = require('xlsx');
const { isValidOrderId } = require('../utils/csvParser');
const { parseOrderXlsxZip } = require('./orderXlsxZip');

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

const XLSX_READ_OPTS = {
  cellStyles: false,
  cellFormula: false,
  cellHTML: false,
  cellNF: false,
  dense: true,
};

function isZipXlsx(buffer) {
  return buffer?.length >= 2 && buffer[0] === 0x50 && buffer[1] === 0x4b;
}

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

function cellValue(cell) {
  if (cell == null || cell === '') return '';
  if (typeof cell === 'object' && 'v' in cell) return cell.v;
  return cell;
}

function lineRefFromRow(row, excelRowNumber) {
  const docNo = String(cellValue(row[COL.DOC_NO]) ?? '').trim();
  return docNo || `row-${excelRowNumber}`;
}

function resolveTargetSheet(buffer, sheetName) {
  const probe = XLSX.read(buffer, { type: 'buffer', bookSheets: true });
  if (sheetName && probe.SheetNames.includes(sheetName)) {
    return sheetName;
  }
  return (
    probe.SheetNames.find((name) => /payment listing/i.test(name)) ||
    probe.SheetNames[0]
  );
}

function parseOrderXlsxLegacy(buffer, { paymentType = 'STRIPE', sheetName } = {}) {
  const targetSheet = resolveTargetSheet(buffer, sheetName);
  const workbook = XLSX.read(buffer, {
    type: 'buffer',
    ...XLSX_READ_OPTS,
    sheet: targetSheet,
  });

  const sheet = workbook.Sheets[targetSheet];
  if (!sheet) {
    throw new Error('No worksheet found in Excel file');
  }

  const rows = Array.isArray(sheet)
    ? sheet
    : XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const filterType = normalizePaymentType(paymentType);

  const orders = [];
  const payments = [];
  let skipped = 0;
  let scanned = 0;

  for (let i = DATA_START_ROW; i < rows.length; i += 1) {
    const row = rows[i];
    if (!row?.length) continue;

    scanned += 1;
    const rowPaymentType = normalizePaymentType(cellValue(row[COL.PAYMENT_TYPE]));
    if (rowPaymentType !== filterType) continue;

    const order_id = String(cellValue(row[COL.ORDER_ID]) ?? '').trim();
    const orderAmount = parseNumber(cellValue(row[COL.AMT]));
    const paymentAmount = parseNumber(cellValue(row[COL.FINAL_PAID]));
    const line_ref = lineRefFromRow(row, i + 1);
    const note = String(cellValue(row[COL.NOTE]) ?? '').trim();
    const comment = String(cellValue(row[COL.COMMENT]) ?? '').trim();

    if (!isValidOrderId(order_id)) {
      skipped += 1;
      continue;
    }

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
    parser: 'xlsx',
  };
}

async function parseOrderXlsx(buffer, options = {}) {
  if (isZipXlsx(buffer)) {
    return parseOrderXlsxZip(buffer, options);
  }
  return parseOrderXlsxLegacy(buffer, options);
}

module.exports = {
  parseOrderXlsx,
  COL,
};
