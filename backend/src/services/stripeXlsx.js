const XLSX = require('xlsx');
const { stripeColDValue } = require('../utils/matchKey');

const DEFAULT_COL = {
  CURRENCY: 0,
  REPORTING_CATEGORY: 1,
  CREATED: 2,
  DESCRIPTION: 3,
  INVOICE_DATE: 4,
  AMT: 5,
  FEE: 6,
  NET: 7,
  RECEIVING_DATE: 8,
};

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function findStripeSheet(workbook, sheetName) {
  if (sheetName && workbook.SheetNames.includes(sheetName)) {
    return sheetName;
  }
  return (
    workbook.SheetNames.find((n) => /stripe advice/i.test(n)) ||
    workbook.SheetNames[0]
  );
}

function findColumnIndex(headerRow, matchers) {
  for (let i = 0; i < headerRow.length; i += 1) {
    const label = String(headerRow[i] ?? '').trim().toLowerCase();
    if (matchers.some((m) => label === m || label.includes(m))) {
      return i;
    }
  }
  return -1;
}

function resolveColumns(rows) {
  const headerRowIndex = rows.findIndex(
    (row) =>
      findColumnIndex(row, ['description']) >= 0 &&
      findColumnIndex(row, ['gross']) >= 0
  );

  if (headerRowIndex < 0) {
    return {
      col: DEFAULT_COL,
      dataStartRow: 1,
    };
  }

  const headerRow = rows[headerRowIndex];
  const descriptionCol = findColumnIndex(headerRow, ['description']);
  const grossCol = findColumnIndex(headerRow, ['gross']);
  const feeCol = findColumnIndex(headerRow, ['fee']);
  const netCol = findColumnIndex(headerRow, ['net']);
  const categoryCol = findColumnIndex(headerRow, ['reporting_category', 'category']);
  const currencyCol = findColumnIndex(headerRow, ['currency']);

  return {
    col: {
      CURRENCY: currencyCol >= 0 ? currencyCol : DEFAULT_COL.CURRENCY,
      DESCRIPTION: descriptionCol,
      AMT: grossCol,
      FEE: feeCol >= 0 ? feeCol : DEFAULT_COL.FEE,
      NET: netCol >= 0 ? netCol : DEFAULT_COL.NET,
      REPORTING_CATEGORY:
        categoryCol >= 0 ? categoryCol : DEFAULT_COL.REPORTING_CATEGORY,
    },
    dataStartRow: headerRowIndex + 1,
  };
}

function readStripeColD(cellValue) {
  const text = stripeColDValue(cellValue);
  return text || null;
}

function parseStripeXlsx(buffer, { sheetName } = {}) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const targetSheet = findStripeSheet(workbook, sheetName);
  const sheet = workbook.Sheets[targetSheet];

  if (!sheet) {
    throw new Error('No worksheet found in Stripe Excel file');
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const { col, dataStartRow } = resolveColumns(rows);
  const payments = [];
  let skipped = 0;
  let scanned = 0;

  for (let i = dataStartRow; i < rows.length; i += 1) {
    const row = rows[i];
    if (!row?.length) continue;

    scanned += 1;
    const description = readStripeColD(row[col.DESCRIPTION]);
    const amount = parseNumber(row[col.AMT]);

    if (!description || Number.isNaN(amount)) {
      skipped += 1;
      continue;
    }

    payments.push({
      description,
      match_key: description,
      amount,
      net: parseNumber(row[col.NET]),
      fee: parseNumber(row[col.FEE]),
      reporting_category:
        String(row[col.REPORTING_CATEGORY] ?? '').trim() || null,
      currency: String(row[col.CURRENCY] ?? '').trim() || null,
      source_row: i + 1,
    });
  }

  return {
    sheetName: targetSheet,
    payments,
    scanned,
    skipped,
  };
}

module.exports = {
  parseStripeXlsx,
  COL: DEFAULT_COL,
};
