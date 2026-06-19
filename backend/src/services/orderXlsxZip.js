const JSZip = require('jszip');
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

const DATA_START_ROW = 3;
const NEEDED_COLS = new Set(Object.values(COL));

function decodeXml(text) {
  return String(text ?? '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function colLetterToIndex(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
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

function parseSharedStrings(xml) {
  if (!xml) return [];
  const strings = [];
  for (const si of xml.match(/<si>[\s\S]*?<\/si>/g) || []) {
    const parts = [...si.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) =>
      decodeXml(m[1])
    );
    strings.push(parts.join(''));
  }
  return strings;
}

function parseCellValue(cellXml, sharedStrings) {
  const type = cellXml.match(/\bt="([^"]+)"/)?.[1];

  if (type === 's') {
    const idx = Number(cellXml.match(/<v>([^<]*)<\/v>/)?.[1]);
    return Number.isFinite(idx) ? sharedStrings[idx] ?? '' : '';
  }

  const inline = cellXml.match(/<is><t(?:\s[^>]*)?>([\s\S]*?)<\/t><\/is>/);
  if (inline) return decodeXml(inline[1]);

  const value = cellXml.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/);
  if (value) return decodeXml(value[1]);

  return '';
}

function parseSheetRows(sheetXml, sharedStrings) {
  const rowsByNumber = new Map();

  for (const rowXml of sheetXml.match(/<row r="\d+"[^>]*>[\s\S]*?<\/row>/g) || []) {
    const rowNum = Number(rowXml.match(/<row r="(\d+)"/)?.[1]);
    if (!Number.isFinite(rowNum)) continue;

    const values = {};
    for (const cellXml of rowXml.match(/<c r="[A-Z]+\d+"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g) || []) {
      const colLetters = cellXml.match(/r="([A-Z]+)/)?.[1];
      if (!colLetters) continue;

      const colIdx = colLetterToIndex(colLetters);
      if (!NEEDED_COLS.has(colIdx)) continue;
      values[colIdx] = parseCellValue(cellXml, sharedStrings);
    }

    rowsByNumber.set(rowNum, values);
  }

  return rowsByNumber;
}

function resolveSheetEntry(workbookXml, relsXml, { sheetName, namePattern }) {
  const relMap = new Map();
  for (const match of relsXml.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relMap.set(match[1], match[2].replace(/^\//, ''));
  }

  const sheets = [
    ...workbookXml.matchAll(/<sheet\b([^>]*)\/?>/g),
  ].map((match) => {
    const attrs = match[1];
    const name = attrs.match(/name="([^"]+)"/)?.[1];
    const relId = attrs.match(/r:id="([^"]+)"/)?.[1];
    return { name, relId };
  });

  if (sheetName) {
    const exact = sheets.find((sheet) => sheet.name === sheetName);
    if (exact?.relId) {
      return { name: exact.name, path: `xl/${relMap.get(exact.relId)}` };
    }
  }

  const patternMatch = sheets.find((sheet) => sheet.name && namePattern.test(sheet.name));
  if (patternMatch?.relId) {
    return { name: patternMatch.name, path: `xl/${relMap.get(patternMatch.relId)}` };
  }

  const first = sheets.find((sheet) => sheet.relId);
  if (!first?.relId) {
    throw new Error('No worksheet found in Excel file');
  }

  return { name: first.name, path: `xl/${relMap.get(first.relId)}` };
}

function lineRefFromValues(values, excelRowNumber) {
  const docNo = String(values[COL.DOC_NO] ?? '').trim();
  return docNo || `row-${excelRowNumber}`;
}

async function parseOrderXlsxZip(buffer, { paymentType = 'STRIPE', sheetName } = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');

  if (!workbookXml || !relsXml) {
    throw new Error('Invalid Excel file: missing workbook metadata');
  }

  const sheetEntry = resolveSheetEntry(workbookXml, relsXml, {
    sheetName,
    namePattern: /payment listing/i,
  });
  const sheetFile = zip.file(sheetEntry.path);
  if (!sheetFile) {
    throw new Error(`Worksheet file not found: ${sheetEntry.name}`);
  }

  const [sheetXml, sharedStringsXml] = await Promise.all([
    sheetFile.async('string'),
    zip.file('xl/sharedStrings.xml')?.async('string').catch(() => null),
  ]);

  const sharedStrings = parseSharedStrings(sharedStringsXml);
  const rowsByNumber = parseSheetRows(sheetXml, sharedStrings);
  const filterType = normalizePaymentType(paymentType);

  const orders = [];
  const payments = [];
  let skipped = 0;
  let scanned = 0;

  const rowNumbers = [...rowsByNumber.keys()].sort((a, b) => a - b);
  for (const rowNum of rowNumbers) {
    if (rowNum < DATA_START_ROW) continue;

    const values = rowsByNumber.get(rowNum);
    if (!values || !Object.keys(values).length) continue;

    scanned += 1;
    const rowPaymentType = normalizePaymentType(values[COL.PAYMENT_TYPE]);
    if (rowPaymentType !== filterType) continue;

    const order_id = String(values[COL.ORDER_ID] ?? '').trim();
    const orderAmount = parseNumber(values[COL.AMT]);
    const paymentAmount = parseNumber(values[COL.FINAL_PAID]);
    const line_ref = lineRefFromValues(values, rowNum);
    const note = String(values[COL.NOTE] ?? '').trim();
    const comment = String(values[COL.COMMENT] ?? '').trim();

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
        sourceRow: rowNum,
      });
    }

    if (!Number.isNaN(paymentAmount)) {
      payments.push({
        order_id,
        amount: paymentAmount,
        line_ref,
        note: note || null,
        comment: comment || null,
        sourceRow: rowNum,
      });
    }

    if (Number.isNaN(orderAmount) && Number.isNaN(paymentAmount)) {
      skipped += 1;
    }
  }

  return {
    sheetName: sheetEntry.name,
    paymentType: filterType,
    orders,
    payments,
    scanned,
    skipped,
    parser: 'zip',
  };
}

module.exports = {
  parseOrderXlsxZip,
  COL,
};
