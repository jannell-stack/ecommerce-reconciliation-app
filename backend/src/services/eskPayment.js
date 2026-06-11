const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { applyColumnBDateStyles } = require('./eskPaymentStyles');

const TEMPLATE_PATH = path.join(__dirname, '..', '..', 'templates', 'ESK-Payment-template.xlsx');
const META_PATH = path.join(__dirname, '..', '..', 'templates', 'esk-payment-meta.json');

const NC_COL = {
  HOLDER: 0,
  DOC_NO: 1,
  RETAIL_INVOICE: 2,
  PAYMENT_TYPE: 3,
  AMT: 4,
  NOTE: 5,
  NOTE_HELPER: 6,
  ORDER_HELPER: 7,
};

const NC_HEADER_ROWS = 3;
const SK_ORDER_ID = 0;
const SK_ITEM_TYPE = 13;
const SK_PRICE_PAID = 18;
const SK_PAYMENT_METHOD = 19;
const SK_LOOKUP_KEY = 33;
const SK_LOOKUP_PRICE = 34;
const SK_LOOKUP_METHOD = 35;
const SK_ORDERS_MIN_COLUMNS = 36;
const XLSX_READ_OPTS = {
  cellStyles: false,
  cellHTML: false,
  cellFormula: false,
};

let cachedProductionMeta = null;

function cellHasContent(cell) {
  if (!cell) return false;
  if (cell.f) return true;
  return cell.v != null && cell.v !== '';
}

const LISTING_DATE_FORMAT = 'm/d/yy';

function setCell(sheet, address, value) {
  if (value && typeof value === 'object' && ('v' in value || 'f' in value)) {
    sheet[address] = value;
    return;
  }
  if (value === '' || value == null) return;
  sheet[address] = {
    t: typeof value === 'number' ? 'n' : 's',
    v: value,
  };
}

function isExcelDateSerial(value) {
  return typeof value === 'number' && Number.isFinite(value) && value >= 30000 && value <= 60000;
}

function setDateCell(sheet, address, serial) {
  if (!isExcelDateSerial(serial)) return;
  const existing = sheet[address] || {};
  sheet[address] = {
    ...existing,
    t: 'n',
    v: serial,
    z: LISTING_DATE_FORMAT,
  };
}

function setColumnBCell(sheet, address, value) {
  if (isExcelDateSerial(value)) {
    setDateCell(sheet, address, value);
    return;
  }
  setCell(sheet, address, value);
}

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function formatNoteFromNc(ncNote) {
  let body = String(ncNote ?? '').trim();
  if (!body) return '';
  body = body.replace(/^Note:\s*/i, '').trim();
  body = body.replace(/^[A-Za-z]{1,3}\s*#?\s*/i, '').trim();
  if (body.startsWith('#')) body = body.slice(1).trim();
  body = body.replace(/^-\s*,/, ',');
  return body;
}

function ncHolderValue(holder) {
  if (holder === '' || holder == null) return ' ';
  return holder;
}

function extractOrderKey(noteText) {
  const text = String(noteText ?? '').trim();
  const match = text.match(/(\d{6}-\d{5,6})/);
  if (match) return match[1];
  return text.slice(0, 13).trim();
}

function extractAllOrderKeys(noteText) {
  const matches = [...String(noteText ?? '').matchAll(/(\d{6}-\d{5,6})/g)].map((m) => m[1]);
  return [...new Set(matches)];
}

function extractTopUpAmount(noteText) {
  const text = String(noteText ?? '');
  let match = text.match(/\$(\d+(?:\.\d+)?)\s*top\s*up/i);
  if (match) return parseNumber(match[1]);
  match = text.match(/top\s*up\s*\$(\d+(?:\.\d+)?)/i);
  return match ? parseNumber(match[1]) : NaN;
}

function hasReferToOrderNote(noteText) {
  const text = String(noteText ?? '');
  return /refer to\s+#\d{6}-\d{5,6}/i.test(text)
    || /\$\d+(?:\.\d+)?\s+refer to\s+#/i.test(text);
}

function hasTopUpMention(noteText) {
  return /top\s*up/i.test(String(noteText ?? ''));
}

function amountsRoughlyEqual(a, b) {
  return Math.abs(parseNumber(a) - parseNumber(b)) < 0.01;
}

function getSkPricedRows(orderId, skLookup) {
  const rows = skLookup.get(String(orderId ?? '').trim()) || [];
  return rows.filter((row) => Number.isFinite(row.pricePaid) && row.pricePaid !== 0);
}

function getSkPrimaryPrice(orderId, skLookup) {
  return getSkPricedRows(orderId, skLookup)[0]?.pricePaid ?? NaN;
}

function splitNoteForGroupedRow(formattedNote, orderKey, index) {
  if (index === 0) return formattedNote;
  return orderKey;
}

function shouldSplitGroupedOrders(orderIds, ncAmt, formattedNote, skLookup) {
  if (orderIds.length < 2 || hasReferToOrderNote(formattedNote)) return false;

  const skAmounts = orderIds.map((id) => getSkPrimaryPrice(id, skLookup));
  if (skAmounts.some((amount) => Number.isNaN(amount) || amount <= 0)) return false;
  if (!amountsRoughlyEqual(skAmounts.reduce((sum, amount) => sum + amount, 0), ncAmt)) {
    return false;
  }

  if (/grouped order/i.test(formattedNote)) return true;
  if (!formattedNote.includes('&')) return false;

  const minAmount = Math.min(...skAmounts);
  const maxAmount = Math.max(...skAmounts);
  const ratio = maxAmount / minAmount;
  return ratio >= 2.925 && ratio <= 5;
}

function expandNcRow(ncRow, skLookup) {
  const formattedNote = formatNoteFromNc(ncRow[NC_COL.NOTE]);
  const ncAmt = parseNumber(ncRow[NC_COL.AMT]);
  const defaultExpansion = {
    ncRow,
    orderKey: extractOrderKey(formattedNote),
    note: formattedNote,
    amount: ncRow[NC_COL.AMT],
  };

  if (Number.isNaN(ncAmt)) return [defaultExpansion];

  const orderIds = extractAllOrderKeys(formattedNote);
  if (shouldSplitGroupedOrders(orderIds, ncAmt, formattedNote, skLookup)) {
    const skAmounts = orderIds.map((id) => getSkPrimaryPrice(id, skLookup));
    return orderIds.map((orderKey, index) => ({
      ncRow,
      orderKey,
      note: splitNoteForGroupedRow(formattedNote, orderKey, index),
      amount: skAmounts[index],
    }));
  }

  const singleOrderId = orderIds[0] || defaultExpansion.orderKey;
  const pricedRows = getSkPricedRows(singleOrderId, skLookup);
  const skSum = pricedRows.reduce((sum, row) => sum + row.pricePaid, 0);
  const topUpAmount = extractTopUpAmount(formattedNote);

  if (
    orderIds.length === 1
    && Number.isFinite(topUpAmount)
    && topUpAmount > 0
    && ncAmt > topUpAmount
    && pricedRows.length >= 2
    && ncAmt > skSum + 0.01
  ) {
    return [
      { ncRow, orderKey: singleOrderId, note: formattedNote, amount: ncAmt - topUpAmount },
      { ncRow, orderKey: singleOrderId, note: formattedNote, amount: topUpAmount },
    ];
  }

  if (
    orderIds.length === 1
    && pricedRows.length === 2
    && amountsRoughlyEqual(skSum, ncAmt)
    && !hasTopUpMention(formattedNote)
  ) {
    const paymentMethods = new Set(pricedRows.map((row) => row.paymentMethod).filter(Boolean));
    if (paymentMethods.size >= 2) {
      return pricedRows.map((row) => ({
        ncRow,
        orderKey: singleOrderId,
        note: formattedNote,
        amount: row.pricePaid,
      }));
    }
  }

  return [defaultExpansion];
}

function extractReferToInvoice(text) {
  const match = String(text ?? '').match(/refer to\s+(LS0ESK\d+)/i);
  return match ? match[1] : '';
}

function parseListingDate(value) {
  if (isExcelDateSerial(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    const monthNames = [
      'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
    ];
    return {
      serial: value,
      sheetLabel: `${monthNames[date.getUTCMonth()]}${String(date.getUTCFullYear()).slice(-2)}`,
    };
  }

  const text = String(value ?? '').trim();
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (Number.isNaN(date.getTime())) return null;

  const epoch = Date.UTC(1899, 11, 30);
  const serial = Math.round((date.getTime() - epoch) / 86400000);
  const monthNames = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
  ];

  return {
    serial,
    sheetLabel: `${monthNames[month - 1]}${String(year).slice(-2)}`,
  };
}

const NC_SUMMARY_LABELS = new Set(['Grand Total', 'Deposit Refund', 'Net Total']);

function isNcDateSeparatorRow(ncRow) {
  const docNo = String(ncRow[NC_COL.DOC_NO] ?? '').trim();
  if (docNo) return false;
  return parseListingDate(ncRow[NC_COL.HOLDER]) != null;
}

function isNcSummaryRow(ncRow) {
  const label = String(ncRow[NC_COL.PAYMENT_TYPE] ?? '').trim();
  return NC_SUMMARY_LABELS.has(label);
}

function applyBlankSpacerRow(sheet, excelRow, valuesOnly) {
  if (valuesOnly) {
    setCell(sheet, `A${excelRow}`, '');
  } else {
    setCell(sheet, `A${excelRow}`, { f: `LEFT(G${excelRow},13)`, v: '', t: 's' });
  }
  setCell(sheet, `B${excelRow}`, ' ');
  for (const column of ['C', 'D', 'E', 'F']) {
    setCell(sheet, `${column}${excelRow}`, ' ');
  }
}

function applyDateSpacerRow(sheet, excelRow, dateSerial, valuesOnly) {
  if (valuesOnly) {
    setCell(sheet, `A${excelRow}`, '');
  } else {
    setCell(sheet, `A${excelRow}`, { f: `LEFT(G${excelRow},13)`, v: '', t: 's' });
  }
  setDateCell(sheet, `B${excelRow}`, dateSerial);
  for (const column of ['C', 'D', 'E', 'F']) {
    setCell(sheet, `${column}${excelRow}`, ' ');
  }
}

function findSheet(workbook, pattern) {
  return workbook.SheetNames.find((name) => pattern.test(name)) || workbook.SheetNames[0];
}

function readSheetRows(buffer, sheetPattern, readOptions = {}) {
  const workbook = XLSX.read(buffer, { type: 'buffer', ...XLSX_READ_OPTS, ...readOptions });
  const sheetName = findSheet(workbook, sheetPattern);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error('No worksheet found in Excel file');
  }

  return {
    sheetName,
    rows: XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }),
  };
}

function addSkLookupEntry(lookup, orderId, pricePaid, paymentMethod) {
  if (!orderId) return;
  if (!lookup.has(orderId)) {
    lookup.set(orderId, []);
  }
  lookup.get(orderId).push({ pricePaid, paymentMethod });
}

function isSkFeeRow(row) {
  const itemType = String(row[SK_ITEM_TYPE] ?? '').trim();
  return itemType === 'Fee';
}

function skRowPricePaid(row) {
  const pricePaid = parseNumber(row[SK_PRICE_PAID]);
  if (Number.isFinite(pricePaid)) return pricePaid;
  return isSkFeeRow(row) ? 0 : NaN;
}

function normalizeSkPaymentMethod(ncPaymentType) {
  const paymentType = String(ncPaymentType ?? '').trim().toUpperCase();
  if (paymentType === 'STRIPE') return 'Stripe';
  if (paymentType.includes('GPAY') || paymentType === 'ECOM-GPAY') return 'Grab';
  if (paymentType === 'SHOPBACK') return 'ShopBack';
  return '';
}

function isTopUpNote(note) {
  return /top\s*up/i.test(String(note ?? ''));
}

function isDuplicateInvoiceNote(note) {
  return /duplicate\s+inv/i.test(String(note ?? ''));
}

function zeroSkMatch(paymentType, paymentMethod) {
  return {
    pricePaid: 0,
    paymentMethod: paymentMethod || normalizeSkPaymentMethod(paymentType) || null,
  };
}

function createSkUsageTracker(skLookup) {
  const usedCounts = new Map();

  function getOrderUsage(orderId) {
    const key = String(orderId ?? '').trim();
    if (!usedCounts.has(key)) {
      usedCounts.set(key, new Map());
    }
    return usedCounts.get(key);
  }

  function countAvailable(orderId, absAmount) {
    const rows = skLookup.get(String(orderId ?? '').trim()) || [];
    return rows.filter((row) => Math.abs(row.pricePaid) === absAmount).length;
  }

  function getUsed(orderId, absAmount) {
    return getOrderUsage(orderId).get(absAmount) || 0;
  }

  function claim(orderId, absAmount) {
    const orderUsage = getOrderUsage(orderId);
    const used = getUsed(orderId, absAmount);
    orderUsage.set(absAmount, used + 1);
    return used;
  }

  return { countAvailable, getUsed, claim };
}

function resolveSkMatch(orderId, skLookup, amount, context, usageTracker) {
  const orderKey = String(orderId ?? '').trim();
  const parsedAmount = parseNumber(amount);

  if (parsedAmount < 0 || isDuplicateInvoiceNote(context.note)) {
    return zeroSkMatch(context.paymentType);
  }

  const candidate = lookupSk(orderKey, skLookup, amount, context);
  if (!candidate) return null;

  const absAmount = Math.abs(parsedAmount);
  if (Number.isNaN(absAmount) || !usageTracker) return candidate;

  const rows = skLookup.get(orderKey) || [];
  const amountMatches = rows.filter((row) => Math.abs(row.pricePaid) === absAmount);
  const isExactAmountMatch = amountMatches.length > 0
    && Math.abs(candidate.pricePaid) === absAmount;

  if (!isExactAmountMatch) return candidate;

  const usedBefore = usageTracker.getUsed(orderKey, absAmount);
  if (usedBefore >= amountMatches.length) {
    return zeroSkMatch(context.paymentType, candidate.paymentMethod);
  }

  usageTracker.claim(orderKey, absAmount);
  return amountMatches[usedBefore] || candidate;
}

function prepareOrdersRowsForLookup(rows) {
  const prepared = rows.map((row) => {
    const copy = [...row];
    while (copy.length < SK_ORDERS_MIN_COLUMNS) {
      copy.push('');
    }
    return copy;
  });
  let currentOrderId = '';

  for (let i = 1; i < prepared.length; i += 1) {
    let orderId = String(prepared[i][SK_ORDER_ID] ?? '').trim();
    if (orderId) {
      prepared[i][SK_ORDER_ID] = orderId;
      currentOrderId = orderId;
      continue;
    }

    const pricePaid = skRowPricePaid(prepared[i]);
    if (currentOrderId && Number.isFinite(pricePaid) && pricePaid !== 0) {
      prepared[i][SK_ORDER_ID] = currentOrderId;
    } else if (currentOrderId && isSkFeeRow(prepared[i])) {
      prepared[i][SK_ORDER_ID] = currentOrderId;
    }
  }

  if (prepared[0]) {
    prepared[0][SK_LOOKUP_KEY] = 'LookupKey';
    prepared[0][SK_LOOKUP_PRICE] = 'LookupPrice';
    prepared[0][SK_LOOKUP_METHOD] = 'LookupMethod';
  }

  for (let i = 1; i < prepared.length; i += 1) {
    const orderId = String(prepared[i][SK_ORDER_ID] ?? '').trim();
    const pricePaid = skRowPricePaid(prepared[i]);
    const paymentMethod = String(prepared[i][SK_PAYMENT_METHOD] ?? '').trim();

    if (!orderId || !Number.isFinite(pricePaid)) continue;

    const absPrice = Math.abs(pricePaid);
    prepared[i][SK_LOOKUP_KEY] = `${orderId}|${absPrice}`;
    prepared[i][SK_LOOKUP_PRICE] = pricePaid;
    prepared[i][SK_LOOKUP_METHOD] = paymentMethod;
  }

  return prepared;
}

function buildSkLookup(rows) {
  const lookup = new Map();
  let currentOrderId = '';

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const orderId = String(row[SK_ORDER_ID] ?? '').trim();
    const pricePaid = skRowPricePaid(row);
    const paymentMethod = String(row[SK_PAYMENT_METHOD] ?? '').trim();

    if (orderId) {
      currentOrderId = orderId;
      addSkLookupEntry(lookup, orderId, pricePaid, paymentMethod);
      continue;
    }

    if (currentOrderId && Number.isFinite(pricePaid) && pricePaid !== 0) {
      addSkLookupEntry(lookup, currentOrderId, pricePaid, paymentMethod);
    } else if (currentOrderId && isSkFeeRow(row) && paymentMethod) {
      addSkLookupEntry(lookup, currentOrderId, 0, paymentMethod);
    }
  }

  return lookup;
}

function buildSkLookupFromBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', ...XLSX_READ_OPTS });
  const sheetName = findSheet(workbook, /orders/i);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet?.['!ref']) {
    throw new Error('No Orders worksheet found in SK Combined file');
  }

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const lookup = new Map();

  for (let row = 1; row <= range.e.r; row += 1) {
    const orderCell = sheet[XLSX.utils.encode_cell({ r: row, c: SK_ORDER_ID })];
    const orderId = String(orderCell?.v ?? '').trim();
    if (!orderId) continue;

    const priceCell = sheet[XLSX.utils.encode_cell({ r: row, c: SK_PRICE_PAID })];
    const methodCell = sheet[XLSX.utils.encode_cell({ r: row, c: SK_PAYMENT_METHOD })];
    const entry = {
      pricePaid: parseNumber(priceCell?.v),
      paymentMethod: String(methodCell?.v ?? '').trim(),
    };

    if (!lookup.has(orderId)) {
      lookup.set(orderId, []);
    }
    lookup.get(orderId).push(entry);
  }

  return { lookup, sheetName };
}

function lookupSk(orderId, skLookup, amount, { paymentType, note } = {}) {
  const rows = skLookup.get(String(orderId ?? '').trim());
  if (!rows?.length) return null;

  const absAmount = Math.abs(parseNumber(amount));
  if (!Number.isNaN(absAmount)) {
    const amountMatch = rows.find((row) => Math.abs(row.pricePaid) === absAmount);
    if (amountMatch) return amountMatch;
  }

  const expectedMethod = normalizeSkPaymentMethod(paymentType);
  if (isTopUpNote(note) && expectedMethod === 'Stripe') {
    const feeRow = rows.find(
      (row) => row.pricePaid === 0 && row.paymentMethod === 'Stripe',
    );
    if (feeRow) return feeRow;
  }

  if (expectedMethod) {
    const methodRows = rows.filter((row) => row.paymentMethod === expectedMethod);
    if (methodRows.length === 1) return methodRows[0];

    const pricedRows = methodRows.filter((row) => row.pricePaid !== 0);
    if (pricedRows.length === 1) return pricedRows[0];

    if (!Number.isNaN(absAmount) && pricedRows.length > 1) {
      return pricedRows.reduce((best, row) => {
        const bestDiff = Math.abs(Math.abs(best.pricePaid) - absAmount);
        const rowDiff = Math.abs(Math.abs(row.pricePaid) - absAmount);
        return rowDiff < bestDiff ? row : best;
      });
    }

    if (pricedRows.length > 0) return pricedRows[0];
    if (methodRows.length > 0) return methodRows[0];
  }

  if (rows.length === 1) return rows[0];
  return rows[0];
}

function buildOrdersLookupRange(ordersLastRow) {
  const lastRow = Math.max(ordersLastRow, 2);
  return {
    fullRange: `Orders!$A$1:$AG$${lastRow}`,
    keyRange: `Orders!$AH$2:$AJ$${lastRow}`,
    lastRow,
  };
}

function skLookupFormula(excelRow, ordersLookup, field) {
  const keyColumn = field === 'price' ? 2 : 3;
  const fallbackColumn = field === 'price' ? 19 : 20;
  const compositeLookup = [
    `VLOOKUP($A${excelRow}&"|"&TEXT(ABS($F${excelRow}),"0"),`,
    `${ordersLookup.keyRange},${keyColumn},FALSE)`,
  ].join('');
  const orderLookup = `VLOOKUP($A${excelRow},${ordersLookup.fullRange},${fallbackColumn},FALSE)`;
  return `IFERROR(${compositeLookup},${orderLookup})`;
}

function shouldWriteLookupFormula(skMatch) {
  if (!skMatch) return false;
  return skMatch.pricePaid !== 0;
}

function loadTemplateWorkbook() {
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error('ESK payment template is missing. Expected backend/templates/ESK-Payment-template.xlsx');
  }

  return XLSX.read(fs.readFileSync(TEMPLATE_PATH), {
    type: 'buffer',
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
  });
}

function loadProductionMeta() {
  if (cachedProductionMeta) return cachedProductionMeta;

  if (!fs.existsSync(META_PATH)) {
    throw new Error('ESK payment metadata is missing. Expected backend/templates/esk-payment-meta.json');
  }

  cachedProductionMeta = JSON.parse(fs.readFileSync(META_PATH, 'utf8'));
  return cachedProductionMeta;
}

function useValuesOnlyOutput() {
  return process.env.ESK_VALUES_ONLY === 'true' || process.env.NODE_ENV === 'production';
}

function getTemplateMainSheetName(workbook) {
  return workbook.SheetNames.find((name) => /payment listing/i.test(name)) || workbook.SheetNames[0];
}

function buildTemplateMetadata(templateMain) {
  const rows = XLSX.utils.sheet_to_json(templateMain, { header: 1, defval: null });
  const patterns = new Map();

  for (let i = 3; i < rows.length; i += 1) {
    const docNo = String(rows[i][2] ?? '').trim();
    if (!docNo) continue;
    const excelRow = i + 1;
    const hCell = templateMain[`H${excelRow}`];
    const iCell = templateMain[`I${excelRow}`];
    const jCell = templateMain[`J${excelRow}`];
    const kCell = templateMain[`K${excelRow}`];

    patterns.set(docNo, {
      hasH: cellHasContent(hCell),
      hasI: cellHasContent(iCell),
      hasJ: cellHasContent(jCell),
      hasK: cellHasContent(kCell),
      hFormula: hCell?.f || null,
      hValue: hCell?.v ?? null,
      jValue: jCell?.v ?? null,
      kValue: kCell?.v ?? null,
      kFormula: kCell?.f || null,
      retailInvoice: rows[i][3] ?? null,
      templateNote: rows[i][6] ?? null,
      templateOrderId: rows[i][0] ?? null,
    });
  }

  return { patterns };
}

function buildNcMap(ncRows) {
  const map = new Map();

  for (let i = NC_HEADER_ROWS; i < ncRows.length; i += 1) {
    const docNo = String(ncRows[i][NC_COL.DOC_NO] ?? '').trim();
    if (!docNo || docNo.toUpperCase() === 'ESK') continue;
    map.set(docNo, ncRows[i]);
  }

  return map;
}

function clearFormulaCells(sheet, excelRow, columns) {
  for (const column of columns) {
    delete sheet[`${column}${excelRow}`];
  }
}

function updateFormulaRow(
  sheet,
  excelRow,
  ncRow,
  pattern,
  skLookup,
  valuesOnly,
  usageTracker,
  expansion = null,
  ordersLookup = null,
) {
  const note = expansion?.note ?? formatNoteFromNc(ncRow[NC_COL.NOTE]);
  const orderKey = expansion?.orderKey ?? extractOrderKey(note);
  const rowAmount = expansion?.amount ?? ncRow[NC_COL.AMT];
  const docNo = String(ncRow[NC_COL.DOC_NO] ?? '').trim();
  const skMatch = resolveSkMatch(orderKey, skLookup, rowAmount, {
    paymentType: ncRow[NC_COL.PAYMENT_TYPE],
    note,
  }, usageTracker);
  const useFormulas = !valuesOnly;

  if (valuesOnly) {
    setCell(sheet, `A${excelRow}`, orderKey);
  } else {
    setCell(sheet, `A${excelRow}`, { f: `LEFT(G${excelRow},13)`, v: orderKey, t: 's' });
  }
  setCell(sheet, `B${excelRow}`, ncHolderValue(ncRow[NC_COL.HOLDER]));
  setCell(sheet, `C${excelRow}`, docNo);
  setCell(sheet, `D${excelRow}`, ncRow[NC_COL.RETAIL_INVOICE]);
  setCell(sheet, `E${excelRow}`, ncRow[NC_COL.PAYMENT_TYPE]);
  setCell(sheet, `F${excelRow}`, rowAmount);
  setCell(sheet, `G${excelRow}`, note);

  clearFormulaCells(sheet, excelRow, ['H', 'I', 'J', 'K']);

  let hValue = null;
  if (pattern.hasH) {
    if (useFormulas && pattern.hFormula && ordersLookup && shouldWriteLookupFormula(skMatch)) {
      hValue = skMatch?.pricePaid ?? null;
      if (hValue != null && Number.isNaN(hValue)) hValue = null;
      setCell(sheet, `H${excelRow}`, {
        f: skLookupFormula(excelRow, ordersLookup, 'price'),
        v: hValue,
        t: 'n',
      });
    } else if (pattern.hFormula || skMatch) {
      hValue = skMatch?.pricePaid ?? null;
      if (hValue != null && Number.isNaN(hValue)) hValue = null;
      setCell(sheet, `H${excelRow}`, hValue);
    } else if (pattern.hValue != null) {
      hValue = pattern.hValue;
      setCell(sheet, `H${excelRow}`, { v: hValue, t: 'n' });
    }

    if (pattern.hasI && hValue != null) {
      const amount = parseNumber(rowAmount);
      const diff = !Number.isNaN(amount) ? amount - hValue : null;
      if (useFormulas) {
        setCell(sheet, `I${excelRow}`, {
          f: `F${excelRow}-H${excelRow}`,
          v: diff,
          t: 'n',
        });
      } else {
        setCell(sheet, `I${excelRow}`, diff);
      }
    }

    if (pattern.hasJ) {
      const paymentMethod = skMatch?.paymentMethod ?? null;
      if (useFormulas && pattern.hasJ && ordersLookup && shouldWriteLookupFormula(skMatch)) {
        setCell(sheet, `J${excelRow}`, {
          f: skLookupFormula(excelRow, ordersLookup, 'method'),
          v: paymentMethod,
          t: 's',
        });
      } else {
        setCell(sheet, `J${excelRow}`, paymentMethod);
      }
    }
  }

  if (pattern.hasK) {
    const referInvoice =
      extractReferToInvoice(note) || extractReferToInvoice(ncRow[NC_COL.NOTE]);
    const kValue = pattern.kValue ?? referInvoice ?? null;
    if (kValue != null || pattern.kFormula) {
      if (valuesOnly) {
        setCell(sheet, `K${excelRow}`, kValue);
      } else {
        setCell(sheet, `K${excelRow}`, {
          f: pattern.kFormula || undefined,
          v: kValue,
          t: typeof kValue === 'number' ? 'n' : 's',
        });
      }
    }
  }
}

function applyHeaderRows(sheet, headerRows) {
  headerRows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      if (value == null || value === '') return;
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      if (columnIndex === 1) {
        setColumnBCell(sheet, address, value);
        return;
      }
      setCell(sheet, address, value);
    });
  });
}

function applySpacerRow(sheet, excelRow, cells) {
  for (const [column, cell] of Object.entries(cells)) {
    const address = `${column}${excelRow}`;
    if (cell.f) {
      const output = { t: cell.t || 's', v: cell.v, f: cell.f };
      if (column === 'B' && isExcelDateSerial(cell.v)) {
        output.z = LISTING_DATE_FORMAT;
      }
      sheet[address] = output;
      continue;
    }
    if (column === 'B') {
      setColumnBCell(sheet, address, cell.v);
      continue;
    }
    setCell(sheet, address, cell.v);
  }
}

function finalizeSheetRange(sheet, rowCount) {
  sheet['!ref'] = `A1:AB${rowCount}`;
}

const DEFAULT_DATA_ROW_PATTERN = {
  hasH: true,
  hasI: true,
  hasJ: true,
  hasK: false,
  hFormula: true,
};

function copyHeaderRowsFromTemplate(templateMain, sheet) {
  const templateRows = XLSX.utils.sheet_to_json(templateMain, { header: 1, defval: null });

  for (let rowIndex = 0; rowIndex < NC_HEADER_ROWS; rowIndex += 1) {
    const row = templateRows[rowIndex] || [];
    row.forEach((value, columnIndex) => {
      if (value == null && value !== 0 && value !== ' ') return;
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      if (columnIndex === 1) {
        setColumnBCell(sheet, address, value);
        return;
      }
      setCell(sheet, address, value);
    });
  }
}

function buildMainSheetFromNcRows(
  ncRows,
  skRows,
  skLookup,
  listingDate,
  valuesOnly,
  { headerRows, headerSource } = {},
) {
  const sheet = {};

  if (headerRows) {
    applyHeaderRows(sheet, headerRows);
  } else if (headerSource) {
    copyHeaderRowsFromTemplate(headerSource, sheet);
  }

  if (listingDate) {
    setDateCell(sheet, 'B3', listingDate.serial);
  }

  let generated = 0;
  let excelRow = NC_HEADER_ROWS;
  const usageTracker = createSkUsageTracker(skLookup);
  const ordersLookup = buildOrdersLookupRange(skRows.length);

  for (let i = NC_HEADER_ROWS; i < ncRows.length; i += 1) {
    const ncRow = ncRows[i];

    if (isNcDateSeparatorRow(ncRow)) {
      const rowDate = parseListingDate(ncRow[NC_COL.HOLDER]);
      excelRow += 1;
      applyBlankSpacerRow(sheet, excelRow, valuesOnly);
      excelRow += 1;
      applyDateSpacerRow(sheet, excelRow, rowDate.serial, valuesOnly);
      continue;
    }

    if (isNcSummaryRow(ncRow)) continue;

    const docNo = String(ncRow[NC_COL.DOC_NO] ?? '').trim();
    if (!docNo || docNo.toUpperCase() === 'ESK') continue;

    const expansions = expandNcRow(ncRow, skLookup);
    for (const expansion of expansions) {
      excelRow += 1;
      updateFormulaRow(
        sheet,
        excelRow,
        ncRow,
        DEFAULT_DATA_ROW_PATTERN,
        skLookup,
        valuesOnly,
        usageTracker,
        expansion,
        ordersLookup,
      );
      generated += 1;
    }
  }

  excelRow += 1;
  applyBlankSpacerRow(sheet, excelRow, valuesOnly);

  finalizeSheetRange(sheet, excelRow);
  return { sheet, generated, rowCount: excelRow };
}

async function createEskPaymentWorkbook(skBuffer, ncBuffer) {
  const valuesOnly = useValuesOnlyOutput();
  const { sheetName: ncSheetName, rows: ncRows } = readSheetRows(ncBuffer, /payment listing/i);

  if (ncRows.length <= NC_HEADER_ROWS) {
    throw new Error('NC payment listing file has no data rows');
  }

  const listingDate = parseListingDate(ncRows[2]?.[0]);
  const ncMap = buildNcMap(ncRows);
  const { sheetName: skSheetName, rows: skRows } = readSheetRows(
    skBuffer,
    /orders/i,
    { cellFormula: true },
  );
  const preparedSkRows = prepareOrdersRowsForLookup(skRows);
  const skLookup = buildSkLookup(preparedSkRows);

  let mainSheet;
  let generated;
  let sheetTitle;

  if (valuesOnly) {
    const meta = loadProductionMeta();
    sheetTitle = listingDate
      ? `Payment Listing-${listingDate.sheetLabel}`
      : meta.mainSheetName;

    ({ sheet: mainSheet, generated } = buildMainSheetFromNcRows(
      ncRows,
      preparedSkRows,
      skLookup,
      listingDate,
      valuesOnly,
      { headerRows: meta.headerRows },
    ));
  } else {
    const template = loadTemplateWorkbook();
    const templateMainName = getTemplateMainSheetName(template);
    const templateMain = template.Sheets[templateMainName];
    sheetTitle = listingDate
      ? `Payment Listing-${listingDate.sheetLabel}`
      : templateMainName;

    ({ sheet: mainSheet, generated } = buildMainSheetFromNcRows(
      ncRows,
      preparedSkRows,
      skLookup,
      listingDate,
      false,
      { headerSource: templateMain },
    ));
  }

  const workbook = {
    SheetNames: [sheetTitle, 'Orders'],
    Sheets: {
      [sheetTitle]: mainSheet,
      Orders: XLSX.utils.aoa_to_sheet(preparedSkRows),
    },
  };

  let buffer = XLSX.write(workbook, {
    type: 'buffer',
    bookType: 'xlsx',
    compression: true,
    cellStyles: true,
    cellFormula: !valuesOnly,
  });
  buffer = await applyColumnBDateStyles(buffer);
  const filename = `ESK-Payment Listing${listingDate ? `-${listingDate.sheetLabel}` : ''}.xlsx`;

  return {
    buffer,
    filename,
    sheetTitle,
    generated,
    skipped: ncMap.size - generated,
    skSheetName,
    ncSheetName,
  };
}

module.exports = {
  createEskPaymentWorkbook,
};
