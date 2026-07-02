const XLSX = require('xlsx');
const { applyColumnBDateStyles } = require('./eskPaymentStyles');
const {
  detectNcListingDate,
  isNcDateSeparatorRow,
  isNcSummaryRow,
  isNcValidDocNo,
  resolveNcLayout,
} = require('../utils/ncLayout');
const SK_ORDER_ID = 0;
const SK_ITEM_TYPE = 13;
const SK_PRICE_PAID = 18;
const SK_PAYMENT_METHOD = 19;
const SK_LOOKUP_KEY = 33;
const SK_LOOKUP_PRICE = 34;
const SK_LOOKUP_METHOD = 35;
const SK_ORDERS_MIN_COLUMNS = 36;

const SK_LEGACY_COLUMNS = {
  ORDER_ID: SK_ORDER_ID,
  ITEM_TYPE: SK_ITEM_TYPE,
  PRICE_PAID: SK_PRICE_PAID,
  PAYMENT_METHOD: SK_PAYMENT_METHOD,
};

function resolveSkColumns(headerRow) {
  const headers = (headerRow || []).map((cell) => String(cell ?? '').trim());
  const pricePaidIndex = headers.indexOf('Order Final Price Paid');
  const productTypeIndex = headers.indexOf('Product Type');

  if (pricePaidIndex >= 0 && productTypeIndex >= 0) {
    return {
      ORDER_ID: Math.max(headers.indexOf('Order ID'), 0),
      ITEM_TYPE: productTypeIndex,
      PRICE_PAID: pricePaidIndex,
      PAYMENT_METHOD: Math.max(headers.indexOf('Payment Method'), 16),
    };
  }

  return SK_LEGACY_COLUMNS;
}
const XLSX_READ_OPTS = {
  cellStyles: false,
  cellHTML: false,
  cellFormula: false,
};

const LISTING_DATE_FORMAT = 'm/d/yy';
const AMOUNT_NUMBER_FORMAT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';

const ESK_DETAIL_HEADERS = [
  'Ord No.',
  'Ord',
  'docdate',
  'BILLNO',
  'RETAIL_INVOICE',
  'PAYMENT',
  'ORIGINALAMT',
  'NOTE',
  'DEPOSIT_REFUND',
  'NET_TOTAL',
  'COSTCODE',
  'Final rounded price paid',
  'diff',
  'pymt type',
  'Comment',
];

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

function setAmountCell(sheet, address, value) {
  if (value === '' || value == null) return;
  const numeric = parseNumber(value);
  if (Number.isNaN(numeric)) {
    setCell(sheet, address, value);
    return;
  }
  sheet[address] = {
    t: 'n',
    v: numeric,
    z: AMOUNT_NUMBER_FORMAT,
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
  body = body.replace(/^Order\s*#\s*/i, '').trim();
  body = body.replace(/^[A-Za-z]{1,3}\s*#?\s*/i, '').trim();
  if (body.startsWith('#')) body = body.slice(1).trim();
  body = body.replace(/^-\s*,/, ',');

  const orderMatch = body.match(/(\d{6}-\d{5,6})/);
  if (orderMatch && !body.startsWith(orderMatch[1])) {
    const orderIndex = body.indexOf(orderMatch[1]);
    if (orderIndex > 0) body = body.slice(orderIndex);
  }

  return body;
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

function isAdminFeeNote(noteText) {
  return /admin fee/i.test(String(noteText ?? ''));
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

function expandNcRow(ncRow, skLookup, layout) {
  const { col: ncCol } = layout;
  const formattedNote = formatNoteFromNc(ncRow[ncCol.NOTE]);
  const ncAmt = parseNumber(ncRow[ncCol.AMT]);
  const defaultExpansion = {
    ncRow,
    orderKey: extractOrderKey(formattedNote),
    note: formattedNote,
    amount: ncRow[ncCol.AMT],
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

function skRowPricePaidFromValues(rawPrice, itemType) {
  const pricePaid = parseNumber(rawPrice);
  if (Number.isFinite(pricePaid)) return pricePaid;
  return itemType === 'Fee' ? 0 : NaN;
}

function buildProductionSkLookup(skBuffer) {
  const workbook = XLSX.read(skBuffer, { type: 'buffer', ...XLSX_READ_OPTS });
  const sheetName = findSheet(workbook, /orders/i);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet?.['!ref']) {
    throw new Error('No Orders worksheet found in SK Combined file');
  }

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const lookup = new Map();
  let currentOrderId = '';

  const cellValue = (row, col) => sheet[XLSX.utils.encode_cell({ r: row, c: col })]?.v;
  const skCol = resolveSkColumns(
    Array.from({ length: range.e.c + 1 }, (_, col) => cellValue(0, col)),
  );

  for (let row = 1; row <= range.e.r; row += 1) {
    const orderId = String(cellValue(row, skCol.ORDER_ID) ?? '').trim();
    const itemType = String(cellValue(row, skCol.ITEM_TYPE) ?? '').trim();
    const pricePaid = skRowPricePaidFromValues(cellValue(row, skCol.PRICE_PAID), itemType);
    const paymentMethod = String(cellValue(row, skCol.PAYMENT_METHOD) ?? '').trim();

    if (orderId) {
      currentOrderId = orderId;
      addSkLookupEntry(lookup, orderId, pricePaid, paymentMethod);
      continue;
    }

    if (currentOrderId && Number.isFinite(pricePaid) && pricePaid !== 0) {
      addSkLookupEntry(lookup, currentOrderId, pricePaid, paymentMethod);
    } else if (currentOrderId && itemType === 'Fee' && paymentMethod) {
      addSkLookupEntry(lookup, currentOrderId, 0, paymentMethod);
    }
  }

  return { lookup, sheetName, rowCount: range.e.r + 1 };
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

function lookupSk(orderId, skLookup, amount, { paymentType, note } = {}) {
  const rows = skLookup.get(String(orderId ?? '').trim());
  if (!rows?.length) return null;

  const absAmount = Math.abs(parseNumber(amount));
  if (!Number.isNaN(absAmount)) {
    const amountMatch = rows.find((row) => Math.abs(row.pricePaid) === absAmount);
    if (amountMatch) return amountMatch;
  }

  const expectedMethod = normalizeSkPaymentMethod(paymentType);

  if (expectedMethod) {
    const pricedRows = rows
      .filter((row) => row.paymentMethod === expectedMethod && row.pricePaid !== 0);

    if (pricedRows.length === 1) return pricedRows[0];

    if (!Number.isNaN(absAmount) && pricedRows.length > 1) {
      return pricedRows.reduce((best, row) => {
        const bestDiff = Math.abs(Math.abs(best.pricePaid) - absAmount);
        const rowDiff = Math.abs(Math.abs(row.pricePaid) - absAmount);
        return rowDiff < bestDiff ? row : best;
      });
    }

    if (pricedRows.length > 0) return pricedRows[0];
  }

  if (rows.length === 1) return rows[0];
  return rows[0];
}

function hasSkStripePricedRow(orderKey, skLookup) {
  const rows = skLookup.get(String(orderKey ?? '').trim()) || [];
  return rows.some((row) => row.paymentMethod === 'Stripe' && row.pricePaid !== 0);
}

function buildNcStripeCompositeMap(ncRows, skLookup, layout) {
  const { col: ncCol, headerRows } = layout;
  const map = new Map();

  for (let i = headerRows; i < ncRows.length; i += 1) {
    const ncRow = ncRows[i];
    if (isNcDateSeparatorRow(ncRow, layout) || isNcSummaryRow(ncRow, layout)) continue;

    const paymentType = String(ncRow[ncCol.PAYMENT_TYPE] ?? '').trim().toUpperCase();
    if (paymentType !== 'STRIPE') continue;

    const orderKey = extractOrderKey(formatNoteFromNc(ncRow[ncCol.NOTE]));
    const amt = parseNumber(ncRow[ncCol.AMT]);
    if (!orderKey || Number.isNaN(amt)) continue;
    if (isAdminFeeNote(formatNoteFromNc(ncRow[ncCol.NOTE]))) continue;

    const absAmt = Math.abs(amt);
    const key = `${orderKey}|${absAmt}`;
    if (map.has(key)) continue;
    if (hasSkStripePricedRow(orderKey, skLookup)) continue;

    map.set(key, {
      key,
      price: absAmt,
      method: 'Stripe',
    });
  }

  return map;
}

function resolveHLookupValue(orderKey, rowAmount, skLookup, ncStripeComposite = null, note = '') {
  if (isAdminFeeNote(note)) return null;

  const orderRows = skLookup.get(String(orderKey ?? '').trim()) || [];
  const absAmount = Math.abs(parseNumber(rowAmount));

  if (!Number.isNaN(absAmount)) {
    const amountMatch = orderRows.find((row) => Math.abs(row.pricePaid) === absAmount);
    if (amountMatch) return amountMatch.pricePaid;

    const composite = ncStripeComposite?.get(`${String(orderKey ?? '').trim()}|${absAmount}`);
    if (composite) return composite.price;
  }

  return orderRows[0]?.pricePaid ?? null;
}

function lookupSkPaymentMethod(orderId, skLookup, amount, ncStripeComposite = null, note = '') {
  if (isAdminFeeNote(note)) return null;

  const orderKey = String(orderId ?? '').trim();
  const rows = skLookup.get(orderKey);
  if (!rows?.length) return null;

  const absAmount = Math.abs(parseNumber(amount));
  if (!Number.isNaN(absAmount)) {
    const amountMatch = rows.find((row) => Math.abs(row.pricePaid) === absAmount);
    if (amountMatch?.paymentMethod) return amountMatch.paymentMethod;

    const composite = ncStripeComposite?.get(`${orderKey}|${absAmount}`);
    if (composite?.method) return composite.method;
  }

  return rows[0]?.paymentMethod ?? null;
}

function buildNcMap(ncRows, layout) {
  const { col: ncCol, headerRows } = layout;
  const map = new Map();

  for (let i = headerRows; i < ncRows.length; i += 1) {
    const docNo = String(ncRows[i][ncCol.DOC_NO] ?? '').trim();
    if (!isNcValidDocNo(docNo)) continue;
    map.set(docNo, ncRows[i]);
  }

  return map;
}

function finalizeSheetRange(sheet, rowCount) {
  sheet['!ref'] = `A1:O${rowCount}`;
}

function copySkOrdersSheet(skBuffer) {
  const workbook = XLSX.read(skBuffer, { type: 'buffer', ...XLSX_READ_OPTS });
  const sheetName = findSheet(workbook, /orders/i);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet?.['!ref']) {
    throw new Error('No Orders worksheet found in SK Combined file');
  }

  return {
    sheetName,
    sheet,
    lastRow: XLSX.utils.decode_range(sheet['!ref']).e.r + 1,
  };
}

function buildEskSheetTitle() {
  return 'Payment Listing-Detail _sort';
}

function buildEskOutputFilename() {
  return 'ESK Payment Details.xlsx';
}

function compareDetailEntries(a, b, ncCol) {
  const dateA = parseNumber(a.ncRow[ncCol.HOLDER]);
  const dateB = parseNumber(b.ncRow[ncCol.HOLDER]);
  const safeDateA = Number.isFinite(dateA) ? dateA : Number.MAX_SAFE_INTEGER;
  const safeDateB = Number.isFinite(dateB) ? dateB : Number.MAX_SAFE_INTEGER;
  if (safeDateA !== safeDateB) return safeDateA - safeDateB;

  return String(a.ncRow[ncCol.DOC_NO] ?? '').localeCompare(String(b.ncRow[ncCol.DOC_NO] ?? ''));
}

function detailOrdersLookupRange(lastRow) {
  return `Orders!$A$1:$Q$${Math.max(lastRow, 2)}`;
}

function writeDetailRow(
  sheet,
  excelRow,
  ncRow,
  layout,
  expansion,
  skLookup,
  ordersLookupRange,
  ncStripeComposite,
  usageTracker,
) {
  const { col: ncCol } = layout;
  const note = expansion?.note ?? formatNoteFromNc(ncRow[ncCol.NOTE]);
  const orderKey = expansion?.orderKey ?? extractOrderKey(note);
  const rowAmount = expansion?.amount ?? ncRow[ncCol.AMT];
  const skMatch = resolveSkMatch(orderKey, skLookup, rowAmount, {
    paymentType: ncRow[ncCol.PAYMENT_TYPE],
    note,
  }, usageTracker);

  setCell(sheet, `H${excelRow}`, note);
  setCell(sheet, `A${excelRow}`, {
    f: `LEFT(H${excelRow},13)`,
    v: orderKey,
    t: 's',
  });
  setCell(sheet, `B${excelRow}`, {
    f: `LEFT(H${excelRow},6)`,
    v: orderKey ? orderKey.slice(0, 6) : '',
    t: 's',
  });

  setDateCell(sheet, `C${excelRow}`, ncRow[ncCol.HOLDER]);
  setCell(sheet, `D${excelRow}`, ncRow[ncCol.DOC_NO]);
  setCell(sheet, `E${excelRow}`, ncRow[ncCol.RETAIL_INVOICE]);
  setCell(sheet, `F${excelRow}`, ncRow[ncCol.PAYMENT_TYPE]);
  setAmountCell(sheet, `G${excelRow}`, rowAmount);
  setCell(sheet, `I${excelRow}`, ncRow[ncCol.DEPOSIT_REFUND] ?? '');
  setCell(sheet, `J${excelRow}`, ncRow[ncCol.NET_TOTAL] ?? rowAmount);
  setCell(sheet, `K${excelRow}`, ncRow[ncCol.COSTCODE] ?? '');

  const hValue = resolveHLookupValue(orderKey, rowAmount, skLookup, ncStripeComposite, note);
  const paymentMethod =
    lookupSkPaymentMethod(orderKey, skLookup, rowAmount, ncStripeComposite, note)
    ?? skMatch?.paymentMethod
    ?? null;
  const amount = parseNumber(rowAmount);
  const diff = !Number.isNaN(amount) && hValue != null && !Number.isNaN(hValue)
    ? amount - hValue
    : null;

  setCell(sheet, `L${excelRow}`, {
    f: `VLOOKUP(A${excelRow},${ordersLookupRange},16,0)`,
    v: hValue ?? 0,
    t: 'n',
  });
  setCell(sheet, `M${excelRow}`, {
    f: `G${excelRow}-L${excelRow}`,
    v: diff,
    t: 'n',
  });
  setCell(sheet, `N${excelRow}`, {
    f: `VLOOKUP(A${excelRow},${ordersLookupRange},17,0)`,
    v: paymentMethod ?? '',
    t: 's',
  });
}

function buildDetailSheetFromNcRows(
  ncRows,
  skLookup,
  layout,
  { ncStripeComposite = null, ordersLastRow = 0 } = {},
) {
  const { col: ncCol, headerRows: dataStartRow } = layout;
  const sheet = {};
  const usageTracker = createSkUsageTracker(skLookup);
  const ordersLookupRange = detailOrdersLookupRange(ordersLastRow);
  const entries = [];

  ESK_DETAIL_HEADERS.forEach((header, columnIndex) => {
    setCell(sheet, XLSX.utils.encode_cell({ r: 0, c: columnIndex }), header);
  });

  for (let i = dataStartRow; i < ncRows.length; i += 1) {
    const ncRow = ncRows[i];
    if (isNcDateSeparatorRow(ncRow, layout) || isNcSummaryRow(ncRow, layout)) continue;

    const docNo = String(ncRow[ncCol.DOC_NO] ?? '').trim();
    if (!isNcValidDocNo(docNo)) continue;

    const expansions = expandNcRow(ncRow, skLookup, layout);
    for (const expansion of expansions) {
      entries.push({ ncRow, expansion });
    }
  }

  entries.sort((a, b) => compareDetailEntries(a, b, ncCol));

  let excelRow = 1;
  for (const entry of entries) {
    excelRow += 1;
    writeDetailRow(
      sheet,
      excelRow,
      entry.ncRow,
      layout,
      entry.expansion,
      skLookup,
      ordersLookupRange,
      ncStripeComposite,
      usageTracker,
    );
  }

  finalizeSheetRange(sheet, Math.max(excelRow, 1));
  return { sheet, generated: entries.length, rowCount: excelRow };
}

async function createEskPaymentWorkbook(skBuffer, ncBuffer, { ncFileName = null } = {}) {
  const ncWorkbook = XLSX.read(ncBuffer, { type: 'buffer', ...XLSX_READ_OPTS });
  const ncSheetName = findSheet(ncWorkbook, /payment listing/i);
  const ncSheet = ncWorkbook.Sheets[ncSheetName];
  if (!ncSheet) {
    throw new Error('No worksheet found in NC payment listing file');
  }

  const ncRows = XLSX.utils.sheet_to_json(ncSheet, { header: 1, defval: '' });
  const ncLayout = resolveNcLayout(ncRows[0]);

  if (ncRows.length <= ncLayout.headerRows) {
    throw new Error('NC payment listing file has no data rows');
  }

  const listingDate = detectNcListingDate(ncWorkbook, ncRows, ncLayout, ncFileName);
  const ncMap = buildNcMap(ncRows, ncLayout);
  const { lookup: skLookup, sheetName: skSheetName } = buildProductionSkLookup(skBuffer);
  const ncStripeComposite = buildNcStripeCompositeMap(ncRows, skLookup, ncLayout);
  const { sheet: ordersSheet, lastRow: ordersLastRow } = copySkOrdersSheet(skBuffer);

  const sheetTitle = buildEskSheetTitle();
  const { sheet: mainSheet, generated } = buildDetailSheetFromNcRows(
    ncRows,
    skLookup,
    ncLayout,
    { ncStripeComposite, ordersLastRow },
  );

  const workbook = {
    SheetNames: [sheetTitle, 'Orders'],
    Sheets: {
      [sheetTitle]: mainSheet,
      Orders: ordersSheet,
    },
  };

  let buffer = XLSX.write(workbook, {
    type: 'buffer',
    bookType: 'xlsx',
    compression: true,
    cellStyles: true,
    cellFormula: true,
  });
  buffer = await applyColumnBDateStyles(buffer);
  const filename = buildEskOutputFilename();

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
