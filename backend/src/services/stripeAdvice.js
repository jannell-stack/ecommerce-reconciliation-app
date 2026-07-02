const XLSX = require('xlsx');
const { applyStripeAdviceWorkbookStyles } = require('./eskPaymentStyles');
const {
  detectNcListingPeriod,
  extractNcOrderKey,
  isExcelDateSerial,
  isNcFooterRow,
  isNcSummaryRow,
  resolveNcLayout,
} = require('../utils/ncLayout');

const ITEMIZED_COL = {
  PAYOUT_EFFECTIVE: 1,
  CREATED: 3,
  AVAILABLE_ON: 4,
  CURRENCY: 5,
  GROSS: 6,
  FEE: 7,
  NET: 8,
  REPORTING_CATEGORY: 9,
  DESCRIPTION: 10,
  ORDER_ID: 11,
};

const STRIPE_ADVICE_HEADERS = [
  'automatic_payout_effective_at',
  'payment_metadata[order_id]',
  'Gross',
  'Comm',
  'Nett',
  'Inv dd',
  'Inv amt',
  'diff',
];

const STRIPE_ADVICE_SHEET_NAME = 'Stripe Advice';
const ESK_LISTING_SHEET_NAME = 'Payment Listing-Detail _sort';
const ESK_ORDERS_SHEET_NAME = 'Orders';
const STRIPE_ADVICE_FILENAME = 'Stripe Advice.xlsx';
const STRIPE_INV_DATE_FORMAT = 'm/d/yy';
const STRIPE_PAYOUT_DATE_FORMAT = 'm/d/yy';
const STRIPE_AMOUNT_FORMAT = '#,##0.00';
const LISTING_LOOKUP_MAX_COL = 'N';
const LISTING_DATE_FORMAT = 'm/d/yy';

function findItemizedSheet(workbook) {
  return (
    workbook.SheetNames.find((name) => /itemized_payout_reconciliation/i.test(name)) ||
    workbook.SheetNames[0]
  );
}

function findNcSheet(workbook, sheetName) {
  if (sheetName && workbook.SheetNames.includes(sheetName)) {
    return sheetName;
  }

  const listingSheet = workbook.SheetNames.find(
    (name) => /payment listing|esk payment details/i.test(name) && !/^orders$/i.test(name),
  );
  if (listingSheet) return listingSheet;

  return workbook.SheetNames.find((name) => !/^orders$/i.test(name)) || workbook.SheetNames[0];
}

function extractOrderKey(noteText) {
  const text = String(noteText ?? '').trim();
  const match = text.match(/(\d{6}-\d{5,6})/);
  if (match) return match[1];
  return text.slice(0, 13).trim();
}

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function retailInvoiceSerial(invoiceNo) {
  const match = String(invoiceNo ?? '').match(/ESK(\d{4})(\d{2})(\d{2})/);
  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const epoch = Date.UTC(1899, 11, 30);
  const serial = Math.round((Date.UTC(year, month - 1, day) - epoch) / 86400000);
  return Number.isFinite(serial) ? serial : null;
}

function resolveNcDocdateSerial(ncRow, ncCol) {
  const holder = ncRow[ncCol.HOLDER];
  if (isExcelDateSerial(holder)) {
    return Math.floor(holder) === holder ? holder : holder;
  }
  return retailInvoiceSerial(ncRow[ncCol.RETAIL_INVOICE]);
}

function resolveNcInvAmount(ncRow, ncCol, layout) {
  if (layout.format === 'eskPayment' && ncCol.FINAL_PRICE_PAID != null) {
    const fromColumnL = parseNumber(ncRow[ncCol.FINAL_PRICE_PAID]);
    if (Number.isFinite(fromColumnL)) return fromColumnL;
  }
  return parseNumber(ncRow[ncCol.AMT]);
}

function buildNcStripeLookupMap(ncRows, layout) {
  const { col: ncCol, headerRows } = layout;
  const map = new Map();

  for (let i = headerRows; i < ncRows.length; i += 1) {
    const row = ncRows[i];
    if (!row?.length || isNcFooterRow(row) || isNcSummaryRow(row, layout)) continue;

    const paymentType = String(row[ncCol.PAYMENT_TYPE] ?? '').trim().toUpperCase();
    if (paymentType !== 'STRIPE') continue;

    const orderKey = extractNcOrderKey(row, layout);
    const invAmount = resolveNcInvAmount(row, ncCol, layout);
    const docdateSerial = resolveNcDocdateSerial(row, ncCol);
    if (!orderKey || Number.isNaN(invAmount)) continue;

    if (map.has(orderKey)) continue;

    map.set(orderKey, {
      docdateSerial: docdateSerial ?? '',
      invAmount,
    });
  }

  return map;
}

function buildTopUpListingLookup(ncRows, layout) {
  const { col: ncCol, headerRows } = layout;
  const byAmount = new Map();
  const bySix = new Map();

  for (let i = headerRows; i < ncRows.length; i += 1) {
    const row = ncRows[i];
    if (!row?.length || isNcFooterRow(row) || isNcSummaryRow(row, layout)) continue;

    const orderKey = extractNcOrderKey(row, layout);
    if (!/^\d{6}-\d{5,6}$/.test(orderKey)) continue;

    const amount = parseNumber(row[ncCol.AMT]);
    if (Number.isNaN(amount)) continue;

    const six = orderKey.slice(0, 6);
    const docdateSerial = resolveNcDocdateSerial(row, ncCol) ?? '';

    if (!bySix.has(six)) {
      bySix.set(six, {
        docdateSerial,
        invAmount: resolveNcInvAmount(row, ncCol, layout),
      });
    }

    const lookupKey = `${six}|${amount}`;
    if (byAmount.has(lookupKey)) continue;

    byAmount.set(lookupKey, {
      docdateSerial,
      invAmount: resolveNcInvAmount(row, ncCol, layout),
    });
  }

  return { byAmount, bySix };
}

function mergeTopUpLookups(ncSources) {
  const byAmount = new Map();
  const bySix = new Map();

  for (const source of ncSources) {
    for (const [key, value] of source.topUpLookup.byAmount) {
      if (!byAmount.has(key)) {
        byAmount.set(key, value);
      }
    }
    for (const [key, value] of source.topUpLookup.bySix) {
      if (!bySix.has(key)) {
        bySix.set(key, value);
      }
    }
  }

  return { byAmount, bySix };
}

function resolveTopUpMatch(orderId, topUpLookup) {
  return topUpLookup.bySix.get(orderId) ?? null;
}

function isFullOrderId(value) {
  return /^\d{6}-\d{5,6}$/.test(String(value ?? '').trim());
}

function isStripeTopUpOrder(row) {
  return /top\s*up\s*order/i.test(String(row[ITEMIZED_COL.DESCRIPTION] ?? ''));
}

function resolveStripeFullOrderId(row) {
  const metadataOrderId = String(row[ITEMIZED_COL.ORDER_ID] ?? '').trim();
  const descriptionOrderId = extractOrderKey(row[ITEMIZED_COL.DESCRIPTION]);

  if (isFullOrderId(descriptionOrderId)) {
    return descriptionOrderId;
  }
  if (isFullOrderId(metadataOrderId)) {
    return metadataOrderId;
  }
  if (metadataOrderId) return metadataOrderId;
  return descriptionOrderId;
}

function resolveStripeAdviceOrderId(row) {
  const fullOrderId = resolveStripeFullOrderId(row);
  if (isStripeTopUpOrder(row) && isFullOrderId(fullOrderId)) {
    return fullOrderId.slice(0, 6);
  }
  return fullOrderId;
}

function payoutEffectiveValue(row) {
  const effective = row[ITEMIZED_COL.PAYOUT_EFFECTIVE];
  if (typeof effective === 'number' && Number.isFinite(effective)) {
    return effective;
  }

  const available = row[ITEMIZED_COL.AVAILABLE_ON];
  if (typeof available === 'number' && Number.isFinite(available)) {
    return available;
  }

  return '';
}

function resolveNcStripeMatch(orderId, ncLookup) {
  return ncLookup.get(String(orderId ?? '').trim()) || null;
}

function shouldIncludeItemizedRow(row) {
  const category = String(row[ITEMIZED_COL.REPORTING_CATEGORY] ?? '')
    .trim()
    .toLowerCase();
  return category !== 'fee';
}

function buildAdviceRow(row, ncLookup, topUpLookup) {
  const fullOrderId = resolveStripeFullOrderId(row);
  const isTopUp = isStripeTopUpOrder(row);
  const orderId = resolveStripeAdviceOrderId(row);
  const gross = parseNumber(row[ITEMIZED_COL.GROSS]);
  const fee = parseNumber(row[ITEMIZED_COL.FEE]);
  const net = parseNumber(row[ITEMIZED_COL.NET]);
  let ncMatch = null;

  if (isTopUp) {
    ncMatch = resolveTopUpMatch(orderId, topUpLookup);
  } else {
    ncMatch = resolveNcStripeMatch(fullOrderId, ncLookup);
  }

  return {
    payoutEffective: payoutEffectiveValue(row),
    orderId,
    isTopUp,
    gross: Number.isNaN(gross) ? '' : gross,
    comm: Number.isNaN(fee) ? '' : Math.abs(fee),
    nett: Number.isNaN(net) ? '' : net,
    invDateSerial: ncMatch?.docdateSerial ?? '',
    invAmount: ncMatch?.invAmount ?? '',
  };
}

function setCell(sheet, address, value) {
  if (value === '' || value == null) return;
  sheet[address] = {
    t: typeof value === 'number' ? 'n' : 's',
    v: value,
  };
}

function setAmountCell(sheet, address, value, formula) {
  if ((value === '' || value == null) && !formula) return;
  const numeric = parseNumber(value);
  const cell = {
    t: Number.isNaN(numeric) ? 's' : 'n',
    v: Number.isNaN(numeric) ? value : numeric,
    z: STRIPE_AMOUNT_FORMAT,
  };
  if (formula) cell.f = formula;
  sheet[address] = cell;
}

function setPayoutEffectiveCell(sheet, address, value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    setCell(sheet, address, value);
    return;
  }
  sheet[address] = {
    t: 'n',
    v: value,
    z: STRIPE_PAYOUT_DATE_FORMAT,
  };
}

function setDateCell(sheet, address, serial, formula) {
  if (!formula && !isExcelDateSerial(serial)) return;
  const cell = {
    t: 'n',
    v: isExcelDateSerial(serial) ? serial : 0,
    z: STRIPE_INV_DATE_FORMAT,
  };
  if (formula) cell.f = formula;
  sheet[address] = cell;
}

function listingLookupRange(lastRow) {
  return `'${ESK_LISTING_SHEET_NAME}'!$A$1:$${LISTING_LOOKUP_MAX_COL}$${Math.max(lastRow, 2)}`;
}

function listingColumnRange(column, lastRow) {
  return `'${ESK_LISTING_SHEET_NAME}'!$${column}$1:$${column}$${Math.max(lastRow, 2)}`;
}

function listingLookupRangeFromColumn(startColumn, lastRow) {
  return `'${ESK_LISTING_SHEET_NAME}'!$${startColumn}$1:$${LISTING_LOOKUP_MAX_COL}$${Math.max(lastRow, 2)}`;
}

function topUpListingLookupRange(lastRow) {
  return listingLookupRangeFromColumn('B', lastRow);
}

function topUpAmountMatchFormula(excelRow, listingLastRow) {
  const bRange = listingColumnRange('B', listingLastRow);
  const gRange = listingColumnRange('G', listingLastRow);
  return `COUNTIFS(${bRange},B${excelRow},${gRange},C${excelRow})`;
}

function invDateFormula(excelRow, listingLastRow, isTopUp) {
  if (isTopUp) {
    const bRange = listingColumnRange('B', listingLastRow);
    const cRange = listingColumnRange('C', listingLastRow);
    const gRange = listingColumnRange('G', listingLastRow);
    const amountMatch = topUpAmountMatchFormula(excelRow, listingLastRow);
    const sumifs = `SUMIFS(${cRange},${bRange},B${excelRow},${gRange},C${excelRow})`;
    const fallback = `VLOOKUP(B${excelRow},${listingLookupRangeFromColumn('B', listingLastRow)},2,0)`;
    return `IF(${amountMatch}=0,${fallback},${sumifs})`;
  }
  return `VLOOKUP(B${excelRow},${listingLookupRange(listingLastRow)},3,0)`;
}

function invAmountFormula(excelRow, listingLastRow, isTopUp) {
  if (isTopUp) {
    return `VLOOKUP(B${excelRow},${topUpListingLookupRange(listingLastRow)},11,0)`;
  }
  return `VLOOKUP(B${excelRow},${listingLookupRange(listingLastRow)},12,0)`;
}

function diffFormula(excelRow) {
  return `G${excelRow}-C${excelRow}`;
}

function buildStripeAdviceSheet(dataRows, listingLastRow) {
  const sheet = {};
  const headerRow = 3;
  const dataStartRow = 4;

  STRIPE_ADVICE_HEADERS.forEach((header, columnIndex) => {
    const address = XLSX.utils.encode_cell({ r: headerRow - 1, c: columnIndex });
    setCell(sheet, address, header);
    if (columnIndex === 5) {
      sheet[address].z = STRIPE_INV_DATE_FORMAT;
    } else if (columnIndex === 2 || columnIndex === 3 || columnIndex === 4 || columnIndex >= 6) {
      sheet[address].z = STRIPE_AMOUNT_FORMAT;
    }
  });

  dataRows.forEach((row, rowIndex) => {
    const excelRow = dataStartRow + rowIndex;

    setPayoutEffectiveCell(sheet, `A${excelRow}`, row.payoutEffective);
    setCell(sheet, `B${excelRow}`, row.orderId);
    setAmountCell(sheet, `C${excelRow}`, row.gross);
    setAmountCell(sheet, `D${excelRow}`, row.comm);
    setAmountCell(sheet, `E${excelRow}`, row.nett);
    setDateCell(
      sheet,
      `F${excelRow}`,
      row.invDateSerial,
      invDateFormula(excelRow, listingLastRow, row.isTopUp),
    );
    setAmountCell(
      sheet,
      `G${excelRow}`,
      row.invAmount,
      invAmountFormula(excelRow, listingLastRow, row.isTopUp),
    );
    const diffValue = !Number.isNaN(parseNumber(row.gross)) && Number.isFinite(row.invAmount)
      ? row.invAmount - row.gross
      : '';
    setAmountCell(sheet, `H${excelRow}`, diffValue, diffFormula(excelRow));
  });

  const lastRow = Math.max(dataStartRow + dataRows.length - 1, headerRow);
  sheet['!ref'] = `A2:H${lastRow}`;
  return sheet;
}

function cloneSheet(sheet) {
  const cloned = {};
  for (const [address, cell] of Object.entries(sheet)) {
    if (address.startsWith('!')) {
      cloned[address] = sheet[address];
      continue;
    }
    const copied = { ...cell };
    if (copied.t === 'z') {
      if (copied.f) {
        copied.t = typeof copied.v === 'number' ? 'n' : 's';
        if (copied.t === 's' && (copied.v == null || copied.v === '')) {
          copied.v = '';
        }
      } else {
        delete copied.t;
      }
    }
    cloned[address] = copied;
  }
  return cloned;
}

function listingOrdersLookupRange(ordersLastRow) {
  return `Orders!$A$1:$Q$${Math.max(ordersLastRow, 2)}`;
}

function normalizeEmbeddedListingSheet(sheet, ordersLastRow) {
  if (!sheet?.['!ref']) return sheet;

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const ordersLookupRange = listingOrdersLookupRange(ordersLastRow);

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const excelRow = rowIndex + 1;
    const addressC = XLSX.utils.encode_cell({ r: rowIndex, c: 2 });
    const cellC = sheet[addressC];
    if (cellC && isExcelDateSerial(cellC.v)) {
      sheet[addressC] = {
        ...cellC,
        t: 'n',
        z: LISTING_DATE_FORMAT,
      };
    }

    if (rowIndex === 0) continue;

    const addressA = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
    const addressG = XLSX.utils.encode_cell({ r: rowIndex, c: 6 });
    const hasRowData = Boolean(sheet[addressA]?.v || sheet[addressA]?.f || sheet[addressG]?.v);
    if (!hasRowData) continue;

    const addressL = XLSX.utils.encode_cell({ r: rowIndex, c: 11 });
    const addressM = XLSX.utils.encode_cell({ r: rowIndex, c: 12 });
    const addressN = XLSX.utils.encode_cell({ r: rowIndex, c: 13 });

    if (!sheet[addressL]?.f) {
      sheet[addressL] = {
        ...(sheet[addressL] || {}),
        t: 'n',
        f: `VLOOKUP(A${excelRow},${ordersLookupRange},16,0)`,
        v: sheet[addressL]?.v ?? 0,
      };
    }

    if (!sheet[addressM]?.f) {
      const gross = parseNumber(sheet[addressG]?.v);
      const finalPrice = parseNumber(sheet[addressL]?.v);
      sheet[addressM] = {
        ...(sheet[addressM] || {}),
        t: 'n',
        f: `G${excelRow}-L${excelRow}`,
        v: !Number.isNaN(gross) && !Number.isNaN(finalPrice) ? gross - finalPrice : sheet[addressM]?.v ?? 0,
      };
    } else if (sheet[addressM].t === 'z') {
      sheet[addressM] = {
        ...sheet[addressM],
        t: typeof sheet[addressM].v === 'number' ? 'n' : 's',
        v: sheet[addressM].v ?? 0,
      };
    }

    if (!sheet[addressN]?.f) {
      sheet[addressN] = {
        ...(sheet[addressN] || {}),
        t: 's',
        f: `VLOOKUP(A${excelRow},${ordersLookupRange},17,0)`,
        v: sheet[addressN]?.v ?? '',
      };
    }
  }

  return sheet;
}

function sheetLastRow(sheet) {
  if (!sheet?.['!ref']) return 1;
  return XLSX.utils.decode_range(sheet['!ref']).e.r + 1;
}

function buildMergedListingSheet(sources) {
  if (sources.length === 1) {
    return cloneSheet(sources[0].listingSheet);
  }

  const merged = {};
  let writeRow = 0;

  for (const source of sources) {
    const rows = XLSX.utils.sheet_to_json(source.listingSheet, { header: 1, defval: '' });
    const startIndex = writeRow === 0 ? 0 : 1;

    for (let i = startIndex; i < rows.length; i += 1) {
      const row = rows[i];
      if (!row?.length) continue;
      row.forEach((value, columnIndex) => {
        if (value === '' || value == null) return;
        const address = XLSX.utils.encode_cell({ r: writeRow, c: columnIndex });
        merged[address] = {
          t: typeof value === 'number' ? 'n' : 's',
          v: value,
        };
      });
      writeRow += 1;
    }
  }

  merged['!ref'] = `A1:${LISTING_LOOKUP_MAX_COL}${Math.max(writeRow, 1)}`;
  return merged;
}

function buildMergedOrdersSheet(sources) {
  const withOrders = sources.filter((source) => source.ordersSheet);
  if (withOrders.length === 0) return null;

  if (withOrders.length === 1) {
    return cloneSheet(withOrders[0].ordersSheet);
  }

  const merged = {};
  let writeRow = 0;

  for (const source of withOrders) {
    const rows = XLSX.utils.sheet_to_json(source.ordersSheet, { header: 1, defval: '' });
    const startIndex = writeRow === 0 ? 0 : 1;

    for (let i = startIndex; i < rows.length; i += 1) {
      const row = rows[i];
      if (!row?.length) continue;
      row.forEach((value, columnIndex) => {
        if (value === '' || value == null) return;
        const address = XLSX.utils.encode_cell({ r: writeRow, c: columnIndex });
        merged[address] = {
          t: typeof value === 'number' ? 'n' : 's',
          v: value,
        };
      });
      writeRow += 1;
    }
  }

  merged['!ref'] = `A1:AJ${Math.max(writeRow, 1)}`;
  return merged;
}

function sortNcSourcesByPeriod(ncSources) {
  return [...ncSources].sort((a, b) => {
    const sortA = a.period?.sortKey ?? Number.MAX_SAFE_INTEGER;
    const sortB = b.period?.sortKey ?? Number.MAX_SAFE_INTEGER;
    return sortA - sortB;
  });
}

function normalizeNcInputs(ncInput, options = {}) {
  const entries = Array.isArray(ncInput) ? ncInput : [ncInput];
  const fileNames = options.ncFileNames || (options.ncFileName ? [options.ncFileName] : []);

  return entries.map((entry, index) => {
    if (entry?.buffer) {
      return {
        buffer: entry.buffer,
        fileName: entry.fileName || fileNames[index] || null,
        sheetName: entry.sheetName || options.ncSheetName || null,
      };
    }

    return {
      buffer: entry,
      fileName: fileNames[index] || options.ncFileName || null,
      sheetName: options.ncSheetName || null,
    };
  });
}

function loadNcStripeSource({ buffer, fileName, sheetName }) {
  const ncWorkbook = XLSX.read(buffer, { type: 'buffer', cellFormula: true });
  const ncSheetNameResolved = findNcSheet(ncWorkbook, sheetName);
  const ncSheet = ncWorkbook.Sheets[ncSheetNameResolved];
  const ordersSheetName = ncWorkbook.SheetNames.find((name) => /^orders$/i.test(name));
  const ordersSheet = ordersSheetName ? ncWorkbook.Sheets[ordersSheetName] : null;

  if (!ncSheet) {
    const label = fileName || ncSheetNameResolved || 'NC file';
    throw new Error(`NC or ESK Payment Details sheet not found in ${label}`);
  }

  const ncRows = XLSX.utils.sheet_to_json(ncSheet, { header: 1, defval: '' });
  const layout = resolveNcLayout(ncRows[0]);
  const period = detectNcListingPeriod(ncWorkbook, ncRows, layout, fileName);

  return {
    map: buildNcStripeLookupMap(ncRows, layout),
    topUpLookup: buildTopUpListingLookup(ncRows, layout),
    listingSheet: ncSheet,
    ordersSheet,
    ordersSheetName: ordersSheetName || null,
    sheetName: ncSheetNameResolved,
    fileName,
    period,
  };
}

function mergeNcStripeLookups(ncSources) {
  const merged = new Map();

  for (const source of ncSources) {
    for (const [key, value] of source.map) {
      if (!merged.has(key)) {
        merged.set(key, value);
      }
    }
  }

  return merged;
}

async function createStripeAdviceWorkbook(
  stripeBuffer,
  ncInput,
  { ncSheetName, ncFileName, ncFileNames, stripeFileName } = {},
) {
  const ncFiles = normalizeNcInputs(ncInput, { ncSheetName, ncFileName, ncFileNames });

  if (ncFiles.length === 0) {
    throw new Error('At least one NC or ESK Payment Details file is required');
  }

  const stripeWorkbook = XLSX.read(stripeBuffer, { type: 'buffer' });
  const itemizedSheetName = findItemizedSheet(stripeWorkbook);
  const itemizedSheet = stripeWorkbook.Sheets[itemizedSheetName];

  if (!itemizedSheet) {
    throw new Error('Itemized_payout_reconciliation sheet not found in Stripe file');
  }

  const ncSources = sortNcSourcesByPeriod(ncFiles.map((file) => loadNcStripeSource(file)));
  const ncLookup = mergeNcStripeLookups(ncSources);
  const topUpLookup = mergeTopUpLookups(ncSources);
  const ordersSheet = buildMergedOrdersSheet(ncSources);
  const ordersLastRow = ordersSheet ? sheetLastRow(ordersSheet) : 0;
  const listingSheet = normalizeEmbeddedListingSheet(
    buildMergedListingSheet(ncSources),
    ordersLastRow,
  );
  const listingLastRow = sheetLastRow(listingSheet);
  const itemizedRows = XLSX.utils.sheet_to_json(itemizedSheet, { header: 1, defval: '' });

  const outputRows = [];
  let generated = 0;
  let linked = 0;

  for (let i = 1; i < itemizedRows.length; i += 1) {
    const row = itemizedRows[i];
    if (!row?.length || !shouldIncludeItemizedRow(row)) continue;

    const fullOrderId = resolveStripeFullOrderId(row);
    const adviceOrderId = resolveStripeAdviceOrderId(row);
    const gross = parseNumber(row[ITEMIZED_COL.GROSS]);
    const isTopUp = isStripeTopUpOrder(row);
    const isLinked = isTopUp
      ? Boolean(resolveTopUpMatch(adviceOrderId, topUpLookup))
      : Boolean(fullOrderId && ncLookup.has(fullOrderId));
    if (isLinked) {
      linked += 1;
    }

    outputRows.push(buildAdviceRow(row, ncLookup, topUpLookup));
    generated += 1;
  }

  if (generated === 0) {
    throw new Error('No billable rows found in Itemized_payout_reconciliation sheet');
  }

  const worksheet = buildStripeAdviceSheet(outputRows, listingLastRow);
  const sheetNames = [STRIPE_ADVICE_SHEET_NAME, ESK_LISTING_SHEET_NAME];
  const sheets = {
    [STRIPE_ADVICE_SHEET_NAME]: worksheet,
    [ESK_LISTING_SHEET_NAME]: listingSheet,
  };

  if (ordersSheet) {
    sheetNames.push(ESK_ORDERS_SHEET_NAME);
    sheets[ESK_ORDERS_SHEET_NAME] = ordersSheet;
  }

  const workbook = {
    SheetNames: sheetNames,
    Sheets: sheets,
  };

  let buffer = XLSX.write(workbook, {
    type: 'buffer',
    bookType: 'xlsx',
    cellStyles: true,
    cellFormula: true,
  });
  buffer = await applyStripeAdviceWorkbookStyles(buffer);

  return {
    buffer,
    filename: STRIPE_ADVICE_FILENAME,
    sheetName: STRIPE_ADVICE_SHEET_NAME,
    generated,
    linked,
    itemizedSheetName,
    ncSheetNames: ncSources.map((source) => source.sheetName),
    ncFileNames: ncSources.map((source) => source.fileName).filter(Boolean),
    stripeFileName,
  };
}

module.exports = {
  createStripeAdviceWorkbook,
  STRIPE_ADVICE_HEADERS,
};
