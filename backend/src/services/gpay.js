const XLSX = require('xlsx');
const { isExcelDateSerial } = require('./eskPaymentStyles');
const {
  ESK_LISTING_SHEET_NAME,
  ESK_ORDERS_SHEET_NAME,
  buildMergedListingSheet,
  buildMergedOrdersSheet,
  listingLookupRange,
  loadEskPaymentSource,
  mergeEskOrderIdSets,
  normalizeEmbeddedListingSheet,
  normalizeNcInputs,
  sheetLastRow,
  sortNcSourcesByPeriod,
} = require('./eskListingWorkbook');
const { applyGpayWorkbookStyles } = require('./eskPaymentStyles');

const GPAY_HEADERS = [
  'Category',
  'Partner transaction ID 2',
  'Payment Method',
  'Settlement ID',
  'Transfer Date',
  'Amount',
  'Offer',
  'Net Sales',
  'Net MDR',
  'Tax on MDR',
  'Ttl Comm',
  'Total',
  'Inv dd',
  'Inv amt',
  'diff',
];

const GPAY_SHEET_NAME = 'GPay';
const GPAY_FILENAME = 'GPay.xlsx';
const GPAY_AMOUNT_FORMAT = '#,##0.00';
const GPAY_INV_DATE_FORMAT = 'm/d/yy';

const GRAB_HEADER = {
  CATEGORY: 'Category',
  STATUS: 'Status',
  PARTNER_TXN_ID_2: 'Partner transaction ID 2',
  PAYMENT_METHOD: 'Payment Method',
  SETTLEMENT_ID: 'Settlement ID',
  TRANSFER_DATE: 'Transfer Date',
  AMOUNT: 'Amount',
  OFFER: 'Offer',
  NET_SALES: 'Net Sales',
  NET_MDR: 'Net MDR',
  TAX_ON_MDR: 'Tax on MDR',
  TOTAL: 'Total',
};

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function findGrabTransactionSheet(workbook) {
  return (
    workbook.SheetNames.find((name) => /transaction_store|grab transaction/i.test(name))
    || workbook.SheetNames[0]
  );
}

function resolveGrabColumns(headerRow) {
  const headers = (headerRow || []).map((cell) => String(cell ?? '').trim());
  const indexOf = (name) => headers.findIndex((header) => header === name);

  const col = {};
  for (const [key, name] of Object.entries(GRAB_HEADER)) {
    const index = indexOf(name);
    if (index < 0) {
      throw new Error(`Grab Transaction file is missing column "${name}"`);
    }
    col[key] = index;
  }
  return col;
}

function formatGrabTransferDate(value) {
  if (isExcelDateSerial(value)) {
    const epoch = Date.UTC(1899, 11, 30);
    const date = new Date(epoch + Math.floor(value) * 86400000);
    const day = String(date.getUTCDate()).padStart(2, '0');
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const year = date.getUTCFullYear();
    return `${day}/${month}/${year}`;
  }

  return String(value ?? '').trim();
}

function isFullOrderId(value) {
  return /^\d{6}-\d{5,6}$/.test(String(value ?? '').trim());
}

const GRAB_STATUS = {
  TRANSFERRED: 'transferred',
  FAILED: 'failed',
};

function normalizeGrabCellText(value) {
  return String(value ?? '').trim();
}

function normalizeGrabStatus(value) {
  return normalizeGrabCellText(value).toLowerCase();
}

function shouldIncludeGrabRow(row, grabCol) {
  const status = normalizeGrabStatus(row[grabCol.STATUS]);
  if (status === GRAB_STATUS.FAILED) return false;
  if (status !== GRAB_STATUS.TRANSFERRED) return false;

  const orderId = normalizeGrabCellText(row[grabCol.PARTNER_TXN_ID_2]);
  return isFullOrderId(orderId);
}

function ttlCommFormula(excelRow) {
  return `H${excelRow}-L${excelRow}`;
}

function invDateFormula(excelRow, listingLastRow) {
  return `VLOOKUP(B${excelRow},${listingLookupRange(listingLastRow)},3,0)`;
}

function invAmountFormula(excelRow, listingLastRow) {
  return `VLOOKUP(B${excelRow},${listingLookupRange(listingLastRow)},12,0)`;
}

function diffFormula(excelRow) {
  return `N${excelRow}-H${excelRow}`;
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
    z: GPAY_AMOUNT_FORMAT,
  };
  if (formula) cell.f = formula;
  sheet[address] = cell;
}

function setDateCell(sheet, address, serial, formula) {
  if (!formula && !isExcelDateSerial(serial)) return;
  const cell = {
    t: 'n',
    v: isExcelDateSerial(serial) ? serial : 0,
    z: GPAY_INV_DATE_FORMAT,
  };
  if (formula) cell.f = formula;
  sheet[address] = cell;
}

function buildGpayRow(row, grabCol) {
  const netSales = parseNumber(row[grabCol.NET_SALES]);
  const total = parseNumber(row[grabCol.TOTAL]);
  const ttlComm = !Number.isNaN(netSales) && !Number.isNaN(total) ? netSales - total : '';

  return {
    category: String(row[grabCol.CATEGORY] ?? ''),
    orderId: String(row[grabCol.PARTNER_TXN_ID_2] ?? '').trim(),
    paymentMethod: String(row[grabCol.PAYMENT_METHOD] ?? '').trim(),
    settlementId: String(row[grabCol.SETTLEMENT_ID] ?? '').trim(),
    transferDate: formatGrabTransferDate(row[grabCol.TRANSFER_DATE]),
    amount: row[grabCol.AMOUNT],
    offer: row[grabCol.OFFER],
    netSales: row[grabCol.NET_SALES],
    netMdr: row[grabCol.NET_MDR],
    taxOnMdr: row[grabCol.TAX_ON_MDR],
    ttlComm,
    total: row[grabCol.TOTAL],
  };
}

function buildGpaySheet(dataRows, listingLastRow) {
  const sheet = {};
  const headerRow = 1;
  const dataStartRow = 2;

  GPAY_HEADERS.forEach((header, columnIndex) => {
    setCell(sheet, XLSX.utils.encode_cell({ r: headerRow - 1, c: columnIndex }), header);
  });

  dataRows.forEach((row, rowIndex) => {
    const excelRow = dataStartRow + rowIndex;

    setCell(sheet, `A${excelRow}`, row.category);
    setCell(sheet, `B${excelRow}`, row.orderId);
    setCell(sheet, `C${excelRow}`, row.paymentMethod);
    setCell(sheet, `D${excelRow}`, row.settlementId);
    setCell(sheet, `E${excelRow}`, row.transferDate);
    setAmountCell(sheet, `F${excelRow}`, row.amount);
    setAmountCell(sheet, `G${excelRow}`, row.offer);
    setAmountCell(sheet, `H${excelRow}`, row.netSales);
    setAmountCell(sheet, `I${excelRow}`, row.netMdr);
    setAmountCell(sheet, `J${excelRow}`, row.taxOnMdr);
    setAmountCell(sheet, `K${excelRow}`, row.ttlComm, ttlCommFormula(excelRow));
    setAmountCell(sheet, `L${excelRow}`, row.total);
    setDateCell(sheet, `M${excelRow}`, '', invDateFormula(excelRow, listingLastRow));
    setAmountCell(sheet, `N${excelRow}`, '', invAmountFormula(excelRow, listingLastRow));
    setAmountCell(sheet, `O${excelRow}`, '', diffFormula(excelRow));
  });

  const lastRow = Math.max(dataStartRow + dataRows.length - 1, headerRow);
  sheet['!ref'] = `A1:O${lastRow}`;
  return sheet;
}

async function createGpayWorkbook(
  grabBuffer,
  ncInput,
  { ncSheetName, ncFileName, ncFileNames, grabFileName } = {},
) {
  const ncFiles = normalizeNcInputs(ncInput, { ncSheetName, ncFileName, ncFileNames });

  if (ncFiles.length === 0) {
    throw new Error('At least one ESK Payment Details file is required');
  }

  const grabWorkbook = XLSX.read(grabBuffer, { type: 'buffer' });
  const grabSheetName = findGrabTransactionSheet(grabWorkbook);
  const grabSheet = grabWorkbook.Sheets[grabSheetName];

  if (!grabSheet) {
    throw new Error('Grab Transaction sheet not found');
  }

  const grabRows = XLSX.utils.sheet_to_json(grabSheet, { header: 1, defval: '' });
  if (grabRows.length < 2) {
    throw new Error('Grab Transaction file has no data rows');
  }

  const grabCol = resolveGrabColumns(grabRows[0]);
  const ncSources = sortNcSourcesByPeriod(ncFiles.map((file) => loadEskPaymentSource(file)));
  const eskOrderIds = mergeEskOrderIdSets(ncSources);
  const ordersSheet = buildMergedOrdersSheet(ncSources);
  const ordersLastRow = ordersSheet ? sheetLastRow(ordersSheet) : 0;
  const listingSheet = normalizeEmbeddedListingSheet(
    buildMergedListingSheet(ncSources),
    ordersLastRow,
  );
  const listingLastRow = sheetLastRow(listingSheet);

  const outputRows = [];
  let generated = 0;
  let linked = 0;

  for (let i = 1; i < grabRows.length; i += 1) {
    const row = grabRows[i];
    if (!row?.length || !shouldIncludeGrabRow(row, grabCol)) continue;

    const orderId = normalizeGrabCellText(row[grabCol.PARTNER_TXN_ID_2]);
    if (eskOrderIds.has(orderId)) linked += 1;

    outputRows.push(buildGpayRow(row, grabCol));
    generated += 1;
  }

  if (generated === 0) {
    throw new Error('No transferred Grab rows found');
  }

  const worksheet = buildGpaySheet(outputRows, listingLastRow);
  const sheetNames = [GPAY_SHEET_NAME, ESK_LISTING_SHEET_NAME];
  const sheets = {
    [GPAY_SHEET_NAME]: worksheet,
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
  buffer = await applyGpayWorkbookStyles(buffer);

  return {
    buffer,
    filename: GPAY_FILENAME,
    sheetName: GPAY_SHEET_NAME,
    generated,
    linked,
    grabSheetName,
    ncSheetNames: ncSources.map((source) => source.sheetName),
    ncFileNames: ncSources.map((source) => source.fileName).filter(Boolean),
    grabFileName,
  };
}

module.exports = {
  createGpayWorkbook,
  GPAY_HEADERS,
};
