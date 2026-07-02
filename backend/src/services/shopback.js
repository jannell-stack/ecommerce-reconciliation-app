const XLSX = require('xlsx');
const { isExcelDateSerial } = require('./eskPaymentStyles');
const {
  ESK_LISTING_SHEET_NAME,
  ESK_ORDERS_SHEET_NAME,
  buildEskOrdSet,
  buildMergedListingSheet,
  buildMergedOrdersSheet,
  listingLookupRangeByOrd,
  loadEskPaymentSource,
  normalizeEmbeddedListingSheet,
  normalizeNcInputs,
  sheetLastRow,
  sortNcSourcesByPeriod,
} = require('./eskListingWorkbook');
const { applyShopbackWorkbookStyles } = require('./eskPaymentStyles');

const SHOPBACK_HEADERS = [
  'Source.Name',
  'Brand',
  'Transaction Date',
  'Reference',
  'Billable Amount',
  'Total Fees',
  'Sales Tax',
  'Total Comm',
  'Net Earnings /(Payable)',
  'inv date',
  'inv amt',
  'diff',
  "Bank r'ced dd",
  '',
];

const SHOPBACK_SHEET_NAME = 'Shopback';
const SHOPBACK_FILENAME = 'Shopback.xlsx';
const SHOPBACK_AMOUNT_FORMAT = '#,##0.00';
const SHOPBACK_DATE_FORMAT = 'm/d/yy';

const ORI_HEADER_TRANSACTION_ID = 'Transaction ID';
const ORI_HEADER_EXPECTED_SETTLEMENT = 'Expected settlement date';

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function normalizeCellText(value) {
  return String(value ?? '').trim();
}

function findShopbackOriSheet(workbook) {
  return (
    workbook.SheetNames.find((name) => /shopback/i.test(name) && /ori/i.test(name))
    || workbook.SheetNames.find((name) => /shopback/i.test(name))
    || workbook.SheetNames[0]
  );
}

function buildOriHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((header, index) => {
    const key = normalizeCellText(header);
    if (key) map[key] = index;
  });
  return map;
}

function getOriValue(row, headerMap, headerName) {
  const index = headerMap[headerName];
  if (index == null || index < 0) return '';
  return row[index];
}

const MONTH_INDEX = {
  jan: 0,
  feb: 1,
  mar: 2,
  apr: 3,
  may: 4,
  jun: 5,
  jul: 6,
  aug: 7,
  sep: 8,
  oct: 9,
  nov: 10,
  dec: 11,
};

function parseShopbackDateText(value) {
  const text = normalizeCellText(value);
  const match = text.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
  if (!match) return null;

  const month = MONTH_INDEX[match[2].slice(0, 3).toLowerCase()];
  if (month == null) return null;

  const epoch = Date.UTC(1899, 11, 30);
  return Math.floor(
    (Date.UTC(Number(match[3]), month, Number(match[1])) - epoch) / 86400000,
  );
}

function formatShopbackSettlementDate(value) {
  const serial = parseShopbackDateText(value);
  if (serial == null) return normalizeCellText(value);

  const epoch = Date.UTC(1899, 11, 30);
  const date = new Date(epoch + serial * 86400000);
  const day = String(date.getUTCDate()).padStart(2, '0');
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const year = date.getUTCFullYear();
  return `${day}/${month}/${year}`;
}

function parseShopbackTransactionDate(value) {
  if (isExcelDateSerial(value)) return Math.floor(Number(value));

  const serial = parseShopbackDateText(value);
  return serial == null ? normalizeCellText(value) : serial;
}

function isShopbackReference(value) {
  return /^\d{6}$/.test(normalizeCellText(value));
}

function isShopbackPurchaseRow(row, headerMap) {
  const transactionType = normalizeCellText(getOriValue(row, headerMap, 'Transaction Type'));
  return /purchase/i.test(transactionType);
}

function parseShopbackOriRows(rows) {
  const settlementBySource = new Map();
  let headerMap = null;
  const outputRows = [];

  for (const row of rows) {
    const label = normalizeCellText(row[1]);
    const sourceName = normalizeCellText(row[0]);

    if (label === ORI_HEADER_EXPECTED_SETTLEMENT) {
      if (sourceName) {
        settlementBySource.set(sourceName, formatShopbackSettlementDate(row[2]));
      }
      continue;
    }

    if (label === ORI_HEADER_TRANSACTION_ID) {
      headerMap = buildOriHeaderMap(row);
      continue;
    }

    if (!headerMap) continue;

    const reference = normalizeCellText(getOriValue(row, headerMap, 'Reference'));
    if (!isShopbackReference(reference) || !isShopbackPurchaseRow(row, headerMap)) continue;

    const rowSourceName = sourceName || normalizeCellText(getOriValue(row, headerMap, 'Source.Name'));
    outputRows.push({
      sourceName: rowSourceName,
      brand: normalizeCellText(getOriValue(row, headerMap, 'Brand')),
      transactionDate: parseShopbackTransactionDate(getOriValue(row, headerMap, 'Transaction Date')),
      reference,
      billableAmount: getOriValue(row, headerMap, 'Billable Amount'),
      totalFees: getOriValue(row, headerMap, 'Total Fees'),
      salesTax: getOriValue(row, headerMap, 'Sales Tax'),
      netEarnings: getOriValue(row, headerMap, 'Net Earnings /(Payable)'),
      bankReceivedDate: settlementBySource.get(rowSourceName) || '',
    });
  }

  return outputRows;
}

function totalCommFormula(excelRow) {
  return `F${excelRow}+G${excelRow}`;
}

function invDateFormula(excelRow, listingLastRow) {
  return `VLOOKUP(D${excelRow},${listingLookupRangeByOrd(listingLastRow, 'N')},2,0)`;
}

function invAmountFormula(excelRow, listingLastRow) {
  return `VLOOKUP(D${excelRow},${listingLookupRangeByOrd(listingLastRow, 'O')},11,0)`;
}

function diffFormula(excelRow) {
  return `E${excelRow}-K${excelRow}`;
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
    z: SHOPBACK_AMOUNT_FORMAT,
  };
  if (formula) cell.f = formula;
  sheet[address] = cell;
}

function setDateCell(sheet, address, value, formula) {
  if (!formula && value === '') return;
  const serial = typeof value === 'number' ? value : parseShopbackTransactionDate(value);
  const cell = {
    t: typeof serial === 'number' ? 'n' : 's',
    v: typeof serial === 'number' ? serial : serial,
    z: SHOPBACK_DATE_FORMAT,
  };
  if (formula) cell.f = formula;
  sheet[address] = cell;
}

function buildShopbackSheet(dataRows, listingLastRow) {
  const sheet = {};
  const headerRow = 1;
  const dataStartRow = 2;

  SHOPBACK_HEADERS.forEach((header, columnIndex) => {
    setCell(sheet, XLSX.utils.encode_cell({ r: headerRow - 1, c: columnIndex }), header);
  });

  dataRows.forEach((row, rowIndex) => {
    const excelRow = dataStartRow + rowIndex;

    setCell(sheet, `A${excelRow}`, row.sourceName);
    setCell(sheet, `B${excelRow}`, row.brand);
    setDateCell(sheet, `C${excelRow}`, row.transactionDate);
    setCell(sheet, `D${excelRow}`, row.reference);
    setAmountCell(sheet, `E${excelRow}`, row.billableAmount);
    setAmountCell(sheet, `F${excelRow}`, row.totalFees);
    setAmountCell(sheet, `G${excelRow}`, row.salesTax);
    setAmountCell(sheet, `H${excelRow}`, '', totalCommFormula(excelRow));
    setAmountCell(sheet, `I${excelRow}`, row.netEarnings);
    setDateCell(sheet, `J${excelRow}`, '', invDateFormula(excelRow, listingLastRow));
    setAmountCell(sheet, `K${excelRow}`, '', invAmountFormula(excelRow, listingLastRow));
    setAmountCell(sheet, `L${excelRow}`, '', diffFormula(excelRow));
    setCell(sheet, `M${excelRow}`, row.bankReceivedDate);
  });

  const lastRow = Math.max(dataStartRow + dataRows.length - 1, headerRow);
  sheet['!ref'] = `A1:N${lastRow}`;
  return sheet;
}

async function createShopbackWorkbook(
  shopbackBuffer,
  ncInput,
  { ncSheetName, ncFileName, ncFileNames, shopbackFileName } = {},
) {
  const ncFiles = normalizeNcInputs(ncInput, { ncSheetName, ncFileName, ncFileNames });

  if (ncFiles.length === 0) {
    throw new Error('At least one ESK Payment Details file is required');
  }

  const shopbackWorkbook = XLSX.read(shopbackBuffer, { type: 'buffer' });
  const shopbackSheetName = findShopbackOriSheet(shopbackWorkbook);
  const shopbackSheet = shopbackWorkbook.Sheets[shopbackSheetName];

  if (!shopbackSheet) {
    throw new Error('Shopback Transaction sheet not found');
  }

  const oriRows = XLSX.utils.sheet_to_json(shopbackSheet, { header: 1, defval: '' });
  if (oriRows.length < 2) {
    throw new Error('Shopback Transaction file has no data rows');
  }

  const outputRows = parseShopbackOriRows(oriRows);
  if (outputRows.length === 0) {
    throw new Error('No Shopback purchase rows found in the transaction file');
  }

  const ncSources = sortNcSourcesByPeriod(ncFiles.map((file) => loadEskPaymentSource(file)));
  const eskOrdSet = buildEskOrdSet(ncSources);
  const ordersSheet = buildMergedOrdersSheet(ncSources);
  const ordersLastRow = ordersSheet ? sheetLastRow(ordersSheet) : 0;
  const listingSheet = normalizeEmbeddedListingSheet(
    buildMergedListingSheet(ncSources),
    ordersLastRow,
  );
  const listingLastRow = sheetLastRow(listingSheet);

  let linked = 0;
  for (const row of outputRows) {
    if (eskOrdSet.has(row.reference)) linked += 1;
  }

  const worksheet = buildShopbackSheet(outputRows, listingLastRow);
  const sheetNames = [SHOPBACK_SHEET_NAME, ESK_LISTING_SHEET_NAME];
  const sheets = {
    [SHOPBACK_SHEET_NAME]: worksheet,
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
  buffer = await applyShopbackWorkbookStyles(buffer);

  return {
    buffer,
    filename: SHOPBACK_FILENAME,
    sheetName: SHOPBACK_SHEET_NAME,
    generated: outputRows.length,
    linked,
    shopbackSheetName,
    ncSheetNames: ncSources.map((source) => source.sheetName),
    ncFileNames: ncSources.map((source) => source.fileName).filter(Boolean),
    shopbackFileName,
  };
}

module.exports = {
  createShopbackWorkbook,
  SHOPBACK_HEADERS,
  parseShopbackOriRows,
};
