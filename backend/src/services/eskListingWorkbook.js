const XLSX = require('xlsx');
const {
  detectNcListingPeriod,
  extractNcOrderKey,
  isExcelDateSerial,
  isNcFooterRow,
  isNcSummaryRow,
  resolveNcLayout,
} = require('../utils/ncLayout');

const ESK_LISTING_SHEET_NAME = 'Payment Listing-Detail _sort';
const ESK_ORDERS_SHEET_NAME = 'Orders';
const LISTING_LOOKUP_MAX_COL = 'N';
const LISTING_DATE_FORMAT = 'm/d/yy';

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

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
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

function sheetLastRow(sheet) {
  if (!sheet?.['!ref']) return 1;
  return XLSX.utils.decode_range(sheet['!ref']).e.r + 1;
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

function loadEskPaymentSource({ buffer, fileName, sheetName }) {
  const ncWorkbook = XLSX.read(buffer, { type: 'buffer', cellFormula: true });
  const ncSheetNameResolved = findNcSheet(ncWorkbook, sheetName);
  const ncSheet = ncWorkbook.Sheets[ncSheetNameResolved];
  const ordersSheetName = ncWorkbook.SheetNames.find((name) => /^orders$/i.test(name));
  const ordersSheet = ordersSheetName ? ncWorkbook.Sheets[ordersSheetName] : null;

  if (!ncSheet) {
    const label = fileName || ncSheetNameResolved || 'ESK file';
    throw new Error(`ESK Payment Details sheet not found in ${label}`);
  }

  const ncRows = XLSX.utils.sheet_to_json(ncSheet, { header: 1, defval: '' });
  const layout = resolveNcLayout(ncRows[0]);
  const period = detectNcListingPeriod(ncWorkbook, ncRows, layout, fileName);

  return {
    listingSheet: ncSheet,
    ordersSheet,
    ordersSheetName: ordersSheetName || null,
    sheetName: ncSheetNameResolved,
    fileName,
    period,
    ncRows,
    layout,
  };
}

function buildEskOrderIdSet(ncRows, layout) {
  const { col: ncCol, headerRows } = layout;
  const orderIds = new Set();

  for (let i = headerRows; i < ncRows.length; i += 1) {
    const row = ncRows[i];
    if (!row?.length || isNcFooterRow(row) || isNcSummaryRow(row, layout)) continue;

    const orderKey = extractNcOrderKey(row, layout);
    if (/^\d{6}-\d{5,6}$/.test(orderKey)) {
      orderIds.add(orderKey);
    }
  }

  return orderIds;
}

function mergeEskOrderIdSets(ncSources) {
  const merged = new Set();
  for (const source of ncSources) {
    for (const orderId of buildEskOrderIdSet(source.ncRows, source.layout)) {
      merged.add(orderId);
    }
  }
  return merged;
}

function listingLookupRange(lastRow) {
  return `'${ESK_LISTING_SHEET_NAME}'!$A$1:$${LISTING_LOOKUP_MAX_COL}$${Math.max(lastRow, 2)}`;
}

function listingLookupRangeByOrd(lastRow, maxCol = LISTING_LOOKUP_MAX_COL) {
  return `'${ESK_LISTING_SHEET_NAME}'!$B$1:$${maxCol}$${Math.max(lastRow, 2)}`;
}

function buildEskOrdSet(ncSources) {
  const ordSet = new Set();

  for (const source of ncSources) {
    const { ncRows, layout } = source;
    const { col: ncCol, headerRows } = layout;

    for (let i = headerRows; i < ncRows.length; i += 1) {
      const row = ncRows[i];
      if (!row?.length || isNcFooterRow(row) || isNcSummaryRow(row, layout)) continue;

      const ord = String(row[ncCol.ORD] ?? '').trim();
      if (/^\d{6}$/.test(ord)) {
        ordSet.add(ord);
      }
    }
  }

  return ordSet;
}

module.exports = {
  ESK_LISTING_SHEET_NAME,
  ESK_ORDERS_SHEET_NAME,
  LISTING_LOOKUP_MAX_COL,
  buildEskOrderIdSet,
  buildEskOrdSet,
  buildMergedListingSheet,
  buildMergedOrdersSheet,
  cloneSheet,
  listingLookupRange,
  listingLookupRangeByOrd,
  loadEskPaymentSource,
  mergeEskOrderIdSets,
  normalizeEmbeddedListingSheet,
  normalizeNcInputs,
  sheetLastRow,
  sortNcSourcesByPeriod,
};
