const NC_LEGACY_COL = {
  HOLDER: 0,
  DOC_NO: 1,
  RETAIL_INVOICE: 2,
  PAYMENT_TYPE: 3,
  AMT: 4,
  NOTE: 5,
  NOTE_HELPER: 6,
  ORDER_HELPER: 7,
};

const NC_LEGACY_LAYOUT = {
  format: 'legacy',
  headerRows: 3,
  col: NC_LEGACY_COL,
};

const NC_SUMMARY_LABELS = new Set(['Grand Total', 'Deposit Refund', 'Net Total']);

const MONTH_INDEX = {
  Jan: 1,
  Feb: 2,
  Mar: 3,
  Apr: 4,
  May: 5,
  Jun: 6,
  Jul: 7,
  Aug: 8,
  Sep: 9,
  Oct: 10,
  Nov: 11,
  Dec: 12,
};

function isExcelDateSerial(value) {
  return typeof value === 'number' && Number.isFinite(value) && value >= 30000 && value <= 60000;
}

function resolveNcLayout(headerRow) {
  const headers = (headerRow || []).map((cell) => String(cell ?? '').trim());
  const indexOf = (name, fallback) => {
    const idx = headers.findIndex((header) => header.toLowerCase() === name.toLowerCase());
    return idx >= 0 ? idx : fallback;
  };

  const hasEskPaymentHeaders =
    headers.includes('Ord No.')
    && headers.includes('RETAIL_INVOICE')
    && headers.includes('PAYMENT')
    && headers.some((header) => header.toLowerCase() === 'docdate');

  if (hasEskPaymentHeaders) {
    return {
      format: 'eskPayment',
      headerRows: 1,
      col: {
        ORDER_ID: indexOf('Ord No.', 0),
        ORD: indexOf('Ord', 1),
        HOLDER: indexOf('docdate', 2),
        DOC_NO: indexOf('BILLNO', 3),
        RETAIL_INVOICE: indexOf('RETAIL_INVOICE', 4),
        PAYMENT_TYPE: indexOf('PAYMENT', 5),
        AMT: indexOf('ORIGINALAMT', 6),
        NOTE: indexOf('NOTE', 7),
      DEPOSIT_REFUND: indexOf('DEPOSIT_REFUND', 8),
      NET_TOTAL: indexOf('NET_TOTAL', 9),
      COSTCODE: indexOf('COSTCODE', 10),
      FINAL_PRICE_PAID: indexOf('Final rounded price paid', 11),
    },
  };
  }

  const hasDetailHeaders =
    headers.includes('RETAIL_INVOICE')
    && headers.includes('BILLNO')
    && headers.some((header) => header.toLowerCase() === 'docdate');

  if (!hasDetailHeaders) {
    return NC_LEGACY_LAYOUT;
  }

  return {
    format: 'detail',
    headerRows: 1,
    col: {
      SN: indexOf('SN', 0),
      HOLDER: indexOf('docdate', 1),
      DOC_NO: indexOf('BILLNO', 2),
      RETAIL_INVOICE: indexOf('RETAIL_INVOICE', 3),
      PAYMENT_TYPE: indexOf('PAYMENT', 4),
      AMT: indexOf('ORIGINALAMT', 5),
      NOTE: indexOf('NOTE', 6),
      DEPOSIT_REFUND: indexOf('DEPOSIT_REFUND', 7),
      NET_TOTAL: indexOf('NET_TOTAL', 8),
      COSTCODE: indexOf('COSTCODE', 9),
    },
  };
}

function extractOrderKeyFromText(text) {
  const match = String(text ?? '').trim().match(/(\d{6}-\d{5,6})/);
  return match ? match[1] : '';
}

function extractNcOrderKey(row, layout) {
  const { col } = layout;
  if (col.ORDER_ID != null && col.ORDER_ID >= 0) {
    const orderId = String(row[col.ORDER_ID] ?? '').trim();
    if (/^\d{6}-\d{5,6}$/.test(orderId)) return orderId;
  }

  let body = String(row[col.NOTE] ?? '').trim();
  body = body.replace(/^Note:\s*/i, '').trim();
  if (body.startsWith('#')) body = body.slice(1).trim();

  const orderKey = extractOrderKeyFromText(body);
  if (orderKey) return orderKey;

  return extractOrderKeyFromText(row[col.ORDER_ID]);
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
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
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

function parsePeriodFromText(text) {
  const match = String(text ?? '').match(
    /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*(\d{2})?\b/i,
  );
  if (!match) return null;

  const month = match[1].slice(0, 1).toUpperCase() + match[1].slice(1, 3).toLowerCase();
  const yearSuffix = match[2] || '26';

  return {
    month,
    yearSuffix,
    sortKey: Number(yearSuffix) * 100 + (MONTH_INDEX[month] || 0),
  };
}

function periodFromSheetLabel(sheetLabel) {
  const match = String(sheetLabel ?? '').match(
    /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{2})$/i,
  );
  if (!match) return null;

  const month = match[1].slice(0, 1).toUpperCase() + match[1].slice(1, 3).toLowerCase();
  const yearSuffix = match[2];

  return {
    month,
    yearSuffix,
    sortKey: Number(yearSuffix) * 100 + (MONTH_INDEX[month] || 0),
  };
}

function isNcValidDocNo(docNo) {
  const value = String(docNo ?? '').trim();
  if (!value || value.toUpperCase() === 'ESK') return false;
  if (/^\d+$/.test(value)) return false;
  return true;
}

function isNcFooterRow(ncRow) {
  const cell0 = String(ncRow?.[0] ?? '').trim();
  return /lines in the current region/i.test(cell0);
}

function isNcDateSeparatorRow(ncRow, layout) {
  if (layout.format === 'detail') return false;

  const docNo = String(ncRow[layout.col.DOC_NO] ?? '').trim();
  if (docNo) return false;
  return parseListingDate(ncRow[layout.col.HOLDER]) != null;
}

function isNcSummaryRow(ncRow, layout) {
  if (isNcFooterRow(ncRow)) return true;

  const label = String(ncRow[layout.col.PAYMENT_TYPE] ?? '').trim();
  return NC_SUMMARY_LABELS.has(label);
}

function detectNcListingDate(ncWorkbook, ncRows, layout, fileName = null) {
  for (const sheetName of ncWorkbook.SheetNames) {
    const fromSheetName = parsePeriodFromText(sheetName);
    if (fromSheetName) {
      const listingDate = parseListingDateFromPeriod(fromSheetName);
      if (listingDate) return listingDate;
    }
  }

  if (fileName) {
    const fromFileName = parsePeriodFromText(fileName);
    if (fromFileName) {
      const listingDate = parseListingDateFromPeriod(fromFileName);
      if (listingDate) return listingDate;
    }
  }

  if (layout.format === 'legacy') {
    return parseListingDate(ncRows[2]?.[layout.col.HOLDER]);
  }

  let earliestSerial = null;
  for (let i = layout.headerRows; i < ncRows.length; i += 1) {
    const ncRow = ncRows[i];
    if (!ncRow?.length || isNcFooterRow(ncRow) || isNcSummaryRow(ncRow, layout)) continue;

    const docNo = String(ncRow[layout.col.DOC_NO] ?? '').trim();
    if (!isNcValidDocNo(docNo)) continue;

    const serial = ncRow[layout.col.HOLDER];
    if (!isExcelDateSerial(serial)) continue;
    if (earliestSerial == null || serial < earliestSerial) {
      earliestSerial = serial;
    }
  }

  return earliestSerial != null ? parseListingDate(earliestSerial) : null;
}

function parseListingDateFromPeriod(period) {
  const monthIndex = MONTH_INDEX[period.month];
  if (!monthIndex) return null;

  const year = Number(`20${period.yearSuffix}`);
  const date = new Date(Date.UTC(year, monthIndex - 1, 1));
  const epoch = Date.UTC(1899, 11, 30);
  const serial = Math.round((date.getTime() - epoch) / 86400000);
  const monthNames = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
  ];

  return {
    serial,
    sheetLabel: `${monthNames[monthIndex - 1]}${period.yearSuffix}`,
  };
}

function detectNcListingPeriod(ncWorkbook, ncRows, layout, fileName = null) {
  for (const sheetName of ncWorkbook.SheetNames) {
    const fromSheetName = parsePeriodFromText(sheetName);
    if (fromSheetName) return fromSheetName;
  }

  if (fileName) {
    const fromFileName = parsePeriodFromText(fileName);
    if (fromFileName) return fromFileName;
  }

  if (layout.format === 'legacy') {
    const listingDate = parseListingDate(ncRows[2]?.[layout.col.HOLDER]);
    if (listingDate) return periodFromSheetLabel(listingDate.sheetLabel);
  }

  for (let i = layout.headerRows; i < Math.min(ncRows.length, layout.headerRows + 20); i += 1) {
    const row = ncRows[i];
    if (!row?.length || isNcFooterRow(row)) continue;

    for (let col = 0; col < Math.min(row.length, 10); col += 1) {
      const cell = row[col];
      const fromText = parsePeriodFromText(String(cell ?? ''));
      if (fromText) return fromText;

      const parsedListingDate = parseListingDate(cell);
      if (parsedListingDate) {
        return periodFromSheetLabel(parsedListingDate.sheetLabel);
      }
    }
  }

  const listingDate = detectNcListingDate(ncWorkbook, ncRows, layout, fileName);
  return listingDate ? periodFromSheetLabel(listingDate.sheetLabel) : null;
}

function ncHolderValue(holder, layout) {
  if (layout.format === 'detail' || layout.format === 'eskPayment') {
    return isExcelDateSerial(holder) ? holder : ' ';
  }
  if (holder === '' || holder == null) return ' ';
  return holder;
}

module.exports = {
  NC_LEGACY_LAYOUT,
  NC_SUMMARY_LABELS,
  detectNcListingDate,
  detectNcListingPeriod,
  extractNcOrderKey,
  isExcelDateSerial,
  isNcDateSeparatorRow,
  isNcFooterRow,
  isNcSummaryRow,
  isNcValidDocNo,
  ncHolderValue,
  parseListingDate,
  parsePeriodFromText,
  periodFromSheetLabel,
  resolveNcLayout,
};
