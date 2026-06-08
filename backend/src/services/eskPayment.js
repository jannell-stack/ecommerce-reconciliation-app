const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const TEMPLATE_PATH = path.join(__dirname, '..', '..', 'templates', 'ESK-Payment-template.xlsx');

const NC_COL = {
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
const SK_PRICE_PAID = 18;
const SK_PAYMENT_METHOD = 19;
const ORG_SUFFIX = ', Org Code:SKJSK,Def1:ESK,Store Code:0ESK';

function cellHasContent(cell) {
  if (!cell) return false;
  if (cell.f) return true;
  return cell.v != null && cell.v !== '';
}

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

function parseNumber(value) {
  if (value === '' || value == null) return NaN;
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function extractOrderId(note, colG, colH) {
  const helper = String(colH ?? '').trim();
  const helperMatch = helper.match(/^(\d{6}-\d{5,6})/);
  if (helperMatch) return helperMatch[1];

  const colGText = String(colG ?? '').trim();
  const colGMatch = colGText.match(/^(\d{6}-\d{5,6})/);
  if (colGMatch) return colGMatch[1];

  const noteMatch = String(note ?? '').match(/#?(\d{6}-\d{5,6})/);
  return noteMatch ? noteMatch[1].replace(/,+$/, '') : '';
}

function formatNote(ncRow, orderId) {
  const raw = String(ncRow[NC_COL.NOTE] ?? '');
  let body = raw.replace(/^Note:\s*/i, '').replace(/^BA\s+#?/i, '').trim();
  if (body.startsWith('#')) body = body.slice(1).trim();

  if (/Org Code:/i.test(body)) {
    const match = body.match(/(\d{6}-\d{5,6})(.*)/);
    if (match) return `${match[1]}${match[2]}`;
    return body;
  }

  const match = body.match(/^(\d{6}-\d{5,6})\s*(.*)$/);
  if (match) {
    const extra = match[2].trim();
    if (extra) return `${match[1]} ${extra}${ORG_SUFFIX}`;
    return `${match[1]}${ORG_SUFFIX}`;
  }

  if (orderId) return `${orderId}${ORG_SUFFIX}`;
  return body;
}

function leftOrderId(noteText) {
  return String(noteText ?? '').slice(0, 13).trim();
}

function extractReferToInvoice(text) {
  const match = String(text ?? '').match(/refer to\s+(LS0ESK\d+)/i);
  return match ? match[1] : '';
}

function mapPaymentType(ncType, skMethod) {
  const normalized = String(ncType ?? '').trim().toUpperCase();
  if (normalized === 'ECOM-GPAY') return 'Grab';
  if (normalized === 'SHOPBACK') return 'ShopBack';
  if (normalized === 'STRIPE') return 'Stripe';

  const sk = String(skMethod ?? '').trim().toLowerCase();
  if (sk.includes('grab')) return 'Grab';
  if (sk.includes('shopback')) return 'ShopBack';
  if (sk.includes('stripe')) return 'Stripe';
  if (!sk) return '';
  return sk.charAt(0).toUpperCase() + sk.slice(1);
}

function parseListingDate(value) {
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

function findSheet(workbook, pattern) {
  return workbook.SheetNames.find((name) => pattern.test(name)) || workbook.SheetNames[0];
}

function readSheetRows(buffer, sheetPattern) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
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

function buildSkLookup(rows) {
  const lookup = new Map();

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const orderId = String(row[SK_ORDER_ID] ?? '').trim();
    if (!orderId) continue;

    const entry = {
      pricePaid: parseNumber(row[SK_PRICE_PAID]),
      paymentMethod: String(row[SK_PAYMENT_METHOD] ?? '').trim(),
    };

    if (!lookup.has(orderId)) {
      lookup.set(orderId, []);
    }
    lookup.get(orderId).push(entry);
  }

  return lookup;
}

function lookupSk(orderId, skLookup, amount) {
  const rows = skLookup.get(String(orderId ?? '').trim());
  if (!rows?.length) return null;
  if (rows.length === 1) return rows[0];

  const absAmount = Math.abs(parseNumber(amount));
  if (!Number.isNaN(absAmount)) {
    const amountMatch = rows.find((row) => row.pricePaid === absAmount);
    if (amountMatch) return amountMatch;
  }

  return rows[0];
}

function loadTemplateWorkbook() {
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error('ESK payment template is missing. Expected backend/templates/ESK-Payment-template.xlsx');
  }

  return XLSX.read(fs.readFileSync(TEMPLATE_PATH), {
    type: 'buffer',
    cellFormula: true,
    cellStyles: false,
  });
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

function updateFormulaRow(sheet, excelRow, ncRow, pattern, skLookup, ordersRange, valuesOnly) {
  const orderId = extractOrderId(
    ncRow[NC_COL.NOTE],
    ncRow[NC_COL.NOTE_HELPER],
    ncRow[NC_COL.ORDER_HELPER],
  );
  const note = pattern.templateNote || formatNote(ncRow, orderId);
  const orderKey = pattern.templateOrderId
    ? String(pattern.templateOrderId).trim()
    : leftOrderId(note);
  const docNo = String(ncRow[NC_COL.DOC_NO] ?? '').trim();

  if (valuesOnly) {
    setCell(sheet, `A${excelRow}`, orderKey);
  } else {
    setCell(sheet, `A${excelRow}`, { f: `LEFT(G${excelRow},13)`, v: orderKey, t: 's' });
  }
  setCell(sheet, `B${excelRow}`, ' ');
  setCell(sheet, `C${excelRow}`, docNo);
  setCell(
    sheet,
    `D${excelRow}`,
    pattern.retailInvoice ?? ncRow[NC_COL.RETAIL_INVOICE],
  );
  setCell(sheet, `E${excelRow}`, ncRow[NC_COL.PAYMENT_TYPE]);
  setCell(sheet, `F${excelRow}`, ncRow[NC_COL.AMT]);
  setCell(sheet, `G${excelRow}`, note);

  clearFormulaCells(sheet, excelRow, ['H', 'I', 'J', 'K']);

  let hValue = null;
  if (pattern.hasH) {
    if (pattern.hFormula) {
      const skMatch = lookupSk(orderKey, skLookup, ncRow[NC_COL.AMT]);
      hValue = skMatch?.pricePaid ?? null;
      if (hValue != null && Number.isNaN(hValue)) hValue = null;
      if (valuesOnly) {
        setCell(sheet, `H${excelRow}`, hValue);
      } else {
        setCell(sheet, `H${excelRow}`, {
          f: `VLOOKUP(A${excelRow},${ordersRange},19,0)`,
          v: hValue,
          t: 'n',
        });
      }
    } else if (pattern.hValue != null) {
      hValue = pattern.hValue;
      setCell(sheet, `H${excelRow}`, { v: hValue, t: 'n' });
    }

    if (pattern.hasI && hValue != null) {
      const amount = parseNumber(ncRow[NC_COL.AMT]);
      const diff = Number.isNaN(amount) ? null : amount - hValue;
      if (valuesOnly) {
        setCell(sheet, `I${excelRow}`, diff);
      } else {
        setCell(sheet, `I${excelRow}`, {
          f: `F${excelRow}-H${excelRow}`,
          v: diff,
          t: 'n',
        });
      }
    }

    if (pattern.hasJ) {
      const skMatch = lookupSk(orderKey, skLookup, ncRow[NC_COL.AMT]);
      const paymentType =
        pattern.jValue ??
        mapPaymentType(ncRow[NC_COL.PAYMENT_TYPE], skMatch?.paymentMethod);
      if (valuesOnly) {
        setCell(sheet, `J${excelRow}`, paymentType);
      } else {
        setCell(sheet, `J${excelRow}`, {
          f: `VLOOKUP(A${excelRow},${ordersRange},20,0)`,
          v: paymentType,
          t: 's',
        });
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

function buildMainSheet(
  templateMain,
  templatePatterns,
  ncMap,
  skRows,
  skLookup,
  listingDate,
  valuesOnly,
) {
  const sheet = templateMain;
  const templateRows = XLSX.utils.sheet_to_json(templateMain, { header: 1, defval: null });
  const ordersLastRow = Math.max(skRows.length, 2);
  const ordersRange = `Orders!$A$1:$AG$${ordersLastRow}`;
  let generated = 0;

  if (listingDate) {
    setCell(sheet, 'B3', listingDate.serial);
  }

  for (let i = 3; i < templateRows.length; i += 1) {
    const docNo = String(templateRows[i][2] ?? '').trim();
    if (!docNo) continue;

    const ncRow = ncMap.get(docNo);
    if (!ncRow) continue;

    const pattern = templatePatterns.get(docNo);
    if (!pattern) continue;

    updateFormulaRow(sheet, i + 1, ncRow, pattern, skLookup, ordersRange, valuesOnly);
    generated += 1;
  }

  return { sheet, generated, rowCount: templateRows.length };
}

function createEskPaymentWorkbook(skBuffer, ncBuffer) {
  const template = loadTemplateWorkbook();
  const templateMainName = getTemplateMainSheetName(template);
  const templateMain = template.Sheets[templateMainName];
  const { patterns: templatePatterns } = buildTemplateMetadata(templateMain);

  const { sheetName: skSheetName, rows: skRows } = readSheetRows(skBuffer, /orders/i);
  const { sheetName: ncSheetName, rows: ncRows } = readSheetRows(ncBuffer, /payment listing/i);

  if (ncRows.length <= NC_HEADER_ROWS) {
    throw new Error('NC payment listing file has no data rows');
  }

  const listingDate = parseListingDate(ncRows[2]?.[0]);
  const sheetTitle = listingDate
    ? `Payment Listing-${listingDate.sheetLabel}`
    : templateMainName;

  const valuesOnly = useValuesOnlyOutput();
  const skLookup = buildSkLookup(skRows);
  const ncMap = buildNcMap(ncRows);
  const { sheet: mainSheet, generated } = buildMainSheet(
    templateMain,
    templatePatterns,
    ncMap,
    skRows,
    skLookup,
    listingDate,
    valuesOnly,
  );

  const workbook = {
    SheetNames: [sheetTitle],
    Sheets: {
      [sheetTitle]: mainSheet,
    },
  };

  if (!valuesOnly) {
    workbook.SheetNames.push('Orders');
    workbook.Sheets.Orders = XLSX.utils.aoa_to_sheet(skRows);
  }

  const buffer = XLSX.write(workbook, {
    type: 'buffer',
    bookType: 'xlsx',
    compression: true,
  });
  const filename = `ESK-Payment${listingDate ? `-${listingDate.sheetLabel}` : ''}.xlsx`;

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
