const XLSX = require('xlsx');
const db = require('../db');
const { runReconciliation } = require('./reconciliation');
const { runReconciliation2 } = require('./reconciliation2');
const { amountsEqual } = require('./reconciliation');
const { orderColAMatchKey, stripeColDValue } = require('../utils/matchKey');

const MALAYSIA_TIME_ZONE = 'Asia/Kuala_Lumpur';

function formatGeneratedAtMalaysia() {
  return new Date().toLocaleString('en-MY', {
    timeZone: MALAYSIA_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  });
}

function buildMetadataIndex() {
  const index = new Map();
  const rows = db
    .prepare(
      `SELECT line_ref, order_id, note, comment FROM orders WHERE line_ref IS NOT NULL
       UNION ALL
       SELECT line_ref, order_id, note, comment FROM payments WHERE line_ref IS NOT NULL`
    )
    .all();

  for (const row of rows) {
    if (!row.line_ref) continue;
    const prev = index.get(row.line_ref) || { note: '', comment: '' };
    index.set(row.line_ref, {
      order_id: row.order_id,
      note: row.note || prev.note || '',
      comment: row.comment || prev.comment || '',
    });
  }

  return index;
}

function enrichRow(row, metaIndex) {
  const meta = row.lineRef ? metaIndex.get(row.lineRef) : null;
  return {
    ...row,
    note: row.note || meta?.note || '',
    comment: row.comment || meta?.comment || '',
  };
}

function baseExportFields(row) {
  return {
    'Order ID': row.orderId,
    'Col F Amount': row.orderAmount ?? row.amount ?? '',
    'Col H Amount': row.paymentAmount ?? '',
    Note: row.note ?? '',
    Comment: row.comment ?? '',
    'Doc No': row.lineRef ?? '',
    Status: row.status,
    Type: row.type ?? '',
    Reason: row.reason ?? '',
  };
}

function rowToExport(row) {
  return baseExportFields(row);
}

function exceptionRowToExport(row) {
  const orderAmount = Number(row.orderAmount ?? row.amount);
  const paymentAmount = Number(row.paymentAmount);
  const difference =
    Number.isFinite(orderAmount) && Number.isFinite(paymentAmount)
      ? Math.round((paymentAmount - orderAmount) * 100) / 100
      : '';

  return {
    ...baseExportFields(row),
    'Difference (C-B)': difference,
  };
}

const STANDARD_HEADERS = [
  'Order ID',
  'Col F Amount',
  'Col H Amount',
  'Note',
  'Comment',
  'Doc No',
  'Status',
  'Type',
  'Reason',
];

const EXCEPTION_HEADERS = [...STANDARD_HEADERS, 'Difference (C-B)'];

function emptyPlaceholderRow(headers) {
  return Object.fromEntries(
    headers.map((header) => [header, header === 'Status' ? 'No records' : ''])
  );
}

function rowsToSheet(rows, mapRow, headers) {
  const data = rows.length ? rows.map(mapRow) : [emptyPlaceholderRow(headers)];
  return XLSX.utils.json_to_sheet(data, { header: headers });
}

function buildSummarySheet(summary, metaCount, { title, matchRule } = {}) {
  const matchedPct =
    summary.totalOrders > 0
      ? Math.round((summary.matchedCount / summary.totalOrders) * 100)
      : 0;

  const rows = [
    [title || 'Reconciliation Summary', ''],
    ['Generated at', formatGeneratedAtMalaysia()],
    ['', ''],
    ['Col F Total Rows', summary.totalOrders],
    ['Col H Total Rows', summary.totalPayments],
    ['Matched', summary.matchedCount],
    ['Matched %', `${matchedPct}%`],
    ['Unmatched', summary.unmatchedCount],
    ['Exceptions', summary.exceptionCount],
    ['Rows with Note/Comment metadata', metaCount],
  ];

  if (matchRule) {
    rows.splice(3, 0, ['Match rule', matchRule]);
  }

  return rows;
}

function buildWorkbook(reconciliation, summaryOptions = {}) {
  const metaIndex = buildMetadataIndex();
  const enrich = (row) => enrichRow(row, metaIndex);

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(
      buildSummarySheet(reconciliation.summary, metaIndex.size, summaryOptions)
    ),
    'Summary'
  );

  const sheets = [
    { name: 'Matched', rows: reconciliation.matched.map(enrich) },
    { name: 'Unmatched', rows: reconciliation.unmatched.map(enrich) },
    { name: 'Exceptions', rows: reconciliation.exception.map(enrich) },
  ];

  for (const { name, rows } of sheets) {
    const mapRow = name === 'Exceptions' ? exceptionRowToExport : rowToExport;
    const headers = name === 'Exceptions' ? EXCEPTION_HEADERS : STANDARD_HEADERS;
    XLSX.utils.book_append_sheet(
      workbook,
      rowsToSheet(rows, mapRow, headers),
      name
    );
  }

  const allRows = [
    ...reconciliation.matched,
    ...reconciliation.unmatched,
    ...reconciliation.exception,
  ].map((row) => rowToExport(enrich(row)));

  XLSX.utils.book_append_sheet(
    workbook,
    rowsToSheet(allRows, (r) => r, STANDARD_HEADERS),
    'All Results'
  );

  return workbook;
}

function writeWorkbookBuffer(workbook, filenamePrefix) {
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const date = new Date().toISOString().slice(0, 10);
  return {
    buffer,
    filename: `${filenamePrefix}-${date}.xlsx`,
  };
}

function exportReconciliationBuffer() {
  const reconciliation = runReconciliation();
  const workbook = buildWorkbook(reconciliation);
  const { buffer, filename } = writeWorkbookBuffer(workbook, 'reconciliation');
  return {
    buffer,
    filename,
    reconciliation,
    metadataRows: buildMetadataIndex().size,
  };
}

const RECON2_ESK_COL_A = 'ESK Column A';
const RECON2_ESK_COL_F = 'ESK Column F';

const RECON2_HEADERS = [
  RECON2_ESK_COL_A,
  'Stripe Col D',
  RECON2_ESK_COL_F,
  'Stripe Col F',
];

function buildStripePaymentLookup() {
  const lookup = new Map();
  const rows = db
    .prepare('SELECT id, description, amount FROM stripe_payments')
    .all();
  for (const row of rows) {
    lookup.set(row.id, row);
  }
  return lookup;
}

function buildStripeByMatchKey() {
  const byKey = new Map();
  const rows = db
    .prepare('SELECT id, description, amount FROM stripe_payments')
    .all();
  for (const row of rows) {
    const key = stripeColDValue(row.description);
    if (!key) continue;
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(row);
  }
  return byKey;
}

function resolveStripeRowForExport(row, stripeLookup, stripeByKey) {
  if (row.stripePaymentId) {
    const linked = stripeLookup.get(row.stripePaymentId);
    if (linked) return linked;
  }

  if (row.type === 'stripe') {
    if (row.stripePaymentId) return stripeLookup.get(row.stripePaymentId) ?? null;
    const key = stripeColDValue(row.stripeDescription);
    return (stripeByKey.get(key) || [])[0] ?? null;
  }

  const orderKey = orderColAMatchKey(row.orderId);
  const candidates = stripeByKey.get(orderKey) || [];
  if (!candidates.length) return null;

  const orderAmount = row.orderAmount ?? row.amount;
  return (
    candidates.find((p) => amountsEqual(p.amount, orderAmount)) || candidates[0]
  );
}

function exportStripeColD(stripeRow) {
  if (!stripeRow?.description && stripeRow?.description !== 0) return '';
  const n = Number(stripeRow.description);
  if (Number.isFinite(n) && /^\d+$/.test(String(stripeRow.description).trim())) {
    return n;
  }
  return String(stripeRow.description).trim();
}

function recon2ExportRow(row, stripeLookup, stripeByKey) {
  const isStripeOnly = row.type === 'stripe';
  const stripeRow = resolveStripeRowForExport(row, stripeLookup, stripeByKey);

  return {
    [RECON2_ESK_COL_A]: isStripeOnly
      ? ''
      : orderColAMatchKey(row.orderId),
    'Stripe Col D': stripeRow ? exportStripeColD(stripeRow) : '',
    [RECON2_ESK_COL_F]: isStripeOnly ? '' : (row.orderAmount ?? row.amount ?? ''),
    'Stripe Col F': stripeRow
      ? stripeRow.amount
      : isStripeOnly
        ? row.amount ?? ''
        : '',
  };
}

function buildRecon2SummarySheet(summary) {
  const matchedPct =
    summary.totalOrders > 0
      ? Math.round((summary.matchedCount / summary.totalOrders) * 100)
      : 0;

  return [
    ['Reconciliation 2 Summary', ''],
    ['Generated at', formatGeneratedAtMalaysia()],
    ['', ''],
    [
      'Match rule',
      'ESK Column A (first 6 chars) → Stripe col D; amounts ESK Column F vs Stripe col F',
    ],
    ['Total ESK', summary.totalOrders],
    ['Total Stripe rows', summary.totalPayments],
    ['Matched', summary.matchedCount],
    ['Matched %', `${matchedPct}%`],
    ['Unmatched', summary.unmatchedCount],
    ['Exceptions', summary.exceptionCount],
    ['', ''],
    ['Export columns', 'ESK Column A, Stripe Col D, ESK Column F, Stripe Col F'],
    [
      'Unmatched rule',
      'Unmatched: no Stripe col D, or no ESK Column A (orphan Stripe rows)',
    ],
    [
      'Exceptions rule',
      'Exceptions: ESK Column A and Stripe col D present but ESK Column F ≠ Stripe col F',
    ],
  ];
}

function exportRowTallies(exportRow) {
  if (isExportOrderColAEmpty(exportRow) || isExportStripeColDEmpty(exportRow)) {
    return false;
  }

  const orderKey = String(exportRow[RECON2_ESK_COL_A]).trim();
  const stripeD = stripeColDValue(exportRow['Stripe Col D']);
  if (!orderKey || !stripeD || orderKey !== stripeD) return false;

  const orderF = exportRow[RECON2_ESK_COL_F];
  const stripeF = exportRow['Stripe Col F'];
  if (orderF === '' || orderF == null || stripeF === '' || stripeF == null) {
    return false;
  }

  return amountsEqual(orderF, stripeF);
}

function isExportStripeColDEmpty(exportRow) {
  const value = exportRow['Stripe Col D'];
  return value === '' || value == null;
}

function isExportOrderColAEmpty(exportRow) {
  const value = exportRow[RECON2_ESK_COL_A];
  return value === '' || value == null;
}

function partitionRecon2ForExport(reconciliation, stripeLookup, stripeByKey) {
  const toExport = (row) => recon2ExportRow(row, stripeLookup, stripeByKey);
  const matched = [...reconciliation.matched];
  const unmatched = [];
  const exception = [];

  for (const row of [...reconciliation.unmatched, ...reconciliation.exception]) {
    const exportRow = toExport(row);
    if (exportRowTallies(exportRow)) {
      matched.push(row);
    } else if (isExportStripeColDEmpty(exportRow) || isExportOrderColAEmpty(exportRow)) {
      unmatched.push(row);
    } else {
      exception.push(row);
    }
  }

  return { matched, unmatched, exception };
}

function buildWorkbook2(reconciliation) {
  const stripeLookup = buildStripePaymentLookup();
  const stripeByKey = buildStripeByMatchKey();
  const toExport = (row) => recon2ExportRow(row, stripeLookup, stripeByKey);
  const { matched, unmatched, exception } = partitionRecon2ForExport(
    reconciliation,
    stripeLookup,
    stripeByKey
  );

  const exportSummary = {
    ...reconciliation.summary,
    matchedCount: matched.length,
    unmatchedCount: unmatched.length,
    exceptionCount: exception.length,
  };

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(buildRecon2SummarySheet(exportSummary)),
    'Summary'
  );

  const sheets = [
    { name: 'Matched', rows: matched },
    { name: 'Unmatched', rows: unmatched },
    { name: 'Exceptions', rows: exception },
  ];

  for (const { name, rows } of sheets) {
    XLSX.utils.book_append_sheet(
      workbook,
      rowsToSheet(rows, toExport, RECON2_HEADERS),
      name
    );
  }

  const allRows = [...matched, ...unmatched, ...exception];

  XLSX.utils.book_append_sheet(
    workbook,
    rowsToSheet(allRows, toExport, RECON2_HEADERS),
    'All Results'
  );

  return workbook;
}

function exportReconciliation2Buffer() {
  const reconciliation = runReconciliation2();
  const workbook = buildWorkbook2(reconciliation);
  const { buffer, filename } = writeWorkbookBuffer(workbook, 'reconciliation-2');
  return {
    buffer,
    filename,
    reconciliation,
  };
}

module.exports = {
  exportReconciliationBuffer,
  exportReconciliation2Buffer,
  buildWorkbook,
  buildWorkbook2,
};
