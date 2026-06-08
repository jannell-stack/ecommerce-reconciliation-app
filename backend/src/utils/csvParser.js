const iconv = require('iconv-lite');

const ORDER_ID_PATTERN = /^[A-Za-z0-9#._\-/ ]{1,64}$/;

function scoreCsvText(text) {
  if (!text || text.includes('\u0000')) return -1;
  const sample = text.slice(0, 2000).toLowerCase();
  let score = 0;
  if (sample.includes('order')) score += 2;
  if (sample.includes('amount')) score += 2;
  if (/[0-9]{3,}-[0-9]{6}/.test(sample)) score += 3;
  if (/ord-/i.test(sample)) score += 1;
  const replacementChars = (text.match(/\uFFFD/g) || []).length;
  score -= replacementChars;
  return score;
}

function decodeCsvBuffer(buffer) {
  if (!buffer?.length) {
    throw new Error('File is empty');
  }

  if (buffer[0] === 0x50 && buffer[1] === 0x4b) {
    throw new Error(
      'This looks like an Excel file (.xlsx). In Excel use File → Save As → CSV UTF-8, then upload.'
    );
  }

  const candidates = [];

  let utf8 = buffer.toString('utf-8');
  if (utf8.charCodeAt(0) === 0xfeff) utf8 = utf8.slice(1);
  candidates.push({ encoding: 'utf-8', text: utf8.trim() });

  if (iconv.encodingExists('gbk')) {
    candidates.push({
      encoding: 'gbk',
      text: iconv.decode(buffer, 'gbk').trim(),
    });
  }

  const best = candidates.sort(
    (a, b) => scoreCsvText(b.text) - scoreCsvText(a.text)
  )[0];

  if (scoreCsvText(best.text) < 0) {
    throw new Error(
      'Could not read this file. Export CSV with columns order_id and amount (UTF-8 or GBK), not Excel .xlsx.'
    );
  }

  return best.text;
}

function parseAmount(value) {
  if (value == null || value === '') return NaN;
  const cleaned = String(value)
    .replace(/[$,\s"￥¥]/g, '')
    .replace(/^-$/, '0')
    .trim();
  if (!cleaned || cleaned === '-') return NaN;
  return parseFloat(cleaned);
}

function isValidOrderId(orderId) {
  if (!orderId || typeof orderId !== 'string') return false;
  const trimmed = orderId.trim();
  if (!trimmed || trimmed.length > 64) return false;
  if (!ORDER_ID_PATTERN.test(trimmed)) return false;
  return !/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/.test(trimmed);
}

function parseCsvRows(buffer) {
  const text = decodeCsvBuffer(buffer);
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (lines.length === 0) return [];

  const header = lines[0].split(',').map((h) => h.trim().toLowerCase());
  const orderIdIdx = header.findIndex(
    (h) =>
      h === 'order_id' ||
      h === 'orderid' ||
      h === 'order id' ||
      h.includes('订单') ||
      h.includes('单号')
  );
  const amountIdx = header.findIndex(
    (h) => h === 'amount' || h.includes('金额') || h.includes('amount')
  );

  const hasHeader = orderIdIdx >= 0 && amountIdx >= 0;
  const start = hasHeader ? 1 : 0;
  const idCol = orderIdIdx >= 0 ? orderIdIdx : 0;
  const amtCol = amountIdx >= 0 ? amountIdx : 1;

  const rows = [];

  for (const line of lines.slice(start)) {
    const cols = line.split(',').map((c) => c.trim().replace(/^"|"$/g, ''));
    const order_id = cols[idCol]?.trim();
    const amount = parseAmount(cols[amtCol]);

    if (!isValidOrderId(order_id) || Number.isNaN(amount)) continue;
    rows.push({ order_id, amount });
  }

  return rows;
}

module.exports = {
  parseCsvRows,
  parseAmount,
  isValidOrderId,
  decodeCsvBuffer,
  ORDER_ID_PATTERN,
};
