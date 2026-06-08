const db = require('../db');

const STATUS = {
  MATCHED: 'matched',
  UNMATCHED: 'unmatched',
  EXCEPTION: 'exception',
};

function amountsEqual(a, b) {
  return Math.abs(Number(a) - Number(b)) < 0.01;
}

function belongsToSameLine(order, payment) {
  if (order.line_ref && payment.line_ref) {
    return order.line_ref === payment.line_ref;
  }
  return order.order_id === payment.order_id;
}

function rowMeta(order, payment) {
  return {
    note: order?.note || payment?.note || '',
    comment: order?.comment || payment?.comment || '',
  };
}

function buildMatchedRecord(order, payment) {
  return {
    orderId: order.order_id,
    orderAmount: order.amount,
    paymentAmount: payment.amount,
    amount: order.amount,
    status: STATUS.MATCHED,
    orderRecordId: order.id,
    paymentRecordId: payment.id,
    lineRef: order.line_ref || payment.line_ref || null,
    type: 'order',
    ...rowMeta(order, payment),
  };
}

function buildExceptionRecord(order, payment) {
  return {
    orderId: order.order_id,
    orderAmount: order.amount,
    paymentAmount: payment.amount,
    amount: order.amount,
    status: STATUS.EXCEPTION,
    orderRecordId: order.id,
    paymentRecordId: payment.id,
    lineRef: order.line_ref || payment.line_ref || null,
    type: 'order',
    reason: 'amount_mismatch',
    ...rowMeta(order, payment),
  };
}

function buildUnmatchedOrderRecord(order) {
  return {
    orderId: order.order_id,
    amount: order.amount,
    status: STATUS.UNMATCHED,
    orderRecordId: order.id,
    lineRef: order.line_ref || null,
    type: 'order',
    reason: 'no_payment',
    ...rowMeta(order),
  };
}

function buildUnmatchedPaymentRecord(payment) {
  return {
    orderId: payment.order_id,
    amount: payment.amount,
    status: STATUS.UNMATCHED,
    paymentRecordId: payment.id,
    lineRef: payment.line_ref || null,
    type: 'payment',
    reason: 'no_order',
    ...rowMeta(null, payment),
  };
}

function deduplicateRecords(records) {
  const seen = new Set();
  const unique = [];

  for (const record of records) {
    const key = record.lineRef
      ? String(record.lineRef)
      : String(record.orderId);
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(record);
  }

  return unique;
}

function runReconciliation() {
  const orders = db.prepare('SELECT * FROM orders ORDER BY id').all();
  const payments = db.prepare('SELECT * FROM payments ORDER BY id').all();

  const usedPaymentIds = new Set();
  const matched = [];
  const unmatched = [];
  const exception = [];

  for (const order of orders) {
    const candidates = payments.filter(
      (p) => !usedPaymentIds.has(p.id) && belongsToSameLine(order, p)
    );

    const exactPayment = candidates.find((p) =>
      amountsEqual(p.amount, order.amount)
    );

    if (exactPayment) {
      usedPaymentIds.add(exactPayment.id);
      matched.push(buildMatchedRecord(order, exactPayment));
      continue;
    }

    if (candidates.length > 0) {
      const mismatchedPayment = candidates[0];
      usedPaymentIds.add(mismatchedPayment.id);
      exception.push(buildExceptionRecord(order, mismatchedPayment));
      continue;
    }

    unmatched.push(buildUnmatchedOrderRecord(order));
  }

  for (const payment of payments) {
    if (!usedPaymentIds.has(payment.id)) {
      unmatched.push(buildUnmatchedPaymentRecord(payment));
    }
  }

  const uniqueUnmatched = deduplicateRecords(unmatched);
  const uniqueException = deduplicateRecords(exception);

  const rows = [...uniqueException, ...uniqueUnmatched, ...matched].sort(
    (a, b) => String(a.orderId).localeCompare(String(b.orderId))
  );

  return {
    summary: {
      totalOrders: orders.length,
      totalPayments: payments.length,
      matchedCount: matched.length,
      unmatchedCount: uniqueUnmatched.length,
      exceptionCount: uniqueException.length,
    },
    matched,
    unmatched: uniqueUnmatched,
    exception: uniqueException,
    rows,
  };
}

module.exports = {
  runReconciliation,
  amountsEqual,
  deduplicateRecords,
  belongsToSameLine,
  STATUS,
};
