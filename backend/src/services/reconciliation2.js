const db = require('../db');
const {
  amountsEqual,
  deduplicateRecords,
  STATUS,
} = require('./reconciliation');
const {
  orderColAMatchKey,
  stripeColDValue,
  orderMapsToStripe,
} = require('../utils/matchKey');

function belongsByPrefix(order, stripePayment) {
  return orderMapsToStripe(order.order_id, stripePayment.description);
}

function orderColF(order) {
  return Number(order.amount);
}

function stripeColF(stripePayment) {
  return Number(stripePayment.amount);
}

function rowMeta(order, stripePayment) {
  return {
    note: order?.note || '',
    comment: order?.comment || '',
    matchKey: order
      ? orderColAMatchKey(order.order_id)
      : stripeColDValue(stripePayment?.description),
  };
}

function buildMatchedRecord(order, stripePayment) {
  return {
    orderId: order.order_id,
    orderAmount: orderColF(order),
    paymentAmount: stripeColF(stripePayment),
    amount: orderColF(order),
    stripeDescription: stripeColDValue(stripePayment.description),
    status: STATUS.MATCHED,
    orderRecordId: order.id,
    stripePaymentId: stripePayment.id,
    lineRef: order.line_ref || null,
    type: 'order',
    ...rowMeta(order, stripePayment),
  };
}

function buildExceptionRecord(order, stripePayment) {
  return {
    orderId: order.order_id,
    orderAmount: orderColF(order),
    paymentAmount: stripeColF(stripePayment),
    amount: orderColF(order),
    stripeDescription: stripeColDValue(stripePayment.description),
    status: STATUS.EXCEPTION,
    orderRecordId: order.id,
    stripePaymentId: stripePayment.id,
    lineRef: order.line_ref || null,
    type: 'order',
    reason: 'amount_mismatch',
    ...rowMeta(order, stripePayment),
  };
}

function buildUnmatchedOrderRecord(order) {
  return {
    orderId: order.order_id,
    amount: orderColF(order),
    status: STATUS.UNMATCHED,
    orderRecordId: order.id,
    lineRef: order.line_ref || null,
    type: 'order',
    reason: 'no_stripe_payment',
    ...rowMeta(order),
  };
}

function buildUnmatchedStripeRecord(stripePayment) {
  const key = stripeColDValue(stripePayment.description);
  return {
    orderId: key,
    amount: stripeColF(stripePayment),
    paymentAmount: stripeColF(stripePayment),
    stripeDescription: key,
    status: STATUS.UNMATCHED,
    stripePaymentId: stripePayment.id,
    lineRef: stripePayment.description,
    type: 'stripe',
    reason: 'no_order',
    ...rowMeta(null, stripePayment),
  };
}

function runReconciliation2() {
  const orders = db.prepare('SELECT * FROM orders ORDER BY id').all();
  const stripePayments = db
    .prepare('SELECT * FROM stripe_payments ORDER BY id')
    .all();

  const usedStripeIds = new Set();
  const matched = [];
  const unmatched = [];
  const exception = [];

  for (const order of orders) {
    const candidates = stripePayments.filter(
      (p) => !usedStripeIds.has(p.id) && belongsByPrefix(order, p)
    );

    const exactPayment = candidates.find((p) =>
      amountsEqual(stripeColF(p), orderColF(order))
    );

    if (exactPayment) {
      usedStripeIds.add(exactPayment.id);
      matched.push(buildMatchedRecord(order, exactPayment));
      continue;
    }

    if (candidates.length > 0) {
      const mismatchedPayment = candidates[0];
      usedStripeIds.add(mismatchedPayment.id);
      exception.push(buildExceptionRecord(order, mismatchedPayment));
      continue;
    }

    unmatched.push(buildUnmatchedOrderRecord(order));
  }

  for (const stripePayment of stripePayments) {
    if (!usedStripeIds.has(stripePayment.id)) {
      unmatched.push(buildUnmatchedStripeRecord(stripePayment));
    }
  }

  const uniqueUnmatched = deduplicateRecords(unmatched);
  const uniqueException = deduplicateRecords(exception);

  const rows = [...uniqueException, ...uniqueUnmatched, ...matched].sort(
    (a, b) => String(a.orderId).localeCompare(String(b.orderId))
  );

  return {
    mode: 'reconciliation2',
    summary: {
      totalOrders: orders.length,
      totalPayments: stripePayments.length,
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
  runReconciliation2,
  orderColAMatchKey,
  belongsByPrefix,
};
