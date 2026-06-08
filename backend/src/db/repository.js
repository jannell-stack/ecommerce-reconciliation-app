const db = require('../db');
const { stripeColDValue } = require('../utils/matchKey');

const orderExistsByAmount = db.prepare(
  'SELECT 1 FROM orders WHERE order_id = ? AND line_ref IS NULL AND ABS(amount - ?) < 0.01 LIMIT 1'
);

const paymentExistsByAmount = db.prepare(
  'SELECT 1 FROM payments WHERE order_id = ? AND line_ref IS NULL AND ABS(amount - ?) < 0.01 LIMIT 1'
);

const orderExistsByLineRef = db.prepare(
  'SELECT 1 FROM orders WHERE line_ref = ? LIMIT 1'
);

const paymentExistsByLineRef = db.prepare(
  'SELECT 1 FROM payments WHERE line_ref = ? LIMIT 1'
);

const updateOrderMeta = db.prepare(
  `UPDATE orders
   SET note = COALESCE(@note, note),
       comment = COALESCE(@comment, comment),
       amount = @amount
   WHERE line_ref = @line_ref`
);

const updatePaymentMeta = db.prepare(
  `UPDATE payments
   SET note = COALESCE(@note, note),
       comment = COALESCE(@comment, comment),
       amount = @amount
   WHERE line_ref = @line_ref`
);

function shouldSkipOrder(order_id, amount, line_ref, dedupe) {
  if (!dedupe) return false;
  if (line_ref) return Boolean(orderExistsByLineRef.get(line_ref));
  return Boolean(orderExistsByAmount.get(order_id, amount));
}

function shouldSkipPayment(order_id, amount, line_ref, dedupe) {
  if (!dedupe) return false;
  if (line_ref) return Boolean(paymentExistsByLineRef.get(line_ref));
  return Boolean(paymentExistsByAmount.get(order_id, amount));
}

function insertOrders(records, { dedupe = false } = {}) {
  const insert = db.prepare(
    `INSERT INTO orders (order_id, amount, line_ref, note, comment)
     VALUES (@order_id, @amount, @line_ref, @note, @comment)`
  );

  let inserted = 0;
  const insertMany = db.transaction((rows) => {
    for (const row of rows) {
      const order_id = String(row.order_id ?? row.orderId);
      const amount = Number(row.amount);
      const line_ref = row.line_ref ?? row.lineRef ?? null;
      if (shouldSkipOrder(order_id, amount, line_ref, dedupe)) {
        if (line_ref) {
          updateOrderMeta.run({
            line_ref,
            amount,
            note: row.note ?? null,
            comment: row.comment ?? null,
          });
        }
        continue;
      }
      insert.run({
        order_id,
        amount,
        line_ref,
        note: row.note ?? null,
        comment: row.comment ?? null,
      });
      inserted += 1;
    }
  });
  insertMany(records);
  return inserted;
}

function insertPayments(records, { dedupe = false } = {}) {
  const insert = db.prepare(
    `INSERT INTO payments (order_id, amount, line_ref, note, comment)
     VALUES (@order_id, @amount, @line_ref, @note, @comment)`
  );

  let inserted = 0;
  const insertMany = db.transaction((rows) => {
    for (const row of rows) {
      const order_id = String(row.order_id ?? row.orderId);
      const amount = Number(row.amount);
      const line_ref = row.line_ref ?? row.lineRef ?? null;
      if (shouldSkipPayment(order_id, amount, line_ref, dedupe)) {
        if (line_ref) {
          updatePaymentMeta.run({
            line_ref,
            amount,
            note: row.note ?? null,
            comment: row.comment ?? null,
          });
        }
        continue;
      }
      insert.run({
        order_id,
        amount,
        line_ref,
        note: row.note ?? null,
        comment: row.comment ?? null,
      });
      inserted += 1;
    }
  });
  insertMany(records);
  return inserted;
}

function clearOrders() {
  return db.prepare('DELETE FROM orders').run().changes;
}

function clearPayments() {
  return db.prepare('DELETE FROM payments').run().changes;
}

const stripeExistsByRow = db.prepare(
  'SELECT 1 FROM stripe_payments WHERE description = ? AND ABS(amount - ?) < 0.01 AND source_row = ? LIMIT 1'
);

function insertStripePayments(records, { dedupe = false } = {}) {
  const insert = db.prepare(
    `INSERT INTO stripe_payments
     (description, match_key, amount, net, fee, reporting_category, currency, source_row)
     VALUES (@description, @match_key, @amount, @net, @fee, @reporting_category, @currency, @source_row)`
  );

  let inserted = 0;
  const insertMany = db.transaction((rows) => {
    for (const row of rows) {
      const description = String(row.description ?? '').trim();
      const amount = Number(row.amount);
      const source_row = row.source_row ?? null;
      if (
        dedupe &&
        stripeExistsByRow.get(description, amount, source_row)
      ) {
        continue;
      }
      insert.run({
        description,
        match_key: stripeColDValue(row.match_key ?? description),
        amount,
        net: Number.isFinite(Number(row.net)) ? Number(row.net) : null,
        fee: Number.isFinite(Number(row.fee)) ? Number(row.fee) : null,
        reporting_category: row.reporting_category ?? null,
        currency: row.currency ?? null,
        source_row,
      });
      inserted += 1;
    }
  });
  insertMany(records);
  return inserted;
}

function clearStripePayments() {
  return db.prepare('DELETE FROM stripe_payments').run().changes;
}

module.exports = {
  insertOrders,
  insertPayments,
  clearOrders,
  clearPayments,
  insertStripePayments,
  clearStripePayments,
};
