const { runReconciliation } = require('./reconciliation');
const {
  syncOrdersFromShopify,
  syncPaymentsFromStripe,
} = require('./syncRunner');
const {
  getLastSync,
  setLastSync,
  getDefaultSince,
} = require('../db/syncMeta');

function log(section, data) {
  console.log(`[cron] ${section}:`, JSON.stringify(data, null, 2));
}

async function runDailySync() {
  const startedAt = new Date().toISOString();
  console.log(`[cron] ===== Daily sync started at ${startedAt} =====`);

  const ordersSince = getLastSync('orders') ?? getDefaultSince();
  const paymentsSince = getLastSync('payments') ?? getDefaultSince();

  log('Fetching orders since', { since: ordersSince });
  const orders = await syncOrdersFromShopify({
    since: ordersSince,
    dedupe: true,
  });
  log('Orders sync result', orders);

  log('Fetching payments since', { since: paymentsSince });
  const payments = await syncPaymentsFromStripe({
    since: paymentsSince,
    dedupe: true,
  });
  log('Payments sync result', payments);

  const reconciliation = runReconciliation();
  log('Reconciliation summary', reconciliation.summary);
  log('Matched', { count: reconciliation.matched.length });
  log('Unmatched', { count: reconciliation.unmatched.length });
  log('Exceptions', { count: reconciliation.exception.length });

  const completedAt = new Date().toISOString();
  setLastSync('orders', completedAt);
  setLastSync('payments', completedAt);
  setLastSync('job', completedAt);

  const result = {
    startedAt,
    completedAt,
    orders,
    payments,
    reconciliation: {
      summary: reconciliation.summary,
      matched: reconciliation.matched.length,
      unmatched: reconciliation.unmatched.length,
      exception: reconciliation.exception.length,
    },
  };

  console.log(`[cron] ===== Daily sync completed at ${completedAt} =====`);
  return result;
}

module.exports = { runDailySync };
