const {
  insertOrders,
  insertPayments,
  clearOrders,
  clearPayments,
} = require('../db/repository');
const { fetchShopifyOrders } = require('./shopify');
const { fetchStripePayments } = require('./stripe');

async function syncOrdersFromShopify({
  replace = false,
  since = null,
  dedupe = false,
} = {}) {
  if (replace) clearOrders();

  const { records, fetched, skipped } = await fetchShopifyOrders({ since });
  const synced = records.length ? insertOrders(records, { dedupe }) : 0;

  return {
    source: 'shopify',
    synced,
    fetched,
    skipped,
    since,
    replaced: replace,
  };
}

async function syncPaymentsFromStripe({
  replace = false,
  since = null,
  dedupe = false,
} = {}) {
  if (replace) clearPayments();

  const { records, fetched, skipped } = await fetchStripePayments({ since });
  const synced = records.length ? insertPayments(records, { dedupe }) : 0;

  return {
    source: 'stripe',
    synced,
    fetched,
    skipped,
    since,
    replaced: replace,
  };
}

module.exports = {
  syncOrdersFromShopify,
  syncPaymentsFromStripe,
};
