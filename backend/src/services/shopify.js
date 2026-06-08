const { getShopifyConfig } = require('../config/env');

function normalizeStoreDomain(domain) {
  const trimmed = domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
  return trimmed.includes('.') ? trimmed : `${trimmed}.myshopify.com`;
}

function mapShopifyOrder(order) {
  const amount = parseFloat(order.current_total_price ?? order.total_price ?? 0);
  if (Number.isNaN(amount)) return null;

  const orderId =
    order.name?.replace(/^#/, '') ||
    String(order.id);

  return {
    order_id: String(orderId),
    amount,
    external_id: String(order.id),
  };
}

async function shopifyFetch(path, config) {
  const domain = normalizeStoreDomain(config.storeDomain);
  const url = `https://${domain}/admin/api/${config.apiVersion}${path}`;

  const res = await fetch(url, {
    headers: {
      'X-Shopify-Access-Token': config.accessToken,
      'Content-Type': 'application/json',
    },
  });

  if (!res.ok) {
    const body = await res.text();
    let detail = body;
    try {
      detail = JSON.parse(body).errors ?? body;
    } catch {
      /* use raw body */
    }
    throw new Error(`Shopify API error (${res.status}): ${detail}`);
  }

  return res.json();
}

async function fetchShopifyOrders({ since } = {}) {
  const config = getShopifyConfig();
  const orders = [];
  const sinceQuery = since
    ? `&created_at_min=${encodeURIComponent(since)}`
    : '';
  let path = `/orders.json?status=any&limit=250${sinceQuery}`;

  while (path) {
    const data = await shopifyFetch(path, config);
    const batch = data.orders || [];
    orders.push(...batch);
    path = null;

    const lastId = batch.length ? batch[batch.length - 1].id : null;
    if (batch.length === 250 && lastId) {
      path = `/orders.json?status=any&limit=250&since_id=${lastId}${sinceQuery}`;
    }
  }

  const mapped = [];
  let skipped = 0;

  for (const order of orders) {
    const row = mapShopifyOrder(order);
    if (row) mapped.push(row);
    else skipped += 1;
  }

  return { records: mapped, fetched: orders.length, skipped };
}

async function fetchAllShopifyOrders() {
  return fetchShopifyOrders();
}

module.exports = {
  fetchShopifyOrders,
  fetchAllShopifyOrders,
  mapShopifyOrder,
};
