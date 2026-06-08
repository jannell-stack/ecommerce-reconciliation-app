const Stripe = require('stripe');
const { getStripeConfig } = require('../config/env');

function extractOrderId(paymentIntent) {
  const metadata = paymentIntent.metadata || {};
  const keys = [
    'order_id',
    'orderId',
    'shopify_order_id',
    'shopify_order_name',
  ];

  for (const key of keys) {
    const value = metadata[key];
    if (value?.trim()) {
      return String(value).replace(/^#/, '').trim();
    }
  }

  if (paymentIntent.description) {
    const match = paymentIntent.description.match(/#?(\d{4,})/);
    if (match) return match[1];
  }

  return null;
}

function mapStripePaymentIntent(pi) {
  const orderId = extractOrderId(pi);
  if (!orderId) return null;

  const amountCents = pi.amount_received ?? pi.amount ?? 0;
  const amount = amountCents / 100;
  if (Number.isNaN(amount) || amount <= 0) return null;

  return {
    order_id: orderId,
    amount,
    external_id: pi.id,
  };
}

async function fetchStripePayments({ since } = {}) {
  const { secretKey } = getStripeConfig();
  const stripe = new Stripe(secretKey);

  const paymentIntents = [];
  let startingAfter;
  const listParams = { limit: 100 };

  if (since) {
    listParams.created = {
      gte: Math.floor(new Date(since).getTime() / 1000),
    };
  }

  do {
    const page = await stripe.paymentIntents.list({
      ...listParams,
      starting_after: startingAfter,
    });

    for (const pi of page.data) {
      if (pi.status === 'succeeded') {
        paymentIntents.push(pi);
      }
    }

    startingAfter =
      page.has_more && page.data.length
        ? page.data[page.data.length - 1].id
        : undefined;
  } while (startingAfter);

  const mapped = [];
  let skipped = 0;

  for (const pi of paymentIntents) {
    const row = mapStripePaymentIntent(pi);
    if (row) mapped.push(row);
    else skipped += 1;
  }

  return {
    records: mapped,
    fetched: paymentIntents.length,
    skipped,
  };
}

async function fetchAllStripePayments() {
  return fetchStripePayments();
}

module.exports = {
  fetchStripePayments,
  fetchAllStripePayments,
  mapStripePaymentIntent,
  extractOrderId,
};
