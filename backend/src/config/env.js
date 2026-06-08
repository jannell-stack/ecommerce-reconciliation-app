function requireEnv(name) {
  const value = process.env[name];
  if (!value?.trim()) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value.trim();
}

function getShopifyConfig() {
  return {
    storeDomain: requireEnv('SHOPIFY_STORE_DOMAIN'),
    accessToken: requireEnv('SHOPIFY_ACCESS_TOKEN'),
    apiVersion: process.env.SHOPIFY_API_VERSION?.trim() || '2024-10',
  };
}

function getStripeConfig() {
  return {
    secretKey: requireEnv('STRIPE_SECRET_KEY'),
  };
}

module.exports = {
  requireEnv,
  getShopifyConfig,
  getStripeConfig,
};
