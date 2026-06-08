/**
 * Reconciliation 2: first 6 characters of Order column A (Ord No)
 * map to the exact value in Stripe column D (description).
 * Example: 864859-290426 → 864859 matches Stripe col D = 864859
 */
function orderColAMatchKey(orderId) {
  return String(orderId ?? '').trim().slice(0, 6);
}

function stripeColDValue(description) {
  if (description === '' || description == null) return '';
  if (typeof description === 'number' && Number.isFinite(description)) {
    return String(Math.trunc(description));
  }
  return String(description).trim();
}

function orderMapsToStripe(orderId, stripeDescription) {
  const orderKey = orderColAMatchKey(orderId);
  const stripeD = stripeColDValue(stripeDescription);
  if (!orderKey || !stripeD) return false;
  return orderKey === stripeD;
}

module.exports = {
  orderColAMatchKey,
  stripeColDValue,
  orderMapsToStripe,
};
