const API_BASE = '/api';

async function request(path, options = {}) {
  let res;
  try {
    res = await fetch(`${API_BASE}${path}`, options);
  } catch (err) {
    throw new Error(err.message || 'Network error — the server may be restarting. Try again in a moment.');
  }

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(
      data.error || res.statusText || `Request failed (${res.status})`
    );
  }
  return data;
}

export function fetchReconciliation() {
  return request('/reconcile');
}

export function runReconciliation() {
  return request('/reconcile', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: '{}',
  });
}

export function fetchReconciliation2() {
  return request('/reconcile/2');
}

export function runReconciliation2() {
  return request('/reconcile/2', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: '{}',
  });
}

export function syncOrders(orders) {
  return request('/orders', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(orders),
  });
}

export function syncPayments(payments) {
  return request('/payments', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payments),
  });
}

export function uploadOrdersCsv(file) {
  const form = new FormData();
  form.append('file', file);
  return request('/orders/upload', { method: 'POST', body: form });
}

export function uploadPaymentsCsv(file) {
  const form = new FormData();
  form.append('file', file);
  return request('/payments/upload', { method: 'POST', body: form });
}

export function clearOrders() {
  return request('/orders', { method: 'DELETE' });
}

export function clearPayments() {
  return request('/payments', { method: 'DELETE' });
}

export function syncOrdersFromShopify(replace = false) {
  return request('/sync/orders', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ replace }),
  });
}

export function syncPaymentsFromStripe(replace = false) {
  return request('/sync/payments', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ replace }),
  });
}

export function fetchInsights() {
  return request('/insights');
}

export function fetchDataHealth() {
  return request('/data/health');
}

export function resetAllData() {
  return request('/data/reset', { method: 'POST' });
}

export function importOrderXlsx(file, { paymentType = 'STRIPE', replace = true } = {}) {
  const form = new FormData();
  form.append('file', file);
  return request(
    `/import/order-xlsx?paymentType=${encodeURIComponent(paymentType)}&replace=${replace}`,
    { method: 'POST', body: form }
  );
}

export function importStripeXlsx(file, { replace = true } = {}) {
  const form = new FormData();
  form.append('file', file);
  return request(
    `/import/stripe-xlsx?replace=${replace}`,
    { method: 'POST', body: form }
  );
}

export function clearNcImport() {
  return request('/import/order-xlsx', { method: 'DELETE' });
}

export function clearStripeImport() {
  return request('/import/stripe-xlsx', { method: 'DELETE' });
}

async function downloadExcelExport(path, fallbackPrefix) {
  const res = await fetch(`${API_BASE}${path}`);
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.error || 'Export failed');
  }

  const blob = await res.blob();
  const disposition = res.headers.get('Content-Disposition') || '';
  const match = disposition.match(/filename="?([^"]+)"?/);
  const filename =
    match?.[1] || `${fallbackPrefix}-${new Date().toISOString().slice(0, 10)}.xlsx`;

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  return { filename };
}

export function downloadReconciliationExcel() {
  return downloadExcelExport('/export/reconciliation', 'reconciliation');
}

export function downloadReconciliationExcel2() {
  return downloadExcelExport('/export/reconciliation-2', 'reconciliation-2');
}

export async function createEskPayment(skFile, ncFile) {
  const form = new FormData();
  form.append('skFile', skFile);
  form.append('ncFile', ncFile);

  const res = await fetch(`${API_BASE}/export/esk-payment`, {
    method: 'POST',
    body: form,
  });

  if (!res.ok) {
    const text = await res.text();
    let message = `Failed to create ESK-Payment file (${res.status})`;
    try {
      const data = JSON.parse(text);
      if (data.error) message = data.error;
    } catch {
      if (text.trim()) message = text.trim().slice(0, 300);
    }
    throw new Error(message);
  }

  const blob = await res.blob();
  const disposition = res.headers.get('Content-Disposition') || '';
  const match = disposition.match(/filename="?([^"]+)"?/);
  const filename = match?.[1] || 'ESK-Payment Listing.xlsx';
  const generated = res.headers.get('X-Esk-Generated-Rows');

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  return { filename, generated: generated ? Number(generated) : null };
}
