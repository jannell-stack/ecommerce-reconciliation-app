import { clearAuth, getStoredPassword, setStoredPassword } from './auth';

const API_BASE = '/api';

function authHeaders(extraHeaders = {}) {
  const password = getStoredPassword();
  const headers = { ...extraHeaders };
  if (password) {
    headers['X-App-Password'] = password;
  }
  return headers;
}

function handleUnauthorized(res) {
  if (res.status === 401) {
    clearAuth();
    window.location.reload();
    throw new Error('Session expired. Please sign in again.');
  }
}

export async function loginWithPassword(password) {
  const res = await fetch(`${API_BASE}/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ password }),
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || 'Invalid password');
  }

  setStoredPassword(password);
  return data;
}

async function request(path, options = {}) {
  let res;
  try {
    res = await fetch(`${API_BASE}${path}`, {
      ...options,
      headers: authHeaders(options.headers),
    });
  } catch (err) {
    throw new Error(err.message || 'Network error — the server may be restarting. Try again in a moment.');
  }

  handleUnauthorized(res);

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
  const res = await fetch(`${API_BASE}${path}`, {
    headers: authHeaders(),
  });
  handleUnauthorized(res);
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
    headers: authHeaders(),
    body: form,
  });

  handleUnauthorized(res);
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
  const filename = match?.[1] || 'ESK Payment Details.xlsx';
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

export async function createStripeAdvice(stripeFile, ncFiles) {
  const ncFileList = Array.isArray(ncFiles) ? ncFiles : [ncFiles];
  if (ncFileList.length === 0) {
    throw new Error('At least one NC Payment Listing file is required');
  }

  const form = new FormData();
  form.append('stripeFile', stripeFile);
  for (const ncFile of ncFileList) {
    form.append('ncFile', ncFile);
  }

  const res = await fetch(`${API_BASE}/export/stripe-advice`, {
    method: 'POST',
    headers: authHeaders(),
    body: form,
  });

  handleUnauthorized(res);
  if (!res.ok) {
    const text = await res.text();
    let message = `Failed to create Stripe Advice file (${res.status})`;
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
  const filename = match?.[1] || 'Stripe Advice.xlsx';
  const generated = res.headers.get('X-Stripe-Advice-Generated-Rows');
  const linked = res.headers.get('X-Stripe-Advice-Linked-Rows');

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  return {
    filename,
    generated: generated ? Number(generated) : null,
    linked: linked ? Number(linked) : null,
  };
}

export async function createGpay(grabFile, eskFiles) {
  const eskFileList = Array.isArray(eskFiles) ? eskFiles : [eskFiles];
  if (eskFileList.length === 0) {
    throw new Error('At least one ESK Payment Details file is required');
  }

  const form = new FormData();
  form.append('grabFile', grabFile);
  for (const eskFile of eskFileList) {
    form.append('eskFile', eskFile);
  }

  const res = await fetch(`${API_BASE}/export/gpay`, {
    method: 'POST',
    headers: authHeaders(),
    body: form,
  });

  handleUnauthorized(res);
  if (!res.ok) {
    const text = await res.text();
    let message = `Failed to create GPay file (${res.status})`;
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
  const filename = match?.[1] || 'GPay.xlsx';
  const generated = res.headers.get('X-GPay-Generated-Rows');
  const linked = res.headers.get('X-GPay-Linked-Rows');

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  return {
    filename,
    generated: generated ? Number(generated) : null,
    linked: linked ? Number(linked) : null,
  };
}

export async function createShopback(shopbackFile, eskFiles) {
  const eskFileList = Array.isArray(eskFiles) ? eskFiles : [eskFiles];
  if (eskFileList.length === 0) {
    throw new Error('At least one ESK Payment Details file is required');
  }

  const form = new FormData();
  form.append('shopbackFile', shopbackFile);
  for (const eskFile of eskFileList) {
    form.append('eskFile', eskFile);
  }

  const res = await fetch(`${API_BASE}/export/shopback`, {
    method: 'POST',
    headers: authHeaders(),
    body: form,
  });

  handleUnauthorized(res);
  if (!res.ok) {
    const text = await res.text();
    let message = `Failed to create Shopback file (${res.status})`;
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
  const filename = match?.[1] || 'Shopback.xlsx';
  const generated = res.headers.get('X-Shopback-Generated-Rows');
  const linked = res.headers.get('X-Shopback-Linked-Rows');

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  return {
    filename,
    generated: generated ? Number(generated) : null,
    linked: linked ? Number(linked) : null,
  };
}
