import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  clearOrders,
  clearPayments,
  fetchReconciliation,
  runReconciliation,
  syncOrders,
  syncPayments,
  uploadOrdersCsv,
  uploadPaymentsCsv,
  syncOrdersFromShopify,
  syncPaymentsFromStripe,
  fetchInsights,
  fetchDataHealth,
  resetAllData,
  importOrderXlsx,
  importStripeXlsx,
  clearNcImport,
  clearStripeImport,
  runReconciliation2,
  downloadReconciliationExcel,
  downloadReconciliationExcel2,
  createEskPayment,
} from '../api';

const ORDER_ID_PATTERN = /^[A-Za-z0-9#._\-/ ]{1,64}$/;

function isReadableOrderId(value) {
  return ORDER_ID_PATTERN.test(String(value ?? '').trim());
}
import './Dashboard.css';

/** Set to true to show the AI insights panel and fetch /api/insights. */
const SHOW_AI_INSIGHTS = false;

const SAMPLE_ORDERS = [
  { order_id: 'ORD-1001', amount: 49.99 },
  { order_id: 'ORD-1002', amount: 120.0 },
  { order_id: 'ORD-1003', amount: 15.5 },
  { order_id: 'ORD-1004', amount: 89.0 },
];

const SAMPLE_PAYMENTS = [
  { order_id: 'ORD-1001', amount: 49.99 },
  { order_id: 'ORD-1002', amount: 120.0 },
  { order_id: 'ORD-1003', amount: 20.0 },
  { order_id: 'ORD-1005', amount: 55.0 },
];

const STATUS_FILTERS = [
  { id: 'all', label: 'All' },
  { id: 'matched', label: 'Matched' },
  { id: 'unmatched', label: 'Unmatched' },
  { id: 'exception', label: 'Exception' },
];

function formatAmount(value) {
  return Number(value).toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function statusLabel(status) {
  if (status === 'matched') return 'Matched';
  if (status === 'exception') return 'Exception';
  return 'Unmatched';
}

function applyData(setters, data) {
  setters.setMatched(data.matched || []);
  setters.setUnmatched([
    ...(data.unmatched || []),
    ...(data.exception || []),
  ]);
  setters.setSummary(data.summary || null);
}

function SummaryCards({ summary }) {
  const matchedPct =
    summary.totalOrders > 0
      ? Math.round((summary.matchedCount / summary.totalOrders) * 100)
      : 0;

  return (
    <div className="summary-cards">
      <article className="summary-card">
        <span className="summary-card__label">Total orders</span>
        <span className="summary-card__value">{summary.totalOrders}</span>
        <span className="summary-card__hint">
          {summary.totalPayments} payments loaded
        </span>
      </article>
      <article className="summary-card summary-card--accent">
        <span className="summary-card__label">Matched %</span>
        <span className="summary-card__value">{matchedPct}%</span>
        <span className="summary-card__hint">
          {summary.matchedCount} of {summary.totalOrders} orders
        </span>
      </article>
      <article className="summary-card summary-card--warn">
        <span className="summary-card__label">Unmatched count</span>
        <span className="summary-card__value">{summary.unmatchedCount}</span>
        <span className="summary-card__hint">
          {summary.exceptionCount ?? 0} amount exceptions
        </span>
      </article>
    </div>
  );
}

function StatusFilter({ value, onChange }) {
  return (
    <div className="status-filter" role="group" aria-label="Filter by status">
      {STATUS_FILTERS.map((f) => (
        <button
          key={f.id}
          type="button"
          className={`status-filter__btn ${value === f.id ? 'status-filter__btn--active' : ''}`}
          onClick={() => onChange(f.id)}
        >
          {f.label}
        </button>
      ))}
    </div>
  );
}

const REASON_ICONS = {
  missing_payment: '○',
  amount_mismatch: '≠',
  duplicate_order: '⧉',
};

function InsightsPanel({ insights, loading, error, onRefresh, issueCount }) {
  if (!insights && !loading && !error && issueCount === 0) return null;

  return (
    <section className="insights-panel">
      <div className="insights-panel__head">
        <div>
          <h2>AI insights</h2>
          <p className="insights-panel__subtitle">
            Analysis of unmatched and exception transactions
          </p>
        </div>
        <button
          type="button"
          className="btn-ghost"
          disabled={loading || issueCount === 0}
          onClick={onRefresh}
        >
          {loading ? 'Analyzing…' : 'Refresh insights'}
        </button>
      </div>

      {error && <div className="alert alert--error">{error}</div>}

      {loading && !insights && (
        <p className="insights-panel__loading">Generating insights…</p>
      )}

      {insights && (
        <>
          <p className="insights-panel__overview">{insights.overview}</p>
          {insights.info && (
            <p className="insights-panel__hint">{insights.info}</p>
          )}
          {insights.warning && (
            <p className="insights-panel__warning">{insights.warning}</p>
          )}
          {insights.truncated && (
            <p className="insights-panel__hint">
              Showing {insights.showing} of {insights.issueCount} issues.
            </p>
          )}
          {insights.insights?.length === 0 ? (
            <p className="insights-panel__empty">No issues to analyze.</p>
          ) : (
            <ul className="insights-list">
              {insights.insights.map((item) => (
                <li
                  key={item.orderId}
                  className={`insight-card insight-card--${item.suggestedReason}`}
                >
                  <div className="insight-card__top">
                    <span className="insight-card__icon" aria-hidden>
                      {REASON_ICONS[item.suggestedReason] || '•'}
                    </span>
                    <div>
                      <span className="insight-card__order">{item.orderId}</span>
                      <span className={`insight-card__reason insight-card__reason--${item.suggestedReason}`}>
                        {item.reasonLabel}
                      </span>
                    </div>
                    {item.confidence && (
                      <span className="insight-card__confidence">
                        {item.confidence}
                      </span>
                    )}
                  </div>
                  <p className="insight-card__text">{item.explanation}</p>
                </li>
              ))}
            </ul>
          )}
        </>
      )}
    </section>
  );
}

function TransactionTable({ title, rows, variant, emptyMessage }) {
  return (
    <section className={`txn-panel txn-panel--${variant}`}>
      <div className="txn-panel__head">
        <h2>{title}</h2>
        <span className="txn-panel__count">{rows.length}</span>
      </div>
      <div className="table-wrap">
        <table className="recon-table">
          <thead>
            <tr>
              <th>Order ID</th>
              <th>Amount</th>
              <th>Status</th>
              <th>Type</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={4} className="empty">
                  {emptyMessage}
                </td>
              </tr>
            ) : (
              rows.map((row, i) => (
                <tr
                  key={`${row.orderId}-${row.status}-${row.type}-${i}`}
                  className={
                    variant === 'issues' ? 'row-issue' : ''
                  }
                >
                  <td className="cell-mono">{row.orderId}</td>
                  <td>
                    {row.status === 'exception'
                      ? `${formatAmount(row.orderAmount)} → ${formatAmount(row.paymentAmount)}`
                      : formatAmount(row.amount)}
                  </td>
                  <td>
                    <span className={`badge badge--${row.status}`}>
                      {statusLabel(row.status)}
                    </span>
                  </td>
                  <td className="cell-muted">{row.type || '—'}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}

export default function Dashboard() {
  const [matched, setMatched] = useState([]);
  const [unmatched, setUnmatched] = useState([]);
  const [summary, setSummary] = useState(null);
  const [statusFilter, setStatusFilter] = useState('all');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [message, setMessage] = useState(null);
  const [insights, setInsights] = useState(null);
  const [insightsLoading, setInsightsLoading] = useState(false);
  const [insightsError, setInsightsError] = useState(null);
  const [dataHealth, setDataHealth] = useState(null);
  const [eskSkFile, setEskSkFile] = useState(null);
  const [eskNcFile, setEskNcFile] = useState(null);
  const [eskCreating, setEskCreating] = useState(false);

  const setters = useMemo(
    () => ({ setMatched, setUnmatched, setSummary }),
    []
  );

  const issueCount = useMemo(() => unmatched.length, [unmatched]);

  const loadInsights = useCallback(async () => {
    if (!SHOW_AI_INSIGHTS) return;
    setInsightsLoading(true);
    setInsightsError(null);
    try {
      const data = await fetchInsights();
      setInsights(data);
    } catch (err) {
      setInsightsError(err.message);
    } finally {
      setInsightsLoading(false);
    }
  }, []);

  const loadReconciliation = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await fetchReconciliation();
      applyData(setters, data);
      const issues = [
        ...(data.unmatched || []),
        ...(data.exception || []),
      ];
      if (SHOW_AI_INSIGHTS && issues.length > 0) {
        setInsightsLoading(true);
        try {
          const insightData = await fetchInsights();
          setInsights(insightData);
          setInsightsError(null);
        } catch (err) {
          setInsightsError(err.message);
        } finally {
          setInsightsLoading(false);
        }
      } else {
        setInsights(null);
        setInsightsError(null);
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }, [setters]);

  const loadDataHealth = useCallback(async () => {
    try {
      const health = await fetchDataHealth();
      setDataHealth(health);
    } catch {
      setDataHealth(null);
    }
  }, []);

  useEffect(() => {
    loadReconciliation();
    loadDataHealth();
  }, [loadReconciliation, loadDataHealth]);

  async function handleAction(action, { runReconcile = '1' } = {}) {
    setLoading(true);
    setError(null);
    setMessage(null);
    try {
      await action();
      if (runReconcile === '1') {
        const data = await runReconciliation();
        applyData(setters, data);
        const issues = [
          ...(data.unmatched || []),
          ...(data.exception || []),
        ];
        if (issues.length > 0) await loadInsights();
        else setInsights(null);
      } else if (runReconcile === '2') {
        const data = await runReconciliation2();
        applyData(setters, data);
        const issues = [
          ...(data.unmatched || []),
          ...(data.exception || []),
        ];
        if (issues.length > 0) await loadInsights();
        else setInsights(null);
      }
      await loadDataHealth();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleReconciliation1() {
    setLoading(true);
    setError(null);
    setMessage(null);
    try {
      const data = await runReconciliation();
      applyData(setters, data);
      setMessage(data.message || 'Reconciliation complete');
      const issues = [
        ...(data.unmatched || []),
        ...(data.exception || []),
      ];
      if (issues.length > 0) await loadInsights();
      else setInsights(null);
      await loadDataHealth();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleReconciliation2() {
    setLoading(true);
    setError(null);
    setMessage(null);
    try {
      const data = await runReconciliation2();
      applyData(setters, data);
      setMessage(data.message || 'Reconciliation 2 complete');
      const issues = [
        ...(data.unmatched || []),
        ...(data.exception || []),
      ];
      if (issues.length > 0) await loadInsights();
      else setInsights(null);
      await loadDataHealth();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  const hasCorruptedRows = useMemo(() => {
    const rows = [...matched, ...unmatched];
    return rows.some((r) => !isReadableOrderId(r.orderId));
  }, [matched, unmatched]);

  const dataLooksCorrupted =
    hasCorruptedRows || dataHealth?.dataHealthy === false;

  const filteredMatched = useMemo(() => {
    if (statusFilter === 'all' || statusFilter === 'matched') return matched;
    return [];
  }, [matched, statusFilter]);

  const filteredUnmatched = useMemo(() => {
    if (statusFilter === 'matched') return [];
    if (statusFilter === 'unmatched') {
      return unmatched.filter((r) => r.status === 'unmatched');
    }
    if (statusFilter === 'exception') {
      return unmatched.filter((r) => r.status === 'exception');
    }
    return unmatched;
  }, [unmatched, statusFilter]);

  const showMatchedTable =
    statusFilter === 'all' || statusFilter === 'matched';
  const showUnmatchedTable =
    statusFilter === 'all' ||
    statusFilter === 'unmatched' ||
    statusFilter === 'exception';

  return (
    <div className="dashboard">
      <header className="dashboard__header">
        <div className="dashboard__brand">
          <h1>Ecommerce Reconciliation</h1>
        </div>
        <div className="dashboard__header-actions">
          <button
            type="button"
            className="btn-export"
            disabled={loading}
            onClick={handleReconciliation1}
          >
            {loading ? 'Running…' : 'Run Reconciliation 1'}
          </button>
          <button
            type="button"
            className="btn-primary btn-recon2"
            disabled={loading}
            onClick={handleReconciliation2}
          >
            {loading ? 'Running…' : 'Run Reconciliation 2'}
          </button>
          <button
            type="button"
            className="btn-ghost"
            disabled={loading}
            onClick={loadReconciliation}
          >
            Refresh
          </button>
          <button
            type="button"
            className="btn-export"
            disabled={loading || !summary}
            onClick={async () => {
              setError(null);
              setMessage(null);
              try {
                const { filename } = await downloadReconciliationExcel();
                setMessage(`Exported ${filename}`);
              } catch (err) {
                setError(err.message);
              }
            }}
          >
            Export Excel 1
          </button>
          <button
            type="button"
            className="btn-export btn-export-2"
            disabled={loading || !summary}
            onClick={async () => {
              setError(null);
              setMessage(null);
              try {
                const { filename } = await downloadReconciliationExcel2();
                setMessage(`Exported ${filename}`);
              } catch (err) {
                setError(err.message);
              }
            }}
          >
            Export Excel 2
          </button>
        </div>
      </header>

      <section className="toolbar">
        <div className="toolbar__group toolbar__group--excel toolbar__group--esk">
          <span className="toolbar__label">ESK Payment</span>
          <div className="toolbar__actions toolbar__actions--stacked">
            <label className="file-btn btn-integration">
              SK Combined
              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                disabled={loading || eskCreating}
                onChange={(e) => {
                  setEskSkFile(e.target.files?.[0] || null);
                  e.target.value = '';
                }}
              />
            </label>
            <label className="file-btn btn-integration">
              NC Listing
              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                disabled={loading || eskCreating}
                onChange={(e) => {
                  setEskNcFile(e.target.files?.[0] || null);
                  e.target.value = '';
                }}
              />
            </label>
            <button
              type="button"
              className="btn-primary btn-esk-create"
              disabled={loading || eskCreating || !eskSkFile || !eskNcFile}
              onClick={async () => {
                setEskCreating(true);
                setError(null);
                setMessage(null);
                try {
                  const result = await createEskPayment(eskSkFile, eskNcFile);
                  setMessage(
                    `Created ${result.filename}${
                      result.generated != null
                        ? ` (${result.generated} row(s))`
                        : ''
                    }`,
                  );
                  setEskSkFile(null);
                  setEskNcFile(null);
                } catch (err) {
                  setError(err.message);
                } finally {
                  setEskCreating(false);
                }
              }}
            >
              {eskCreating ? 'Creating…' : 'Create ESK-Payment'}
            </button>
          </div>
          <span className="toolbar__hint">
            {eskSkFile || eskNcFile
              ? `${eskSkFile?.name || 'SK file missing'} + ${eskNcFile?.name || 'NC file missing'}`
              : 'SK Combined + NC Payment Listing → ESK-Payment Listing.xlsx'}
          </span>
        </div>
        <div className="toolbar__group toolbar__group--excel">
          <span className="toolbar__label">ESK.xlsx</span>
          <div className="toolbar__actions">
            <label className="file-btn btn-integration">
              Import ESK
              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                disabled={loading}
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (!file) return;
                  handleAction(
                    async () => {
                      const r = await importOrderXlsx(file, {
                        paymentType: 'STRIPE',
                        replace: true,
                      });
                      setMessage(r.message);
                    },
                    { runReconcile: false }
                  );
                  e.target.value = '';
                }}
              />
            </label>
            <button
              type="button"
              className="btn-ghost"
              disabled={loading}
              onClick={() =>
                handleAction(async () => {
                  const r = await clearNcImport();
                  setMessage(r.message);
                })
              }
            >
              Clear
            </button>
          </div>
          <span className="toolbar__hint">
            Col A = Order ID, F = AMT, H = Paid, E = STRIPE
          </span>
        </div>
        <div className="toolbar__group toolbar__group--excel">
          <span className="toolbar__label">Stripe advice</span>
          <div className="toolbar__actions">
            <label className="file-btn btn-integration">
              Import Stripe
              <input
                type="file"
                accept=".xlsx,.xls"
                hidden
                disabled={loading}
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (!file) return;
                  handleAction(
                    async () => {
                      const r = await importStripeXlsx(file, { replace: true });
                      setMessage(
                        `${r.message} — ${r.stripeSynced} row(s). Click "Run Reconciliation 2" to match.`
                      );
                    },
                    { runReconcile: false }
                  );
                  e.target.value = '';
                }}
              />
            </label>
            <button
              type="button"
              className="btn-ghost"
              disabled={loading}
              onClick={() =>
                handleAction(
                  async () => {
                    const r = await clearStripeImport();
                    setMessage(r.message);
                  },
                  { runReconcile: false }
                )
              }
            >
              Clear
            </button>
          </div>
          <span className="toolbar__hint">
            Order col A (first 6 chars, e.g. 864859) → Stripe col D; amounts col F vs col F
          </span>
        </div>
        <div className="toolbar__group">
          <span className="toolbar__label">Orders</span>
          <button
            type="button"
            className="btn-integration"
            disabled={loading}
            onClick={() =>
              handleAction(async () => {
                const r = await syncOrdersFromShopify(true);
                setMessage(r.message);
              })
            }
          >
            Shopify sync
          </button>
          <button
            type="button"
            disabled={loading}
            onClick={() =>
              handleAction(async () => {
                const r = await syncOrders(SAMPLE_ORDERS);
                setMessage(r.message);
              })
            }
          >
            Sync sample
          </button>
          <label className="file-btn">
            Upload CSV
            <input
              type="file"
              accept=".csv"
              hidden
              disabled={loading}
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (!file) return;
                handleAction(async () => {
                  const r = await uploadOrdersCsv(file);
                  setMessage(r.message);
                });
                e.target.value = '';
              }}
            />
          </label>
          <button
            type="button"
            className="btn-ghost"
            disabled={loading}
            onClick={() =>
              handleAction(async () => {
                await clearOrders();
                setMessage('Orders cleared');
              })
            }
          >
            Clear
          </button>
        </div>
        <div className="toolbar__group">
          <span className="toolbar__label">Payments</span>
          <button
            type="button"
            className="btn-integration"
            disabled={loading}
            onClick={() =>
              handleAction(async () => {
                const r = await syncPaymentsFromStripe(true);
                setMessage(r.message);
              })
            }
          >
            Stripe sync
          </button>
          <button
            type="button"
            disabled={loading}
            onClick={() =>
              handleAction(async () => {
                const r = await syncPayments(SAMPLE_PAYMENTS);
                setMessage(r.message);
              })
            }
          >
            Sync sample
          </button>
          <label className="file-btn">
            Upload CSV
            <input
              type="file"
              accept=".csv"
              hidden
              disabled={loading}
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (!file) return;
                handleAction(async () => {
                  const r = await uploadPaymentsCsv(file);
                  setMessage(r.message);
                });
                e.target.value = '';
              }}
            />
          </label>
          <button
            type="button"
            className="btn-ghost"
            disabled={loading}
            onClick={() =>
              handleAction(async () => {
                await clearPayments();
                setMessage('Payments cleared');
              })
            }
          >
            Clear
          </button>
        </div>
      </section>

      {dataLooksCorrupted && (
        <div className="alert alert--corrupt">
          <div>
            <strong>Unreadable data detected.</strong> Order IDs look corrupted
            (often from uploading an Excel .xlsx file instead of CSV). Clear all
            data, then upload UTF-8 or GBK CSV files with columns{' '}
            <code>order_id</code> and <code>amount</code>.
            {dataHealth && (
              <span className="alert__meta">
                {' '}
                ({dataHealth.corruptedOrders} bad orders,{' '}
                {dataHealth.corruptedPayments} bad payments)
              </span>
            )}
          </div>
          <button
            type="button"
            className="btn-primary"
            disabled={loading}
            onClick={() =>
              handleAction(async () => {
                await resetAllData();
                setInsights(null);
                setMessage('All data cleared. Upload CSV files again.');
              })
            }
          >
            Reset all data
          </button>
        </div>
      )}

      {message && <div className="alert alert--success">{message}</div>}
      {error && <div className="alert alert--error">{error}</div>}

      {summary && <SummaryCards summary={summary} />}

      {SHOW_AI_INSIGHTS && (
        <InsightsPanel
          insights={insights}
          loading={insightsLoading}
          error={insightsError}
          issueCount={issueCount}
          onRefresh={loadInsights}
        />
      )}

      <section className="results">
        <div className="results__toolbar">
          <h2 className="results__title">Transactions</h2>
          <StatusFilter value={statusFilter} onChange={setStatusFilter} />
        </div>

        <div
          className={`results__grid ${
            showMatchedTable && showUnmatchedTable
              ? 'results__grid--dual'
              : 'results__grid--single'
          }`}
        >
          {showMatchedTable && (
            <TransactionTable
              title="Matched transactions"
              rows={filteredMatched}
              variant="matched"
              emptyMessage={
                statusFilter === 'matched'
                  ? 'No matched transactions for this filter.'
                  : 'No matched transactions yet. Run reconciliation after loading data.'
              }
            />
          )}
          {showUnmatchedTable && (
            <TransactionTable
              title="Unmatched transactions"
              rows={filteredUnmatched}
              variant="issues"
              emptyMessage={
                statusFilter === 'exception'
                  ? 'No amount exceptions.'
                  : statusFilter === 'unmatched'
                    ? 'No unmatched transactions.'
                    : 'No unmatched transactions yet.'
              }
            />
          )}
        </div>
      </section>
    </div>
  );
}
