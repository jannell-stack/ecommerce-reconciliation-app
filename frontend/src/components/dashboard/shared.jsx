export const ORDER_ID_PATTERN = /^[A-Za-z0-9#._\-/ ]{1,64}$/;

export const SHOW_AI_INSIGHTS = false;

export const SAMPLE_ORDERS = [
  { order_id: 'ORD-1001', amount: 49.99 },
  { order_id: 'ORD-1002', amount: 120.0 },
  { order_id: 'ORD-1003', amount: 15.5 },
  { order_id: 'ORD-1004', amount: 89.0 },
];

export const SAMPLE_PAYMENTS = [
  { order_id: 'ORD-1001', amount: 49.99 },
  { order_id: 'ORD-1002', amount: 120.0 },
  { order_id: 'ORD-1003', amount: 20.0 },
  { order_id: 'ORD-1005', amount: 55.0 },
];

export const STATUS_FILTERS = [
  { id: 'all', label: 'All' },
  { id: 'matched', label: 'Matched' },
  { id: 'unmatched', label: 'Unmatched' },
  { id: 'exception', label: 'Exception' },
];

export function isReadableOrderId(value) {
  return ORDER_ID_PATTERN.test(String(value ?? '').trim());
}

export function formatAmount(value) {
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

export function applyData(setters, data) {
  setters.setMatched(data.matched || []);
  setters.setUnmatched([
    ...(data.unmatched || []),
    ...(data.exception || []),
  ]);
  setters.setSummary(data.summary || null);
}

export function appendUniqueFiles(prev, picked) {
  const next = [...prev];
  for (const file of picked) {
    const duplicate = next.some(
      (existing) => existing.name === file.name && existing.size === file.size,
    );
    if (!duplicate) next.push(file);
  }
  return next;
}

export function SummaryCards({ summary }) {
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

export function StatusFilter({ value, onChange }) {
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

export function InsightsPanel({ insights, loading, error, onRefresh, issueCount }) {
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

export function TransactionTable({ title, rows, variant, emptyMessage }) {
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
                  className={variant === 'issues' ? 'row-issue' : ''}
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
