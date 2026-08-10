import {
  InsightsPanel,
  SHOW_AI_INSIGHTS,
  StatusFilter,
  SummaryCards,
  TransactionTable,
} from '../dashboard/shared';
import { useDashboard } from '../../context/DashboardContext';

export default function DashboardAlertsAndResults() {
  const {
    loading,
    message,
    error,
    dataLooksCorrupted,
    dataHealth,
    handleAction,
    resetAllData,
    setInsights,
    setMessage,
    summary,
    insights,
    insightsLoading,
    insightsError,
    issueCount,
    loadInsights,
    statusFilter,
    setStatusFilter,
    showMatchedTable,
    showUnmatchedTable,
    filteredMatched,
    filteredUnmatched,
  } = useDashboard();

  return (
    <>
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
    </>
  );
}

function PageHero({ title, subtitle, step }) {
  return (
    <header className="page-hero">
      <h2 className="page-hero__title">{title}</h2>
      {subtitle ? <p className="page-hero__subtitle">{subtitle}</p> : null}
      {step ? <span className="page-hero__step">{step}</span> : null}
    </header>
  );
}

export { PageHero };
