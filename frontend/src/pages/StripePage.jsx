import { appendUniqueFiles } from '../components/dashboard/shared';
import { PageHero } from '../components/dashboard/DashboardAlertsAndResults';
import { useDashboard } from '../context/DashboardContext';
export default function StripePage() {
  const d = useDashboard();

  return (
    <div className="dashboard">
      <PageHero
        title="Generate Stripe Payment Advice File"
        step="A) Generate Stripe Payment advice"
      />

      <section className="workflow-card">
        <h3 className="workflow-card__title">Stripe advice generation</h3>
        <div className="toolbar__actions page-actions">
          <label className="file-btn btn-integration">
            Stripe Payout
            <input
              type="file"
              accept=".xlsx,.xls"
              hidden
              disabled={d.loading || d.stripeAdviceCreating}
              onChange={(e) => {
                d.setStripeAdviceStripeFile(e.target.files?.[0] || null);
                e.target.value = '';
              }}
            />
          </label>
          <label className="file-btn btn-integration">
            ESK for Stripe
            {d.stripeAdviceNcFiles.length > 0
              ? ` (${d.stripeAdviceNcFiles.length})`
              : ''}
            <input
              type="file"
              accept=".xlsx,.xls"
              multiple
              hidden
              disabled={d.loading || d.stripeAdviceCreating}
              onChange={(e) => {
                const picked = Array.from(e.target.files || []);
                if (picked.length === 0) return;
                d.setStripeAdviceNcFiles((prev) => appendUniqueFiles(prev, picked));
                e.target.value = '';
              }}
            />
          </label>
          {d.stripeAdviceNcFiles.length > 0 ? (
            <button
              type="button"
              className="btn-ghost"
              disabled={d.loading || d.stripeAdviceCreating}
              onClick={() => d.setStripeAdviceNcFiles([])}
            >
              Clear NC
            </button>
          ) : null}
          <button
            type="button"
            className="btn-primary btn-stripe-advice-create"
            disabled={
              d.loading
              || d.stripeAdviceCreating
              || !d.stripeAdviceStripeFile
              || d.stripeAdviceNcFiles.length === 0
            }
            onClick={async () => {
              d.setStripeAdviceCreating(true);
              d.setError(null);
              d.setMessage(null);
              try {
                const result = await d.createStripeAdvice(
                  d.stripeAdviceStripeFile,
                  d.stripeAdviceNcFiles,
                );
                d.setMessage(
                  `Created ${result.filename}${
                    result.generated != null
                      ? ` (${result.generated} row(s)`
                      : ''
                  }${
                    result.linked != null ? `, ${result.linked} linked` : ''
                  }${result.generated != null ? ')' : ''}`,
                );
                d.setStripeAdviceStripeFile(null);
                d.setStripeAdviceNcFiles([]);
              } catch (err) {
                d.setError(err.message);
              } finally {
                d.setStripeAdviceCreating(false);
              }
            }}
          >
            {d.stripeAdviceCreating ? 'Creating…' : 'Create Stripe Advice'}
          </button>
        </div>
        <span className="toolbar__hint">
          {d.stripeAdviceStripeFile || d.stripeAdviceNcFiles.length > 0
            ? `${d.stripeAdviceStripeFile?.name || 'Stripe file missing'} + ${
                d.stripeAdviceNcFiles.length > 0
                  ? d.stripeAdviceNcFiles.map((file) => file.name).join(', ')
                  : 'NC file(s) missing'
              }`
            : 'Stripe Itemized + ESK Payment Details (multi-file) → Stripe Advice'}
        </span>
      </section>

      {d.message && <div className="alert alert--success">{d.message}</div>}
      {d.error && <div className="alert alert--error">{d.error}</div>}
    </div>
  );
}