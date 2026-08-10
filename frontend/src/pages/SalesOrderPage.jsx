import { PageHero } from '../components/dashboard/DashboardAlertsAndResults';
import { useDashboard } from '../context/DashboardContext';
export default function SalesOrderPage() {
  const d = useDashboard();

  return (
    <div className="dashboard">
      <PageHero
        title="Generate ESK Payment File"
        step="A) Generate Sales Order & Payment Listing Validation file"
      />

      <section className="workflow-card">
        <h3 className="workflow-card__title">ESK payment generation</h3>
        <div className="toolbar__actions page-actions">
          <label className="file-btn btn-integration btn-esk-file">
            SK Verification
            <input
              type="file"
              accept=".xlsx,.xls"
              hidden
              disabled={d.loading || d.eskCreating}
              onChange={(e) => {
                d.setEskSkFile(e.target.files?.[0] || null);
                e.target.value = '';
              }}
            />
          </label>
          <label className="file-btn btn-integration btn-esk-file">
            NC Listing
            <input
              type="file"
              accept=".xlsx,.xls"
              hidden
              disabled={d.loading || d.eskCreating}
              onChange={(e) => {
                d.setEskNcFile(e.target.files?.[0] || null);
                e.target.value = '';
              }}
            />
          </label>
          <button
            type="button"
            className="btn-primary btn-esk-create"
            disabled={d.loading || d.eskCreating || !d.eskSkFile || !d.eskNcFile}
            onClick={async () => {
              d.setEskCreating(true);
              d.setError(null);
              d.setMessage(null);
              try {
                const result = await d.createEskPayment(d.eskSkFile, d.eskNcFile);
                d.setMessage(
                  `Created ${result.filename}${
                    result.generated != null
                      ? ` (${result.generated} row(s))`
                      : ''
                  }`,
                );
                d.setEskSkFile(null);
                d.setEskNcFile(null);
              } catch (err) {
                d.setError(err.message);
              } finally {
                d.setEskCreating(false);
              }
            }}
          >
            {d.eskCreating ? 'Creating…' : 'Create ESK-Payment'}
          </button>
        </div>
        <span className="toolbar__hint">
          {d.eskSkFile || d.eskNcFile
            ? `${d.eskSkFile?.name || 'SK file missing'} + ${d.eskNcFile?.name || 'NC file missing'}`
            : 'Upload SK Verification and NC Listing, then create ESK Payment Details.'}
        </span>
      </section>

      {d.message && <div className="alert alert--success">{d.message}</div>}
      {d.error && <div className="alert alert--error">{d.error}</div>}
    </div>
  );
}