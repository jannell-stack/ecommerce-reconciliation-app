import { appendUniqueFiles } from '../components/dashboard/shared';
import { useDashboard } from '../context/DashboardContext';
import { PageHero } from '../components/dashboard/DashboardAlertsAndResults';

export default function GPayPage() {
  const d = useDashboard();

  return (
    <div className="dashboard">
      <PageHero
        title="Generate GPay Advice File"
        step="A) Generate GPay Payment advice"
      />

      <section className="workflow-card">
        <h3 className="workflow-card__title">GPay advice generation</h3>
        <div className="toolbar__actions page-actions">
          <label className="file-btn btn-integration btn-gpay-file">
            Grab Transaction
            <input
              type="file"
              accept=".xlsx,.xls"
              hidden
              disabled={d.loading || d.gpayCreating}
              onChange={(e) => {
                d.setGpayGrabFile(e.target.files?.[0] || null);
                e.target.value = '';
              }}
            />
          </label>
          <label className="file-btn btn-integration btn-gpay-file">
            ESK for Grab
            {d.gpayEskFiles.length > 0 ? ` (${d.gpayEskFiles.length})` : ''}
            <input
              type="file"
              accept=".xlsx,.xls"
              multiple
              hidden
              disabled={d.loading || d.gpayCreating}
              onChange={(e) => {
                const picked = Array.from(e.target.files || []);
                if (picked.length === 0) return;
                d.setGpayEskFiles((prev) => appendUniqueFiles(prev, picked));
                e.target.value = '';
              }}
            />
          </label>
          {d.gpayEskFiles.length > 0 ? (
            <button
              type="button"
              className="btn-ghost"
              disabled={d.loading || d.gpayCreating}
              onClick={() => d.setGpayEskFiles([])}
            >
              Clear ESK
            </button>
          ) : null}
          <button
            type="button"
            className="btn-primary btn-gpay-create"
            disabled={
              d.loading
              || d.gpayCreating
              || !d.gpayGrabFile
              || d.gpayEskFiles.length === 0
            }
            onClick={async () => {
              d.setGpayCreating(true);
              d.setError(null);
              d.setMessage(null);
              try {
                const result = await d.createGpay(d.gpayGrabFile, d.gpayEskFiles);
                d.setMessage(
                  `Created ${result.filename}${
                    result.generated != null
                      ? ` (${result.generated} row(s)`
                      : ''
                  }${
                    result.linked != null ? `, ${result.linked} linked` : ''
                  }${result.generated != null ? ')' : ''}`,
                );
                d.setGpayGrabFile(null);
                d.setGpayEskFiles([]);
              } catch (err) {
                d.setError(err.message);
              } finally {
                d.setGpayCreating(false);
              }
            }}
          >
            {d.gpayCreating ? 'Creating…' : 'Create GPay'}
          </button>
        </div>
        <span className="toolbar__hint">
          {d.gpayGrabFile || d.gpayEskFiles.length > 0
            ? `${d.gpayGrabFile?.name || 'Grab file missing'} + ${
                d.gpayEskFiles.length > 0
                  ? d.gpayEskFiles.map((f) => f.name).join(', ')
                  : 'ESK file(s) missing'
              }`
            : 'Grab Transaction + ESK Payment Details → GPay.xlsx'}
        </span>
      </section>

      {d.message && <div className="alert alert--success">{d.message}</div>}
      {d.error && <div className="alert alert--error">{d.error}</div>}
    </div>
  );
}
