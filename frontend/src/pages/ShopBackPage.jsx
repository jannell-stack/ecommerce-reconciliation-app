import { appendUniqueFiles } from '../components/dashboard/shared';
import { useDashboard } from '../context/DashboardContext';
import { PageHero } from '../components/dashboard/DashboardAlertsAndResults';

export default function ShopBackPage() {
  const d = useDashboard();

  return (
    <div className="dashboard">
      <PageHero
        title="Generate Shopback Advice File"
        step="A) Generate ShopBack Payment advice"
      />

      <section className="workflow-card">
        <h3 className="workflow-card__title">Shopback advice generation</h3>
        <div className="toolbar__actions page-actions">
          <label className="file-btn btn-integration btn-shopback-file">
            Shopback Transaction
            <input
              type="file"
              accept=".xlsx,.xls"
              hidden
              disabled={d.loading || d.shopbackCreating}
              onChange={(e) => {
                d.setShopbackTransactionFile(e.target.files?.[0] || null);
                e.target.value = '';
              }}
            />
          </label>
          <label className="file-btn btn-integration btn-shopback-file">
            ESK for Shopback
            {d.shopbackEskFiles.length > 0 ? ` (${d.shopbackEskFiles.length})` : ''}
            <input
              type="file"
              accept=".xlsx,.xls"
              multiple
              hidden
              disabled={d.loading || d.shopbackCreating}
              onChange={(e) => {
                const picked = Array.from(e.target.files || []);
                if (picked.length === 0) return;
                d.setShopbackEskFiles((prev) => appendUniqueFiles(prev, picked));
                e.target.value = '';
              }}
            />
          </label>
          {d.shopbackEskFiles.length > 0 ? (
            <button
              type="button"
              className="btn-ghost"
              disabled={d.loading || d.shopbackCreating}
              onClick={() => d.setShopbackEskFiles([])}
            >
              Clear ESK
            </button>
          ) : null}
          <button
            type="button"
            className="btn-primary btn-shopback-create"
            disabled={
              d.loading
              || d.shopbackCreating
              || !d.shopbackTransactionFile
              || d.shopbackEskFiles.length === 0
            }
            onClick={async () => {
              d.setShopbackCreating(true);
              d.setError(null);
              d.setMessage(null);
              try {
                const result = await d.createShopback(
                  d.shopbackTransactionFile,
                  d.shopbackEskFiles,
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
                d.setShopbackTransactionFile(null);
                d.setShopbackEskFiles([]);
              } catch (err) {
                d.setError(err.message);
              } finally {
                d.setShopbackCreating(false);
              }
            }}
          >
            {d.shopbackCreating ? 'Creating…' : 'Create Shopback'}
          </button>
        </div>
        <span className="toolbar__hint">
          {d.shopbackTransactionFile || d.shopbackEskFiles.length > 0
            ? `${d.shopbackTransactionFile?.name || 'Shopback file missing'} + ${
                d.shopbackEskFiles.length > 0
                  ? d.shopbackEskFiles.map((f) => f.name).join(', ')
                  : 'ESK file(s) missing'
              }`
            : 'Shopback Ori + ESK Payment Details → Shopback.xlsx'}
        </span>
      </section>

      {d.message && <div className="alert alert--success">{d.message}</div>}
      {d.error && <div className="alert alert--error">{d.error}</div>}
    </div>
  );
}
