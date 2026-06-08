# Ecommerce Reconciliation App

Full-stack web app to upload or sync ecommerce **orders** and **payments**, then reconcile them by **Order ID** and **Amount**.

## Project structure

```
ecommerce-reconciliation-app/
├── backend/
│   ├── data/                 # SQLite database (created on first run)
│   ├── src/
│   │   ├── index.js          # Express server
│   │   ├── db.js             # SQLite setup
│   │   ├── routes/
│   │   │   ├── orders.js
│   │   │   ├── payments.js
│   │   │   ├── reconcile.js
│   │   │   └── sync.js
│   │   └── services/
│   │       ├── reconciliation.js
│   │       ├── shopify.js
│   │       └── stripe.js
│   ├── .env.example
│   └── package.json
├── frontend/
│   ├── src/
│   │   ├── api.js
│   │   ├── App.jsx
│   │   └── components/
│   │       └── Dashboard.jsx
│   └── package.json
├── sample-data/
│   ├── orders.csv
│   └── payments.csv
└── README.md
```

## Prerequisites

- [Node.js](https://nodejs.org/) 18 or later
- npm (included with Node.js)

## Run locally

### 1. Install dependencies

```bash
cd backend
npm install

cd ../frontend
npm install
```

### 2. Configure API keys

Copy the example env file and add your credentials:

```bash
cd backend
copy .env.example .env   # Windows
# cp .env.example .env  # macOS/Linux
```

Edit `backend/.env`:

| Variable | Source |
|----------|--------|
| `SHOPIFY_STORE_DOMAIN` | Your shop domain, e.g. `my-store.myshopify.com` |
| `SHOPIFY_ACCESS_TOKEN` | Shopify Admin → custom app → Admin API access token |
| `STRIPE_SECRET_KEY` | Stripe Dashboard → Developers → API keys (secret key) |
| `OPENAI_API_KEY` | OpenAI platform → API keys (for `/insights`) |

Optional: `SHOPIFY_API_VERSION` (default `2024-10`), `PORT` (default `3001`).

### 3. Start the backend API

```bash
cd backend
npm run dev
```

API runs at **http://localhost:3001**

### 4. Start the frontend (new terminal)

```bash
cd frontend
npm run dev
```

Dashboard runs at **http://localhost:5173** (proxies `/api` to the backend).

### 5. Use the dashboard

1. Click **Sync sample orders** and **Sync sample payments** (or upload CSV files from `sample-data/`).
2. Click **Run reconciliation**.
3. Review the table: **Order ID | Amount | Match Status**. Unmatched rows are highlighted in red.

## API endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/orders` | List all orders |
| POST | `/orders` | Add order(s) JSON: `{ order_id, amount }` or array |
| POST | `/orders/upload` | Upload CSV (`file` field) |
| DELETE | `/orders` | Clear all orders |
| GET | `/payments` | List all payments |
| POST | `/payments` | Add payment(s) JSON |
| POST | `/payments/upload` | Upload CSV |
| DELETE | `/payments` | Clear all payments |
| GET | `/reconcile` | Run reconciliation and return results |
| POST | `/reconcile` | Same as GET (explicit run) |
| POST | `/sync/orders` | Fetch orders from **Shopify** and store in DB |
| POST | `/sync/payments` | Fetch payments from **Stripe** and store in DB |
| GET | `/insights` | AI analysis of unmatched/exception transactions (OpenAI) |
| GET | `/export/reconciliation` | Download reconciliation results as `.xlsx` |

### External sync

**Shopify** (`POST /sync/orders`): loads all orders via Admin API. `order_id` is the order name without `#` (e.g. `1001`). Amount is `current_total_price`.

**Stripe** (`POST /sync/payments`): loads succeeded PaymentIntents. `order_id` is read from metadata (`order_id`, `shopify_order_id`, etc.). Amount is in dollars.

Optional body/query: `{ "replace": true }` clears existing rows before insert.

```bash
curl -X POST http://localhost:3001/sync/orders
curl -X POST "http://localhost:3001/sync/payments?replace=true"
```

For reliable matching, set Stripe PaymentIntent metadata `order_id` (or `shopify_order_id`) to the same value as the Shopify order name/number.

## AI insights

`GET /insights` runs reconciliation, then analyzes unmatched and exception rows using the **OpenAI API**.

Suggested reasons per transaction:

- **Missing payment** — no matching payment or order
- **Amount mismatch** — order ID matches but amounts differ
- **Duplicate order** — same Order ID appears multiple times in orders

Set `OPENAI_API_KEY` in `backend/.env` (optional `OPENAI_MODEL`, default `gpt-4o-mini`). If OpenAI is unavailable, the API falls back to rule-based insights.

Insights appear automatically on the dashboard when issues exist.

## Scheduled daily job

Uses **node-cron** when the backend starts. Each run:

1. Fetches **new** Shopify orders since the last sync (`created_at_min`)
2. Fetches **new** Stripe payments since the last sync (`created` filter)
3. Runs reconciliation
4. Logs results to the console with `[cron]` prefix

Configure in `backend/.env`:

| Variable | Default | Description |
|----------|---------|-------------|
| `CRON_ENABLED` | `true` | Set to `false` to disable |
| `CRON_SCHEDULE` | `0 2 * * *` | Cron expression (daily at 2:00 AM) |
| `CRON_TIMEZONE` | — | Optional IANA timezone |

Run the job manually:

```bash
cd backend
npm run job:run
```

### Example: sync orders via curl

```bash
curl -X POST http://localhost:3001/orders \
  -H "Content-Type: application/json" \
  -d '[{"order_id":"ORD-1001","amount":49.99}]'
```

### Example: reconcile

```bash
curl http://localhost:3001/reconcile
```

## Matching rules

Reconciliation runs in `backend/src/services/reconciliation.js` and returns three groups:

| Status | Condition |
|--------|-----------|
| **Matched** | Same `order_id` and same `amount` (within $0.01) |
| **Unmatched** | Order has no payment for that `order_id`, or payment has no order |
| **Exception** | Payment exists for the `order_id` but `amount` does not match |

API response shape:

```json
{
  "summary": { "matchedCount", "unmatchedCount", "exceptionCount" },
  "matched": [],
  "unmatched": [],
  "exception": []
}
```

## CSV format

```csv
order_id,amount
ORD-1001,49.99
ORD-1002,120.00
```

Header row is optional; columns are `order_id` and `amount`.

**Order.xlsx (Stripe):** Use **Import Stripe (F vs H)** on the dashboard, or:

```bash
curl -X POST "http://localhost:3001/import/order-xlsx?paymentType=STRIPE&replace=true" \
  -F "file=@C:/Users/you/Desktop/Order.xlsx"
```

Maps: Column **A** = Order ID, **E** = Payment Type (filter `STRIPE`), **F** = AMT (orders), **H** = Final rounded price paid (payments).

**CSV uploads:** Upload `.csv` files only — not Excel `.xlsx` unless using the Order.xlsx import above. In Excel use **File → Save As → CSV UTF-8**. GBK-encoded CSV (common on Chinese Windows) is also supported. If order IDs show as garbled symbols, click **Reset all data** on the dashboard and re-upload.

## Production build (frontend)

```bash
cd frontend
npm run build
npm run preview
```

Set `VITE_API_URL` or configure your reverse proxy if the API is not on the same host.
