require('dotenv').config();

process.on('uncaughtException', (err) => {
  console.error('Uncaught exception:', err);
  if (process.env.NODE_ENV !== 'production') {
    process.exit(1);
  }
});

process.on('unhandledRejection', (reason) => {
  console.error('Unhandled rejection:', reason);
  if (process.env.NODE_ENV !== 'production') {
    process.exit(1);
  }
});

const path = require('path');
const express = require('express');
const cors = require('cors');

require('./db');

const ordersRouter = require('./routes/orders');
const paymentsRouter = require('./routes/payments');
const reconcileRouter = require('./routes/reconcile');
const syncRouter = require('./routes/sync');
const insightsRouter = require('./routes/insights');
const dataRouter = require('./routes/data');
const importRouter = require('./routes/import');
const exportRouter = require('./routes/export');
const authRouter = require('./routes/auth');
const { requireAppAuth } = require('./middleware/appAuth');
const { startScheduler } = require('./scheduler');

const app = express();
const PORT = process.env.PORT || 3001;
const isProduction = process.env.NODE_ENV === 'production';

app.use(cors());
app.use(express.json());

const api = express.Router();

api.get('/health', (_req, res) => {
  res.json({ status: 'ok' });
});

api.use('/auth', authRouter);
api.use(requireAppAuth);

api.use('/orders', ordersRouter);
api.use('/payments', paymentsRouter);
api.use('/reconcile', reconcileRouter);
api.use('/sync', syncRouter);
api.use('/insights', insightsRouter);
api.use('/data', dataRouter);
api.use('/import', importRouter);
api.use('/export', exportRouter);

app.use('/api', api);

if (!isProduction) {
  app.get('/health', (_req, res) => {
    res.json({ status: 'ok' });
  });
  app.use('/auth', authRouter);
  app.use(requireAppAuth);
  app.use('/orders', ordersRouter);
  app.use('/payments', paymentsRouter);
  app.use('/reconcile', reconcileRouter);
  app.use('/sync', syncRouter);
  app.use('/insights', insightsRouter);
  app.use('/data', dataRouter);
  app.use('/import', importRouter);
  app.use('/export', exportRouter);
}

if (isProduction) {
  const distPath = path.join(__dirname, '..', '..', 'frontend', 'dist');
  app.use(express.static(distPath));
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api')) {
      return next();
    }
    res.sendFile(path.join(distPath, 'index.html'), (err) => {
      if (err) next(err);
    });
  });
}

app.use('/api', (_req, res) => {
  res.status(404).json({ error: 'API route not found' });
});

app.use((err, _req, res, _next) => {
  console.error(err);
  res.status(500).json({ error: 'Internal server error' });
});

const server = app.listen(PORT, '0.0.0.0', () => {
  console.log(
    isProduction
      ? `Reconciliation app running at http://0.0.0.0:${PORT}`
      : `Reconciliation API running at http://localhost:${PORT}`,
  );
  startScheduler();
});

server.keepAliveTimeout = 120000;
server.headersTimeout = 125000;
