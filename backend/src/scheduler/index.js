const cron = require('node-cron');
const { runDailySync } = require('../services/dailyJob');

let isRunning = false;

async function executeJob() {
  if (isRunning) {
    console.warn('[cron] Skipping run — previous daily sync still in progress');
    return;
  }

  isRunning = true;
  try {
    await runDailySync();
  } catch (err) {
    console.error('[cron] Daily sync failed:', err.message);
    if (err.stack) console.error(err.stack);
  } finally {
    isRunning = false;
  }
}

function startScheduler() {
  const enabled = process.env.CRON_ENABLED !== 'false';
  const schedule = process.env.CRON_SCHEDULE || '0 2 * * *';

  if (!enabled) {
    console.log('[cron] Scheduler disabled (set CRON_ENABLED=true to enable)');
    return;
  }

  if (!cron.validate(schedule)) {
    console.error(`[cron] Invalid CRON_SCHEDULE: "${schedule}"`);
    return;
  }

  cron.schedule(schedule, executeJob, {
    timezone: process.env.CRON_TIMEZONE || undefined,
  });

  console.log(`[cron] Daily job scheduled: "${schedule}"`);
}

module.exports = { startScheduler, executeJob, runDailySync };
