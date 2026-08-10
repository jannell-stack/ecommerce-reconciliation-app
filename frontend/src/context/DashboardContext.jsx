import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
} from 'react';
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
  createGpay,
  createShopback,
  createStripeAdvice,
} from '../api';
import {
  applyData,
  isReadableOrderId,
  SAMPLE_ORDERS,
  SAMPLE_PAYMENTS,
  SHOW_AI_INSIGHTS,
} from '../components/dashboard/shared';

const DashboardContext = createContext(null);

export function DashboardProvider({ children }) {
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
  const [stripeAdviceStripeFile, setStripeAdviceStripeFile] = useState(null);
  const [stripeAdviceNcFiles, setStripeAdviceNcFiles] = useState([]);
  const [stripeAdviceCreating, setStripeAdviceCreating] = useState(false);
  const [gpayGrabFile, setGpayGrabFile] = useState(null);
  const [gpayEskFiles, setGpayEskFiles] = useState([]);
  const [gpayCreating, setGpayCreating] = useState(false);
  const [shopbackTransactionFile, setShopbackTransactionFile] = useState(null);
  const [shopbackEskFiles, setShopbackEskFiles] = useState([]);
  const [shopbackCreating, setShopbackCreating] = useState(false);

  const setters = useMemo(
    () => ({ setMatched, setUnmatched, setSummary }),
    [],
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
      if (SHOW_AI_INSIGHTS) {
        const issues = [
          ...(data.unmatched || []),
          ...(data.exception || []),
        ];
        if (issues.length > 0) {
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

  const handleAction = useCallback(
    async (action, { runReconcile = '1' } = {}) => {
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
    },
    [loadInsights, setters, loadDataHealth],
  );

  const handleReconciliation1 = useCallback(async () => {
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
  }, [loadInsights, setters, loadDataHealth]);

  const handleReconciliation2 = useCallback(async () => {
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
  }, [loadInsights, setters, loadDataHealth]);

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

  const value = {
    matched,
    unmatched,
    summary,
    statusFilter,
    setStatusFilter,
    loading,
    error,
    setError,
    message,
    setMessage,
    insights,
    setInsights,
    insightsLoading,
    insightsError,
    dataHealth,
    eskSkFile,
    setEskSkFile,
    eskNcFile,
    setEskNcFile,
    eskCreating,
    setEskCreating,
    stripeAdviceStripeFile,
    setStripeAdviceStripeFile,
    stripeAdviceNcFiles,
    setStripeAdviceNcFiles,
    stripeAdviceCreating,
    setStripeAdviceCreating,
    gpayGrabFile,
    setGpayGrabFile,
    gpayEskFiles,
    setGpayEskFiles,
    gpayCreating,
    setGpayCreating,
    shopbackTransactionFile,
    setShopbackTransactionFile,
    shopbackEskFiles,
    setShopbackEskFiles,
    shopbackCreating,
    setShopbackCreating,
    issueCount,
    loadInsights,
    loadReconciliation,
    handleAction,
    handleReconciliation1,
    handleReconciliation2,
    dataLooksCorrupted,
    filteredMatched,
    filteredUnmatched,
    showMatchedTable,
    showUnmatchedTable,
    SAMPLE_ORDERS,
    SAMPLE_PAYMENTS,
    clearOrders,
    clearPayments,
    resetAllData,
    importOrderXlsx,
    importStripeXlsx,
    clearNcImport,
    clearStripeImport,
    syncOrdersFromShopify,
    syncPaymentsFromStripe,
    syncOrders,
    syncPayments,
    uploadOrdersCsv,
    uploadPaymentsCsv,
    downloadReconciliationExcel,
    downloadReconciliationExcel2,
    createEskPayment,
    createGpay,
    createShopback,
    createStripeAdvice,
  };

  return (
    <DashboardContext.Provider value={value}>
      {children}
    </DashboardContext.Provider>
  );
}

export function useDashboard() {
  const ctx = useContext(DashboardContext);
  if (!ctx) {
    throw new Error('useDashboard must be used within DashboardProvider');
  }
  return ctx;
}
