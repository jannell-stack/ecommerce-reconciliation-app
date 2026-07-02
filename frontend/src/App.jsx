import ErrorBoundary from './components/ErrorBoundary';
import Dashboard from './components/Dashboard';

export default function App() {
  return (
    <ErrorBoundary>
      <Dashboard />
    </ErrorBoundary>
  );
}
