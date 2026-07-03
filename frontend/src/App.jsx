import ErrorBoundary from './components/ErrorBoundary';
import LoginGate from './components/LoginGate';
import Dashboard from './components/Dashboard';

export default function App() {
  return (
    <ErrorBoundary>
      <LoginGate>
        <Dashboard />
      </LoginGate>
    </ErrorBoundary>
  );
}
