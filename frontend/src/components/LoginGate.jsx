import { useState } from 'react';
import { isAuthenticated } from '../auth';
import { loginWithPassword } from '../api';
import './LoginGate.css';

export default function LoginGate({ children }) {
  const [authenticated, setAuthenticated] = useState(isAuthenticated);
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);

  if (authenticated) {
    return children;
  }

  async function handleSubmit(event) {
    event.preventDefault();
    setSubmitting(true);
    setError('');
    try {
      await loginWithPassword(password);
      setAuthenticated(true);
    } catch (err) {
      setError(err.message || 'Invalid password');
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="login-gate">
      <form className="login-gate__card" onSubmit={handleSubmit}>
        <h1 className="login-gate__title">Ecommerce Reconciliation</h1>
        <p className="login-gate__subtitle">
          Enter the password to continue. Signing in clears any previously imported data.
        </p>
        <label className="login-gate__label" htmlFor="app-password">
          Password
        </label>
        <input
          id="app-password"
          className="login-gate__input"
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          autoComplete="current-password"
          autoFocus
          required
        />
        {error ? <p className="login-gate__error">{error}</p> : null}
        <button className="login-gate__button" type="submit" disabled={submitting}>
          {submitting ? 'Signing in…' : 'Sign in'}
        </button>
      </form>
    </div>
  );
}
