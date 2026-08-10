import { NavLink, Outlet } from 'react-router-dom';
import './AppShell.css';

const NAV_ITEMS = [
  { to: '/', label: 'Overview', end: true },
  { to: '/sales-order', label: 'Sales Order', end: false },
  { to: '/stripe', label: 'Stripe', end: false },
  { to: '/gpay', label: 'GPay', end: false },
  { to: '/shopback', label: 'ShopBack', end: false },
];

export default function AppShell() {
  return (
    <div className="app-shell">
      <header className="app-shell__header">
        <div className="app-shell__brand">
          <span className="app-shell__eyebrow">J / SK</span>
          <h1 className="app-shell__title">Ecommerce Reconciliation</h1>
        </div>
        <nav className="app-shell__nav" aria-label="Main">
          {NAV_ITEMS.map((item) => (
            <NavLink
              key={item.to}
              to={item.to}
              end={item.end}
              className={({ isActive }) =>
                `app-shell__nav-link${isActive ? ' app-shell__nav-link--active' : ''}`
              }
            >
              {item.label}
            </NavLink>
          ))}
        </nav>
      </header>
      <main className="app-shell__main">
        <Outlet />
      </main>
    </div>
  );
}
