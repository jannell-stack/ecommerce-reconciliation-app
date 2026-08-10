import { BrowserRouter, Navigate, Route, Routes } from 'react-router-dom';
import { DashboardProvider } from '../context/DashboardContext';
import AppShell from './layout/AppShell';
import OverviewPage from '../pages/OverviewPage';
import SalesOrderPage from '../pages/SalesOrderPage';
import StripePage from '../pages/StripePage';
import GPayPage from '../pages/GPayPage';
import ShopBackPage from '../pages/ShopBackPage';
import './Dashboard.css';

export default function Dashboard() {
  return (
    <DashboardProvider>
      <BrowserRouter>
        <Routes>
          <Route element={<AppShell />}>
            <Route index element={<OverviewPage />} />
            <Route path="sales-order" element={<SalesOrderPage />} />
            <Route path="stripe" element={<StripePage />} />
            <Route path="gpay" element={<GPayPage />} />
            <Route path="shopback" element={<ShopBackPage />} />
            <Route path="*" element={<Navigate to="/" replace />} />
          </Route>
        </Routes>
      </BrowserRouter>
    </DashboardProvider>
  );
}
