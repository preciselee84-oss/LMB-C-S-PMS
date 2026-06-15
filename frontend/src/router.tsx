import { createBrowserRouter, Navigate } from 'react-router-dom';

import { AppLayout } from './ui/AppLayout';
import { DashboardPage } from './pages/DashboardPage';
import { LoginPage } from './pages/LoginPage';
import { useAuthStore } from './stores/authStore';

function ProtectedRoute() {
  const token = useAuthStore.getState().token;
  return token ? <AppLayout /> : <Navigate to="/login" replace />;
}

export const router = createBrowserRouter([
  { path: '/login', element: <LoginPage /> },
  {
    path: '/',
    element: <ProtectedRoute />,
    children: [
      { index: true, element: <DashboardPage /> },
    ],
  },
]);
