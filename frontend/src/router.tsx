import { createBrowserRouter, Navigate } from 'react-router-dom';

import { AppLayout } from './ui/AppLayout';
import { DashboardPage } from './pages/DashboardPage';

export const router = createBrowserRouter([
  {
    path: '/',
    element: <AppLayout />,
    children: [
      { index: true, element: <DashboardPage /> },
    ],
  },
  { path: '*', element: <Navigate to="/" replace /> },
]);
