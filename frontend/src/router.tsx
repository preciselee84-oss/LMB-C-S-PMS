import { createBrowserRouter, Navigate } from 'react-router-dom';

import { AppLayout } from './ui/AppLayout';
import { DashboardPage } from './pages/DashboardPage';
import { VisitVocPage } from './pages/VisitVocPage';

export const router = createBrowserRouter([
  {
    path: '/',
    element: <AppLayout />,
    children: [
      { index: true, element: <DashboardPage /> },
      { path: 'crm/visit-voc', element: <VisitVocPage /> },
    ],
  },
  { path: '*', element: <Navigate to="/" replace /> },
]);
