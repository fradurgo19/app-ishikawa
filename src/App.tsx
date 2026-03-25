import React, { lazy, Suspense } from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { AuthProvider, useAuth } from './context/AuthContext';
import { Navbar } from './organisms/Navbar';

const queryClient = new QueryClient();
const LoginPage = lazy(() => import('./pages/LoginPage').then((module) => ({ default: module.LoginPage })));
const SelectorPage = lazy(() =>
  import('./pages/SelectorPage').then((module) => ({ default: module.SelectorPage }))
);
const NewRecordPage = lazy(() =>
  import('./pages/NewRecordPage').then((module) => ({ default: module.NewRecordPage }))
);
const FishbonePage = lazy(() =>
  import('./pages/FishbonePage').then((module) => ({ default: module.FishbonePage }))
);
const DataTablePage = lazy(() =>
  import('./pages/DataTablePage').then((module) => ({ default: module.DataTablePage }))
);

const FullScreenLoader: React.FC = () => (
  <div className="min-h-screen bg-gray-50 flex items-center justify-center">
    <div className="animate-spin rounded-full h-8 w-8 border-2 border-red-600 border-t-transparent"></div>
  </div>
);

const ProtectedRoute: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { isAuthenticated, loading } = useAuth();

  if (loading) {
    return <FullScreenLoader />;
  }

  if (!isAuthenticated) {
    return <Navigate to="/login" replace />;
  }

  return (
    <>
      <Navbar />
      {children}
    </>
  );
};

const RootRoute: React.FC = () => {
  const { isAuthenticated, loading } = useAuth();

  if (loading) {
    return <FullScreenLoader />;
  }

  if (isAuthenticated) {
    return <Navigate to="/selector" replace />;
  }

  return <Navigate to="/login" replace />;
};

const AppRoutes: React.FC = () => {
  return (
    <Suspense fallback={<FullScreenLoader />}>
      <Routes>
        <Route path="/login" element={<LoginPage />} />
        <Route
          path="/selector"
          element={
            <ProtectedRoute>
              <SelectorPage />
            </ProtectedRoute>
          }
        />
        <Route
          path="/new-record"
          element={
            <ProtectedRoute>
              <NewRecordPage />
            </ProtectedRoute>
          }
        />
        <Route
          path="/fishbone"
          element={
            <ProtectedRoute>
              <FishbonePage />
            </ProtectedRoute>
          }
        />
        <Route
          path="/data-table"
          element={
            <ProtectedRoute>
              <DataTablePage />
            </ProtectedRoute>
          }
        />
        <Route path="/" element={<RootRoute />} />
      </Routes>
    </Suspense>
  );
};

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <AuthProvider>
        <Router>
          <div className="min-h-screen w-full bg-gray-50">
            <AppRoutes />
          </div>
        </Router>
      </AuthProvider>
    </QueryClientProvider>
  );
}

export default App;