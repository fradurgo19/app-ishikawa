import React from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { AuthProvider, useAuth } from './context/AuthContext';
import { Navbar } from './organisms/Navbar';
import { LoginPage } from './pages/LoginPage';
import { SelectorPage } from './pages/SelectorPage';
import { NewRecordPage } from './pages/NewRecordPage';
import { FishbonePage } from './pages/FishbonePage';
import { DataTablePage } from './pages/DataTablePage';

const queryClient = new QueryClient();

const ProtectedRoute: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { isAuthenticated, loading } = useAuth();

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="animate-spin rounded-full h-8 w-8 border-2 border-red-600 border-t-transparent"></div>
      </div>
    );
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

const AppRoutes: React.FC = () => {
  return (
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
      <Route path="/" element={<Navigate to="/selector" replace />} />
    </Routes>
  );
};

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <AuthProvider>
        <Router>
          <div className="min-h-screen bg-gray-50">
            <AppRoutes />
          </div>
        </Router>
      </AuthProvider>
    </QueryClientProvider>
  );
}

export default App;