import React from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { Button } from '../atoms/Button';
import { Settings, LogOut, Home, Plus, BarChart3, GitBranch } from 'lucide-react';

export const Navbar: React.FC = () => {
  const { user, logout } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();

  const handleLogout = () => {
    logout();
    navigate('/login');
  };

  const navItems = [
    { path: '/selector', label: 'Inicio', icon: Home },
    { path: '/new-record', label: 'Nuevo Registro', icon: Plus, coordinatorOnly: true },
    { path: '/data-table', label: 'Tabla de Datos', icon: BarChart3 },
    { path: '/fishbone', label: 'Ishikawa', icon: GitBranch },
  ];

  const isActive = (path: string) => location.pathname === path;

  return (
    <nav className="w-full bg-white shadow-md border-b border-gray-200">
      <div className="w-full px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-16">
          <div className="flex items-center gap-8">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-red-100 rounded-lg">
                <Settings className="h-6 w-6 text-red-600" />
              </div>
              <span className="text-xl font-bold text-gray-900">Plataforma de Maquinaria Pesada</span>
            </div>

            <div className="hidden md:flex items-center space-x-4">
              {navItems
                .filter(item => !item.coordinatorOnly || user?.role === 'coordinador')
                .map((item) => (
                  <button
                    key={item.path}
                    onClick={() => navigate(item.path)}
                    className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium transition-colors duration-200 ${
                      isActive(item.path)
                        ? 'bg-red-100 text-red-700'
                        : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
                    }`}
                  >
                    <item.icon size={16} />
                    {item.label}
                  </button>
                ))}
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-sm text-gray-600">
              Bienvenido, <span className="font-medium">{user?.name}</span>
              <span className="ml-2 px-2 py-1 bg-gray-100 text-xs rounded-full capitalize">
                {user?.role}
              </span>
            </div>
            
            <Button
              variant="ghost"
              icon={LogOut}
              onClick={handleLogout}
            >
              Cerrar Sesión
            </Button>
          </div>
        </div>
      </div>
    </nav>
  );
};