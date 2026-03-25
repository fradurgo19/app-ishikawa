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
        <div className="flex h-16 items-center gap-4">
          <div className="flex min-w-0 flex-1 items-center gap-6 lg:gap-8">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-red-100 rounded-lg">
                <Settings className="h-6 w-6 text-red-600" />
              </div>
              <span className="truncate text-xl font-bold text-gray-900">
                Plataforma de Maquinaria Pesada
              </span>
            </div>

            <div className="hidden min-w-0 md:flex md:items-center md:space-x-4">
              {navItems
                .filter(item => !item.coordinatorOnly || user?.role === 'coordinador')
                .map((item) => (
                  <button
                    key={item.path}
                    onClick={() => navigate(item.path)}
                    className={`flex shrink-0 items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium transition-colors duration-200 ${
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

          <div className="flex shrink-0 items-center gap-3 border-l border-gray-200 pl-4 md:gap-5 md:pl-6 lg:pl-8">
            <div className="flex max-w-[min(22rem,40vw)] flex-col items-end gap-1.5 text-right sm:max-w-none sm:flex-row sm:items-center sm:gap-3">
              <p className="text-sm leading-snug text-gray-600">
                <span className="text-gray-500">Bienvenido,</span>{' '}
                <span className="font-medium text-gray-900">{user?.name}</span>
              </p>
              <span className="inline-flex shrink-0 rounded-full bg-gray-100 px-2.5 py-1 text-xs capitalize text-gray-700">
                {user?.role}
              </span>
            </div>

            <Button variant="ghost" icon={LogOut} onClick={handleLogout}>
              Cerrar Sesión
            </Button>
          </div>
        </div>
      </div>
    </nav>
  );
};