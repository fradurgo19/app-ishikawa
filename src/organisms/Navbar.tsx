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
        <div className="grid w-full grid-cols-1 items-center gap-3 py-3 md:grid-cols-[minmax(0,1fr)_auto] md:gap-6 md:py-0 lg:h-16 lg:min-h-0">
          <div className="flex min-h-0 min-w-0 flex-col gap-3 md:flex-row md:items-center md:gap-4 lg:gap-8">
            <div className="flex min-w-0 shrink-0 items-center gap-3">
              <div className="shrink-0 p-2 bg-red-100 rounded-lg">
                <Settings className="h-6 w-6 text-red-600" />
              </div>
              <span className="min-w-0 truncate text-xl font-bold text-gray-900">
                Plataforma de Maquinaria Pesada
              </span>
            </div>

            <div className="hidden min-w-0 md:flex md:min-h-0 md:flex-1 md:items-center md:overflow-x-auto md:overflow-y-visible md:pb-0 md:[scrollbar-width:thin]">
              <div className="flex w-max min-w-0 items-center gap-2 pr-1 md:gap-3">
                {navItems
                  .filter(item => !item.coordinatorOnly || user?.role === 'coordinador')
                  .map((item) => (
                    <button
                      key={item.path}
                      type="button"
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
          </div>

          <div className="flex min-w-0 flex-col items-stretch gap-3 border-t border-gray-200 pt-3 sm:flex-row sm:items-center sm:justify-end sm:gap-4 md:border-l md:border-t-0 md:pl-6 md:pt-0 lg:pl-8">
            <div className="flex min-w-0 flex-col items-end gap-1.5 text-right sm:flex-row sm:items-center sm:gap-3 sm:pl-2">
              <p className="max-w-full text-sm leading-snug text-gray-600">
                <span className="text-gray-500">Bienvenido,</span>{' '}
                <span
                  className="inline-block max-w-[min(100%,12rem)] truncate align-bottom font-medium text-gray-900 sm:max-w-[16rem] md:max-w-[20rem] lg:max-w-[24rem]"
                  title={user?.name}
                >
                  {user?.name}
                </span>
              </p>
              <span className="inline-flex shrink-0 rounded-full bg-gray-100 px-2.5 py-1 text-xs capitalize text-gray-700">
                {user?.role}
              </span>
            </div>

            <Button variant="ghost" icon={LogOut} onClick={handleLogout} className="shrink-0">
              Cerrar Sesión
            </Button>
          </div>
        </div>
      </div>
    </nav>
  );
};