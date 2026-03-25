import React, { useCallback, useEffect, useRef, useState } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { Button } from '../atoms/Button';
import { Settings, LogOut, Home, Plus, BarChart3, GitBranch, ChevronDown } from 'lucide-react';

export const Navbar: React.FC = () => {
  const { user, logout } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();
  const [menuOpen, setMenuOpen] = useState(false);
  const menuContainerRef = useRef<HTMLDivElement>(null);

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

  const visibleNavItems = navItems.filter(
    (item) => !item.coordinatorOnly || user?.role === 'coordinador'
  );

  const isActive = (path: string) => location.pathname === path;

  const closeMenu = useCallback(() => {
    setMenuOpen(false);
  }, []);

  useEffect(() => {
    closeMenu();
  }, [location.pathname, closeMenu]);

  useEffect(() => {
    if (!menuOpen) {
      return;
    }

    const handlePointerDown = (event: MouseEvent) => {
      const node = menuContainerRef.current;
      if (node && !node.contains(event.target as Node)) {
        setMenuOpen(false);
      }
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setMenuOpen(false);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [menuOpen]);

  const handleNavigate = (path: string) => {
    navigate(path);
    setMenuOpen(false);
  };

  const currentLabel =
    visibleNavItems.find((item) => isActive(item.path))?.label ?? 'Menú';

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

            <div ref={menuContainerRef} className="relative min-w-0 md:max-w-xs md:flex-1">
              <button
                type="button"
                id="navbar-pages-menu-button"
                aria-haspopup="true"
                aria-expanded={menuOpen}
                aria-controls="navbar-pages-menu"
                onClick={() => setMenuOpen((open) => !open)}
                className="flex w-full items-center justify-between gap-2 rounded-lg border border-gray-200 bg-white px-4 py-2.5 text-left text-sm font-medium text-gray-800 shadow-sm transition hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-offset-2 md:w-auto md:min-w-[14rem]"
              >
                <span className="truncate">
                  <span className="text-gray-500">Ir a:</span>{' '}
                  <span className="font-semibold text-gray-900">{currentLabel}</span>
                </span>
                <ChevronDown
                  className={`h-5 w-5 shrink-0 text-gray-500 transition-transform ${menuOpen ? 'rotate-180' : ''}`}
                  aria-hidden
                />
              </button>

              {menuOpen ? (
                <div
                  id="navbar-pages-menu"
                  role="menu"
                  aria-labelledby="navbar-pages-menu-button"
                  className="absolute left-0 right-0 top-full z-50 mt-1 max-h-[min(70vh,24rem)] overflow-y-auto rounded-lg border border-gray-200 bg-white py-1 shadow-lg md:right-auto md:min-w-[14rem]"
                >
                  {visibleNavItems.map((item) => {
                    const active = isActive(item.path);
                    return (
                      <button
                        key={item.path}
                        type="button"
                        role="menuitem"
                        onClick={() => handleNavigate(item.path)}
                        className={`flex w-full items-center gap-3 px-4 py-3 text-left text-sm transition ${
                          active
                            ? 'bg-red-50 font-semibold text-red-700'
                            : 'text-gray-700 hover:bg-gray-50'
                        }`}
                      >
                        <item.icon size={18} className="shrink-0 opacity-80" aria-hidden />
                        {item.label}
                      </button>
                    );
                  })}
                </div>
              ) : null}
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
