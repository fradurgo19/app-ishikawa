import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { Button } from '../atoms/Button';
import { Input } from '../atoms/Input';
import { Card } from '../atoms/Card';
import { Settings } from 'lucide-react';

const INVALID_CREDENTIALS_MESSAGE =
  'Credenciales inválidas. Prueba "admin" o "tecnico" con contraseña "password123"';

export const LoginPage: React.FC = () => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const { login, loading } = useAuth();
  const navigate = useNavigate();

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError('');
    
    try {
      await login(username.trim(), password);
      navigate('/selector');
    } catch {
      setError(INVALID_CREDENTIALS_MESSAGE);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center px-4">
      <div className="max-w-md w-full">
        <Card className="p-8">
          <div className="text-center mb-8">
            <div className="inline-flex p-3 bg-red-100 rounded-full mb-4">
              <Settings className="h-8 w-8 text-red-600" />
            </div>
            <h1 className="text-2xl font-bold text-gray-900">Plataforma de Maquinaria Pesada</h1>
            <p className="text-gray-600 mt-2">Inicia sesión para acceder al Sistema de Diagrama Ishikawa</p>
          </div>

          <form onSubmit={handleSubmit} className="space-y-4">
            <Input
              label="Usuario"
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Ingresa tu usuario"
              required
            />
            
            <Input
              label="Contraseña"
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Ingresa tu contraseña"
              required
            />

            {error && (
              <div className="p-3 bg-red-50 border border-red-200 rounded-lg">
                <p className="text-sm text-red-600">{error}</p>
              </div>
            )}

            <Button
              type="submit"
              loading={loading}
              fullWidth
            >
              Iniciar Sesión
            </Button>
          </form>

          <div className="mt-6 p-4 bg-gray-50 rounded-lg">
            <h3 className="text-sm font-medium text-gray-700 mb-2">Credenciales de Demostración:</h3>
            <div className="text-xs text-gray-600 space-y-1">
              <p><strong>Administrador:</strong> admin / password123</p>
              <p><strong>Usuario Técnico:</strong> tecnico / password123</p>
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
};