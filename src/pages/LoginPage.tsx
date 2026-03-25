import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { Button } from '../atoms/Button';
import { Card } from '../atoms/Card';
import { LogIn, Settings } from 'lucide-react';

const MICROSOFT_LOGIN_ERROR_MESSAGE =
  'No fue posible iniciar sesión con Microsoft. Intenta nuevamente.';

export const LoginPage: React.FC = () => {
  const [error, setError] = useState('');
  const { login, loading, isAuthenticated, isMicrosoftAuthEnabled } = useAuth();
  const navigate = useNavigate();

  useEffect(() => {
    if (isAuthenticated) {
      navigate('/selector');
    }
  }, [isAuthenticated, navigate]);

  const handleMicrosoftLogin = async () => {
    setError('');

    try {
      await login();
      navigate('/selector');
    } catch (unknownError) {
      setError(getErrorMessage(unknownError));
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
            <p className="text-gray-600 mt-2">
              Inicia sesión con tu cuenta corporativa Microsoft para acceder al Sistema de Diagrama
              Ishikawa
            </p>
          </div>

          <div className="space-y-4">
            <Button
              type="button"
              loading={loading}
              fullWidth
              icon={LogIn}
              onClick={handleMicrosoftLogin}
              disabled={!isMicrosoftAuthEnabled}
            >
              Iniciar sesión con Microsoft
            </Button>

            {error && (
              <div className="p-3 bg-red-50 border border-red-200 rounded-lg">
                <p className="text-sm text-red-600">{error}</p>
              </div>
            )}

            {!isMicrosoftAuthEnabled && (
              <div className="p-3 bg-yellow-50 border border-yellow-200 rounded-lg">
                <p className="text-sm text-yellow-700">
                  Falta configurar Microsoft Entra ID. Define `VITE_MSAL_CLIENT_ID` y
                  `VITE_MSAL_TENANT_ID` para habilitar el inicio de sesión.
                </p>
              </div>
            )}
          </div>
        </Card>
      </div>
    </div>
  );
};

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim()) {
    return error.message;
  }

  return MICROSOFT_LOGIN_ERROR_MESSAGE;
}