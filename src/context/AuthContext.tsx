import React, { createContext, useCallback, useContext, useEffect, useReducer, type ReactNode } from 'react';
import { type AccountInfo } from '@azure/msal-browser';
import { coordinatorEmails } from '../config/authConfig';
import { authService } from '../services/authService';
import { User } from '../types';

interface AuthState {
  user: User | null;
  isAuthenticated: boolean;
  loading: boolean;
}

interface AuthContextType extends AuthState {
  login: () => Promise<void>;
  logout: () => void;
  isMicrosoftAuthEnabled: boolean;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

type AuthAction =
  | { type: 'AUTH_LOADING' }
  | { type: 'AUTHENTICATED'; payload: User }
  | { type: 'AUTH_LOGOUT' };

const authReducer = (state: AuthState, action: AuthAction): AuthState => {
  switch (action.type) {
    case 'AUTH_LOADING':
      return { ...state, loading: true };
    case 'AUTHENTICATED':
      return { ...state, loading: false, isAuthenticated: true, user: action.payload };
    case 'AUTH_LOGOUT':
      return { ...state, loading: false, isAuthenticated: false, user: null };
    default:
      return state;
  }
};

const MICROSOFT_AUTH_MISCONFIGURED_MESSAGE =
  'Microsoft authentication is not configured. Set VITE_MSAL_CLIENT_ID and VITE_MSAL_TENANT_ID.';

export const AuthProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [state, dispatch] = useReducer(authReducer, {
    user: null,
    isAuthenticated: false,
    loading: true,
  });

  const syncSession = useCallback(async () => {
    if (!authService.isMicrosoftAuthEnabled) {
      dispatch({ type: 'AUTH_LOGOUT' });
      return;
    }

    dispatch({ type: 'AUTH_LOADING' });

    try {
      await authService.initializeAuth();
      const account = await authService.getAccountWithRetry();
      if (!account) {
        dispatch({ type: 'AUTH_LOGOUT' });
        return;
      }

      dispatch({ type: 'AUTHENTICATED', payload: mapAccountToUser(account) });
    } catch (error) {
      console.error('Error sincronizando sesión de Microsoft:', error);
      dispatch({ type: 'AUTH_LOGOUT' });
    }
  }, []);

  useEffect(() => {
    void syncSession();
  }, [syncSession]);

  const login = useCallback(async () => {
    if (!authService.isMicrosoftAuthEnabled) {
      throw new Error(MICROSOFT_AUTH_MISCONFIGURED_MESSAGE);
    }

    dispatch({ type: 'AUTH_LOADING' });

    try {
      const loginResult = await authService.login();
      const account = loginResult.account ?? (await authService.getAccountWithRetry());
      if (!account) {
        throw new Error('Microsoft account information was not returned by Entra ID.');
      }

      dispatch({ type: 'AUTHENTICATED', payload: mapAccountToUser(account) });
    } catch (error) {
      dispatch({ type: 'AUTH_LOGOUT' });
      throw error;
    }
  }, []);

  const logout = useCallback(() => {
    dispatch({ type: 'AUTH_LOGOUT' });

    if (!authService.isMicrosoftAuthEnabled) {
      return;
    }

    void authService
      .logout()
      .catch((error) => {
        console.error('Error cerrando sesión de Microsoft:', error);
      });
  }, []);

  return (
    <AuthContext.Provider
      value={{ ...state, login, logout, isMicrosoftAuthEnabled: authService.isMicrosoftAuthEnabled }}
    >
      {children}
    </AuthContext.Provider>
  );
};

function mapAccountToUser(account: AccountInfo): User {
  const normalizedEmail = normalizeAccountEmail(account);
  const normalizedUsername = normalizedEmail.includes('@')
    ? normalizedEmail.split('@')[0]
    : normalizedEmail;

  return {
    id: account.localAccountId || account.homeAccountId || normalizedEmail,
    username: normalizedUsername || 'usuario',
    role: resolveUserRole(normalizedEmail),
    name: normalizeAccountDisplayName(account, normalizedUsername),
  };
}

function normalizeAccountEmail(account: AccountInfo): string {
  const accountUsername = normalizeEnvValue(account.username);
  if (accountUsername) {
    return accountUsername.toLowerCase();
  }

  const idTokenClaims = account.idTokenClaims as Record<string, unknown> | undefined;
  const preferredUsername = normalizeEnvValue(idTokenClaims?.preferred_username);
  if (preferredUsername) {
    return preferredUsername.toLowerCase();
  }

  const uniqueName = normalizeEnvValue(idTokenClaims?.unique_name);
  if (uniqueName) {
    return uniqueName.toLowerCase();
  }

  return account.homeAccountId.toLowerCase();
}

function normalizeAccountDisplayName(account: AccountInfo, fallbackUsername: string): string {
  const accountName = normalizeEnvValue(account.name);
  if (accountName) {
    return accountName;
  }

  return fallbackUsername || 'Usuario';
}

function resolveUserRole(email: string): User['role'] {
  if (!coordinatorEmails.length) {
    return 'coordinador';
  }

  return coordinatorEmails.includes(email.toLowerCase()) ? 'coordinador' : 'basico';
}

function normalizeEnvValue(value: unknown): string {
  if (typeof value !== 'string') {
    return '';
  }

  return value.trim();
}

// eslint-disable-next-line react-refresh/only-export-components
export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth debe ser usado dentro de un AuthProvider');
  }
  return context;
};