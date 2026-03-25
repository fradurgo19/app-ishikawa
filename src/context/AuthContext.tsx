import React, {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useReducer,
  type ReactNode,
} from 'react';
import {
  BrowserCacheLocation,
  PublicClientApplication,
  type AccountInfo,
} from '@azure/msal-browser';
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

const MSAL_CLIENT_ID = normalizeEnvValue(import.meta.env.VITE_MSAL_CLIENT_ID);
const MSAL_TENANT_ID = normalizeEnvValue(import.meta.env.VITE_MSAL_TENANT_ID);
const MSAL_REDIRECT_URI = normalizeEnvValue(import.meta.env.VITE_MSAL_REDIRECT_URI);
const COORDINATOR_EMAILS = parseCoordinatorEmails(import.meta.env.VITE_COORDINATOR_EMAILS);
const MSAL_SCOPES = ['User.Read'];
const MICROSOFT_AUTH_MISCONFIGURED_MESSAGE =
  'Microsoft authentication is not configured. Set VITE_MSAL_CLIENT_ID and VITE_MSAL_TENANT_ID.';

const isMicrosoftAuthEnabled = Boolean(MSAL_CLIENT_ID && MSAL_TENANT_ID);

export const AuthProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const msalClient = useMemo(() => {
    if (!isMicrosoftAuthEnabled) {
      return null;
    }

    return new PublicClientApplication({
      auth: {
        clientId: MSAL_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
        redirectUri: MSAL_REDIRECT_URI || getDefaultRedirectUri(),
      },
      cache: {
        cacheLocation: BrowserCacheLocation.SessionStorage,
      },
    });
  }, []);

  const [state, dispatch] = useReducer(authReducer, {
    user: null,
    isAuthenticated: false,
    loading: true,
  });

  const syncSession = useCallback(async () => {
    if (!msalClient) {
      dispatch({ type: 'AUTH_LOGOUT' });
      return;
    }

    dispatch({ type: 'AUTH_LOADING' });

    try {
      await msalClient.initialize();
      await msalClient.handleRedirectPromise();

      const account = resolveActiveAccount(msalClient);
      if (!account) {
        dispatch({ type: 'AUTH_LOGOUT' });
        return;
      }

      dispatch({ type: 'AUTHENTICATED', payload: mapAccountToUser(account) });
    } catch (error) {
      console.error('Error sincronizando sesión de Microsoft:', error);
      dispatch({ type: 'AUTH_LOGOUT' });
    }
  }, [msalClient]);

  useEffect(() => {
    void syncSession();
  }, [syncSession]);

  const login = useCallback(async () => {
    if (!msalClient) {
      throw new Error(MICROSOFT_AUTH_MISCONFIGURED_MESSAGE);
    }

    dispatch({ type: 'AUTH_LOADING' });

    try {
      await msalClient.initialize();

      const loginResult = await msalClient.loginPopup({
        scopes: MSAL_SCOPES,
        prompt: 'select_account',
      });

      const account = loginResult.account ?? resolveActiveAccount(msalClient);
      if (!account) {
        throw new Error('Microsoft account information was not returned by Entra ID.');
      }

      msalClient.setActiveAccount(account);
      dispatch({ type: 'AUTHENTICATED', payload: mapAccountToUser(account) });
    } catch (error) {
      dispatch({ type: 'AUTH_LOGOUT' });
      throw error;
    }
  }, [msalClient]);

  const logout = useCallback(() => {
    dispatch({ type: 'AUTH_LOGOUT' });

    if (!msalClient) {
      return;
    }

    void msalClient
      .logoutPopup({
        mainWindowRedirectUri: getDefaultRedirectUri(),
      })
      .catch((error) => {
        console.error('Error cerrando sesión de Microsoft:', error);
      });
  }, [msalClient]);

  return (
    <AuthContext.Provider value={{ ...state, login, logout, isMicrosoftAuthEnabled }}>
      {children}
    </AuthContext.Provider>
  );
};

function resolveActiveAccount(msalClient: PublicClientApplication): AccountInfo | null {
  const activeAccount = msalClient.getActiveAccount();
  if (activeAccount) {
    return activeAccount;
  }

  const [firstAccount] = msalClient.getAllAccounts();
  if (firstAccount) {
    msalClient.setActiveAccount(firstAccount);
    return firstAccount;
  }

  return null;
}

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
  if (!COORDINATOR_EMAILS.length) {
    return 'coordinador';
  }

  return COORDINATOR_EMAILS.includes(email.toLowerCase()) ? 'coordinador' : 'basico';
}

function parseCoordinatorEmails(rawValue: string | undefined): string[] {
  const normalizedValue = normalizeEnvValue(rawValue);
  if (!normalizedValue) {
    return [];
  }

  return normalizedValue
    .split(',')
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean);
}

function normalizeEnvValue(value: unknown): string {
  if (typeof value !== 'string') {
    return '';
  }

  return value.trim();
}

function getDefaultRedirectUri(): string {
  if (typeof window === 'undefined') {
    return '/';
  }

  return window.location.origin;
}

// eslint-disable-next-line react-refresh/only-export-components
export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth debe ser usado dentro de un AuthProvider');
  }
  return context;
};