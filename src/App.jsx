import { useEffect, useMemo, useState } from "react";
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
  useMsal
} from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { loginRequest, ssoConfig } from "./authConfig.js";

const formatDate = (value) => {
  if (!value) return "";
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toISOString();
};

const serializeTokenResult = (result) => {
  if (!result) return "";
  const base = {
    uniqueId: result.uniqueId,
    tenantId: result.tenantId,
    scopes: result.scopes,
    expiresOn: formatDate(result.expiresOn),
    extExpiresOn: formatDate(result.extExpiresOn),
    fromCache: result.fromCache,
    idToken: result.idToken,
    accessToken: result.accessToken,
    account: result.account,
    idTokenClaims: result.idTokenClaims
  };
  return JSON.stringify(base, null, 2);
};

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const account = accounts[0] || null;
  const [tokenResult, setTokenResult] = useState(null);
  const [tokenError, setTokenError] = useState("");

  const missingConfig = useMemo(() => {
    return !ssoConfig.clientId || !ssoConfig.authority;
  }, []);

  const login = () => {
    if (missingConfig) return;
    setTokenError("");
    instance.loginRedirect(loginRequest);
  };

  const logout = () => {
    instance.logoutRedirect({
      postLogoutRedirectUri: ssoConfig.postLogoutRedirectUri
    });
  };

  const acquireToken = async () => {
    if (!account || missingConfig) return;
    setTokenError("");
    try {
      const result = await instance.acquireTokenSilent({
        ...loginRequest,
        account
      });
      setTokenResult(result);
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        instance.acquireTokenRedirect({
          ...loginRequest,
          account
        });
        return;
      }
      setTokenError(error?.message || String(error));
    }
  };

  useEffect(() => {
    if (!account) {
      setTokenResult(null);
      return;
    }
    void acquireToken();
  }, [account?.homeAccountId]);

  return (
    <div className="page">
      <header className="hero">
        <div>
          <p className="eyebrow">MSAL External ID SSO Sample</p>
          <h1>Frontend PKCE (Auth Code) Sample</h1>
          <p className="subtle">
            Callback: <code>http://localhost:3000/auth</code>
          </p>
        </div>
        <div className="actions">
          <button
            type="button"
            onClick={login}
            disabled={missingConfig || inProgress !== "none"}
          >
            Login (Redirect)
          </button>
          <button
            type="button"
            className="ghost"
            onClick={logout}
            disabled={!account}
          >
            Logout
          </button>
        </div>
      </header>

      {missingConfig && (
        <section className="card warning">
          <h2>Config is empty</h2>
          <p>
            .env の <code>VITE_CLIENT_ID</code> と <code>VITE_AUTHORITY</code> が空のため、
            ログインできません。値を設定して再起動してください。
          </p>
        </section>
      )}

      <section className="grid">
        <div className="card">
          <h2>SSO Config</h2>
          <pre className="code">
            {JSON.stringify(ssoConfig, null, 2)}
          </pre>
        </div>
        <div className="card">
          <h2>Runtime</h2>
          <pre className="code">
            {JSON.stringify(
              {
                path: window.location.pathname,
                inProgress,
                accountCount: accounts.length
              },
              null,
              2
            )}
          </pre>
        </div>
      </section>

      <AuthenticatedTemplate>
        <section className="card">
          <div className="row">
            <h2>Account</h2>
            <button
              type="button"
              className="ghost"
              onClick={acquireToken}
              disabled={inProgress !== "none"}
            >
              Acquire Token
            </button>
          </div>
          <pre className="code">
            {JSON.stringify(account, null, 2)}
          </pre>
        </section>

        <section className="card">
          <h2>Token Result (all)</h2>
          {tokenError && (
            <p className="error">{tokenError}</p>
          )}
          <pre className="code">{serializeTokenResult(tokenResult)}</pre>
        </section>
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <section className="card">
          <h2>Not authenticated</h2>
          <p>
            ログイン後にトークン情報が表示されます。
          </p>
        </section>
      </UnauthenticatedTemplate>
    </div>
  );
}
