import { useEffect, useMemo, useState } from "react";
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
  useMsal
} from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { createRemoteJWKSet, jwtVerify } from "jose";
import { loginRequest, ssoConfig, tokenValidationConfig } from "./authConfig.js";

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

const decodeJwt = (token) => {
  if (!token || typeof token !== "string") return null;
  const parts = token.split(".");
  if (parts.length < 2) return null;
  try {
    const payload = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = payload.padEnd(Math.ceil(payload.length / 4) * 4, "=");
    return JSON.parse(atob(padded));
  } catch (error) {
    return null;
  }
};

const getExpectedAudience = (scopes) => {
  if (!Array.isArray(scopes)) return "";
  const apiScope = scopes.find((scope) => scope.startsWith("api://"));
  if (!apiScope) return "";
  const parts = apiScope.split("/");
  return parts.length >= 3 ? parts.slice(0, 3).join("/") : apiScope;
};

const normalizeAud = (aud) => {
  if (!aud) return [];
  return Array.isArray(aud) ? aud : [aud];
};

const validateAccessToken = (claims, validationConfig, scopes) => {
  if (!claims) {
    return {
      ok: false,
      checks: [{ name: "token", ok: false, detail: "No access token" }]
    };
  }

  const nowSeconds = Math.floor(Date.now() / 1000);
  const expectedAudience = validationConfig.expectedAudience || "";
  const audienceList = normalizeAud(claims.aud);
  const expectedIssuer = validationConfig.expectedIssuer || "";

  const checks = [
    {
      name: "exp",
      ok: typeof claims.exp === "number" && nowSeconds < claims.exp,
      detail: `exp=${claims.exp ?? "missing"} now=${nowSeconds}`
    },
    {
      name: "nbf",
      ok: typeof claims.nbf !== "number" || nowSeconds >= claims.nbf,
      detail: `nbf=${claims.nbf ?? "missing"} now=${nowSeconds}`
    },
    {
      name: "iss",
      ok: expectedIssuer ? String(claims.iss || "") === expectedIssuer : true,
      detail: `iss=${claims.iss ?? "missing"}`
    },
    {
      name: "aud",
      ok: expectedAudience ? audienceList.includes(expectedAudience) : true,
      detail: `aud=${audienceList.join(", ") || "missing"}`
    }
  ];

  return {
    ok: checks.every((check) => check.ok),
    expected: {
      expectedIssuer: expectedIssuer || "(no check)",
      expectedAudience: expectedAudience || "(no check)"
    },
    checks
  };
};

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const account = accounts[0] || null;
  const [tokenResult, setTokenResult] = useState(null);
  const [tokenError, setTokenError] = useState("");
  const [jwksStatus, setJwksStatus] = useState(null);
  const accessTokenClaims = useMemo(
    () => decodeJwt(tokenResult?.accessToken),
    [tokenResult?.accessToken]
  );

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

  useEffect(() => {
    const run = async () => {
      if (!tokenResult?.accessToken) {
        setJwksStatus({ ok: false, detail: "No access token" });
        return;
      }
      if (!tokenValidationConfig.openIdConfigUrl) {
        setJwksStatus({
          ok: false,
          detail: "VITE_OPENID_CONFIG_URL is empty"
        });
        return;
      }
      try {
        const configResponse = await fetch(tokenValidationConfig.openIdConfigUrl);
        if (!configResponse.ok) {
          throw new Error(
            `OpenID config fetch failed: ${configResponse.status}`
          );
        }
        const openIdConfig = await configResponse.json();
        if (!openIdConfig?.jwks_uri) {
          throw new Error("jwks_uri is missing in OpenID config");
        }
        const jwks = createRemoteJWKSet(new URL(openIdConfig.jwks_uri));
        const expectedAudience = tokenValidationConfig.expectedAudience || "";
        const expectedIssuer = tokenValidationConfig.expectedIssuer || "";
        const options = {
          audience: expectedAudience || undefined,
          issuer: expectedIssuer || undefined
        };
        const { protectedHeader } = await jwtVerify(
          tokenResult.accessToken,
          jwks,
          options
        );
        setJwksStatus({
          ok: true,
          detail: "Signature verified",
          header: protectedHeader,
          jwksUri: openIdConfig.jwks_uri
        });
      } catch (error) {
        setJwksStatus({ ok: false, detail: error?.message || String(error) });
      }
    };
    void run();
  }, [tokenResult?.accessToken]);

  const accessTokenValidation = useMemo(() => {
    return validateAccessToken(
      accessTokenClaims,
      tokenValidationConfig,
      loginRequest.scopes
    );
  }, [accessTokenClaims]);

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

        <section className="card">
          <h2>Access Token Validation (client-only)</h2>
          <p className="subtle">
            本来はバックエンドで実施。ここでは簡易チェックと JWKS 検証のデモのみ。
          </p>
          <pre className="code">
            {JSON.stringify(accessTokenValidation, null, 2)}
          </pre>
        </section>

        <section className="card">
          <h2>JWKS Signature Check (client demo)</h2>
          <p className="subtle">
            本番ではバックエンドで署名検証すること。
          </p>
          <pre className="code">{JSON.stringify(jwksStatus, null, 2)}</pre>
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
