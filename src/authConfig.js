const rawScopes = (import.meta.env.VITE_SCOPES || "")
  .split(",")
  .map((scope) => scope.trim())
  .filter(Boolean);

const defaultScopes = ["openid", "profile", "email"];

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID || "",
    authority: import.meta.env.VITE_AUTHORITY || "",
    redirectUri:
      import.meta.env.VITE_REDIRECT_URI || "http://localhost:3000/auth",
    postLogoutRedirectUri:
      import.meta.env.VITE_POST_LOGOUT_REDIRECT_URI || "http://localhost:3000",
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

export const loginRequest = {
  scopes: rawScopes.length > 0 ? rawScopes : defaultScopes
};

export const tokenValidationConfig = {
  expectedIssuer: import.meta.env.VITE_EXPECTED_ISSUER || "",
  expectedAudience: import.meta.env.VITE_EXPECTED_AUDIENCE || "",
  openIdConfigUrl: import.meta.env.VITE_OPENID_CONFIG_URL || ""
};

export const ssoConfig = {
  clientId: msalConfig.auth.clientId,
  authority: msalConfig.auth.authority,
  redirectUri: msalConfig.auth.redirectUri,
  postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
  scopes: loginRequest.scopes,
  tokenValidation: tokenValidationConfig
};
