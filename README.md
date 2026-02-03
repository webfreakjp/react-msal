# react-msal

MSAL (External ID / Entra ID External ID) を使ったフロントエンド SSO サンプルです。Auth Code + PKCE を使う SPA 方式で、
`http://localhost:3000/auth` をコールバックにしています。

## 使い方

```bash
npm install
npm run dev
```

## 設定

SSO 設定は `.env`

```bash
VITE_CLIENT_ID=
VITE_AUTHORITY=
VITE_SCOPES=
VITE_REDIRECT_URI=
VITE_POST_LOGOUT_REDIRECT_URI=
```

- `VITE_CLIENT_ID`: アプリの Client ID
- `VITE_AUTHORITY`: External ID の authority
- `VITE_SCOPES`: 例: `openid,profile,email,api://<app-id>/access_as_user`
- `VITE_REDIRECT_URI`: 未指定時は `http://localhost:3000/auth`
- `VITE_POST_LOGOUT_REDIRECT_URI`: 未指定時は `http://localhost:3000`

## Entra側の設定
- SPA用アプリ（Client）
  - MSAL (Auth Code + PKCE) によるサインイン
  - APIの delegated permission（api://<API_APP_ID_URI>/access_as_user）を要求してアクセストークン取得
- API用アプリ（Resource）
  - Expose an API
    - Application ID URI
    - access_as_user スコープ作成
  - App roles（Admin/Reader等）作成
  - Enterprise applications（API側）でユーザー/グループにロール割り当て
  - APIは access token の aud/iss/署名 を検証し、roles claim で認可
