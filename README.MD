# Project: BankID + SharePoint examples

This repository contains two small TypeScript Express examples demonstrating:

- BankID (OIDC) sandbox authentication using Authorization Code + PKCE (`src/bankId.ts`).
- A SharePoint / Microsoft Graph file CRUD API using app-only auth (`src/sharePoint.ts`).

Both apps are runnable locally. The BankID example runs over HTTPS (self-signed certs included under `certs/`), and the SharePoint API runs over HTTP (default port 3001).

## Prerequisites

- Node.js (recommended 18+)
- pnpm (the project uses pnpm but npm/yarn may work)

Install dependencies:

```bash
pnpm install
```

## Top-level scripts

Key scripts in `package.json`:

- `pnpm dev`  Run the BankID example (`src/bankId.ts`) with `tsx` for fast dev.
- `pnpm dev:sharepoint`  Run the SharePoint example (`src/sharePoint.ts`) with `tsx`.
- `pnpm build`  Transpile TypeScript to `dist/`.
- `pnpm start`  Run the built BankID server from `dist/`.

Run either service in separate terminals. Example (zsh):

```bash
# BankID (HTTPS) server
pnpm dev

# SharePoint API server (port 3001)
pnpm dev:sharepoint
```

## BankID example (src/bankId.ts)

- Starts an HTTPS Express server on https://localhost:3000 using `certs/localhost-key.pem` and `certs/localhost-cert.pem`.
- Uses `openid-client` to discover the sandbox provider and perform Authorization Code + PKCE.
- Session is used to store PKCE verifier, state and nonce; ID token claims and optional UserInfo are saved in session.

Important notes:

- The current code hardcodes `CLIENT_ID`, `CLIENT_SECRET`, and `SERVER_URL` (sandbox values). Do NOT use those in production.
- Recommended: move sensitive values to environment variables (see `.env.example`).
- Session secret is currently a placeholder; replace with a secure random value for real deployments.

Routes (BankID app):

- GET /  Status; shows `loginUrl` when not authenticated and profile when authenticated.
- GET /auth/bankid  Start OIDC flow (redirects to provider).
- GET /auth/bankid/callback  Callback that exchanges code for tokens and fetches userinfo.
- GET /user  Returns combined profile data (ID token claims + userinfo).
- GET /logout  Destroys session.
- GET /auth/failure  Authentication error page.

TLS / certs

The BankID example expects the TLS files at `certs/localhost-key.pem` and `certs/localhost-cert.pem`. They are self-signed for local testing and will trigger browser warnings. Use mkcert or a trusted cert in production.

## SharePoint / Microsoft Graph example (src/sharePoint.ts)

- App-only authentication using `@azure/msal-node` (Client Credentials flow) and `@microsoft/microsoft-graph-client`.
- Implements basic CRUD endpoints for files and folders in a SharePoint document library (target: `Dokumenty` / `Documents`).

Required environment variables (see `.env.example`):

- TENANT_ID
- CLIENT_ID
- CLIENT_SECRET
- SITE_ID
- SHAREPOINT_SITE_URL

SharePoint routes (examples):

- GET /  API overview
- GET /health  Health check (verifies Graph access to the target site)
- GET /files  List files in the Dokumenty/Documents library
- GET /files/:fileName  Get metadata for a file
- GET /files/:fileName/download  Download a file
- POST /files/upload  Upload via multipart/form-data (field `file`)
- POST /files  Upload small files via JSON (name + content)
- PUT /files/:fileName  Replace file content
- DELETE /files/:fileName  Delete a file
- GET /folders  List folders
- POST /folders  Create a folder
- DELETE /folders/:folderName  Delete a folder

Upload behavior:

- Files smaller than 4MB use a simple PUT to `/content`.
- Files >= 4MB use a resumable upload session (chunked upload).
- The server limits uploads to ~250MB (configured in Multer). Adjust as needed.

## Environment examples

Create a `.env` (not checked into source) based on `.env.example` and set the required values before starting the SharePoint app. Example variables used by the repository:

- For SharePoint (`src/sharePoint.ts`): TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID, SHAREPOINT_SITE_URL
- For BankID it's recommended to add: BANKID_CLIENT_ID, BANKID_CLIENT_SECRET, LISTEN_ORIGIN (or edit `src/bankId.ts` directly)

## Troubleshooting

- If BankID server fails to start because TLS files are missing, verify `certs/localhost-key.pem` and `certs/localhost-cert.pem` exist.
- If the OpenID provider rejects the client or redirect URI, ensure registered client configuration matches the `redirect_uris` used by `src/bankId.ts` (default `https://localhost:3000/auth/bankid/callback`).
- For state/nonce mismatches, ensure your browser accepts cookies; sessions store PKCE verifier, state and nonce.
- For SharePoint Graph errors, verify app registration credentials and that the app has appropriate Graph API permissions (application-level) and SITE_ID is correct.
