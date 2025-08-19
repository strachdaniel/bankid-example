# BankID Sandbox — Example

A small Node + TypeScript example that demonstrates using the BankID OIDC sandbox (Authorization Code flow + PKCE) to authenticate a user and fetch profile information.

This project starts an HTTPS Express server (self-signed certs are included under `certs/`) and uses `openid-client` to perform discovery, authorization, token exchange, and optional userinfo fetch.

## What this does

- Discovers the BankID sandbox OIDC provider.
- Initiates an authorization request (Authorization Code + PKCE).
- Exchanges the code for tokens and validates the ID token (nonce/state).
- Optionally calls the UserInfo endpoint and exposes combined profile data via routes.

The implementation is in `src/bankId.ts`.

## Prerequisites

- Node.js (recommended 18+)
- pnpm (the project was created with pnpm — `pnpm install` will work). The `package.json` lists `pnpm@10.13.1` as the project package manager but any recent pnpm should work.

## Install

Open a terminal (zsh) in the project root and run:

```bash
pnpm install
```

## TLS / certs

This example serves HTTPS on `https://localhost:3000` using files in `certs/`:

- `certs/localhost-key.pem`
- `certs/localhost-cert.pem`

They are self-signed for local development. Your browser will warn about the certificate — you can accept the risk for local testing.

If you want to generate new certs, use a tool such as mkcert or openssl (not included here).

## Configuration

The example currently hardcodes a `CLIENT_ID`, `CLIENT_SECRET`, and `SERVER_URL` inside `src/bankId.ts` for the sandbox. Do NOT use these values in production.

Recommended minimal changes before sharing or deploying:

- Move client credentials and sensitive values to environment variables.
- Use a properly issued TLS certificate for production.
- Replace the session secret with a secure, random value and configure secure cookie settings.

Example environment variables you may add (and load with your preferred method):

- BANKID_CLIENT_ID
- BANKID_CLIENT_SECRET
- LISTEN_ORIGIN (e.g. https://localhost:3000)

Note: The example is intentionally minimal and focuses on demonstrating OIDC flows.

## Scripts

The `package.json` includes these useful scripts:

- `pnpm dev` — Run the TypeScript file with `tsx` (fast local dev).
- `pnpm build` — Transpile TypeScript to `dist/` using `tsc`.
- `pnpm start` — Run the built JS from `dist/`.

Use them from the project root. Example (zsh):

```bash
pnpm dev
```

This starts the HTTPS server on port 3000.

## HTTP Routes

The example exposes a small set of routes useful for testing the flow:

- GET / — JSON status page; shows login URL when not authenticated and basic profile when authenticated.
- GET /auth/bankid — Starts the BankID authorization flow (redirects to the provider).
- GET /auth/bankid/callback — Callback endpoint used by the provider to redirect back with the authorization code.
- GET /user — Returns a more detailed JSON profile (requires session).
- GET /logout — Logs out the local session.
- GET /auth/failure — Shown on authentication errors.

Typical flow:

1. Visit `https://localhost:3000/` in your browser.
2. Click/visit the `loginUrl` (`/auth/bankid`) to be redirected to the BankID sandbox.
3. Authenticate with the sandbox provider and return to the callback which will set a session.
4. View `/user` to inspect combined claims + userinfo.

## Troubleshooting

- If the server fails to start because of TLS files, verify the files exist in `certs/` and are readable.
- If the provider rejects the client or redirect URI, check the registered client configuration on the BankID sandbox dashboard and ensure `redirect_uris` and client credentials match.
- If you see state/nonce mismatches, verify your browser allows cookies (sessions are used to store PKCE/verifier, state and nonce).
