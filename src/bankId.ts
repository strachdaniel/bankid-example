import express from "express";
import session from "express-session";
import https from "https";
import fs from "fs";
import path from "path";
import crypto from "crypto";
// We'll dynamically import openid-client inside main() to avoid ESM interop/runtime issues

const CLIENT_ID = "0caeebae-4ec0-4ba7-95ba-f9a89d47ad40";
const CLIENT_SECRET =
  "AIrOBb5HAn-2GdVdYLuqhPyOezACPStZ7p79h35llsOy1NTv0atO0EkdnbOrOLo8ohYtFfMP5m886X8w4ABcqDU";
const SCOPE =
  "profile.birthnumber profile.phonenumber profile.zoneinfo profile.gender openid profile.titles notification.claims_updated profile.name profile.birthplaceNationality profile.locale profile.idcards profile.maritalstatus profile.legalstatus profile.email profile.paymentAccounts profile.addresses profile.birthdate profile.updatedat";
const SERVER_URL = "https://oidc.sandbox.bankid.cz";
const LISTEN_ORIGIN = "https://localhost:3000"; // Changed to HTTPS
const REDIRECT_URI = `${LISTEN_ORIGIN}/auth/bankid/callback`;

// Minimal augmentation of Express session to hold our values (avoids extra types for now)
declare module "express-session" {
  interface SessionData {
    code_verifier?: string;
    user?: any;
    state?: string;
    nonce?: string;
  }
}

async function main() {
  // Dynamically import openid-client so runtime exports resolve correctly
  const oidc = (await import("openid-client")) as any;

  // Discover the provider and create a client configuration using openid-client v6 API
  const clientConfig = await oidc.discovery(new URL(SERVER_URL), CLIENT_ID, {
    client_secret: CLIENT_SECRET,
    redirect_uris: [REDIRECT_URI],
    response_types: ["code"],
  });

  const app = express();

  app.use(
    session({
      secret: "change_this_to_a_real_secret",
      resave: false,
      saveUninitialized: false,
    })
  );

  // Home route - shows auth state
  app.get("/", (req, res) => {
    if ((req.session as any).user) {
      const user = (req.session as any).user;
      const claims = user.claims || {};
      const userinfo = user.userinfo || {};

      // Combine claims and userinfo to get all available user data
      const allUserData = { ...claims, ...userinfo };

      // Extract common user information
      const userProfile = {
        name:
          allUserData.name ||
          allUserData.given_name + " " + allUserData.family_name ||
          "N/A",
        email: allUserData.email || "N/A",
        given_name: allUserData.given_name || "N/A",
        family_name: allUserData.family_name || "N/A",
        birthdate: allUserData.birthdate || "N/A",
        gender: allUserData.gender || "N/A",
        phone_number: allUserData.phone_number || "N/A",
        locale: allUserData.locale || "N/A",
        zoneinfo: allUserData.zoneinfo || "N/A",
        updated_at: allUserData.updated_at || "N/A",
        sub: allUserData.sub || "N/A",
      };

      res.json({
        message: "Authenticated",
        userProfile,
        rawClaims: claims,
        rawUserInfo: userinfo,
        logoutUrl: "/logout",
      });
    } else {
      res.json({ message: "Not authenticated", loginUrl: "/auth/bankid" });
    }
  });

  // Initiate authentication with BankID (Authorization Code + PKCE)
  app.get("/auth/bankid", async (req, res) => {
    // generate PKCE verifier/challenge
    const code_verifier = oidc.randomPKCECodeVerifier
      ? oidc.randomPKCECodeVerifier()
      : oidc.randomPKCECodeVerifier();
    console.log("Generated PKCE code verifier:", code_verifier);
    const code_challenge = oidc.calculatePKCECodeChallenge
      ? await oidc.calculatePKCECodeChallenge(code_verifier)
      : await oidc.calculatePKCECodeChallenge(code_verifier);
    console.log("Generated PKCE code challenge:", code_challenge);

    // store verifier in session for callback exchange
    req.session.code_verifier = code_verifier;

    // generate state and nonce and store in session for verification
    const state = crypto.randomUUID();
    const nonce = crypto.randomUUID();
    req.session.state = state;
    req.session.nonce = nonce;

    const authUrl = oidc.buildAuthorizationUrl(clientConfig, {
      scope: SCOPE,
      code_challenge,
      code_challenge_method: "S256",
      redirect_uri: REDIRECT_URI,
      state,
      nonce,
      prompt: "login",
      display: "page",
      acr_values: "loa3",
    });

    console.log("Generated authorization URL:", authUrl);

    console.info("Redirecting to BankID for authentication");
    res.redirect(authUrl);
  });

  // Callback route - exchange code for tokens and fetch userinfo
  app.get("/auth/bankid/callback", async (req, res) => {
    try {
      // Validate state parameter
      const receivedState = req.query.state;
      if (receivedState !== req.session.state) {
        console.error("State parameter mismatch");
        return res.redirect("/auth/failure");
      }

      // Exchange the authorization code for tokens using openid-client v6 helper
      // Note: openid-client v6 expects the Configuration/client as the first arg
      // and a URL object constructed from the current request URL as the second arg
      // We need to remove the state parameter from the URL since we validate it manually
      const currentUrl = new URL(
        `${req.protocol}://${req.get("host")}${req.originalUrl}`
      );
      currentUrl.searchParams.delete("state"); // Remove state param for openid-client

      const result = await oidc.authorizationCodeGrant(
        clientConfig,
        currentUrl,
        {
          pkceCodeVerifier: req.session.code_verifier,
          expectedNonce: req.session.nonce, // Provide expected nonce for validation
        }
      );

      // result has helpers like claims()
      const claims = result.claims ? result.claims() : {};
      console.log("ID Token claims:", claims);

      // Optionally fetch userinfo
      let userinfo = {};
      try {
        if (result.access_token && claims.sub) {
          userinfo = await oidc.fetchUserInfo(
            clientConfig,
            result.access_token,
            claims.sub // Pass the subject from ID token
          );
          console.log("UserInfo response:", userinfo);
        }
      } catch (e) {
        console.warn(
          "fetchUserInfo failed, continuing with ID token claims",
          e
        );
      }

      req.session.user = { claims, userinfo, tokenSet: result };

      res.redirect("/");
    } catch (err) {
      console.error("Callback error:", err);
      res.redirect("/auth/failure");
    }
  });

  // User info route - detailed user information
  app.get("/user", (req, res) => {
    if (!(req.session as any).user) {
      return res
        .status(401)
        .json({ error: "Not authenticated", loginUrl: "/auth/bankid" });
    }

    const user = (req.session as any).user;
    const claims = user.claims || {};
    const userinfo = user.userinfo || {};

    // Combine all available user data
    const allUserData = { ...claims, ...userinfo };

    res.json({
      profile: {
        // Basic Info
        name: allUserData.name,
        given_name: allUserData.given_name,
        family_name: allUserData.family_name,
        email: allUserData.email,

        // Personal Details
        birthdate: allUserData.birthdate,
        birthnumber: allUserData.birthnumber,
        gender: allUserData.gender,

        // Contact Info
        phone_number: allUserData.phone_number,
        addresses: allUserData.addresses,

        // Identity Documents
        idcards: allUserData.idcards,

        // Additional Info
        titles: allUserData.titles,
        maritalstatus: allUserData.maritalstatus,
        legalstatus: allUserData.legalstatus,
        locale: allUserData.locale,
        zoneinfo: allUserData.zoneinfo,
        birthplaceNationality: allUserData.birthplaceNationality,

        // Financial
        paymentAccounts: allUserData.paymentAccounts,

        // Metadata
        updated_at: allUserData.updated_at,
        sub: allUserData.sub,
      },
      raw: {
        claims,
        userinfo,
      },
    });
  });

  // Authentication failure route
  app.get("/auth/failure", (req, res) => {
    res.status(400).json({ message: "Authentication failed" });
  });

  // Logout route
  app.get("/logout", (req, res) => {
    req.session.destroy((err) => {
      if (err) {
        return res.status(500).json({ error: "Logout failed" });
      }
      res.json({ message: "Logged out successfully" });
    });
  });

  // Create HTTPS server
  const sslOptions = {
    key: fs.readFileSync(
      path.join(process.cwd(), "certs", "localhost-key.pem")
    ),
    cert: fs.readFileSync(
      path.join(process.cwd(), "certs", "localhost-cert.pem")
    ),
  };

  const httpsServer = https.createServer(sslOptions, app);

  httpsServer.listen(3000, () => {
    console.info("HTTPS Server listening on https://localhost:3000");
    console.info(
      "Note: You may need to accept the self-signed certificate in your browser"
    );
  });
}

main().catch((err) => {
  console.error("Fatal error starting app:", err);
  process.exit(1);
});
