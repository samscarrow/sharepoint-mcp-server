#!/usr/bin/env node

/**
 * One-time OAuth2 login to get a delegated refresh token for email access.
 * Run: npm run oauth-login
 *
 * Prerequisites:
 *   1. Azure app registration has http://localhost:3000/callback as a Web redirect URI
 *   2. Delegated permissions Mail.Read + Mail.Send are added (no admin consent needed)
 *   3. Env vars TENANT_ID, CLIENT_ID, CLIENT_SECRET are set
 */

import http from "node:http";
import { URL } from "node:url";
import fs from "node:fs";
import path from "node:path";
import { execSync } from "node:child_process";

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  console.error("Required env vars: TENANT_ID, CLIENT_ID, CLIENT_SECRET");
  process.exit(1);
}

const REDIRECT_URI = "http://localhost:3000/callback";
const SCOPES = "Mail.ReadWrite Mail.Send Files.Read.All offline_access";
const TOKEN_FILE =
  process.env.TOKEN_FILE ||
  path.join(process.env.HOME || "/tmp", ".config", "bay-view-graph", "tokens.json");

// Build authorize URL
const authUrl = new URL(
  `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`
);
authUrl.searchParams.set("client_id", CLIENT_ID);
authUrl.searchParams.set("response_type", "code");
authUrl.searchParams.set("redirect_uri", REDIRECT_URI);
authUrl.searchParams.set("scope", SCOPES);
authUrl.searchParams.set("response_mode", "query");

async function exchangeCode(code: string): Promise<any> {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: CLIENT_ID!,
    client_secret: CLIENT_SECRET!,
    code,
    redirect_uri: REDIRECT_URI,
    grant_type: "authorization_code",
    scope: SCOPES,
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params,
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Token exchange failed: ${response.status} ${err}`);
  }

  return response.json();
}

function saveTokens(data: any): void {
  const dir = path.dirname(TOKEN_FILE);
  fs.mkdirSync(dir, { recursive: true });

  const tokens = {
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    expires_at: Date.now() + data.expires_in * 1000 - 60000,
  };

  fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2), { mode: 0o600 });
  console.log(`\nTokens saved to ${TOKEN_FILE}`);
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url!, "http://localhost:3000");

  if (url.pathname === "/callback") {
    const code = url.searchParams.get("code");
    const error = url.searchParams.get("error");

    if (error) {
      const desc = url.searchParams.get("error_description") || "";
      res.writeHead(400, { "Content-Type": "text/html" });
      res.end(`<h1>Error</h1><p>${error}: ${desc}</p>`);
      server.close();
      process.exit(1);
    }

    if (code) {
      try {
        const tokens = await exchangeCode(code);
        saveTokens(tokens);
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end(
          "<h1>Success!</h1><p>Tokens saved. You can close this tab and restart the MCP server.</p>"
        );
      } catch (err) {
        res.writeHead(500, { "Content-Type": "text/html" });
        res.end(`<h1>Error</h1><pre>${err}</pre>`);
      }
      server.close();
      process.exit(0);
    }
  }

  res.writeHead(404);
  res.end();
});

server.listen(3000, () => {
  console.log("Listening on http://localhost:3000/callback\n");
  console.log("Opening browser for sign-in...\n");
  console.log(authUrl.toString());

  try {
    execSync(`xdg-open "${authUrl.toString()}"`, { stdio: "ignore" });
  } catch {
    console.log("\nCould not open browser automatically. Open the URL above manually.");
  }
});
