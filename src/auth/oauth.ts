import { google } from "googleapis";
import { createServer } from "http";
import open from "open";
import { saveTokens, loadTokens, StoredTokens } from "./token-store.js";

const SCOPES = [
  "https://www.googleapis.com/auth/presentations",
  "https://www.googleapis.com/auth/drive.file",
];

const REDIRECT_PORT = 3847;
const REDIRECT_URI = `http://localhost:${REDIRECT_PORT}/callback`;

export async function getAuthenticatedClient() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    throw new Error(
      "Missing GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET environment variables"
    );
  }

  const oauth2Client = new google.auth.OAuth2(
    clientId,
    clientSecret,
    REDIRECT_URI
  );

  const tokens = await loadTokens();

  if (tokens) {
    const isExpired = Date.now() > tokens.expiry_date - 5 * 60 * 1000;

    if (isExpired && tokens.refresh_token) {
      oauth2Client.setCredentials({ refresh_token: tokens.refresh_token });
      const { credentials } = await oauth2Client.refreshAccessToken();

      const updatedTokens: StoredTokens = {
        access_token: credentials.access_token!,
        refresh_token: credentials.refresh_token || tokens.refresh_token,
        expiry_date: credentials.expiry_date!,
        scope: tokens.scope,
      };

      await saveTokens(updatedTokens);
      oauth2Client.setCredentials(updatedTokens);
    } else {
      oauth2Client.setCredentials(tokens);
    }

    return oauth2Client;
  }

  throw new Error(
    "No authentication tokens found. Run 'claude-slides --auth' first."
  );
}

export async function runAuthFlow(): Promise<void> {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    throw new Error("Missing GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET");
  }

  const oauth2Client = new google.auth.OAuth2(
    clientId,
    clientSecret,
    REDIRECT_URI
  );

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
    prompt: "consent",
  });

  console.log("Opening browser for authentication...");

  const code = await new Promise<string>((resolve, reject) => {
    const server = createServer((req, res) => {
      const url = new URL(req.url!, `http://localhost:${REDIRECT_PORT}`);
      const code = url.searchParams.get("code");

      if (code) {
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end(
          "<h1>Authentication successful!</h1><p>You can close this window.</p>"
        );
        server.close();
        resolve(code);
      } else {
        res.writeHead(400);
        res.end("Missing code parameter");
        server.close();
        reject(new Error("No authorization code received"));
      }
    });

    server.listen(REDIRECT_PORT, () => {
      open(authUrl);
    });

    setTimeout(() => {
      server.close();
      reject(new Error("Authentication timed out"));
    }, 120000);
  });

  const { tokens } = await oauth2Client.getToken(code);

  await saveTokens({
    access_token: tokens.access_token!,
    refresh_token: tokens.refresh_token!,
    expiry_date: tokens.expiry_date!,
    scope: SCOPES.join(" "),
  });

  console.log("Authentication successful! Tokens saved to ~/.claude-slides/");
}
