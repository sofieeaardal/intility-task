// Delegated "Sign in with Microsoft" flow (authorization code grant), used
// only by the owner, once. Application permissions (app-only, no user) were
// ruled out in favor of this narrower per-user-consent model: the owner's
// own delegated User.ReadBasic.All grant is then reused to look up any
// visitor's basic profile — see graph.js.
const crypto = require("crypto");
const { getGraphAuth, setGraphAuth } = require("./db");

const SCOPE = "openid profile offline_access User.ReadBasic.All";
const REDIRECT_URI = "https://intility-task.intility.minato.intilitest.no/auth/callback";

// Short-lived CSRF state store: login redirects and callbacks happen within
// the same browser session seconds apart, so an in-memory Map is enough.
const pendingStates = new Map(); // state -> expiresAt
function createState() {
  const state = crypto.randomUUID();
  pendingStates.set(state, Date.now() + 5 * 60 * 1000);
  return state;
}
function consumeState(state) {
  const expiresAt = pendingStates.get(state);
  pendingStates.delete(state);
  return !!expiresAt && Date.now() < expiresAt;
}

function getAuthorizeUrl() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    response_mode: "query",
    scope: SCOPE,
    state: createState(),
  });
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params}`;
}

async function handleCallback(code, state) {
  if (!consumeState(state)) throw new Error("Invalid or expired state");

  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "authorization_code",
      code,
      redirect_uri: REDIRECT_URI,
      scope: SCOPE,
    }),
  });
  if (!res.ok) throw new Error(`Token exchange failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  await setGraphAuth({
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresAt: Date.now() + (data.expires_in - 60) * 1000,
  });
}

async function refresh(refreshToken) {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: SCOPE,
    }),
  });
  if (!res.ok) throw new Error(`Token refresh failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  await setGraphAuth({
    accessToken: data.access_token,
    refreshToken: data.refresh_token || refreshToken,
    expiresAt: Date.now() + (data.expires_in - 60) * 1000,
  });
  return data.access_token;
}

// Returns a valid access token, refreshing if needed, or null if the owner
// has never connected (or the refresh token has died and needs reconnecting).
async function getValidAccessToken() {
  const auth = await getGraphAuth();
  if (!auth) return null;
  if (Date.now() < auth.expires_at) return auth.access_token;
  try {
    return await refresh(auth.refresh_token);
  } catch (err) {
    console.error("Graph token refresh failed, owner needs to reconnect:", err.message);
    return null;
  }
}

async function isConnected() {
  return !!(await getGraphAuth());
}

module.exports = { getAuthorizeUrl, handleCallback, getValidAccessToken, isConnected };
