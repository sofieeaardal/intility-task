// Looks up a visitor's job title, department, and photo from Microsoft Graph,
// using the app's own identity (client credentials) rather than the visiting
// user's — SSO only gives us name/email/ID, not the richer directory fields.
// Requires AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET and the
// Graph "User.Read.All" application permission with admin consent granted.

let cachedToken = null;
let cachedTokenExpiry = 0;

async function getGraphToken() {
  if (cachedToken && Date.now() < cachedTokenExpiry) return cachedToken;

  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Azure credentials not configured");
  }

  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  });
  if (!res.ok) throw new Error(`Graph token request failed: ${res.status}`);
  const data = await res.json();
  cachedToken = data.access_token;
  // Refresh a minute early so a request never races an expiring token.
  cachedTokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
  return cachedToken;
}

// Profile fields rarely change; cache per user for an hour to avoid hitting
// Graph on every page load from every visitor.
const profileCache = new Map(); // email -> { data, expiresAt }
const PROFILE_TTL_MS = 60 * 60 * 1000;

async function getGraphProfile(email) {
  const cached = profileCache.get(email);
  if (cached && Date.now() < cached.expiresAt) return cached.data;

  try {
    const token = await getGraphToken();
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}?$select=jobTitle,department`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = res.ok ? await res.json() : null;
    profileCache.set(email, { data, expiresAt: Date.now() + PROFILE_TTL_MS });
    return data;
  } catch (err) {
    console.error("Graph profile lookup failed:", err.message);
    return null;
  }
}

const photoCache = new Map(); // email -> { buffer, expiresAt }

async function getGraphPhoto(email) {
  const cached = photoCache.get(email);
  if (cached && Date.now() < cached.expiresAt) return cached.buffer;

  try {
    const token = await getGraphToken();
    const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}/photo/$value`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const buffer = res.ok ? Buffer.from(await res.arrayBuffer()) : null;
    photoCache.set(email, { buffer, expiresAt: Date.now() + PROFILE_TTL_MS });
    return buffer;
  } catch (err) {
    console.error("Graph photo lookup failed:", err.message);
    return null;
  }
}

module.exports = { getGraphProfile, getGraphPhoto };
