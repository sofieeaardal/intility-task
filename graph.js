// Looks up a visitor's profile photo from Microsoft Graph, using the owner's
// own delegated User.ReadBasic.All grant (see msftAuth.js) rather than an
// app-only Application permission — that grant lets the owner's token read
// any org member's basic profile, so only the owner needs to have connected
// once. Note: ReadBasic.All only exposes name/email/photo, not job title or
// department, so there's no directory-sourced subtitle — just the photo.
const { getValidAccessToken } = require("./msftAuth");

const photoCache = new Map(); // email -> { buffer, expiresAt }
const PHOTO_TTL_MS = 60 * 60 * 1000;

async function getGraphPhoto(email) {
  const cached = photoCache.get(email);
  if (cached && Date.now() < cached.expiresAt) return cached.buffer;

  const token = await getValidAccessToken();
  if (!token) return null;

  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}/photo/$value`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      // Not cached: a fix (e.g. reconnecting) should take effect on the next
      // request, not be masked by an hour-old failed lookup.
      console.error(`Graph photo request failed: ${res.status} ${await res.text()}`);
      return null;
    }
    const buffer = Buffer.from(await res.arrayBuffer());
    photoCache.set(email, { buffer, expiresAt: Date.now() + PHOTO_TTL_MS });
    return buffer;
  } catch (err) {
    console.error("Graph photo lookup failed:", err.message);
    return null;
  }
}

module.exports = { getGraphPhoto };
