const express = require("express");
const path = require("path");
const { initDb, getState, setState, EMPTY_STATE, searchEmployees } = require("./db");

const app = express();

// Minato's gateway enforces tenant SSO on every request and stamps a trusted
// identity header, but any signed-in tenant user could otherwise open this
// app. This is a personal task list, not a shared one: everyone can load the
// page, but only the configured owner's requests see or change the real
// data. Everyone else always gets a fresh, empty, non-persisted state.
// See minato://docs/reference/end-user-identity.
function identify(req, res, next) {
  const uid = req.headers["x-auth-oid"] || req.headers["x-auth-subject"];
  if (!uid) return res.sendStatus(401);

  const owner = process.env.OWNER_OID;
  if (!owner) {
    // Bootstrap mode: OWNER_OID isn't set yet. Log the visiting identity so
    // it can be captured once and pinned via the OWNER_OID env var. Treat
    // nobody as the owner in the meantime, so no data is exposed before
    // OWNER_OID is actually configured.
    console.log(
      "OWNER_OID not configured yet. Visiting identity:",
      JSON.stringify({ oid: req.headers["x-auth-oid"], subject: req.headers["x-auth-subject"], email: req.headers["x-auth-user"] })
    );
  }
  req.isOwner = !!owner && uid === owner;
  next();
}
app.use(identify);

// Generous limit: state includes base64-encoded file attachments.
app.use(express.json({ limit: "50mb" }));
app.use(express.static(path.join(__dirname, "public")));

// Shown only for the owner. Minato's SSO contract gives every app a
// visitor's name and email, but not role/department/photo for anyone
// (see minato://docs/reference/end-user-identity) — those aren't available
// to fetch per-user, so they're hardcoded for the one person they're true for.
const OWNER_SUBTITLE = "Management Graduate • Cloud Infrastructure • Intility AS";

app.get("/api/whoami", (req, res) => {
  const fullName = req.headers["x-auth-name"] || "there";
  const firstName = fullName.split(" ")[0];
  res.json({
    firstName,
    subtitle: req.isOwner ? OWNER_SUBTITLE : null,
    hasAvatar: req.isOwner,
  });
});

app.get("/api/avatar", (req, res) => {
  if (!req.isOwner) return res.sendStatus(404);
  res.sendFile(path.join(__dirname, "private", "avatar.jpg"));
});

app.get("/api/state", async (req, res) => {
  if (!req.isOwner) return res.json(EMPTY_STATE);
  try {
    res.json(await getState());
  } catch (err) {
    console.error("GET /api/state failed:", err);
    res.status(500).json({ error: "Failed to load state" });
  }
});

// Powers the project "invite" picker. Owner-only for now, same as everything
// else — sharing/membership management isn't built yet, this just lets the
// picker search the company directory ahead of that.
app.get("/api/employees", async (req, res) => {
  if (!req.isOwner) return res.json([]);
  const q = String(req.query.q || "").trim();
  if (!q) return res.json([]);
  try {
    res.json(await searchEmployees(q));
  } catch (err) {
    console.error("GET /api/employees failed:", err);
    res.status(500).json({ error: "Search failed" });
  }
});

app.put("/api/state", async (req, res) => {
  if (!req.isOwner) return res.json({ ok: true }); // silently no-op for everyone else
  try {
    if (!req.body || typeof req.body !== "object") {
      return res.status(400).json({ error: "Body must be a JSON object" });
    }
    await setState(req.body);
    res.json({ ok: true });
  } catch (err) {
    console.error("PUT /api/state failed:", err);
    res.status(500).json({ error: "Failed to save state" });
  }
});

const PORT = process.env.PORT || 3000;

// Listen immediately so the platform's readiness check succeeds even if the
// database isn't reachable yet (e.g. no database attached, or it's still
// provisioning). DB setup happens in the background and retries on failure
// instead of crashing the process.
app.listen(PORT, "0.0.0.0", () => {
  console.log(`Intility Task listening on 0.0.0.0:${PORT}`);
});

function initDbWithRetry() {
  initDb().catch((err) => {
    console.error("Database not ready yet, retrying in 10s:", err.message);
    setTimeout(initDbWithRetry, 10000);
  });
}
initDbWithRetry();
