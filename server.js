const express = require("express");
const path = require("path");
const { initDb, getState, setState } = require("./db");

const app = express();
// Generous limit: state includes base64-encoded file attachments.
app.use(express.json({ limit: "50mb" }));
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/state", async (req, res) => {
  try {
    res.json(await getState());
  } catch (err) {
    console.error("GET /api/state failed:", err);
    res.status(500).json({ error: "Failed to load state" });
  }
});

app.put("/api/state", async (req, res) => {
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
