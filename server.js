const express = require("express");
const path = require("path");
const { initDb, getState, setState, findProjectOwner } = require("./db");
const { getGraphPhoto } = require("./graph");
const { getAuthorizeUrl, handleCallback, isConnected, getValidAccessToken } = require("./msftAuth");

const app = express();

// Minato's gateway enforces tenant SSO on every request and stamps a trusted
// identity header, but any signed-in tenant user could otherwise open this
// app. State is stored per user (keyed by that identity), so everyone who
// opens the app gets their own private task list that persists across
// refreshes. The configured owner additionally gets the "connect Microsoft"
// action for profile photos. See minato://docs/reference/end-user-identity.
function identify(req, res, next) {
  const uid = req.headers["x-auth-oid"] || req.headers["x-auth-subject"];
  if (!uid) {
    // Reached without the SSO gateway's identity headers (e.g. the editor's
    // Preview proxy). Fall back to a single local user instead of rejecting,
    // so the app is usable in the Preview. Behind the gateway, per-user
    // storage still applies as before.
    // Reached without the SSO gateway's identity headers (e.g. the editor's
    // Preview proxy). Fall back to a single local user instead of rejecting,
    // so the app is usable in the Preview. Behind the gateway, per-user
    // storage still applies as before. Treated as owner while OWNER_OID is
    // unset (bootstrap mode) so the Microsoft connect actions are reachable.
    req.uid = "preview-user";
    req.isOwner = !process.env.OWNER_OID;
    return next();
  }

  const owner = process.env.OWNER_OID;
  if (!owner) {
    // Bootstrap mode: OWNER_OID isn't set yet. Log the visiting identity so
    // it can be captured once and pinned via the OWNER_OID env var. The app
    // works fine without it (per-user storage needs no owner); it only gates
    // the "connect Microsoft" profile-photo integration.
    console.log(
      "OWNER_OID not configured yet. Visiting identity:",
      JSON.stringify({ oid: req.headers["x-auth-oid"], subject: req.headers["x-auth-subject"], email: req.headers["x-auth-user"] })
    );
  }
  req.uid = uid;
  req.isOwner = !!owner && uid === owner;
  next();
}
app.use(identify);

// Generous limit: state includes base64-encoded file attachments.
app.use(express.json({ limit: "50mb" }));

// Minato's SSO contract gives every app a visitor's name and email, but not
// a photo (see minato://docs/reference/end-user-identity) — that comes from
// Microsoft Graph instead, using the owner's own delegated grant (see
// msftAuth.js), looked up by the visitor's own email.
app.get("/api/whoami", async (req, res) => {
  const fullName = req.headers["x-auth-name"] || "there";
  const firstName = fullName.split(" ")[0];
  const email = req.headers["x-auth-user"];

  const hasAvatar = email ? !!(await getGraphPhoto(email)) : false;
  res.json({
    firstName,
    subtitle: null,
    hasAvatar,
    // Only the owner can see/use the "connect Microsoft" action.
    canConnectGraph: req.isOwner,
    graphConnected: req.isOwner ? await isConnected() : false,
  });
});

app.get("/api/avatar", async (req, res) => {
  const email = req.headers["x-auth-user"];
  if (!email) return res.sendStatus(404);
  const photo = await getGraphPhoto(email);
  if (!photo) return res.sendStatus(404);
  res.set("Content-Type", "image/jpeg");
  res.send(photo);
});

// Build the /auth/callback URL as the browser sees this app, so sign-in
// works both from the deployed URL and from the editor Preview. The Preview
// proxies requests under a /proxy/<port>/ path prefix that is stripped before
// reaching us, so reconstruct the public base from the Referer the browser
// sends when clicking the link (or the workspace proxy template), falling
// back to the forwarded host. Azure only accepts registered URIs.
function callbackUri(req) {
  const referer = req.headers["referer"];
  if (referer) {
    try {
      const u = new URL(referer);
      // Keep whatever path prefix the browser is actually on (/proxy/3000/…)
      const m = u.pathname.match(/^(\/proxy\/\d+)/);
      return `${u.origin}${m ? m[1] : ""}/auth/callback`;
    } catch (e) { /* fall through */ }
  }
  if (process.env.VSCODE_PROXY_URI) {
    const base = process.env.VSCODE_PROXY_URI.replace("{{port}}", "3000").replace(/\/$/, "");
    return `${base}/auth/callback`;
  }
  const host = req.headers["x-forwarded-host"] || req.headers["host"];
  const proto = req.headers["x-forwarded-proto"] || "https";
  if (host) return `${proto}://${host}/auth/callback`;
  return undefined;
}

app.get("/auth/login", (req, res) => {
  if (!req.isOwner) return res.sendStatus(403);
  const url = getAuthorizeUrl(callbackUri(req));
  console.log("Auth login →", url);
  res.redirect(url);
});

app.get("/auth/callback", async (req, res) => {
  if (!req.isOwner) return res.sendStatus(403);
  try {
    await handleCallback(req.query.code, req.query.state, callbackUri(req));
    res.redirect("/?graphConnected=1");
  } catch (err) {
    console.error("Microsoft auth callback failed:", err.message);
    res.status(500).send("Could not complete sign-in with Microsoft. Please try again.");
  }
});

// Read-only calendar sync: the owner's Outlook/Teams meetings for a date
// range, so they can plan todos around them. Requires the Calendars.Read
// delegated permission and a completed "connect Microsoft" sign-in.
app.get("/api/events", async (req, res) => {
  try {
    const token = await getValidAccessToken();
    if (!token) return res.json({ connected: false, events: [] });
    const start = req.query.start;
    const end = req.query.end;
    if (!start || !end) return res.status(400).json({ error: "start and end required" });
    const url = `https://graph.microsoft.com/v1.0/me/calendarview` +
      `?startDateTime=${encodeURIComponent(start)}` +
      `&endDateTime=${encodeURIComponent(end)}` +
      `&$orderby=start/dateTime&$top=100` +
      `&$select=subject,start,end,isAllDay,showAs`;
    const graphRes = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.timezone="Europe/Oslo"',
      },
    });
    if (!graphRes.ok) {
      console.error("Graph calendarview failed:", graphRes.status, await graphRes.text());
      return res.status(502).json({ error: "Failed to fetch calendar" });
    }
    const data = await graphRes.json();
    const events = (data.value || []).map(e => ({
      subject: e.subject || "(uten tittel)",
      start: e.start?.dateTime || null,
      end: e.end?.dateTime || null,
      allDay: !!e.isAllDay,
      busy: e.showAs !== "free",
    }));
    res.json({ connected: true, events });
  } catch (err) {
    console.error("GET /api/events failed:", err);
    res.status(500).json({ error: "Failed to load calendar" });
  }
});

// AI-generated weekly summary: what got done (completed todos), what was
// planned but not done, and (if the owner connected their calendar) which
// meetings filled the week. Runs on the workspace's Intility-hosted LLM.
app.get("/api/weekly-summary", async (req, res) => {
  try {
    const state = await getState(req.uid);
    const now = Date.now();
    const weekAgo = now - 7 * 24 * 3600000;
    const done = state.todos.filter(t => t.completed && !t.archived && t.completedAt && t.completedAt >= weekAgo)
      .map(t => ({ title: t.title, completedAt: new Date(t.completedAt).toISOString().slice(0, 10) }));
    const missed = state.todos.filter(t => !t.completed && !t.archived && t.planned && new Date(t.planned + "T23:59") < new Date())
      .map(t => ({ title: t.title, planned: t.planned, postponed: !!t.postponedUntil }));
    const upcoming = state.todos.filter(t => !t.completed && !t.archived && t.planned && new Date(t.planned + "T00:00") >= new Date())
      .map(t => ({ title: t.title, planned: t.planned }));

    let meetings = [];
    const token = await getValidAccessToken();
    if (token) {
      try {
        const startISO = new Date(weekAgo).toISOString();
        const url = `https://graph.microsoft.com/v1.0/me/calendarview` +
          `?startDateTime=${encodeURIComponent(startISO)}` +
          `&endDateTime=${encodeURIComponent(new Date(now).toISOString())}` +
          `&$top=100&$select=subject,start`;
        const graphRes = await fetch(url, {
          headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.timezone="Europe/Oslo"' },
        });
        if (graphRes.ok) {
          const data = await graphRes.json();
          meetings = (data.value || []).map(e => ({ subject: e.subject, start: e.start?.dateTime || null }));
        }
      } catch (e) { /* calendar optional */ }
    }

    const llmBase = process.env.MINATO_LLM_BASE_URL;
    const llmKey = process.env.MINATO_LLM_API_KEY;
    const llmModel = process.env.MINATO_LLM_MODELS || "glm-5-3";
    if (!llmBase || !llmKey) return res.status(503).json({ error: "AI service not configured" });

    const prompt = `Du er en vennlig norskspråklig assistent som skriver korte, profesjonelle arbeidsrapporter. Lag et kort sammendrag for brukeren basert på dataene under, strukturert i nøyaktig disse fire seksjonene (bruk --- som skille mellom seksjonene, ingen andre markeringer, ingen emojier, ingen "anbefaling"-seksjon):

FULLFØRT
(korte punkter med oppgavenes titler, eller "Ingen fullførte oppgaver denne uken")

MØTER
(korte punkter med hva møtene handlet om, eller "Ingen møter registrert")

UFERDIG
(korte punkter med planlagte oppgaver som ikke ble fullført, eller "Alt planlagt ble fullført")

KOMMENDE
(korte punkter med hva som venter, eller "Ingenting planlagt")

Regler: maks 4 punkter per seksjon, korte fraser, bruk oppgavenes faktiske titler, ingen spekulasjoner eller råd, kun fakta fra dataene. Tekst rett ut uten markdown-tegn som ** eller #.

Fullførte oppgaver (tittel, dato):
${done.length ? done.map(d => `- ${d.title} (${d.completedAt})`).join("\n") : "(ingen)"}

Møter siste uken (tittel, tidspunkt):
${meetings.length ? meetings.map(m => `- ${m.subject} (${m.start})`).join("\n") : "(ingen)"}

Planlagt men IKKE fullført (tittel, planlagt dato):
${missed.length ? missed.map(m => `- ${m.title} (planlagt ${m.planned})`).join("\n") : "(ingen)"}

Planlagt fremover:
${upcoming.length ? upcoming.map(u => `- ${u.title} (${u.planned})`).join("\n") : "(ingen)"}`;

    const llmRes = await fetch(`${llmBase}/chat/completions`, {
      method: "POST",
      headers: { Authorization: `Bearer ${llmKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model: llmModel,
        messages: [{ role: "user", content: prompt }],
        max_tokens: 800,
      }),
    });
    if (!llmRes.ok) {
      console.error("LLM call failed:", llmRes.status, await llmRes.text());
      return res.status(502).json({ error: "AI summary failed" });
    }
    const llmData = await llmRes.json();
    const text = llmData.choices?.[0]?.message?.content || "";
    res.json({ summary: text.trim(), stats: { done: done.length, missed: missed.length, meetings: meetings.length, upcoming: upcoming.length } });
  } catch (err) {
    console.error("GET /api/weekly-summary failed:", err);
    res.status(500).json({ error: "Failed to build summary" });
  }
});

app.get("/api/state", async (req, res) => {
  try {
    res.json(await getState(req.uid));
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
    await setState(req.uid, req.body);
    res.json({ ok: true });
  } catch (err) {
    console.error("PUT /api/state failed:", err);
    res.status(500).json({ error: "Failed to save state" });
  }
});

// Anyone signed in (any Intility employee) can follow a project's share link
// to be recorded as a member. The invite lives in the project creator's state,
// wherever in the per-user table that is. This only records membership today —
// the data model still needs to become project-aware before a member can
// actually see or edit the shared project's tasks. That's the next step.
app.get("/join/:projectId/:token", async (req, res) => {
  try {
    const found = await findProjectOwner(req.params.projectId, req.params.token);
    if (!found) {
      return res.status(404).send("This invite link is invalid or has expired.");
    }
    const state = await getState(found.userId);
    const project = state.projects.find(p => p.id === req.params.projectId);
    if (!project || !project.shareToken || project.shareToken !== req.params.token) {
      return res.status(404).send("This invite link is invalid or has expired.");
    }
    if (!project.members) project.members = [];
    if (!project.members.find(m => m.oid === req.uid)) {
      project.members.push({
        oid: req.uid,
        email: req.headers["x-auth-user"] || "",
        name: req.headers["x-auth-name"] || "",
        joinedAt: Date.now(),
      });
      await setState(found.userId, state);
    }
    res.redirect("/?joined=" + encodeURIComponent(project.name));
  } catch (err) {
    console.error("Join failed:", err);
    res.status(500).send("Something went wrong joining this project.");
  }
});

app.use(express.static(path.join(__dirname, "public")));

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
