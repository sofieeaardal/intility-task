const { Pool } = require("pg");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL || "postgres://dev:dev@localhost:5432/dev",
});

// One JSONB document per signed-in user (keyed by the SSO oid), holding that
// user's todos, projects, pinned notes and id counters. Still a document model —
// there is no multi-row relational access pattern to justify normalizing — but
// keyed per user, so every colleague's data persists across refreshes, not just
// the owner's.
const LEGACY_USER_ID = "legacy-owner"; // pre-multi-user single row; claimed by OWNER_OID once configured

const EMPTY_STATE = {
  todos: [],
  projects: [],
  pinnedNotes: [],
  view: "all",
  sort: "priority",
  nextId: 1,
  nextSubId: 1,
};

async function initDb() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_state (
      user_id TEXT PRIMARY KEY,
      data JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    )
  `);
  // One-time migration from the original single-user shape (one row keyed by an
  // id column). Every step is skipped when it has already run.
  const legacyShape = await pool.query(
    "SELECT 1 FROM information_schema.columns WHERE table_name = 'app_state' AND table_schema = current_schema() AND column_name = 'id'"
  );
  if (legacyShape.rows.length) {
    await pool.query("ALTER TABLE app_state ADD COLUMN IF NOT EXISTS user_id TEXT");
    await pool.query(
      "UPDATE app_state SET user_id = $1 WHERE user_id IS NULL",
      [process.env.OWNER_OID || LEGACY_USER_ID]
    );
    await pool.query("ALTER TABLE app_state ALTER COLUMN user_id SET NOT NULL");
    await pool.query("ALTER TABLE app_state DROP CONSTRAINT IF EXISTS app_state_pkey");
    await pool.query("ALTER TABLE app_state DROP COLUMN id");
    await pool.query("ALTER TABLE app_state ADD PRIMARY KEY (user_id)");
  }
  // Once OWNER_OID is configured, fold the unclaimed legacy row into the
  // owner's own state so their pre-multi-user data reappears.
  const owner = process.env.OWNER_OID;
  if (owner) {
    await pool.query(
      `UPDATE app_state SET user_id = $1 WHERE user_id = $2
       AND NOT EXISTS (SELECT 1 FROM app_state WHERE user_id = $1)`,
      [owner, LEGACY_USER_ID]
    );
  }
  // Holds the owner's delegated Microsoft Graph token (User.ReadBasic.All).
  // Only the owner ever connects; that one delegated grant is then used to
  // look up any visitor's basic profile for the greeting.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS graph_auth (
      id INTEGER PRIMARY KEY DEFAULT 1 CHECK (id = 1),
      access_token TEXT NOT NULL,
      refresh_token TEXT NOT NULL,
      expires_at BIGINT NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    )
  `);
}

async function getGraphAuth() {
  const { rows } = await pool.query("SELECT access_token, refresh_token, expires_at FROM graph_auth WHERE id = 1");
  return rows[0] || null;
}

async function setGraphAuth({ accessToken, refreshToken, expiresAt }) {
  await pool.query(
    `INSERT INTO graph_auth (id, access_token, refresh_token, expires_at) VALUES (1, $1, $2, $3)
     ON CONFLICT (id) DO UPDATE SET access_token = EXCLUDED.access_token, refresh_token = EXCLUDED.refresh_token,
       expires_at = EXCLUDED.expires_at, updated_at = now()`,
    [accessToken, refreshToken, expiresAt]
  );
}

async function getState(userId) {
  const { rows } = await pool.query("SELECT data FROM app_state WHERE user_id = $1", [userId]);
  if (!rows.length) return EMPTY_STATE;
  // Merge over EMPTY_STATE so older/partial saved shapes still get any new default keys.
  return { ...EMPTY_STATE, ...rows[0].data };
}

async function setState(userId, data) {
  await pool.query(
    `INSERT INTO app_state (user_id, data, updated_at) VALUES ($1, $2, now())
     ON CONFLICT (user_id) DO UPDATE SET data = EXCLUDED.data, updated_at = now()`,
    [userId, data]
  );
}

// Join links may point at a project living in any user's state (normally the
// owner's). Scan only the small, cheap `projects` slice of each row — never
// the whole documents, which can embed large base64 attachments.
async function findProjectOwner(projectId, token) {
  const { rows } = await pool.query(
    "SELECT user_id, data->'projects' AS projects FROM app_state WHERE data->'projects' IS NOT NULL"
  );
  for (const row of rows) {
    const project = (row.projects || []).find(
      (p) => p.id === projectId && p.shareToken && p.shareToken === token
    );
    if (project) return { userId: row.user_id, project };
  }
  return null;
}

module.exports = { pool, initDb, getState, setState, findProjectOwner, EMPTY_STATE, getGraphAuth, setGraphAuth };
