const { Pool } = require("pg");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL || "postgres://dev:dev@localhost:5432/dev",
});

// The whole app state (todos, projects, pinned notes, id counters) is kept as a single
// JSONB document. This is a personal, single-user task app — there is no multi-row
// relational access pattern to justify normalizing it, and a document keeps the server
// a thin, low-risk persistence layer around client logic that already owns the shape.
async function initDb() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_state (
      id INTEGER PRIMARY KEY DEFAULT 1 CHECK (id = 1),
      data JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    )
  `);
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
    `INSERT INTO graph_auth (id, access_token, refresh_token, expires_at, updated_at) VALUES (1, $1, $2, $3, now())
     ON CONFLICT (id) DO UPDATE SET access_token = EXCLUDED.access_token, refresh_token = EXCLUDED.refresh_token,
       expires_at = EXCLUDED.expires_at, updated_at = now()`,
    [accessToken, refreshToken, expiresAt]
  );
}

const EMPTY_STATE = {
  todos: [],
  projects: [],
  pinnedNotes: [],
  view: "all",
  sort: "priority",
  nextId: 1,
  nextSubId: 1,
};

async function getState() {
  const { rows } = await pool.query("SELECT data FROM app_state WHERE id = 1");
  if (!rows.length) return EMPTY_STATE;
  // Merge over EMPTY_STATE so older/partial saved shapes still get any new default keys.
  return { ...EMPTY_STATE, ...rows[0].data };
}

async function setState(data) {
  await pool.query(
    `INSERT INTO app_state (id, data, updated_at) VALUES (1, $1, now())
     ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data, updated_at = now()`,
    [data]
  );
}

module.exports = { pool, initDb, getState, setState, EMPTY_STATE, getGraphAuth, setGraphAuth };
