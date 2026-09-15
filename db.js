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

module.exports = { pool, initDb, getState, setState, EMPTY_STATE };
