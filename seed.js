// One-off script: loads this deployment's real data into Postgres.
// A fresh clone of this repo (fresh DB, no seed run) starts empty for new users —
// only running `npm run seed` here populates *this* instance with real data.
const fs = require("fs");
const path = require("path");
const { initDb, setState, pool } = require("./db");

async function main() {
  const dataPath = path.join(__dirname, "seed-data.json");
  const raw = fs.readFileSync(dataPath, "utf8");
  const data = JSON.parse(raw);

  await initDb();
  await setState(data);
  console.log(`Seeded ${data.todos.length} todos across ${data.projects.length} projects.`);
  await pool.end();
}

main().catch((err) => {
  console.error("Seeding failed:", err);
  process.exit(1);
});
