// One-off script: loads the company directory export into Postgres so the
// project "invite" picker can search it. Run again any time the export is
// refreshed — it fully replaces the table, so stale entries don't linger.
const fs = require("fs");
const path = require("path");
const { initDb, replaceEmployees, pool } = require("./db");

async function main() {
  const dataPath = path.join(__dirname, "employees-data.json");
  const raw = fs.readFileSync(dataPath, "utf8");
  const employees = JSON.parse(raw);

  await initDb();
  await replaceEmployees(employees);
  console.log(`Loaded ${employees.length} employees.`);
  await pool.end();
}

main().catch((err) => {
  console.error("Loading employees failed:", err);
  process.exit(1);
});
