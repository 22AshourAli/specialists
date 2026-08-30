// Vercel serverless endpoint: publishes the uploaded dashboard data to the repo
// so every visitor sees the latest figures. Endpoint: POST /api/update with JSON
// { pin, data } where data is the serialized DASHBOARD_DATA object.
//
// NOTE: this writes to the deployment's own file system (/var/task), which is
// ephemeral — it surfaces the update on the active deployment instance as a
// cache-busting live patch. Full persistence still happens by committing the
// regenerated data.js via git (تحديث الداشبورد.bat).
//
// The pin is a light anti-accidental-write gate only (it is visible in the
// client bundle by design, so it is NOT a security control).

const fs = require("fs");
const path = require("path");

// Must match PUBLISH_PIN in assets/app.js.
const PUBLISH_PIN = "mokh-2026!ash";

// The deployment's writable root (Vercel) or the repo root in other runtimes.
function baseDir() {
  if (process.env.DASH_DIR) return process.env.DASH_DIR;
  if (fs.existsSync("/vercel/path0")) return "/vercel/path0";
  return path.join(__dirname, "..");
}

module.exports = async function handler(req, res) {
  if (req.method === "GET") {
    return res.status(200).json({
      ok: true,
      hint: 'POST JSON { pin, data } with the DASHBOARD_DATA payload to publish.',
    });
  }
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Method not allowed" });
  }

  let body;
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body) : (req.body || {});
  } catch (e) {
    return res.status(400).json({ ok: false, error: "Invalid JSON body" });
  }
  if (body.pin !== PUBLISH_PIN) {
    return res.status(403).json({ ok: false, error: "Wrong pin" });
  }
  const text = body.data;
  if (typeof text !== "string") {
    return res.status(400).json({ ok: false, error: "Missing data string" });
  }
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    return res.status(400).json({ ok: false, error: "Data is not valid JSON" });
  }
  if (!parsed || !parsed.meta || !parsed.punchcard) {
    return res.status(400).json({ ok: false, error: "Not a dashboard payload" });
  }

  try {
    const base = baseDir();
    const stamp = String(Date.now());

    // Atomically replace data.js so fully-loaded pages always read consistent content.
    const dataPath = path.join(base, "data.js");
    const tmp = dataPath + ".tmp-" + stamp;
    fs.writeFileSync(tmp, "window.DASHBOARD_DATA = " + text + ";\n", "utf8");
    fs.renameSync(tmp, dataPath);

    // Bump the cache-busting version stamps on index.html so browsers refetch data.js.
    const idxPath = path.join(base, "index.html");
    if (fs.existsSync(idxPath)) {
      const html = fs.readFileSync(idxPath, "utf8");
      const updated = html.replace(
        /((?:src|href)=")(assets\/(?:style\.css|app\.js)|data\.js)(\?v=\d+)?(")/g,
        "$1$2?v=" + stamp + "$4",
      );
      fs.writeFileSync(idxPath, updated, "utf8");
    }

    return res.status(200).json({ ok: true, size: text.length, stamped: stamp });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String((e && e.message) || e) });
  }
};