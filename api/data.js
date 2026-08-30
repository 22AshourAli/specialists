// Vercel serverless endpoint: serves the LATEST published data.js straight from the
// GitHub repository, so visitors always see the newest globally-published figures
// (works even before the automatic Vercel redeploy finishes). GET /api/data
//
// Requires GH_PAT / GH_OWNER / GH_REPO (same env as api/update.js). Without a token
// it falls back to the locally bundled data.js from the current deployment, which is
// correct because that deployment was built from the same repo.

const fs = require("fs");
const path = require("path");

function config() {
  return {
    token: process.env.GH_PAT || "",
    owner: process.env.GH_OWNER || "22AshourAli",
    repo: process.env.GH_REPO || "specialists",
    branch: process.env.GH_BRANCH || "master",
  };
}

function baseDir() {
  if (process.env.DASH_DIR) return process.env.DASH_DIR;
  if (fs.existsSync("/vercel/path0")) return "/vercel/path0";
  return path.join(__dirname, "..");
}

module.exports = async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store, max-age=0");
  res.setHeader("Content-Type", "application/javascript; charset=utf-8");

  const c = config();
  if (c.token) {
    try {
      // raw.githubusercontent is fast and CDN-backed; the trailing query is a no-op that
      // keeps the fresh header distinct from cached responses.
      const raw = await fetch(`https://raw.githubusercontent.com/${c.owner}/${c.repo}/${c.branch}/data.js?ts=${Date.now()}`);
      if (raw.ok) {
        return res.status(200).send(await raw.text());
      }
    } catch (e) {
      // fall through to the bundled copy below
    }
  }

  try {
    fs.createReadStream(path.join(baseDir(), "data.js")).pipe(res);
  } catch (e) {
    return res.status(404).json({ ok: false, error: "data.js not found" });
  }
};