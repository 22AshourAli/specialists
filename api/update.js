// Vercel serverless endpoint: publishes the uploaded dashboard data to the GitHub
// repository, so every visitor (via api/data + the automatic Vercel redeploy) sees
// the latest figures. Endpoint: POST /api/update with JSON { pin, data } where
// `data` is the serialized DASHBOARD_DATA object.
//
// Requires these environment variables in Vercel (set once under Project > Settings
// > Environment Variables):
//   - GH_PAT   : GitHub Personal Access Token (classic, scope: "repo"; or a
//                fine-grained token with Contents:Read/Write on this repository)
//   - GH_OWNER : repository owner (defaults to 22AshourAli)
//   - GH_REPO  : repository name    (defaults to specialists)
//   - GH_BRANCH: target branch      (defaults to master)
//
// When GH_PAT is missing the endpoint answers 503 with a clear message, keeping the
// dashboard fully functional locally.

const PUBLISH_PIN = "mokh-2026!ash"; // must match PUBLISH_PIN in assets/app.js

function config() {
  return {
    token: process.env.GH_PAT || "",
    owner: process.env.GH_OWNER || "22AshourAli",
    repo: process.env.GH_REPO || "specialists",
    branch: process.env.GH_BRANCH || "master",
  };
}

const GH_HEADERS = (token) => ({
  Authorization: "Bearer " + token,
  Accept: "application/vnd.github+json",
});

module.exports = async function handler(req, res) {
  if (req.method === "GET") {
    return res.status(200).json({ ok: true, hint: "POST { pin, data } with the DASHBOARD_DATA payload to publish it globally." });
  }
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Method not allowed" });
  }

  const c = config();
  if (!c.token) {
    return res.status(503).json({
      ok: false,
      error: "Global publishing is not configured on the server yet — set GH_PAT (and GITHUB owner/repo/branch if needed) under Vercel environment variables.",
    });
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
  const text = String(body.data || "");
  try {
    const parsed = JSON.parse(text);
    if (!parsed || !parsed.meta || !parsed.punchcard) {
      return res.status(400).json({ ok: false, error: "Not a dashboard payload" });
    }
  } catch (e) {
    return res.status(400).json({ ok: false, error: "Data is not valid JSON" });
  }

  const script = "window.DASHBOARD_DATA = " + text + ";\n";
  const base = `https://api.github.com/repos/${c.owner}/${c.repo}/contents/data.js`;

  try {
    // Fetch the current file (to keep the sha for an atomic overwrite).
    let sha = null;
    const curRes = await fetch(base, { headers: GH_HEADERS(c.token) });
    if (curRes.status === 200) {
      const cur = await curRes.json();
      sha = cur.sha || null;
    }

    const payload = {
      message: "auto-publish dashboard data (" + new Date().toISOString() + ")",
      content: Buffer.from(script, "utf8").toString("base64"),
      branch: c.branch,
    };
    if (sha) payload.sha = sha;

    const upRes = await fetch(base, {
      method: "PUT",
      headers: Object.assign({ "Content-Type": "application/json" }, GH_HEADERS(c.token)),
      body: JSON.stringify(payload),
    });
    if (!upRes.ok) {
      const errText = (await upRes.text()).slice(0, 300);
      return res.status(502).json({ ok: false, error: "GitHub refused the write (" + upRes.status + "): " + errText });
    }
    const up = await upRes.json();
    return res.status(200).json({
      ok: true,
      size: Buffer.byteLength(script, "utf8"),
      commit: up.commit && up.commit.sha,
      html_url: up.commit && up.commit.html_url,
    });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String((e && e.message) || e) });
  }
};