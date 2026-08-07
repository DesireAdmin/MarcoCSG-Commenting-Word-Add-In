const express = require("express");
const cors = require("cors");
const httpntlm = require("httpntlm");
const nodemailer = require("nodemailer");
require("dotenv").config();

const app = express();

// Without these, an error httpntlm surfaces as an emitted 'error' event
// (rather than through our callback/Promise wrapper) crashes the whole
// process with no useful log line — IIS then just reports a bare 502.
// Logging here turns that into a visible stack trace instead.
process.on("uncaughtException", (err) => {
  console.error("[FATAL] Uncaught exception:", err && err.stack ? err.stack : err);
});
process.on("unhandledRejection", (err) => {
  console.error("[FATAL] Unhandled rejection:", err && err.stack ? err.stack : err);
});

// Same-origin under IISNode in production; harmless to also allow cross-origin
// callers (e.g. localhost:3000 during local dev-server testing).
app.use(cors());
app.use(express.json());

// Shared service account for the SharePoint NTLM relay. Must be granted at
// least Read access on every site collection listed in SITE_COLLECTIONS below
// — being a Windows domain account is not enough by itself.
const SP_DOMAIN = process.env.SP_DOMAIN;
const SP_USERNAME = process.env.SP_USERNAME;
const SP_PASSWORD = process.env.SP_PASSWORD;

if (!SP_DOMAIN || !SP_USERNAME || !SP_PASSWORD) {
  console.error("[CRITICAL] Missing service account configuration in your .env file!");
  process.exit(1);
}

// Known SharePoint site collections this add-in can pull a mention list from.
// The taskpane detects which one the open document belongs to (from its own
// URL) and sends back a short `site` id — never a raw URL — so this list is
// also the allowlist that request gets checked against server-side. Add an
// entry here for each new site collection; the shared service account above
// must also be granted Read access on it. `isDefault` is used whenever the
// open document's site can't be recognized (e.g. a local, non-SharePoint file).
const SITE_COLLECTIONS = [
  {
    id: "desire",
    url: process.env.SP_SITE_URL,
    isDefault: true,
  },
  // { id: "finance", url: "http://spse01:8081/sites/finance" },
];

if (!SITE_COLLECTIONS.some((s) => s.url)) {
  console.error("[CRITICAL] Missing SP_SITE_URL (or SITE_COLLECTIONS has no valid entries) in your .env file!");
  process.exit(1);
}

function resolveSiteCollection(siteId) {
  if (siteId) {
    const found = SITE_COLLECTIONS.find((s) => s.id === siteId);
    if (found) return found;
  }
  return SITE_COLLECTIONS.find((s) => s.isDefault) || SITE_COLLECTIONS[0];
}

// Email is sent via SMTP (nodemailer), not through SharePoint's SendEmail
// REST utility — that path depends on the farm's Outgoing E-Mail config and
// its legacy System.Net.Mail.SmtpClient, which has known TLS 1.2 negotiation
// problems.
//
// Some relays (e.g. an internal smart host trusted by source IP) accept
// anonymous submission, so auth is only attached when both SMTP_USER and
// SMTP_PASS are actually provided.
const smtpPort = parseInt(process.env.SMTP_PORT, 10) || 587;
const smtpUser = process.env.SMTP_USER;
const smtpPass = process.env.SMTP_PASS;
const mailTransporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: smtpPort,
  secure: process.env.SMTP_SECURE === "true" || smtpPort === 465,
  ...(smtpUser && smtpPass ? { auth: { user: smtpUser, pass: smtpPass } } : {}),
  connectionTimeout: parseInt(process.env.SMTP_TIMEOUT, 10) || 30000,
  tls: {
    rejectUnauthorized: false, // Prevents local network handshake interruptions.
  },
});

mailTransporter.verify((err) => {
  if (err) {
    console.error("[SMTP] Connection verification failed:", err.message);
  } else {
    console.log("[SMTP] Server is ready to send messages.");
  }
});

// Promise wrapper around httpntlm so the shared service-account identity
// (NTLM) is used for every call this relay makes — no cookies/CORS involved,
// since this all happens server-side.
function ntlmRequest(method, url, { headers, body } = {}) {
  return new Promise((resolve, reject) => {
    httpntlm[method](
      {
        url,
        username: SP_USERNAME,
        password: SP_PASSWORD,
        domain: SP_DOMAIN,
        headers: headers || {},
        body,
      },
      (err, res) => (err ? reject(err) : resolve(res))
    );
  });
}

// Lightweight health check.
app.get("/api/health", (req, res) => {
  res.json({ status: "ok" });
});

// Exposes the known site collections (id + url only, no credentials) so the
// taskpane can match the open document's URL against them client-side and
// tell us which one it belongs to.
app.get("/api/site-collections", (req, res) => {
  res.json({
    success: true,
    siteCollections: SITE_COLLECTIONS.map(({ id, url }) => ({ id, url })),
  });
});

// Returns the site's user directory via the SharePoint REST API, authenticated
// server-side with the NTLM service account (avoids the browser CORS/claims
// issues that block this call when made directly from the taskpane).
//
// `?site=<id>` selects which known site collection to query — the id is
// checked against SITE_COLLECTIONS above (never a client-supplied URL), so a
// caller can't redirect this service account's NTLM request anywhere else.
app.get("/api/siteusers", async (req, res) => {
  const site = resolveSiteCollection(req.query.site);

  try {
    console.log(`[NTLM] Fetching site users from ${site.url} (site="${site.id}")...`);
    const spRes = await ntlmRequest("get", `${site.url}/_api/web/siteusers`, {
      headers: { Accept: "application/json;odata=verbose" },
    });

    if (spRes.statusCode < 200 || spRes.statusCode >= 300) {
      console.error(`[NTLM] siteusers request failed: HTTP ${spRes.statusCode}`);
      return res
        .status(spRes.statusCode)
        .json({ error: `SharePoint responded ${spRes.statusCode}` });
    }

    const data = JSON.parse(spRes.body);
    const results = (data && data.d && data.d.results) || [];
    console.log(`[NTLM] siteusers OK — ${results.length} user(s) loaded from "${site.id}".`);
    return res.json({ success: true, users: results, site: site.id });
  } catch (error) {
    console.error("[NTLM] siteusers request failed:", error.message || error);
    return res.status(502).json({ error: error.message || "NTLM request to SharePoint failed" });
  }
});

// Sends the mention notification via SMTP.
app.post("/api/send-email", async (req, res) => {
  const { to, subject, html } = req.body;

  if (!to || !subject || !html) {
    return res
      .status(400)
      .json({ error: "Missing required payload parameters (to, subject, html)" });
  }

  try {
    console.log(`[SMTP] Dispatching mail to ${to}...`);
    const info = await mailTransporter.sendMail({
      from: `"Mention Notifier" <${process.env.SMTP_FROM || process.env.SMTP_USER}>`,
      to,
      subject,
      html,
    });
    console.log(`[SMTP] Message sent successfully: ${info.messageId}`);
    return res.status(200).json({ success: true, messageId: info.messageId });
  } catch (error) {
    console.error("[SMTP ERROR] Mail delivery failed:", error.message || error);
    return res.status(502).json({ error: error.message || "SMTP server rejected transmission" });
  }
});

// IISNode sets process.env.PORT to a named pipe address in production; the
// fallback below only applies when running standalone (e.g. local testing).
const BACKEND_PORT = process.env.PORT || 5000;
app.listen(BACKEND_PORT, () => {
  console.log("====================================================");
  console.log(` Mention Notifier backend active on port: ${BACKEND_PORT}`);
  console.log(` Known site collections: ${SITE_COLLECTIONS.map((s) => `${s.id} (${s.url})`).join(", ")}`);
  console.log("====================================================");
});
