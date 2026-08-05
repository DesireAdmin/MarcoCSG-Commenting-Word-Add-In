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

const SP_SITE_URL = process.env.SP_SITE_URL;
const SP_DOMAIN = process.env.SP_DOMAIN;
const SP_USERNAME = process.env.SP_USERNAME;
const SP_PASSWORD = process.env.SP_PASSWORD;

if (!SP_SITE_URL || !SP_DOMAIN || !SP_USERNAME || !SP_PASSWORD) {
  console.error("[CRITICAL] Missing SharePoint service account configuration in your .env file!");
  process.exit(1);
}

// Email is sent via SMTP (nodemailer), not through SharePoint's SendEmail
// REST utility — that path depends on the farm's Outgoing E-Mail config and
// its legacy System.Net.Mail.SmtpClient, which has known TLS 1.2 negotiation
// problems. Only the site-users lookup below still needs SharePoint/NTLM,
// since that's genuinely SharePoint data.
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

// Promise wrapper around httpntlm so the SharePoint service-account identity
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

// Returns the site's user directory via the SharePoint REST API, authenticated
// server-side with the NTLM service account (avoids the browser CORS/claims
// issues that block this call when made directly from the taskpane).
app.get("/api/siteusers", async (req, res) => {
  try {
    console.log(`[NTLM] Fetching site users from ${SP_SITE_URL}...`);
    const spRes = await ntlmRequest("get", `${SP_SITE_URL}/_api/web/siteusers`, {
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
    console.log(`[NTLM] siteusers OK — ${results.length} user(s) loaded.`);
    return res.json({ success: true, users: results });
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
  console.log(` SharePoint NTLM Relay active on port: ${BACKEND_PORT}`);
  console.log(` Target site: ${SP_SITE_URL}`);
  console.log("====================================================");
});
