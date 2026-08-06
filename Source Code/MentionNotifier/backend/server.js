const express = require("express");
const cors = require("cors");
const httpntlm = require("httpntlm");
const { Client: LdapClient } = require("ldapts");
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

// One shared service account authenticates both the SharePoint NTLM relay and
// the Active Directory LDAP bind below — deliberately not two separate
// identities, since both already live in this same .env file and splitting
// them wouldn't reduce the blast radius of a leak. Keep this account scoped
// to read-only access (SharePoint site Read + default AD user-read rights),
// never the domain Administrator.
const SP_SITE_URL = process.env.SP_SITE_URL;
const SP_DOMAIN = process.env.SP_DOMAIN;
const SP_USERNAME = process.env.SP_USERNAME;
const SP_PASSWORD = process.env.SP_PASSWORD;

if (!SP_SITE_URL || !SP_DOMAIN || !SP_USERNAME || !SP_PASSWORD) {
  console.error("[CRITICAL] Missing service account configuration in your .env file!");
  process.exit(1);
}

// LDAP connection target — not credentials (those are the SP_* vars above).
// AD accepts DOMAIN\username as a simple-bind identity, so the bind DN is
// derived from the same SP_DOMAIN/SP_USERNAME rather than duplicated in .env.
const LDAP_URL = process.env.LDAP_URL;
const LDAP_BASE_DN = process.env.LDAP_BASE_DN;

if (!LDAP_URL || !LDAP_BASE_DN) {
  console.error("[CRITICAL] Missing Active Directory (LDAP) configuration in your .env file!");
  process.exit(1);
}

const LDAP_BIND_DN = `${SP_DOMAIN}\\${SP_USERNAME}`;

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
// since this all happens server-side. Only used by the ALTERNATE (commented
// out) SharePoint block in /api/siteusers below.
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

// (&(objectCategory=person)(objectClass=user)...) is the standard AD idiom for
// "real user accounts" as opposed to computers/groups. The userAccountControl
// bitmask check excludes disabled accounts, and requiring `mail` weeds out
// built-in/service accounts (krbtgt, sp_farm, etc.) that have no mailbox —
// those can't sensibly be @mentioned or emailed anyway.
const AD_USER_FILTER =
  "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";

async function searchAdUsers() {
  const client = new LdapClient({ url: LDAP_URL });

  try {
    await client.bind(LDAP_BIND_DN, SP_PASSWORD);

    const { searchEntries } = await client.search(LDAP_BASE_DN, {
      scope: "sub",
      filter: AD_USER_FILTER,
      attributes: ["sAMAccountName", "displayName", "mail"],
      paged: true,
    });

    // ldapts returns a requested-but-empty attribute as [] rather than
    // omitting it, so a plain `|| null` check doesn't catch that case.
    const singleValue = (value) => (Array.isArray(value) ? value[0] || null : value || null);

    // Reshaped to match the SharePoint siteusers REST fields the frontend's
    // normalizeSiteUsers() already expects (Id/Title/Email/LoginName/
    // PrincipalType) — this is what lets the ACTIVE/ALTERNATE blocks in
    // /api/siteusers below swap without any frontend change.
    return searchEntries.map((entry, index) => {
      const samAccountName = singleValue(entry.sAMAccountName);
      return {
        Id: index + 1,
        Title: singleValue(entry.displayName) || samAccountName,
        Email: singleValue(entry.mail),
        LoginName: `${SP_DOMAIN}\\${samAccountName}`,
        PrincipalType: 1,
      };
    });
  } finally {
    await client.unbind();
  }
}

// Lightweight health check.
app.get("/api/health", (req, res) => {
  res.json({ status: "ok" });
});

// ============================================================================
// Mention-list source for /api/siteusers — ACTIVE: Active Directory (LDAP).
//
// To switch back to the SharePoint-site-scoped version instead: comment out
// this block and uncomment the ALTERNATE block further below. The response
// shape (`{ success, users }`, with Id/Title/Email/LoginName/PrincipalType
// per user) is identical either way, so nothing else needs to change.
// ============================================================================
app.get("/api/siteusers", async (req, res) => {
  try {
    console.log(`[LDAP] Fetching users from ${LDAP_BASE_DN}...`);
    const users = await searchAdUsers();
    console.log(`[LDAP] AD user search OK — ${users.length} user(s) loaded.`);
    return res.json({ success: true, users });
  } catch (error) {
    console.error("[LDAP] AD user search failed:", error.message || error);
    return res.status(502).json({ error: error.message || "LDAP request to Active Directory failed" });
  }
});

// ============================================================================
// Mention-list source for /api/siteusers — ALTERNATE (commented out):
// SharePoint site users, scoped to SP_SITE_URL's own site collection only.
//
// NOTE: if you re-enable this, the shared service account (SP_USERNAME) must
// actually have Read access granted on SP_SITE_URL's site collection — the
// account currently configured may not have that yet.
// ============================================================================
// app.get("/api/siteusers", async (req, res) => {
//   try {
//     console.log(`[NTLM] Fetching site users from ${SP_SITE_URL}...`);
//     const spRes = await ntlmRequest("get", `${SP_SITE_URL}/_api/web/siteusers`, {
//       headers: { Accept: "application/json;odata=verbose" },
//     });
//
//     if (spRes.statusCode < 200 || spRes.statusCode >= 300) {
//       console.error(`[NTLM] siteusers request failed: HTTP ${spRes.statusCode}`);
//       return res
//         .status(spRes.statusCode)
//         .json({ error: `SharePoint responded ${spRes.statusCode}` });
//     }
//
//     const data = JSON.parse(spRes.body);
//     const results = (data && data.d && data.d.results) || [];
//     console.log(`[NTLM] siteusers OK — ${results.length} user(s) loaded.`);
//     return res.json({ success: true, users: results });
//   } catch (error) {
//     console.error("[NTLM] siteusers request failed:", error.message || error);
//     return res.status(502).json({ error: error.message || "NTLM request to SharePoint failed" });
//   }
// });

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
  console.log(` Mention-list source: Active Directory (${LDAP_BASE_DN})`);
  console.log(` SharePoint site (unused unless ALTERNATE block is active): ${SP_SITE_URL}`);
  console.log("====================================================");
});
