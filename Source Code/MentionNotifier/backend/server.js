const express = require("express");
const cors = require("cors");
const httpntlm = require("httpntlm");
const nodemailer = require("nodemailer");
require("dotenv").config();

const app = express();

// small timestamped logger, makes the IISNode log files actually readable
function logInfo(...args) {
  console.log(`[${new Date().toISOString()}] [INFO]`, ...args);
}
function logError(...args) {
  console.error(`[${new Date().toISOString()}] [ERROR]`, ...args);
}

// httpntlm sometimes throws via an 'error' event instead of the callback,
// which otherwise kills the process with nothing in the logs. At least this
// way we get a stack trace before it goes down.
process.on("uncaughtException", (err) => {
  logError("[FATAL] Uncaught exception:", err && err.stack ? err.stack : err);
});
process.on("unhandledRejection", (err) => {
  logError("[FATAL] Unhandled rejection:", err && err.stack ? err.stack : err);
});

// same origin once this runs under IISNode, cors() here is really just for
// hitting the API from the localhost:3000 dev server
app.use(cors());
app.use(express.json());

// basic access log - method, path, status, how long it took. Doesn't matter
// which route handles the request, this always fires
app.use((req, res, next) => {
  const startedAt = Date.now();
  res.on("finish", () => {
    logInfo(`${req.method} ${req.originalUrl} -> ${res.statusCode} (${Date.now() - startedAt}ms)`);
  });
  next();
});

// service account for the SharePoint NTLM calls. It needs Read access
// wherever the actual documents live - if there are a lot of site
// collections, do this with a Web Application User Policy in Central Admin
// rather than granting access site by site.
const SP_DOMAIN = process.env.SP_DOMAIN;
const SP_USERNAME = process.env.SP_USERNAME;
const SP_PASSWORD = process.env.SP_PASSWORD;

if (!SP_DOMAIN || !SP_USERNAME || !SP_PASSWORD) {
  logError("[CRITICAL] Missing service account configuration in your .env file!");
  process.exit(1);
}

// hostnames of the actual SharePoint front-end servers we trust. This is the
// only allowlist we keep - not site collections, not web apps, not managed
// paths. Without it /api/siteusers would happily take any URL a caller sends
// and make our service account hit it, which is an SSRF hole - and worse
// with NTLM specifically, since an attacker-controlled server on the other
// end could capture the auth handshake. Server hostnames barely change even
// as site collections pile up, so this doesn't bring back the per-site
// maintenance problem we were trying to get away from.
const TRUSTED_SP_HOSTS = (process.env.TRUSTED_SP_HOSTS || "")
  .split(",")
  .map((h) => h.trim().toLowerCase())
  .filter(Boolean);

if (TRUSTED_SP_HOSTS.length === 0) {
  logError("[CRITICAL] Missing TRUSTED_SP_HOSTS in your .env file!");
  process.exit(1);
}

// just checks the url is http(s) and the host is one we actually trust
function isTrustedSharePointUrl(rawUrl) {
  try {
    const parsed = new URL(rawUrl);
    if (parsed.protocol !== "http:" && parsed.protocol !== "https:") return false;
    return TRUSTED_SP_HOSTS.includes(parsed.hostname.toLowerCase());
  } catch (err) {
    return false;
  }
}

// Works out which SharePoint site a document's URL actually belongs to by
// calling _api/web/siteusers at the full path first, then trimming one
// segment off the end and trying again, and so on, until something answers.
// Turns out _api/web only resolves cleanly when it's called right after an
// actual site/web URL - not a document library, not a folder - so just
// stripping the filename and hoping for the best isn't enough (that's what
// caused the 404s we saw earlier against MySiteB). Rather than hardcode
// managed paths or try to guess how deep a site sits, this just asks
// SharePoint directly at each level and stops at the first one that works.
//
// A 401/403 along the way means we DID find a real site, just no
// permission there, so we stop immediately instead of keep trimming - a
// shorter URL isn't going to fix a permissions problem.
async function findSiteUsersForUrl(rawUrl) {
  const parsed = new URL(rawUrl);
  const segments = parsed.pathname.split("/").filter(Boolean);

  // the last segment of a document URL is always the filename, and a file
  // can never be a site, so skip trying it - saves a call that's guaranteed
  // to fail (and each attempt here is really 2 requests once NTLM's
  // challenge/response is factored in)
  const lastSegment = segments[segments.length - 1] || "";
  const startIndex = /\.[a-z0-9]{2,5}$/i.test(lastSegment) ? segments.length - 1 : segments.length;

  for (let i = startIndex; i >= 0; i--) {
    const candidatePath = segments.slice(0, i).join("/");
    const candidateUrl = `${parsed.protocol}//${parsed.host}${candidatePath ? "/" + candidatePath : ""}`;

    let spRes;
    try {
      spRes = await ntlmRequest("get", `${candidateUrl}/_api/web/siteusers`, {
        headers: { Accept: "application/json;odata=verbose" },
      });
    } catch (err) {
      logError(`[NTLM] Candidate ${candidateUrl} — request failed: ${err.message || err}`);
      continue; // couldn't even reach this one, try a shorter path
    }

    logInfo(`[NTLM] Candidate ${candidateUrl} -> HTTP ${spRes.statusCode}`);

    if (spRes.statusCode === 401 || spRes.statusCode === 403) {
      const err = new Error(`Access denied by SharePoint (HTTP ${spRes.statusCode}) at ${candidateUrl}`);
      err.code = "no_access";
      throw err;
    }

    if (spRes.statusCode >= 200 && spRes.statusCode < 300) {
      const data = JSON.parse(spRes.body);
      const results = (data && data.d && data.d.results) || [];
      return { results, resolvedUrl: candidateUrl };
    }
    // anything else (404 etc) - not a site, keep trimming and try again
  }

  const err = new Error(`No SharePoint site resolved for any prefix of ${rawUrl}`);
  err.code = "not_found";
  throw err;
}

// Mail goes out straight through SMTP (Gmail in this deployment) instead of
// SharePoint's own SendEmail utility - that one depends on the farm's
// Outgoing E-Mail setup and its old System.Net.Mail.SmtpClient has known
// TLS 1.2 issues talking to Gmail. Some internal relays allow anonymous
// submission, so we only attach auth if a user/pass is actually configured.
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
    rejectUnauthorized: false, // avoids handshake failures on our network
  },
});

mailTransporter.verify((err) => {
  if (err) {
    logError("[SMTP] Connection verification failed:", err.stack || err.message);
  } else {
    logInfo("[SMTP] Server is ready to send messages.");
  }
});

// small promise wrapper around httpntlm so we can just await it like
// everything else, using the shared service account
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

// quick way to confirm the backend is even running before digging into logs
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", trustedHosts: TRUSTED_SP_HOSTS });
});

// Builds the @mention list for whichever SharePoint site the open document
// actually belongs to. `docUrl` is Office.context.document.url from the
// taskpane - we check it's on a host we trust, then let
// findSiteUsersForUrl work out the real site and fetch its users. Nothing
// here needs to know about site collections, web apps or managed paths in
// advance, so new sites just work without touching this file.
app.get("/api/siteusers", async (req, res) => {
  const docUrl = req.query.docUrl;

  if (!docUrl || !isTrustedSharePointUrl(docUrl)) {
    logError(`[NTLM] Rejected request — not a trusted SharePoint URL: ${docUrl}`);
    return res.status(400).json({
      error: "not_sharepoint",
      message: "This document isn't recognized as opened from a SharePoint site — @mentions are unavailable.",
    });
  }

  try {
    logInfo(`[NTLM] Resolving site for ${docUrl}...`);
    const { results, resolvedUrl } = await findSiteUsersForUrl(docUrl);
    logInfo(`[NTLM] siteusers OK — ${results.length} user(s) loaded from ${resolvedUrl}.`);
    return res.json({ success: true, users: results });
  } catch (error) {
    if (error.code === "no_access") {
      logError(`[NTLM] ${error.message}`);
      return res.status(403).json({
        error: "no_access",
        message: "Please contact your admin to enable @mentions for this document's location.",
      });
    }

    if (error.code === "not_found") {
      logError(`[NTLM] ${error.message}`);
      return res.status(404).json({
        error: "not_found",
        message: "No SharePoint site could be found for this document — @mentions are unavailable.",
      });
    }

    // full error goes to the log for us to debug, client just gets a plain
    // message - no point leaking the service account name or raw NTLM
    // errors to whoever's looking at the browser console
    logError("[NTLM] siteusers request failed:", error.stack || error.message || error);
    return res.status(502).json({ error: "SharePoint request failed — see backend logs for details." });
  }
});

// sends the actual mention notification email
app.post("/api/send-email", async (req, res) => {
  const { to, subject, html } = req.body;

  if (!to || !subject || !html) {
    logError(`[SMTP] Rejected request — missing required field(s) (to="${to}", subject="${subject}")`);
    return res
      .status(400)
      .json({ error: "Missing required payload parameters (to, subject, html)" });
  }

  try {
    logInfo(`[SMTP] Dispatching mail to ${to}...`);
    const info = await mailTransporter.sendMail({
      from: `"Mention Notifier" <${process.env.SMTP_FROM || process.env.SMTP_USER}>`,
      to,
      subject,
      html,
    });
    logInfo(`[SMTP] Message sent successfully: ${info.messageId}`);
    return res.status(200).json({ success: true, messageId: info.messageId });
  } catch (error) {
    logError("[SMTP ERROR] Mail delivery failed:", error.stack || error.message || error);
    return res.status(502).json({ error: "SMTP server rejected transmission — see backend logs for details." });
  }
});

// catches anything that slips past the routes above (bad JSON body, a
// thrown error we didn't wrap in try/catch, etc) so it ends up in the logs
// with a real stack trace instead of just a bare IIS 500
app.use((err, req, res, next) => {
  logError(`[EXPRESS] Unhandled error on ${req.method} ${req.originalUrl}:`, err.stack || err.message || err);
  res.status(500).json({ error: "Unexpected server error — see backend logs for details." });
});

// IISNode gives us a named pipe for PORT in production - 5000 only matters
// when running this standalone for local testing
const BACKEND_PORT = process.env.PORT || 5000;
app.listen(BACKEND_PORT, () => {
  logInfo("====================================================");
  logInfo(` Mention Notifier backend active on port: ${BACKEND_PORT}`);
  logInfo(` Trusted SharePoint hosts: ${TRUSTED_SP_HOSTS.join(", ")}`);
  logInfo("====================================================");
});
