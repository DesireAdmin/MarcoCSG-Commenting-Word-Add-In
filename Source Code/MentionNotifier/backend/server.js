const express = require("express");
const nodemailer = require("nodemailer");
const cors = require("cors");
require("dotenv").config();

const app = express();

// Enable CORS so the add-in (served from localhost:3000 or an ngrok URL) can call this relay.
app.use(cors());
app.use(express.json());

// Fail fast if credentials are not loaded.
if (!process.env.SMTP_USER || !process.env.SMTP_PASS) {
  console.error("[CRITICAL] Missing SMTP configuration in your .env file!");
  process.exit(1);
}

// SMTP transport built from the EmailSettings config.
// Port 587 uses STARTTLS, so `secure` must be false; port 465 uses implicit TLS (secure: true).
const smtpPort = parseInt(process.env.SMTP_PORT, 10) || 587;
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: smtpPort,
  secure: process.env.SMTP_SECURE === "true" || smtpPort === 465,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
  connectionTimeout: parseInt(process.env.SMTP_TIMEOUT, 10) || 30000,
  tls: {
    rejectUnauthorized: false, // Prevents local network handshake interruptions.
  },
});

// Verify the SMTP connection at startup so credential problems surface immediately.
transporter.verify((err) => {
  if (err) {
    console.error("[SMTP] Connection verification failed:", err.message);
  } else {
    console.log("[SMTP] Server is ready to send messages.");
  }
});

// Lightweight health check.
app.get("/api/health", (req, res) => {
  res.json({ status: "ok" });
});

// Secure email route.
app.post("/api/send-email", async (req, res) => {
  const { to, subject, html } = req.body;

  if (!to || !subject || !html) {
    return res.status(400).json({ error: "Missing required payload parameters (to, subject, html)" });
  }

  try {
    console.log(`[SMTP] Attempting dispatch to: ${to}`);

    const info = await transporter.sendMail({
      from: `"Mention Notifier" <${process.env.SMTP_USER}>`,
      to,
      subject,
      html,
    });

    console.log(`[SMTP] Message sent successfully: ${info.messageId}`);
    return res.status(200).json({ success: true, messageId: info.messageId });
  } catch (error) {
    console.error("[SMTP ERROR] Mail delivery failure:", error);
    return res.status(500).json({ error: "SMTP server rejected transmission payload.", details: error.message });
  }
});

const BACKEND_PORT = process.env.PORT || 5000;
app.listen(BACKEND_PORT, () => {
  console.log("====================================================");
  console.log(` Secure Email Relay active on port: ${BACKEND_PORT}`);
  console.log("====================================================");
});
