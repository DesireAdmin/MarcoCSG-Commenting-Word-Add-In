/* =============================================================
   MentionNotifier — OOXML Comment Reader for OOS On-Premises
   v2026-06-08-native-ooxml-cors-fix
   - Reads comments locally via body.getOoxml() (WordApi 1.1 compliant)
   - No external file fetch API or JSZip dependencies needed!
   - Writes [ref:] anchor to document body via Word.run
   - Tracks processed comments via local in-memory Set to bypass CORS errors
   ============================================================= */

console.log("[MentionNotifier] taskpane.js loaded — Native Flat OPC OOXML version (CORS Fix)");
const emp = require("./emp.json");

// ── Global state ─────────────────────────────────────────────────────────────
const processedCommentIds = new Set();
const emailCache = new Map();
let scanLoopActive = false;
let docSiteUrl = "";
let docFilePath = "";

const CONFIG = {
  adDomain: "RegDocs365",
  scanInterval: 10000,
};

// ── Entry point ──────────────────────────────────────────────────────────────
Office.onReady(function (info) {
  console.log("[MentionNotifier] Office.onReady fired. Host:", info.host);

  if (info.host !== Office.HostType.Word) {
    console.warn("[MentionNotifier] Not Word host — aborting.");
    return;
  }

  const docUrl = Office.context.document.url;
  console.log("[MentionNotifier] Document URL:", docUrl);

  docSiteUrl = extractSiteUrl(docUrl);
  docFilePath = extractServerRelativePath(docUrl);

  console.log("[MentionNotifier] Site URL:", docSiteUrl);
  console.log("[MentionNotifier] Server-relative path:", docFilePath);

  updateStatus("Connected. Initialising background scanning...");

  // First scan after 3s to let Office runtime context stabilize
  setTimeout(function () {
    console.log("[MentionNotifier] Starting first scan...");
    runScanCycle();
  }, 3000);

  // Set up periodic scanning loop execution
  setInterval(function () {
    console.log("[MentionNotifier] Interval scan triggered.");
    runScanCycle();
  }, CONFIG.scanInterval);
});

// ── URL utilities ─────────────────────────────────────────────────────────────
function extractSiteUrl(docUrl) {
  try {
    const parsed = new URL(docUrl);
    return parsed.origin;
  } catch (err) {
    console.error("[URL] extractSiteUrl failed:", err);
    return docUrl.split("/").slice(0, 3).join("/");
  }
}

function extractServerRelativePath(docUrl) {
  try {
    const parsed = new URL(docUrl);
    return decodeURIComponent(parsed.pathname);
  } catch (err) {
    console.error("[URL] extractServerRelativePath failed:", err);
    const origin = docUrl.split("/").slice(0, 3).join("/");
    return docUrl.replace(origin, "");
  }
}

// ── Scan orchestrator ─────────────────────────────────────────────────────────
async function runScanCycle() {
  if (scanLoopActive) return;
  scanLoopActive = true;

  try {
    updateStatus("Scanning document comments via native OOXML...");
    await scanAllCommentsViaOOXML();
    updateStatus("Scan complete.");
  } catch (err) {
    console.error("[Scan] Cycle error:", err);
    updateStatus("Scan error — check console.");
  } finally {
    scanLoopActive = false;
  }
}

// ── Helper: Safe isolation of XML nodes by localName ─────────────────────────
function getNodesByLocalName(parent, localName) {
  const results = [];
  const nodes = parent.getElementsByTagName("*");
  for (let i = 0; i < nodes.length; i++) {
    if (nodes[i].localName === localName) {
      results.push(nodes[i]);
    }
  }
  return results;
}

// ── Native Extraction & Inline Comment Injection Loop ──────────────────────────
async function scanAllCommentsViaOOXML() {
  // Skip execution if document is in a restricted state
  if (Office.context.document.mode === Office.DocumentMode.ReadOnly) return;

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      // getOoxml returns a Flat OPC XML package containing all parts of the document
      const ooxmlData = body.getOoxml();
      await context.sync();

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(ooxmlData.value, "text/xml");

      // Isolate the package parts safely using localName to prevent namespace drops
      const parts = getNodesByLocalName(xmlDoc, "part");
      let commentsPart = null;

      for (let i = 0; i < parts.length; i++) {
        if (parts[i].getAttribute("pkg:name") === "/word/comments.xml") {
          commentsPart = parts[i];
          break;
        }
      }

      // Exit early if the document contains no comments
      if (!commentsPart) {
        console.log("[MentionNotifier] No comment parts found in document structure.");
        return;
      }

      const commentNodes = getNodesByLocalName(commentsPart, "comment");
      console.log(`[MentionNotifier] Parsing ${commentNodes.length} active XML comments...`);

      let docIsModified = false;

      for (let j = 0; j < commentNodes.length; j++) {
        const node = commentNodes[j];
        const commentId = node.getAttribute("w:id");
        const author = node.getAttribute("w:author") || "Unknown User";

        // Extract and aggregate text strings from the comment body
        const textNodes = getNodesByLocalName(node, "t");
        let commentText = "";
        for (let k = 0; k < textNodes.length; k++) {
          commentText += textNodes[k].textContent;
        }

        const alreadyDone = processedCommentIds.has(commentId);
        const hasRef = commentText.includes("[ref:");
        const hasMention = /@[\w][\w.]*/.test(commentText);

        // Filter out comments that are already tracked or don't have mentions
        if (alreadyDone || hasRef || !hasMention) {
          continue;
        }

        console.log(
          `[MentionNotifier] Processing unhandled mention in comment ID ${commentId} by ${author}`
        );

        const mentionRegex = /@([\w][\w.]*[\w]|[\w])/g;
        const matches = [...commentText.matchAll(mentionRegex)];
        const uniqueUsers = [...new Set(matches.map((m) => m[1]))];

        if (uniqueUsers.length === 0) {
          processedCommentIds.add(commentId);
          continue;
        }

        const refId = Math.random().toString(36).substring(2, 8).toUpperCase();
        const anchor = "[ref:" + refId + "]";

        // ── INLINE XML INJECTION: Insert anchor at the end of the comment box ──
        const pNodes = getNodesByLocalName(node, "p");
        if (pNodes.length > 0) {
          const lastP = pNodes[pNodes.length - 1]; // Target the last paragraph block inside the comment box

          const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
          const xmlNamespace = "http://www.w3.org/XML/1998/namespace";

          // Create a native Word processing Run (<w:r>)
          const newR = xmlDoc.createElementNS(wNamespace, "w:r");
          // Create a native Word processing Text node (<w:t>)
          const newT = xmlDoc.createElementNS(wNamespace, "w:t");

          // Force Word to preserve the leading space before the tracking anchor string
          newT.setAttributeNS(xmlNamespace, "xml:space", "preserve");
          newT.textContent = " " + anchor;

          newR.appendChild(newT);
          lastP.appendChild(newR); // Append directly into the comment box UI block
          docIsModified = true;
        }

        // Track state locally
        processedCommentIds.add(commentId);

        const preview = commentText
          .replace(/@[\w.]+/g, "")
          .replace(/\[ref:[^\]]+\]/g, "")
          .trim()
          .substring(0, 150);

        // Dispatch notification routine asynchronously without halting the loop execution
        uniqueUsers.forEach((username) => {
          sendNotification(username, anchor, preview, author, docSiteUrl, docFilePath);
        });
      }

      // If changes were made to any comment box DOM trees, serialize and commit back into Word
      if (docIsModified) {
        const serializer = new XMLSerializer();
        const updatedOoxml = serializer.serializeToString(xmlDoc);

        // Replaces the complete structural package including modified comment blocks
        body.insertOoxml(updatedOoxml, "Replace");
        await context.sync();
        console.log(
          "[MentionNotifier] OOXML updated with comment box inline anchors successfully."
        );
      }
    });
  } catch (err) {
    console.error("[MentionNotifier] Critical failure in OOXML Extraction/Injection Loop:", err);
  }
}

// ── Email notification ────────────────────────────────────────────────────────
async function sendNotification(
  username,
  anchor,
  preview,
  commentAuthor,
  siteUrl,
  serverRelativePath
) {
  let email;
  try {
    email = await resolveUserEmail(username, siteUrl);
  } catch (err) {
    console.error("[Email] Could not resolve email for", username, "—", err.message);
    return;
  }

  const docUrl = Office.context.document.url;
  const anchorUrl = docUrl + "#" + anchor;

  const emailBody = [
    "Hello,",
    "",
    "You were mentioned in a document comment by " + commentAuthor + ".",
    "",
    'Comment: "' + preview + '"',
    "",
    "Open the document here:",
    anchorUrl,
  ].join("\n");

  console.log("[Email] SEND SKIPPED (commented out). Would send to:", email);
}

// ── User resolution (Using your offline emp.json payload) ─────────────────────
async function resolveUserEmail(username, siteUrl) {
  if (emailCache.has(username)) return emailCache.get(username);

  const candidates = [
    "i:0#.w|" + CONFIG.adDomain + "\\" + username,
    CONFIG.adDomain + "\\" + username,
  ];

  for (const loginName of candidates) {
    const email = await fetchEmailByLoginName(siteUrl, loginName.toLowerCase());
    if (email) {
      emailCache.set(username, email);
      return email;
    }
  }

  throw new Error("[UserResolve] No email found for: " + username);
}

async function fetchEmailByLoginName(siteUrl, loginName) {
  try {
    const Email = emp.feed.entry.find(
      (e) => e.content.properties.LoginName.__text === "i:0#.w|" + loginName
    )?.content.properties.Email.__text;
    return Email || null;
  } catch (err) {
    return null;
  }
}

function updateStatus(msg) {
  console.log("[Status]", msg);
  const el = document.getElementById("statusLog");
  if (el) el.innerText = "Status: " + msg;
}
