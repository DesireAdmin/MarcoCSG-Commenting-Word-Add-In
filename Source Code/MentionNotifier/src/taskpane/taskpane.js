/* ========================================================================
   MentionNotifier Engine (Local Sandbox & Native Word Comments Mapping)
   ======================================================================== */

let allUsersCache = []; // Global in-memory cache loaded from local JSON feed
const processedIds = new Set();
const emailCache = new Map();

const CONFIG = {
  adDomain: "RegDocs365",
  scanInterval: 8000,
};

const emp = require("./emp.json");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    updateStatus("Connected to Word. Loading local sandbox directories...");

    // 1. Preload user profiles from local/relative json file to avoid CORS blocks
    preloadLocalUsers();

    // 2. Attach UI event handlers to input and submission controls
    initAutocomplete();

    // 3. Kick off native document comment scan background loop
    setTimeout(scanAllComments, 2000);
    setInterval(scanAllComments, CONFIG.scanInterval);
  }
});

// Pre-fetches real user profiles from local emp.json file asset
async function preloadLocalUsers() {
  try {
    if (emp && emp.feed && emp.feed.entry) {
      allUsersCache = emp.feed.entry;
      updateStatus(`Operational (${allUsersCache.length} static profiles loaded from emp.json).`);
    } else {
      updateStatus("Directory structure mismatch inside emp.json.");
    }
  } catch (err) {
    console.error("[MentionNotifier] Failed preloading local emp.json file:", err);
    updateStatus("Failed to access local user accounts directory asset.");
  }
}

// Binds event listeners to capture typing patterns and the submit button action
function initAutocomplete() {
  const input = document.getElementById("commentInput");
  const box = document.getElementById("suggestionsBox");
  const addBtn = document.getElementById("addCommentBtn");

  if (!input || !box) return;

  // Handle autocomplete as user types
  input.addEventListener("input", function (e) {
    const text = e.target.value;
    const match = text.match(/@([\w.]*)$/);

    if (match) {
      const query = match[1].toLowerCase();

      const matchingUsers = allUsersCache.filter((u) => {
        const props = u.content && u.content.properties;
        if (!props) return false;

        const title = props.Title && props.Title.__text ? props.Title.__text.toLowerCase() : "";
        const email = props.Email && props.Email.__text ? props.Email.__text.toLowerCase() : "";
        const loginName =
          props.LoginName && props.LoginName.__text ? props.LoginName.__text.toLowerCase() : "";

        return title.includes(query) || email.includes(query) || loginName.includes(query);
      });

      renderSuggestions(matchingUsers, match.index, match[0].length);
    } else {
      box.style.display = "none";
    }
  });

  // Close dropdown if user clicks away
  document.addEventListener("click", function (e) {
    if (e.target !== input) box.style.display = "none";
  });

  // Attach submission click action to create the comment
  if (addBtn) {
    addBtn.onclick = createNativeCommentFromTaskpane;
  }
}

// Renders filtered matches dynamically in the taskpane frame
function renderSuggestions(users, startIdx, matchLen) {
  const box = document.getElementById("suggestionsBox");
  const input = document.getElementById("commentInput");
  box.innerHTML = "";

  if (users.length === 0) {
    box.style.display = "none";
    return;
  }

  box.style.display = "block";
  users.slice(0, 5).forEach((user) => {
    const props = user.content && user.content.properties;
    if (!props) return;

    const title = props.Title && props.Title.__text ? props.Title.__text : "Unknown User";
    const email = props.Email && props.Email.__text ? props.Email.__text : "No Email Provided";
    const loginName = props.LoginName && props.LoginName.__text ? props.LoginName.__text : "";

    const div = document.createElement("div");
    div.className = "suggestion-item";
    div.style.padding = "6px";
    div.style.cursor = "pointer";
    div.style.borderBottom = "1px solid #eee";
    div.innerText = `${title} (${email})`;

    div.onclick = function () {
      const fullText = input.value;
      const cleanUsername = extractUsernameFromLogin(loginName);

      const textBefore = fullText.substring(0, startIdx);
      const textAfter = fullText.substring(startIdx + matchLen);

      input.value = `${textBefore}@${cleanUsername}${textAfter}`;
      box.style.display = "none";
      input.focus();
    };

    box.appendChild(div);
  });
}

/**
 * NEW FUNCTIONALITY: Programmatically creates a native comment card
 * at the user's current cursor location inside the active document text body.
 */
async function createNativeCommentFromTaskpane() {
  const input = document.getElementById("commentInput");
  if (!input || !input.value.trim()) {
    updateStatus("Cannot insert an empty comment.");
    return;
  }

  const commentText = input.value.trim();

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // Safety Fallback Check: insertComment requires WordApi 1.2+
      if (typeof selection.insertComment === "function") {
        selection.insertComment(commentText);
        updateStatus("Native comment card generated.");
        input.value = ""; // Clean input panel on success
      } else {
        // If deployed to older OOS environments restricted strictly to WordApi 1.1
        selection.insertText(` [Comment: ${commentText}]`, "End");
        updateStatus("Inserted inline (Word API insertComment is unavailable).");
      }

      await context.sync();
    });
  } catch (err) {
    console.error("[MentionNotifier] Failed to write native comment control:", err);
    updateStatus("Error writing comment to document.");
  }
}

function extractUsernameFromLogin(loginName) {
  if (!loginName) return "";
  if (loginName.includes("\\")) return loginName.split("\\")[1];
  if (loginName.includes("|")) return loginName.split("|").pop();
  return loginName;
}

function updateStatus(msg) {
  const log = document.getElementById("statusLog");
  if (log) log.innerText = `Status: ${msg}`;
}

/* ========================================================================
   Word Native Document Scanning (getComments API Approach)
   ======================================================================== */

async function scanAllComments() {
  if (Office.context.document.mode === Office.DocumentMode.ReadOnly) return;

  try {
    await Word.run(async (context) => {
      const comments = context.document.body.getComments();
      comments.load("items/id,items/content,items/resolved,items/author");
      await context.sync();

      for (const comment of comments.items) {
        try {
          await processComment(comment, context);
        } catch (err) {
          console.error(`[MentionNotifier] Error processing comment ID ${comment.id}:`, err);
          processedIds.add(comment.id);
        }
      }
    });
  } catch (err) {
    console.error("[MentionNotifier] Scanner daemon error stack:", err);
  }
}

async function processComment(comment, context) {
  if (comment.resolved || processedIds.has(comment.id)) return;

  const text = comment.content || "";
  if (text.includes("[ref:")) {
    processedIds.add(comment.id);
    return;
  }

  const mentionRegex = /@([\w][\w.]*[\w]|[\w])/g;
  const matches = [...text.matchAll(mentionRegex)];
  if (matches.length === 0) return;

  const uniqueId = Math.random().toString(36).substring(2, 8).toUpperCase();
  const anchor = `[ref:${uniqueId}]`;

  comment.content = text.trim() + "  " + anchor;
  await context.sync();

  processedIds.add(comment.id);
  const uniqueUsers = [...new Set(matches.map((m) => m[1]))];
  const authorName = comment.author || "A collaborator";

  for (const username of uniqueUsers) {
    await sendNotificationSandbox(username, anchor, text, authorName);
  }
}

async function sendNotificationSandbox(username, anchor, originalText, authorName) {
  const docUrl = Office.context.document.url;
  const cleanDocUrl = buildCleanDocUrl(docUrl);
  const email = await resolveUserEmailLocal(username);

  const commentPreview = originalText
    .replace(/@[\w.]+/g, "")
    .replace(/\[ref:[^\]]+\]/g, "")
    .trim()
    .substring(0, 150);

  const emailBody = [
    "=========================================================================",
    "MOCK EMAIL NOTIFICATION DISPATCH (SANDBOX TRACE LOG)",
    "=========================================================================",
    `To: ${email}`,
    `Subject: Attention: You were mentioned in a comment - ${anchor}`,
    "-------------------------------------------------------------------------",
    "Hello,",
    "",
    `You have been mentioned by ${authorName} in a document comment.`,
    "",
    `Comment Snippet: "${commentPreview}"`,
    "",
    "👉 FULL DOCUMENT URL:",
    cleanDocUrl,
    "",
    "👉 HOW TO FIND THIS COMMENT:",
    "Please copy the reference tag below, open the document, and use the Find feature (Ctrl + F)",
    "to look for this tag inside the document or its comments pane:",
    anchor,
    "",
    "Note: This reference tag is uniquely generated to track and locate this comment transaction.",
    "=========================================================================",
  ].join("\n");

  console.log(emailBody);
}

async function resolveUserEmailLocal(username) {
  if (emailCache.has(username)) return emailCache.get(username);

  const lowerUser = username.toLowerCase();
  const foundUser = allUsersCache.find((u) => {
    const props = u.content && u.content.properties;
    if (!props) return false;

    const loginName = props.LoginName && props.LoginName.__text ? props.LoginName.__text : "";
    const title = props.Title && props.Title.__text ? props.Title.__text : "";

    const cleanLogin = extractUsernameFromLogin(loginName).toLowerCase();
    const cleanTitle = title.toLowerCase();

    return cleanLogin === lowerUser || cleanTitle === lowerUser;
  });

  let email = `${username}@domain.local`;
  if (
    foundUser &&
    foundUser.content &&
    foundUser.content.properties &&
    foundUser.content.properties.Email &&
    foundUser.content.properties.Email.__text
  ) {
    email = foundUser.content.properties.Email.__text;
  }

  emailCache.set(username, email);
  return email;
}

function buildCleanDocUrl(wopiUrl) {
  try {
    const url = new URL(wopiUrl);
    const source = url.searchParams.get("source");
    return source ? decodeURIComponent(source) : `${url.origin}${url.pathname}`;
  } catch {
    return wopiUrl;
  }
}
