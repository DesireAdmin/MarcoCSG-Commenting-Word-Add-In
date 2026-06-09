/* ========================================================================
   MentionNotifier Engine (Local Sandbox & Native getComments API View)
   ======================================================================== */

let allUsersCache = []; // Global in-memory cache loaded from local JSON
const processedIds = new Set();
const emailCache = new Map();

const CONFIG = {
  adDomain: "RegDocs365",
  scanInterval: 8000,
};

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    updateStatus("Connected to Word. Loading local sandbox directories...");

    // 1. Preload user profiles from local/relative json file to avoid CORS blocks
    preloadLocalUsers();

    // 2. Attach UI event handlers to input control elements
    initAutocomplete();

    // 3. Kick off native document comment scan background loop
    setTimeout(scanAllComments, 2000);
    setInterval(scanAllComments, CONFIG.scanInterval);
  }
});

// Pre-fetches real user profiles from local emp.json file asset
async function preloadLocalUsers() {
  try {
    // Relative fetch ensures no CORS errors during standard local development/testing
    const resp = await fetch("emp.json");

    if (resp.ok) {
      allUsersCache = await resp.json();
      updateStatus(`Operational (${allUsersCache.length} static profiles loaded from emp.json).`);
    } else {
      updateStatus(`Directory asset warning. Status code: ${resp.status}`);
    }
  } catch (err) {
    console.error("[MentionNotifier] Failed preloading local emp.json file:", err);
    updateStatus("Failed to access local user accounts directory asset.");
  }
}

// Binds event listeners to capture the user typing text
function initAutocomplete() {
  const input = document.getElementById("commentInput");
  const box = document.getElementById("suggestionsBox");

  if (!input || !box) return;

  input.addEventListener("input", function (e) {
    const text = e.target.value;
    const match = text.match(/@([\w.]*)$/);

    if (match) {
      const query = match[1].toLowerCase();
      const matchingUsers = allUsersCache.filter(
        (u) =>
          (u.Title && u.Title.toLowerCase().includes(query)) ||
          (u.Email && u.Email.toLowerCase().includes(query)) ||
          (u.LoginName && u.LoginName.toLowerCase().includes(query))
      );

      renderSuggestions(matchingUsers, match.index, match[0].length);
    } else {
      box.style.display = "none";
    }
  });

  document.addEventListener("click", function (e) {
    if (e.target !== input) box.style.display = "none";
  });
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
    const div = document.createElement("div");
    div.className = "suggestion-item";
    div.innerText = `${user.Title} (${user.Email})`;

    div.onclick = function () {
      const fullText = input.value;
      const cleanUsername = extractUsernameFromLogin(user.LoginName || "");

      const textBefore = fullText.substring(0, startIdx);
      const textAfter = fullText.substring(startIdx + matchLen);

      input.value = `${textBefore}@${cleanUsername}${textAfter}`;
      box.style.display = "none";
      input.focus();
    };

    box.appendChild(div);
  });
}

function extractUsernameFromLogin(loginName) {
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
      // Pull native text comment elements inside the operational document body
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

  // Generate random tracking code signature
  const uniqueId = Math.random().toString(36).substring(2, 8).toUpperCase();
  const anchor = `[ref:${uniqueId}]`;

  // Append reference code onto the comment box content native structure
  comment.content = text.trim() + "  " + anchor;
  await context.sync();

  processedIds.add(comment.id);
  const uniqueUsers = [...new Set(matches.map((m) => m[1]))];
  const authorName = comment.author || "A collaborator";

  // Dispatch individual sandbox console alerts for matched usernames
  for (const username of uniqueUsers) {
    await sendNotificationSandbox(username, anchor, text, authorName);
  }
}

// Sandbox notification: purely maps data locally and dumps body onto console logs
async function sendNotificationSandbox(username, anchor, originalText, authorName) {
  const docUrl =
    Office.context.document.url || "https://local-testing-environment/mock_document.docx";
  const cleanDocUrl = buildCleanDocUrl(docUrl);
  const email = await resolveUserEmailLocal(username);

  const commentPreview = originalText
    .replace(/@[\w.]+/g, "")
    .replace(/\[ref:[^\]]+\]/g, "")
    .trim()
    .substring(0, 150);

  // Email format printing full URL and text tag instruction explicitly
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

  // Output email structure details strictly onto context dev consoles
  console.log(emailBody);
}

// Safe resolution directly parsing global cache loaded via emp.json
async function resolveUserEmailLocal(username) {
  if (emailCache.has(username)) return emailCache.get(username);

  const lowerUser = username.toLowerCase();
  const foundUser = allUsersCache.find((u) => {
    const cleanLogin = extractUsernameFromLogin(u.LoginName || "").toLowerCase();
    const cleanTitle = (u.Title || "").toLowerCase();
    return cleanLogin === lowerUser || cleanTitle === lowerUser;
  });

  // Fallback pattern to make testing fluid even if user is missing from emp.json
  const email = foundUser ? foundUser.Email : `${username}@domain.local`;
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
