/* ========================================================================
   MentionNotifier Engine with Live SharePoint Typeahead Autocomplete Cache
   ======================================================================== */

let allUsersCache = []; // Global in-memory cache for suggestion engine
const processedIds = new Set();
const emailCache = new Map();
let digestValue = null;
let digestExpiry = 0;

const CONFIG = {
  adDomain: "RegDocs365",
  scanInterval: 8000,
  authMode: "AD",
};

console.log(`[MentionNotifier] Initializing 3/6/2026`);

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    updateStatus("Connected to Word. Preloading directories...");

    // 1. Immediately cache all available site users for autocomplete lookup
    preloadSharePointUsers();

    // 2. Attach UI event handlers to input control elements
    initAutocomplete();

    // 3. Kick off background loop daemon scanning natively
    setTimeout(scanAllComments, 2000);
    setInterval(scanAllComments, CONFIG.scanInterval);
  }
});

// Pre-fetches real user profiles from SharePoint site collection
async function preloadSharePointUsers() {
  try {
    const docUrl = Office.context.document.url;
    if (!docUrl) return;

    const siteUrl = extractSiteUrl(docUrl);
    // PrincipalType eq 1 targets people objects only, skipping security/distribution groups
    const url = `${siteUrl}/_api/web/siteusers?$select=Title,LoginName,Email&$filter=PrincipalType eq 1 and Email ne null`;

    const resp = await fetch(url, {
      credentials: "include",
      headers: { Accept: "application/json;odata=verbose" },
    });

    if (resp.ok) {
      const data = await resp.json();
      allUsersCache = data?.d?.results || [];
      updateStatus(`Autocomplete operational (${allUsersCache.length} profiles cached).`);
    } else {
      updateStatus(`Directory initialization warning. Status code: ${resp.status}`);
    }
  } catch (err) {
    console.error("[MentionNotifier] Failed preloading directories:", err);
    updateStatus("Failed to access user accounts directory.");
  }
}

// Binds event listeners to capture the user typing text
function initAutocomplete() {
  const input = document.getElementById("commentInput");
  const box = document.getElementById("suggestionsBox");

  if (!input || !box) return;

  input.addEventListener("input", function (e) {
    const text = e.target.value;
    // Captures the text immediately following the last typed @ token
    const match = text.match(/@([\w.]*)$/);

    if (match) {
      const query = match[1].toLowerCase();
      // Match text against Title, Email, or raw Account Username properties
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

  // Hides suggestion list if user clicks away
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
  // Limit output elements to 5 items max for display layout optimization
  users.slice(0, 5).forEach((user) => {
    const div = document.createElement("div");
    div.className = "suggestion-item";
    div.innerText = `${user.Title} (${user.Email})`;

    div.onclick = function () {
      const fullText = input.value;
      const cleanUsername = extractUsernameFromLogin(user.LoginName);

      const textBefore = fullText.substring(0, startIdx);
      const textAfter = fullText.substring(startIdx + matchLen);

      // Auto-inserts the formatted username directly inside the user text frame
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
   Word Native Document Scanning Background Logic Loop
   ======================================================================== */

async function scanAllComments() {
  if (Office.context.document.mode === Office.DocumentMode.ReadOnly) return;

  try {
    await Word.run(async (context) => {
      const comments = context.document.body.getComments();
      comments.load("items/id,items/content,items/resolved");
      await context.sync();

      for (const comment of comments.items) {
        try {
          await processComment(comment, context);
        } catch (err) {
          console.error(
            `[MentionNotifier] Error running processing on comment ID ${comment.id}:`,
            err
          );
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
  const bookmarkName = `Ref_${uniqueId}`;
  const anchor = `[ref:${bookmarkName}]`;

  // Inject tracking bookmark directly over the highlighted target area text string
  const commentRange = comment.getRange();
  commentRange.insertBookmark(bookmarkName);

  // Append reference metadata tracking string to the comment content
  comment.content = text.trim() + "  " + anchor;
  await context.sync();

  processedIds.add(comment.id);
  const uniqueUsers = [...new Set(matches.map((m) => m[1]))];

  await Promise.allSettled(
    uniqueUsers.map((username) => sendNotification(username, bookmarkName, text))
  );
}

async function sendNotification(username, bookmarkName, originalText) {
  const docUrl = Office.context.document.url;
  const siteUrl = extractSiteUrl(docUrl);
  const cleanDocUrl = buildCleanDocUrl(docUrl);
  const directCommentUrl = `${cleanDocUrl}#${bookmarkName}`;

  const email = await resolveUserEmail(username);
  const digest = await getFormDigest(siteUrl);

  const commentPreview = originalText
    .replace(/@[\w.]+/g, "")
    .replace(/\[ref:[^\]]+\]/g, "")
    .trim()
    .substring(0, 150);

  const emailBody = [
    "Hello,",
    "",
    `You have been tagged in a document comment.`,
    "",
    `Comment Snippet: "${commentPreview}"`,
    "",
    "👉 CLICK THE LINK BELOW TO OPEN THE DOCUMENT DIRECTLY AT THIS COMMENT LOCATION:",
    directCommentUrl,
    "",
    "Note: Clicking this link opens Word Desktop directly to the referenced page.",
  ].join("\n");

  const emailPayload = {
    properties: {
      __metadata: { type: "SP.Utilities.EmailProperties" },
      To: { results: [email] },
      Subject: "Direct Link: You were mentioned in a document comment",
      Body: emailBody,
    },
  };

  // const sendResp = await fetch(`${siteUrl}/_api/SP.Utilities.Utility.SendEmail`, {
  //   method: "POST",
  //   credentials: "include",
  //   headers: {
  //     "Content-Type": "application/json;odata=verbose",
  //     "X-RequestDigest": digest,
  //     Accept: "application/json;odata=verbose",
  //   },
  //   body: JSON.stringify(emailPayload),
  // });

  if (!sendResp.ok) {
    throw new Error(`SharePoint Mail server rejected processing context lookup path.`);
  }
}

async function resolveUserEmail(username) {
  if (emailCache.has(username)) return emailCache.get(username);
  const siteUrl = extractSiteUrl(Office.context.document.url);
  const loginNameWithPrefix = `i:0#.w|${CONFIG.adDomain}\\${username}`;
  const loginNamePlain = `${CONFIG.adDomain}\\${username}`;

  let email = await fetchEmailByLoginName(siteUrl, loginNameWithPrefix);
  if (!email) email = await fetchEmailByLoginName(siteUrl, loginNamePlain);

  if (!email) throw new Error(`Active Directory profile resolution failed for: ${username}`);

  emailCache.set(username, email);
  return email;
}

async function fetchEmailByLoginName(siteUrl, loginName) {
  try {
    const encoded = encodeURIComponent(loginName);
    const url = `${siteUrl}/_api/web/siteusers?$filter=LoginName eq '${encoded}'&$select=Email`;

    const resp = await fetch(url, {
      credentials: "include",
      headers: { Accept: "application/json;odata=verbose", "Access-Control-Allow-Origin": "*" },
    });

    if (!resp.ok) return null;
    const data = await resp.json();
    return data?.d?.results?.[0]?.Email || null;
  } catch (err) {
    console.error(`[MentionNotifier] Error fetching email for login ${loginName}:`, err);
    return null;
  }
}

async function getFormDigest(siteUrl) {
  const now = Date.now();
  if (digestValue && now < digestExpiry) return digestValue;

  const resp = await fetch(`${siteUrl}/_api/contextinfo`, {
    method: "POST",
    credentials: "include",
    headers: { Accept: "application/json;odata=verbose" },
  });

  if (!resp.ok) throw new Error("Digest validation acquisition failed.");

  const data = await resp.json();
  digestValue = data.d.GetContextWebInformation.FormDigestValue;
  digestExpiry = now + 20 * 60 * 1000;
  return digestValue;
}

function extractSiteUrl(docUrl) {
  const layoutsIdx = docUrl.indexOf("/_layouts");
  if (layoutsIdx > -1) return docUrl.substring(0, layoutsIdx);
  return new URL(docUrl).origin;
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
