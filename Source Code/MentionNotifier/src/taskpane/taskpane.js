/* ========================================================================
   MentionNotifier Engine (Local Sandbox & Native Word Comments Mapping)
   ======================================================================== */

let allUsersCache = []; // Global in-memory cache loaded from the backend relay
const emailCache = new Map();

const notifiedMentions = new Map(); // Structure: Map<commentId, Set<username>>
let cachedDocumentComments = []; // Hierarchical collection tree storage array

let editingCommentId = null; // Identifies targeted modification node
let replyingToCommentId = null; // Identifies active parent response target
let expandedCommentId = null; // Preserves open layout button drawers
let deletingCommentId = null; // Preserves active delete confirmation prompt sub-state
let highlightedSuggestionIndex = -1; // Keyboard-selected row within #suggestionsBox

// CONCURRENCY LOCK: Prevents overlapping parallel executions from double-notifying
let isScanning = false;

const CONFIG = {
  scanInterval: 8000,
  // Same-origin ("") once deployed under IISNode alongside the add-in files.
  // For local dev-server testing (webpack serve on :3000) point this at the
  // relay's own address, e.g. "http://localhost:5000".
  relayBaseUrl: "",
};

Office.onReady(function (info) {
  // Report environment/API support immediately, even before the Word checks,
  // so we can tell whether the Comments API is usable in this host build.
  // runDiagnostics(info);

  if (info.host === Office.HostType.Word) {
    // 1. Preload the real site user directory from SharePoint via the backend relay
    preloadSiteUsers();

    // 2. Attach UI event handlers to input and submission controls
    initAutocomplete();

    initManagementUI();

    // 3. Kick off native document comment scan background loop
    setTimeout(scanAllComments, 2000);
    setInterval(scanAllComments, CONFIG.scanInterval);
  }
});

// Normalizes a raw siteusers REST result array to a consistent shape
// (content.properties.<Field>.__text) used throughout the rest of the file.
function normalizeSiteUsers(results) {
  return (results || []).map((u) => ({
    content: {
      properties: {
        Id: { __text: String(u.Id) },
        Title: { __text: u.Title || "" },
        Email: { __text: u.Email || "" },
        LoginName: { __text: u.LoginName || "" },
        PrincipalType: { __text: String(u.PrincipalType) },
      },
    },
  }));
}

// Calls /api/siteusers on the backend relay (that's what actually talks to
// SharePoint over NTLM, keeps us out of browser CORS/claims issues). We just
// hand it the document's own URL and let it figure out which SharePoint site
// that maps to and whether we're even allowed to query it - nothing to
// configure here on the client side.
async function fetchSiteUsersViaRelay() {
  const docUrl = (Office.context.document && Office.context.document.url) || "";
  const query = docUrl ? `?docUrl=${encodeURIComponent(docUrl)}` : "";
  const res = await fetch(`${CONFIG.relayBaseUrl}/api/siteusers${query}`);

  if (!res.ok) {
    // backend already gives us a readable message for the common cases
    // (not a SharePoint doc, or no access) - just show that instead of a
    // generic HTTP error
    let message = `HTTP ${res.status}`;
    try {
      const e = await res.json();
      if (e && e.message) message = e.message;
      else if (e && e.error) message = `${message} — ${e.error}`;
    } catch (ignore) {
      /* response had no JSON body */
    }
    throw new Error(message);
  }

  const data = await res.json();
  return normalizeSiteUsers(data && data.users);
}

// Loads the mention list for whatever SharePoint site the current document
// lives in. If this fails for any reason we just leave allUsersCache empty
// and show whatever the backend told us (wrong site, no access, etc) - no
// hardcoded fallback list.
async function preloadSiteUsers() {
  try {
    allUsersCache = await fetchSiteUsersViaRelay();
    console.log(
      `[MentionNotifier] siteusers: loaded ${allUsersCache.length} user(s) from backend relay.`
    );
  } catch (err) {
    console.error("[MentionNotifier] siteusers: backend relay failed —", err);
    showError(
      err.message ||
        "Unable to load the user directory. @mentions and notifications are unavailable."
    );
  }
}

function initAutocomplete() {
  const input = document.getElementById("commentInput");
  const box = document.getElementById("suggestionsBox");
  const addBtn = document.getElementById("addCommentBtn");

  if (!input || !box) return;

  input.addEventListener("input", function (e) {
    const text = e.target.value;
    // \w alone doesn't cover hyphens or apostrophes, both common in real
    // login/last names ("svc-DesireAdmin", "O'Brien") - without them here,
    // typing past either character would just close the suggestion box
    const match = text.match(/@([\w.'-]*)$/);

    // don't let a bare "@" with nothing typed yet turn into an empty-string
    // search - "".includes("") is always true, so it'd match everyone in
    // allUsersCache and just show whichever 5 happened to come back first
    if (match && match[1].length > 0) {
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

  input.addEventListener("keydown", function (e) {
    if (box.style.display !== "block" || box.children.length === 0) return;

    if (e.key === "ArrowDown") {
      e.preventDefault();
      highlightedSuggestionIndex = (highlightedSuggestionIndex + 1) % box.children.length;
      applySuggestionHighlight();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      highlightedSuggestionIndex =
        (highlightedSuggestionIndex - 1 + box.children.length) % box.children.length;
      applySuggestionHighlight();
    } else if (e.key === "Enter") {
      // stop it from adding a newline in the textarea - while the
      // suggestions box is open Enter should pick the highlighted user, not
      // both at once
      e.preventDefault();
      if (highlightedSuggestionIndex >= 0) {
        box.children[highlightedSuggestionIndex].click();
      }
    } else if (e.key === "Escape") {
      box.style.display = "none";
    }
  });

  document.addEventListener("click", function (e) {
    if (e.target !== input) box.style.display = "none";
  });

  if (addBtn) {
    addBtn.onclick = handleFormSubmission;
  }
}

function applySuggestionHighlight() {
  const box = document.getElementById("suggestionsBox");
  [...box.children].forEach((child, idx) => {
    child.classList.toggle("active", idx === highlightedSuggestionIndex);
    if (idx === highlightedSuggestionIndex) {
      child.scrollIntoView({ block: "nearest" });
    }
  });
}

function initManagementUI() {
  const searchInput = document.getElementById("commentSearchInput");
  const cancelBtn = document.getElementById("cancelFormStateBtn");

  if (searchInput) {
    searchInput.addEventListener("input", renderCommentsList);
  }
  if (cancelBtn) {
    cancelBtn.onclick = resetFormState;
  }
}

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

  highlightedSuggestionIndex = 0;
  applySuggestionHighlight();
}

async function handleFormSubmission() {
  const input = document.getElementById("commentInput");
  if (!input || !input.value.trim()) {
    updateStatus("Cannot process empty text inputs.");
    return;
  }

  const textPayload = input.value.trim();

  if (editingCommentId) {
    await updateNativeCommentInWord(editingCommentId, textPayload);
  } else if (replyingToCommentId) {
    await createReplyToCommentInWord(replyingToCommentId, textPayload);
  } else {
    await createNativeCommentFromTaskpane(textPayload);
  }
}

async function createNativeCommentFromTaskpane(commentText) {
  const input = document.getElementById("commentInput");
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      if (typeof selection.insertComment === "function") {
        selection.insertComment(commentText);
        updateStatus("Native comment generated.");
        input.value = "";
      } else {
        selection.insertText(` [Comment: ${commentText}]`, "End");
        updateStatus("Inserted inline (Word API insertComment is unavailable).");
      }
      await context.sync();
    });
    await scanAllComments();
  } catch (err) {
    console.error("[MentionNotifier] Failed to write native comment control:", err);
    updateStatus("Error writing comment to document.");
  }
}

async function createReplyToCommentInWord(parentCommentId, replyText) {
  try {
    updateStatus("Posting thread reply to document...");
    await Word.run(async (context) => {
      const comments = context.document.body.getComments();
      comments.load("items/id");
      await context.sync();

      const parentComment = comments.items.find((c) => c.id === parentCommentId);
      if (parentComment) {
        parentComment.reply(replyText);
        updateStatus("Reply appended successfully.");
      } else {
        updateStatus("Parent comment thread not located.");
      }
      await context.sync();
    });

    resetFormState();
    await scanAllComments();
  } catch (err) {
    console.error("[MentionNotifier] Failed creating thread reply:", err);
    updateStatus("Failed to submit thread reply.");
  }
}

async function updateNativeCommentInWord(commentId, newText) {
  try {
    updateStatus("Updating comment on document thread...");
    await Word.run(async (context) => {
      const comments = context.document.body.getComments();
      comments.load("items/id,items/content,items/replies/items/id,items/replies/items/content");
      await context.sync();

      let matchedNode = null;
      for (const item of comments.items) {
        if (item.id === commentId) {
          matchedNode = item;
          break;
        }
        if (item.replies && item.replies.items) {
          for (const reply of item.replies.items) {
            if (reply.id === commentId) {
              matchedNode = reply;
              break;
            }
          }
        }
        if (matchedNode) break;
      }

      if (matchedNode) {
        const refRegex = /\[ref:([^\]]+)\]/;
        const match = matchedNode.content.match(refRegex);
        let absolutePayload = newText;

        if (match && !newText.includes("[ref:")) {
          absolutePayload = `${newText.trim()}  [ref:${match[1]}]`;
        }

        matchedNode.content = absolutePayload;
        updateStatus("Comment modified successfully.");
      } else {
        updateStatus("Target comment node missing.");
      }
      await context.sync();
    });

    resetFormState();
    await scanAllComments();
  } catch (err) {
    console.error("[MentionNotifier] Comment modification error:", err);
    updateStatus("Failed to save comment edits.");
  }
}

async function deleteNativeCommentFromWord(commentId) {
  try {
    updateStatus("Removing comment node from document tree...");
    await Word.run(async (context) => {
      const comments = context.document.body.getComments();
      comments.load("items/id,items/replies/items/id");
      await context.sync();

      let matchedNode = null;
      for (const item of comments.items) {
        if (item.id === commentId) {
          matchedNode = item;
          break;
        }
        if (item.replies && item.replies.items) {
          for (const reply of item.replies.items) {
            if (reply.id === commentId) {
              matchedNode = reply;
              break;
            }
          }
        }
        if (matchedNode) break;
      }

      if (matchedNode) {
        matchedNode.delete();
        updateStatus("Comment element deleted.");
      } else {
        updateStatus("Comment already missing or dropped.");
      }
      await context.sync();
    });

    if (editingCommentId === commentId || replyingToCommentId === commentId) {
      resetFormState();
    }
    await scanAllComments();
  } catch (err) {
    console.error("[MentionNotifier] Comment deletion error:", err);
    updateStatus("Failed to execute deletion process.");
  }
}

async function navigateToCommentInDoc(commentId) {
  try {
    updateStatus("Redirecting context window to comment anchor...");
    await Word.run(async (context) => {
      const comments = context.document.body.getComments();
      comments.load("items/id,items/replies/items/id");
      await context.sync();

      let targetAnchorNode = null;
      for (const item of comments.items) {
        if (item.id === commentId) {
          targetAnchorNode = item;
          break;
        }
        if (item.replies && item.replies.items) {
          const matchReply = item.replies.items.find((r) => r.id === commentId);
          if (matchReply) {
            targetAnchorNode = item;
            break;
          }
        }
      }

      if (targetAnchorNode) {
        const textRange = targetAnchorNode.getRange();
        textRange.select("Select");
        await context.sync();
        updateStatus("Viewport aligned to comment area.");
      } else {
        updateStatus("Thread coordinates not found.");
      }
    });
  } catch (err) {
    console.error("[MentionNotifier] Navigation error:", err);
    updateStatus("Navigation routing failure.");
  }
}

function enterEditMode(commentId, cleanText) {
  resetFormState();
  editingCommentId = commentId;

  const input = document.getElementById("commentInput");
  const addBtn = document.getElementById("addCommentBtn");
  const cancelBtn = document.getElementById("cancelFormStateBtn");
  const label = document.getElementById("inputPanelLabel");

  if (input) input.value = cleanText;
  if (addBtn) addBtn.innerText = "Update Comment";
  if (cancelBtn) cancelBtn.style.display = "block";
  if (label) label.innerText = "Edit Active Comment Workspace";

  updateStatus("Edit mode active.");
  renderCommentsList();
}

function enterReplyMode(parentCommentId, authorLabel, refLabel) {
  resetFormState();
  replyingToCommentId = parentCommentId;

  const addBtn = document.getElementById("addCommentBtn");
  const cancelBtn = document.getElementById("cancelFormStateBtn");
  const label = document.getElementById("inputPanelLabel");
  const input = document.getElementById("commentInput");

  if (addBtn) addBtn.innerText = "Submit Thread Reply";
  if (cancelBtn) cancelBtn.style.display = "block";
  if (label) label.innerText = `Reply to: ${authorLabel} (${refLabel})`;
  if (input) {
    input.value = "";
    input.focus();
  }

  updateStatus("Reply mode active.");
  renderCommentsList();
}

function resetFormState() {
  editingCommentId = null;
  replyingToCommentId = null;

  const input = document.getElementById("commentInput");
  const addBtn = document.getElementById("addCommentBtn");
  const cancelBtn = document.getElementById("cancelFormStateBtn");
  const label = document.getElementById("inputPanelLabel");

  if (input) input.value = "";
  if (addBtn) addBtn.innerText = "Insert Native Comment";
  if (cancelBtn) cancelBtn.style.display = "none";
  if (label) label.innerText = "Draft Comment";

  renderCommentsList();
}

function formatCommentTimestamp(dateInput) {
  if (!dateInput) return "Just now";
  const dateObj = new Date(dateInput);
  if (isNaN(dateObj.getTime())) return "Just now";

  const monthNames = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  const month = monthNames[dateObj.getMonth()];
  const day = dateObj.getDate();
  const year = dateObj.getFullYear();

  let hours = dateObj.getHours();
  const minutes = String(dateObj.getMinutes()).padStart(2, "0");

  const ampm = hours >= 12 ? "PM" : "AM";

  hours = hours % 12;
  hours = hours ? hours : 12;

  return `${month} ${day}, ${year} at ${hours}:${minutes} ${ampm}`;
}

// Works out what badge (if any) a comment should show. Only shows a ref code
// or "PENDING" if there's an actual real-user mention behind it - if the only
// "@word" in there doesn't match anyone, it should just look like a normal
// comment, not something stuck waiting forever.
function computeRefLabel(content) {
  const refMatch = content.match(/\[ref:([^\]]+)\]/);
  if (refMatch) return refMatch[1];

  const hasKnownMention = [...content.matchAll(/@([\w][\w.'-]*[\w]|[\w])/g)].some((m) =>
    isKnownUsername(m[1])
  );
  return hasKnownMention ? "PENDING" : null;
}

function buildCommentNodeHTML(comment, isReplyNode, parentRef) {
  let stateClass = "";
  if (editingCommentId === comment.id) stateClass = "active-edit";
  else if (replyingToCommentId === comment.id) stateClass = "active-reply";

  const refNumber = parentRef || computeRefLabel(comment.content);

  const displayPayloadText = comment.content.replace(/\[ref:[^\]]+\]/g, "").trim();

  // same idea here - only color an "@word" like a mention if it's an actual
  // known user, otherwise leave it as plain text
  const inlineHighlightedText = displayPayloadText.replace(
    /@([\w][\w.'-]*[\w]|[\w])/g,
    (fullMatch, username) =>
      isKnownUsername(username)
        ? `<span style="color: #0078d4; font-weight: 600;">@${username}</span>`
        : fullMatch
  );

  const isDrawerOpen = expandedCommentId === comment.id;
  const isConfirmDeleteOpen = deletingCommentId === comment.id;
  const timestampText = formatCommentTimestamp(comment.creationDate);

  return `
    <div class="comment-card ${stateClass}" data-id="${comment.id}" data-text="${encodeURIComponent(displayPayloadText)}" data-author="${encodeURIComponent(comment.author)}" data-ref="${refNumber || ""}">
      <div class="card-header-line">
        <span><span class="author-name">${comment.author}</span>${refNumber ? ` — <span class="ref-badge">${refNumber}</span>` : ""}</span>
        ${isReplyNode ? `<span style="color: #666; font-size: 10px; background: #e1dfdd; padding: 1px 4px; border-radius:2px;">Reply</span>` : ""}
      </div>
      <div class="card-body-text">${inlineHighlightedText || "<em>[No text content]</em>"}</div>
      <div class="card-timestamp-line">${timestampText}</div>
      
      <div class="card-action-drawer" style="display: ${isDrawerOpen ? "flex" : "none"};">
        <div class="standard-actions-group" style="display: ${isConfirmDeleteOpen ? "none" : "flex"}; gap: 6px;">
          ${!isReplyNode ? `<button class="action-btn btn-reply">Reply</button>` : ""}
          <button class="action-btn btn-edit">Edit</button>
          <button class="action-btn btn-delete">Delete</button>
        </div>
        <div class="inline-confirm-group" style="display: ${isConfirmDeleteOpen ? "flex" : "none"}; align-items: center; gap: 6px; font-size: 11px; color: #a80000; font-weight: 600;">
          <span>Confirm?</span>
          <button class="action-btn btn-delete action-delete-confirm">Yes</button>
          <button class="action-btn btn-cancel action-delete-cancel">No</button>
        </div>
      </div>
    </div>
  `;
}

function renderCommentsList() {
  const container = document.getElementById("commentsListContainer");
  if (!container) return;

  const searchInput = document.getElementById("commentSearchInput");
  const query = searchInput ? searchInput.value.toLowerCase().trim() : "";

  container.innerHTML = "";

  const filteredTree = cachedDocumentComments.filter((parent) => {
    const parentMatches =
      parent.content.toLowerCase().includes(query) || parent.author.toLowerCase().includes(query);
    const childMatches = parent.replies.some(
      (reply) =>
        reply.content.toLowerCase().includes(query) || reply.author.toLowerCase().includes(query)
    );
    return parentMatches || childMatches;
  });

  if (filteredTree.length === 0) {
    container.innerHTML = `<div style="font-size: 12px; color: #605e5c; padding: 20px; text-align: center;">No active comments found.</div>`;
    return;
  }

  filteredTree.forEach((thread) => {
    const threadGroupWrapper = document.createElement("div");
    threadGroupWrapper.className = "comment-card-thread-group";

    const parentHTML = buildCommentNodeHTML(thread, false, null);
    threadGroupWrapper.innerHTML = parentHTML;

    thread.replies.forEach((reply) => {
      const replyOuterFlexWrapper = document.createElement("div");
      replyOuterFlexWrapper.className = "reply-nested-wrapper";

      const parentRefNumber = computeRefLabel(thread.content);

      replyOuterFlexWrapper.innerHTML = `
        <div class="thread-line-gutter"></div>
        ${buildCommentNodeHTML(reply, true, parentRefNumber)}
      `;
      threadGroupWrapper.appendChild(replyOuterFlexWrapper);
    });

    threadGroupWrapper.querySelectorAll(".comment-card").forEach((cardNode) => {
      const cardId = cardNode.getAttribute("data-id");
      const cleanText = decodeURIComponent(cardNode.getAttribute("data-text"));
      const authorLabel = decodeURIComponent(cardNode.getAttribute("data-author"));
      const refLabel = cardNode.getAttribute("data-ref");

      cardNode.addEventListener("click", function (e) {
        if (e.target.tagName === "BUTTON") return;
        expandedCommentId = expandedCommentId === cardId ? null : cardId;
        deletingCommentId = null;
        navigateToCommentInDoc(cardId);
        renderCommentsList();
      });

      const replyBtn = cardNode.querySelector(".btn-reply");
      if (replyBtn) {
        replyBtn.onclick = function (e) {
          e.stopPropagation();
          enterReplyMode(cardId, authorLabel, refLabel);
        };
      }

      cardNode.querySelector(".btn-edit").onclick = function (e) {
        e.stopPropagation();
        enterEditMode(cardId, cleanText);
      };

      cardNode.querySelector(".btn-delete").onclick = function (e) {
        e.stopPropagation();
        deletingCommentId = cardId;
        renderCommentsList();
      };

      const cancelDel = cardNode.querySelector(".action-delete-cancel");
      if (cancelDel) {
        cancelDel.onclick = function (e) {
          e.stopPropagation();
          deletingCommentId = null;
          renderCommentsList();
        };
      }

      const confirmDel = cardNode.querySelector(".action-delete-confirm");
      if (confirmDel) {
        confirmDel.onclick = function (e) {
          e.stopPropagation();
          deleteNativeCommentFromWord(cardId);
        };
      }
    });

    container.appendChild(threadGroupWrapper);
  });
}

function extractUsernameFromLogin(loginName) {
  if (!loginName) return "";
  if (loginName.includes("\\")) return loginName.split("\\")[1];
  if (loginName.includes("|")) return loginName.split("|").pop();
  return loginName;
}

// No-op: there is no visible status line in the UI anymore (removed from
// taskpane.html). Kept as a harmless stub so the many call sites elsewhere
// in this file (comment insert/edit/reply/delete flows) don't need touching.
function updateStatus() {}
function flashStatus() {}

// Shows an error message in the #errorLog line below the action buttons.
// Pass `ms` for a transient error (auto-hides after that many milliseconds);
// omit it for a persistent error that stays until the next showError() call.
function showError(msg, ms) {
  const el = document.getElementById("errorLog");
  if (!el) return;
  el.innerText = msg;
  el.style.display = "block";
  if (ms) {
    setTimeout(() => {
      if (el.innerText === msg) el.style.display = "none";
    }, ms);
  }
}

/* ========================================================================
   Word Native Document Scanning (getComments API Approach)
   ======================================================================== */

async function scanAllComments() {
  if (Office.context.document.mode === Office.DocumentMode.ReadOnly) return;

  // Concurrency check blocks overlapping interval loops from double-notifying
  if (isScanning) return;
  isScanning = true;

  try {
    await Word.run(async (context) => {
      const comments = context.document.body.getComments();
      comments.load(
        "items/id,items/content,items/resolved,items/authorName,items/creationDate,items/replies/items/id,items/replies/items/content,items/replies/items/resolved,items/replies/items/authorName,items/replies/items/creationDate"
      );
      await context.sync();

      const workingSyncCache = [];

      for (const comment of comments.items) {
        try {
          await processComment(comment, context);

          let parsedAuthor = "Collaborator";
          if (comment.authorName && String(comment.authorName).trim().length > 0) {
            parsedAuthor = String(comment.authorName).trim();
          }

          if (!comment.resolved) {
            const parentNode = {
              id: comment.id,
              content: comment.content || "",
              author: parsedAuthor,
              creationDate: comment.creationDate,
              isReply: false,
              replies: [],
            };

            if (comment.replies && comment.replies.items) {
              for (const reply of comment.replies.items) {
                await processComment(reply, context);

                let parsedReplyAuthor = "Collaborator";
                if (reply.authorName && String(reply.authorName).trim().length > 0) {
                  parsedReplyAuthor = String(reply.authorName).trim();
                }

                if (!reply.resolved) {
                  parentNode.replies.push({
                    id: reply.id,
                    content: reply.content || "",
                    author: parsedReplyAuthor,
                    creationDate: reply.creationDate,
                    isReply: true,
                  });
                }
              }
            }

            workingSyncCache.push(parentNode);
          }
        } catch (err) {
          console.error(
            `[MentionNotifier] Error running background thread lookup element ${comment.id}:`,
            err
          );
        }
      }

      // Batch all modifications together at the very end to prevent mid-loop index corruption
      await context.sync();

      cachedDocumentComments = workingSyncCache;
      renderCommentsList();
    });
  } catch (err) {
    console.error("[MentionNotifier] Scanner daemon error stack:", err);
  } finally {
    isScanning = false; // Always clear lock context
  }
}

// Same matching logic as resolveUserEmailLocal below, just a plain yes/no
// check for "does this username exist at all". Used so a typo'd @mention
// doesn't get treated as a real one and throw a confusing "no email on
// file" error for someone who was never a real user to begin with.
function isKnownUsername(username) {
  const lowerUser = username.toLowerCase();
  return allUsersCache.some((u) => {
    const props = u.content && u.content.properties;
    if (!props) return false;

    const loginName = props.LoginName && props.LoginName.__text ? props.LoginName.__text : "";
    const title = props.Title && props.Title.__text ? props.Title.__text : "";

    const cleanLogin = extractUsernameFromLogin(loginName).toLowerCase();
    const cleanTitle = extractUsernameFromLogin(title).toLowerCase();

    return cleanLogin === lowerUser || cleanTitle === lowerUser;
  });
}

async function processComment(comment, context) {
  if (comment.resolved) return;

  const text = comment.content || "";
  const mentionRegex = /@([\w][\w.'-]*[\w]|[\w])/g;
  const matches = [...text.matchAll(mentionRegex)];
  const currentMentions = [...new Set(matches.map((m) => m[1]))].filter(isKnownUsername);

  if (currentMentions.length === 0) return;

  const refRegex = /\[ref:([^\]]+)\]/;
  const refMatch = text.match(refRegex);
  let anchor = "";
  let isNewOrMissingRef = false;

  if (refMatch) {
    anchor = `[ref:${refMatch[1]}]`;
  } else {
    const uniqueId = Math.random().toString(36).substring(2, 8).toUpperCase();
    anchor = `[ref:${uniqueId}]`;
    isNewOrMissingRef = true;
  }

  if (!notifiedMentions.has(comment.id)) {
    notifiedMentions.set(comment.id, new Set());

    if (!isNewOrMissingRef) {
      currentMentions.forEach((username) => notifiedMentions.get(comment.id).add(username));
      return;
    }
  }

  const notifiedSet = notifiedMentions.get(comment.id);
  const newMentions = currentMentions.filter((username) => !notifiedSet.has(username));

  if (newMentions.length === 0 && !isNewOrMissingRef) return;

  if (isNewOrMissingRef) {
    comment.content = text.trim() + "  " + anchor;
    // Internal context.sync() is removed from here. Handled atomically at loop completion.
  }

  const authorName = comment.authorName || "A collaborator";

  for (const username of newMentions) {
    await sendNotificationSandbox(username, anchor, text, authorName);
    notifiedSet.add(username);
  }
}

async function sendNotificationSandbox(username, anchor, originalText, authorName) {
  const email = await resolveUserEmailLocal(username);

  if (!email) {
    console.error(
      `[MentionNotifier] No email address found for "${username}" in SharePoint siteusers data — notification not sent.`
    );
    showError(`Email not sent — "${username}" has no email address on file.`, 4000);
    return;
  }

  const docUrl = Office.context.document.url;
  const cleanDocUrl = buildCleanDocUrl(docUrl);

  const commentPreview = originalText
    .replace(/@[\w.'-]+/g, "")
    .replace(/\[ref:[^\]]+\]/g, "")
    .trim()
    .substring(0, 150);

  // Updated email template highlighting our interactive taskpane search dashboard
  const emailBodyHTML = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #323130; max-width: 550px; line-height: 1.5; border: 1px solid #edebe9; padding: 20px; border-radius: 6px; background-color: #ffffff;">
      <h2 style="color: #0078d4; margin-top: 0; font-weight: 600; font-size: 18px;">Document Mention Alert</h2>
      <p>Hello,</p>
      <p>You have been mentioned by <strong>${authorName}</strong> in a document comment thread.</p>
      
      <div style="background: #faf9f8; border-left: 4px solid #0078d4; padding: 12px 16px; margin: 15px 0; font-style: italic; border-radius: 2px;">
        "${commentPreview}"
      </div>
      
      <p style="margin-top: 20px;"><strong>👉 OPEN DOCUMENT:</strong><br>
      <a href="${cleanDocUrl}" target="_blank" style="color: #0078d4; text-decoration: none; word-break: break-all;">${cleanDocUrl}</a></p>
      
      <p style="margin-top: 15px;"><strong>👉 HOW TO LOCATE THIS COMMENT:</strong><br>
      Open the document and paste the unique reference token below directly into the <strong>Mention Notifier Add-in Search Bar</strong>. You can then click the comment card to automatically scroll and center the document focus exactly on this comment thread location:</p>
      
      <div style="background: #eff6fc; display: inline-block; padding: 6px 14px; font-weight: bold; border-radius: 4px; color: #0078d4; font-family: monospace; font-size: 14px; letter-spacing: 0.5px;">
        ${anchor}
      </div>
      
      <hr style="border: none; border-top: 1px solid #edebe9; margin-top: 25px; margin-bottom: 15px;" />
      <p style="font-size: 11px; color: #8a8886; margin: 0;">This transaction notice was automatically transmitted. Please do not reply directly to this inbox.</p>
    </div>
  `;
  const subject = `Attention: You were mentioned in a comment - ${anchor}`;

  // Sends via the backend NTLM relay (SharePoint REST SendEmail utility,
  // authenticated server-side with the service account). No fallback —
  // if this fails, the user sees an error instead of a silent no-op.
  try {
    await sendEmailViaRelay(email, subject, emailBodyHTML);
    console.log(`[MentionNotifier] send-email: notification sent to ${email}.`);
  } catch (err) {
    console.error(`[MentionNotifier] send-email: failed to notify ${email} —`, err);
    showError(`Failed to send notification email to ${email}.`, 4000);
  }
}

// POST /api/send-email to the backend NTLM relay — authenticates server-side
// with the service account.
async function sendEmailViaRelay(to, subject, html) {
  const response = await fetch(`${CONFIG.relayBaseUrl}/api/send-email`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ to, subject, html }),
  });

  const result = await response.json();
  if (!response.ok || !result.success) {
    throw new Error(result.error || `HTTP error server status: ${response.status}`);
  }
}

// Resolves a tagged user's email strictly from the SharePoint siteusers data
// (allUsersCache). Returns null if the user has no email on file — callers
// must not send to a fabricated address.
async function resolveUserEmailLocal(username) {
  if (emailCache.has(username)) return emailCache.get(username);

  const lowerUser = username.toLowerCase();
  const candidates = allUsersCache.filter((u) => {
    const props = u.content && u.content.properties;
    if (!props) return false;

    const loginName = props.LoginName && props.LoginName.__text ? props.LoginName.__text : "";
    const title = props.Title && props.Title.__text ? props.Title.__text : "";

    const cleanLogin = extractUsernameFromLogin(loginName).toLowerCase();
    const cleanTitle = extractUsernameFromLogin(title).toLowerCase();

    return cleanLogin === lowerUser || cleanTitle === lowerUser;
  });

  // SharePoint can carry more than one user record for the same person (e.g.
  // a legacy non-claims duplicate from a Central Admin web app policy grant,
  // alongside the real claims-based record). Prefer whichever match actually
  // has an email on file instead of just taking the first one.
  const foundUser =
    candidates.find((u) => u.content.properties.Email && u.content.properties.Email.__text) ||
    candidates[0];

  const email =
    (foundUser &&
      foundUser.content &&
      foundUser.content.properties &&
      foundUser.content.properties.Email &&
      foundUser.content.properties.Email.__text) ||
    null;

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
