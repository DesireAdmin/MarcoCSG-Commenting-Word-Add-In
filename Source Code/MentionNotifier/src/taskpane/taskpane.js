/* ========================================================================
   MentionNotifier Engine (Local Sandbox & Native Word Comments Mapping)
   ======================================================================== */

let allUsersCache = []; // Global in-memory cache loaded from local JSON feed
const emailCache = new Map();

/**
 * Stateful storage objects mapping parents and thread subsets.
 * State elements survive refresh interval clean cycles seamlessly.
 */
const notifiedMentions = new Map(); // Structure: Map<commentId, Set<username>>
let cachedDocumentComments = []; // Hierarchical collection tree storage array

let editingCommentId = null; // Identifies targeted modification node
let replyingToCommentId = null; // Identifies active parent response target
let expandedCommentId = null; // Preserves open layout button drawers
let deletingCommentId = null; // Preserves active delete confirmation prompt sub-state

// CONCURRENCY LOCK: Prevents overlapping parallel executions from double-notifying
let isScanning = false;

const CONFIG = {
  adDomain: "RegDocs365",
  scanInterval: 8000,
};

const emp = require("./emp.json");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    updateStatus("Connected to Word. Loading local sandbox directories...");

    preloadLocalUsers();
    initAutocomplete();
    initManagementUI();

    setTimeout(scanAllComments, 2000);
    setInterval(scanAllComments, CONFIG.scanInterval);
  }
});

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

function initAutocomplete() {
  const input = document.getElementById("commentInput");
  const box = document.getElementById("suggestionsBox");
  const addBtn = document.getElementById("addCommentBtn");

  if (!input || !box) return;

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

  document.addEventListener("click", function (e) {
    if (e.target !== input) box.style.display = "none";
  });

  if (addBtn) {
    addBtn.onclick = handleFormSubmission;
  }
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

function buildCommentNodeHTML(comment, isReplyNode, parentRef) {
  let stateClass = "";
  if (editingCommentId === comment.id) stateClass = "active-edit";
  else if (replyingToCommentId === comment.id) stateClass = "active-reply";

  const refRegex = /\[ref:([^\]]+)\]/;
  const refMatch = comment.content.match(refRegex);
  const refNumber = refMatch ? refMatch[1] : parentRef || "PENDING";

  const displayPayloadText = comment.content.replace(/\[ref:[^\]]+\]/g, "").trim();

  const inlineHighlightedText = displayPayloadText.replace(
    /@([\w][\w.]*[\w]|[\w])/g,
    `<span style="color: #0078d4; font-weight: 600;">@$1</span>`
  );

  const isDrawerOpen = expandedCommentId === comment.id;
  const isConfirmDeleteOpen = deletingCommentId === comment.id;
  const timestampText = formatCommentTimestamp(comment.creationDate);

  return `
    <div class="comment-card ${stateClass}" data-id="${comment.id}" data-text="${encodeURIComponent(displayPayloadText)}" data-author="${encodeURIComponent(comment.author)}" data-ref="${refNumber}">
      <div class="card-header-line">
        <span><span class="author-name">${comment.author}</span> — <span class="ref-badge">${refNumber}</span></span>
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

      const parentRefNumber = thread.content.match(/\[ref:([^\]]+)\]/)?.[1] || "PENDING";

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

function updateStatus(msg) {
  const log = document.getElementById("statusLog");
  if (log) log.innerText = `Status: ${msg}`;
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

async function processComment(comment, context) {
  if (comment.resolved) return;

  const text = comment.content || "";
  const mentionRegex = /@([\w][\w.]*[\w]|[\w])/g;
  const matches = [...text.matchAll(mentionRegex)];
  const currentMentions = [...new Set(matches.map((m) => m[1]))];

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
  const docUrl = Office.context.document.url;
  const cleanDocUrl = buildCleanDocUrl(docUrl);
  const email = await resolveUserEmailLocal(username);

  const commentPreview = originalText
    .replace(/@[\w.]+/g, "")
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

  try {
    updateStatus("Transmitting email via Secure Backend Relay...");

    // Replace with your active ngrok secure backend URL link if testing in Word Online
    const BACKEND_URL = "http://localhost:5000/api/send-email";

    const response = await fetch(BACKEND_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        to: email,
        subject: `Attention: You were mentioned in a comment - ${anchor}`,
        html: emailBodyHTML,
      }),
    });

    const result = await response.json();

    if (response.ok && result.success) {
      updateStatus(`Live email notification dispatched to: ${email}`);
    } else {
      throw new Error(result.error || `HTTP error server status: ${response.status}`);
    }
  } catch (err) {
    console.error("[MentionNotifier] Secure Backend Relay post failed:", err);
    updateStatus("Email notification delivery failure.");
  }
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
