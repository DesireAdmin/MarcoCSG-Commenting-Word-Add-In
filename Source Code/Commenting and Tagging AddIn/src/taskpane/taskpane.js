/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
console.log("VERSION-2026-06-01");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  try {
    // Add these debug lines temporarily
    console.log("Office.js loaded:", typeof Office !== "undefined");
    console.log("Word supported:", Office.context.requirements.isSetSupported("WordApi", "1.1"));
    console.log(
      "Word 1.3 supported:",
      Office.context.requirements.isSetSupported("WordApi", "1.3")
    );

    await Word.run(async function (context) {
      /* ── Insert 'Hello World!' at the end of the document ── */
      /* body.insertParagraph is WordApi 1.1 — safe for OOS     */
      var body = context.document.body;
      var newParagraph = body.insertParagraph("Hello World!", Word.InsertLocation.end);

      /* ── Apply basic formatting (all WordApi 1.1) ── */
      newParagraph.font.bold = true;
      newParagraph.font.color = "#bdf327";
      newParagraph.font.size = 14;

      /* ── Flush all queued commands to the document ── */
      await context.sync();

      /* ── Confirm success in the taskpane ── */
      showMessage("Hello World inserted successfully!", "success");
    });
  } catch (error) {
    showMessage("Error: " + error.message, "error");
    console.error(error);
  }
}

function showMessage(text, type) {
  var el = document.getElementById("message");
  if (el) {
    el.textContent = text;
    el.style.color = type === "error" ? "#9C0006" : "#375623";
  }
}
