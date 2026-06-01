/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  // Add these debug lines temporarily
  console.log("Office.js loaded:", typeof Office !== "undefined");
  console.log("Word supported:", Office.context.requirements.isSetSupported("WordApi", "1.1"));
  console.log("Word 1.3 supported:", Office.context.requirements.isSetSupported("WordApi", "1.3"));

  try {
    await Word.run(async function (context) {
      /* ── Insert 'Hello World!' at the end of the document ── */
      /* body.insertParagraph is WordApi 1.1 — safe for OOS     */
      var body = context.document.body;
      var newParagraph = body.insertParagraph("Hello World!", "End");

      /* ── Apply basic formatting (all WordApi 1.1) ── */
      newParagraph.font.bold = true;
      newParagraph.font.color = "#cf8e15";
      newParagraph.font.size = 14;

      /* ── Flush all queued commands to the document ── */
      await context.sync();

      /* ── Confirm success in the taskpane ── */
      showMessage("Hello World inserted successfully!", "success");
    });
  } catch (error) {
    showMessage("Error: " + error.message, "error");
    console.error("Full error object:", error);
    // Log the error code specifically — this tells us exactly which API failed
    console.error("Error code:", error.code);
    console.error("Error debug info:", error.debugInfo);
  }
}

function showMessage(text, type) {
  var el = document.getElementById("message");
  if (el) {
    el.textContent = text;
    el.style.color = type === "error" ? "#9C0006" : "#375623";
  }
}
