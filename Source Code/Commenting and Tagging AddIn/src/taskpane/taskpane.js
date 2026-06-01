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

function run() {
  // Add these debug lines temporarily
  console.log("Office.js loaded:", typeof Office !== "undefined");
  console.log("Word supported:", Office.context.requirements.isSetSupported("WordApi", "1.1"));
  console.log("Word 1.3 supported:", Office.context.requirements.isSetSupported("WordApi", "1.3"));

  Word.run(function (context) {
    var body = context.document.body;

    // Step 1: Insert the paragraph — don't use the return value
    body.insertParagraph("Hello World!", "End");

    // Step 2: Sync the insertion first
    return context.sync().then(function () {
      // Step 3: Now fetch the last paragraph separately
      var paragraphs = body.paragraphs;
      context.load(paragraphs, "items");

      return context.sync().then(function () {
        // Step 4: Get the last paragraph and format it
        var lastParagraph = paragraphs.items[paragraphs.items.length - 1];
        context.load(lastParagraph, "font");

        lastParagraph.font.bold = true;
        lastParagraph.font.color = "#50cf15";
        lastParagraph.font.size = 14;

        return context.sync().then(function () {
          showMessage("Hello World inserted successfully!", "success");
        });
      });
    });
  }).catch(function (error) {
    showMessage("Error: " + error.message, "error");
    console.error("Full error object:", error);
    // Log the error code specifically — this tells us exactly which API failed
    console.error("Error code:", error.code);
    console.error("Error debug info:", error.debugInfo);
  });
}

function showMessage(text, type) {
  var el = document.getElementById("message");
  if (el) {
    el.textContent = text;
    el.style.color = type === "error" ? "#9C0006" : "#375623";
  }
}
