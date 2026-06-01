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
  console.log("Office.js loaded::", typeof Office !== "undefined");
  console.log("Word supported::", Office.context.requirements.isSetSupported("WordApi", "1.1"));
  console.log("Word 1.3 supported::", Office.context.requirements.isSetSupported("WordApi", "1.3"));

  Word.run(function (context) {
    // CHECKPOINT 1 — can we even access document body?
    var body = context.document.body;
    context.load(body, "text");

    return context
      .sync()
      .then(function () {
        console.log("CHECKPOINT 1 PASSED — body.text:", body.text);

        // CHECKPOINT 2 — can we insert anything?
        body.insertText("TEST", "End");
        return context.sync();
      })
      .then(function () {
        console.log("CHECKPOINT 2 PASSED — insertText worked");
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
