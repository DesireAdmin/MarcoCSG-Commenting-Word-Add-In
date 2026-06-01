// NO imports at the top — Office and Word are globals injected by office.js
// DELETE any lines like:
// import * as OfficeHelpers from "@microsoft/office-js-helpers";
// import "office-ui-fabric-js/dist/js/fabric.js";

console.log("DEPLOYED VERSION: 2026-06-01-v4");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

function run() {
  Word.run(function (context) {
    var body = context.document.body;
    context.load(body, "text");

    return context
      .sync()
      .then(function () {
        console.log("CHECKPOINT 1 PASSED body.text length:", body.text.length);
        body.insertText(" Hello World!", "End");
        return context.sync();
      })
      .then(function () {
        console.log("CHECKPOINT 2 PASSED insert worked");
        showMessage("Hello World inserted!", "success");
      });
  }).catch(function (error) {
    console.error("FAILED:", error.code, error.message);
    console.error("debugInfo:", JSON.stringify(error.debugInfo, null, 2));
    showMessage("Error: " + error.message, "error");
  });
}

function showMessage(text, type) {
  var el = document.getElementById("message");
  if (el) {
    el.textContent = text;
    el.style.color = type === "error" ? "#9C0006" : "#375623";
  }
}
