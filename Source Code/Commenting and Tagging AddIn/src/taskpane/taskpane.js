console.log("DEPLOYED VERSION: REQUIREMENT SET CHECK");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = checkRequirements;
  }
});

function checkRequirements() {
  // Comments API requires WordApi 1.4
  const hasWordApi11 = Office.context.requirements.isSetSupported("WordApi", "1.1");
  const hasWordApi14 = Office.context.requirements.isSetSupported("WordApi", "1.4");

  console.log("WordApi 1.1 Supported:", hasWordApi11);
  console.log("WordApi 1.4 (Required for Comments) Supported:", hasWordApi14);

  if (!hasWordApi14) {
    showMessage(
      "CRITICAL: Your Office Online Server version does NOT support the Native Comments API (Requires WordApi 1.4).",
      "error"
    );
  } else {
    showMessage("WordApi 1.4 is supported! The issue is a script syntax error.", "success");
  }
}

function showMessage(text, type) {
  var el = document.getElementById("message");
  if (el) {
    el.textContent = text;
    el.style.color = type === "error" ? "#9C0006" : "#375623";
  }
}
