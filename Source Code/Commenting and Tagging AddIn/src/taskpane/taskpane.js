console.log("DEPLOYED VERSION: COMMON API TEST v1");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

function run() {
  // Using the Office Common API for on-premises OOS compatibility
  Office.context.document.setSelectedDataAsync(
    " Hello World!",
    { coercionType: Office.CoercionType.Text },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("FAILED:", asyncResult.error.message);
        showMessage("Error: " + asyncResult.error.message, "error");
      } else {
        console.log("Insert worked using Common API!");
        showMessage("Hello World inserted!", "success");
      }
    }
  );
}

function showMessage(text, type) {
  var el = document.getElementById("message");
  if (el) {
    el.textContent = text;
    el.style.color = type === "error" ? "#9C0006" : "#375623";
  }
}
