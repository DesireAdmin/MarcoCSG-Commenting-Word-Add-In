console.log("DEPLOYED VERSION: CLEAR CACHE TEST");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // We are NOT loading the body.
      // We are NOT logging body.text.length.
      // We are just inserting text and syncing.
      body.insertText(" Hello Maker!", "End");

      await context.sync();

      console.log("Insert worked!");
      showMessage("Hello Maker inserted!", "success");
    });
  } catch (error) {
    console.error("FAILED:", error);
    showMessage("Error: " + error.message, "error");
  }
}
function showMessage(text, type) {
  var el = document.getElementById("message");
  if (el) {
    el.textContent = text;
    el.style.color = type === "error" ? "#9C0006" : "#375623";
  }
}
