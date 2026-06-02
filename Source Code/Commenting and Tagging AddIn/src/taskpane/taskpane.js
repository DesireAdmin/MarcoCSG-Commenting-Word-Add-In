console.log("DEPLOYED VERSION: 2026-06-02-v5");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // No body.load() needed at all!
      body.insertText(" Hello World!", "End");

      await context.sync();
      showMessage("Hello World inserted!", "success");
    });
  } catch (error) {
    console.error("FAILED:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", JSON.stringify(error.debugInfo));
    }
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
