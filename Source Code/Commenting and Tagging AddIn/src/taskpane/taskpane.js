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

      // Preferred syntax over context.load(body, "text")
      body.load("text");

      // Wait for the sync to completely finish before doing anything else
      await context.sync();

      // Now it is safe to read the property
      console.log("CHECKPOINT 1 PASSED body.text length:", body.text.length);

      body.insertText(" Hello Guys!", "End");

      // Sync again to execute the insert action
      await context.sync();

      console.log("CHECKPOINT 2 PASSED insert worked");
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
