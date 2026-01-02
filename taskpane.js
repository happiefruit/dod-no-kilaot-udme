/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // 1. Get a reference to the current email
  const item = Office.context.mailbox.item;
  const statusLabel = document.getElementById("item-subject");

  statusLabel.innerHTML = "Reading email...";

  // 2. Read the body of the email
  item.body.getAsync("text", async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;
      
      statusLabel.innerHTML = "Thinking (Talking to GPT)...";

      // --- REAL OPENAI CONNECTION ---
      // IMPORTANT: Paste your sk- key inside the quotes below
      const apiKey = "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"; 

      try {
        const response = await fetch("https://api.openai.com/v1/chat/completions", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + apiKey
          },
          body: JSON.stringify({
            model: "gpt-4o-mini", // You can also use "gpt-3.5-turbo"
            messages: [
              { role: "system", content: "You are a helpful email assistant. Draft a professional and polite reply to this email." },
              { role: "user", content: emailBody }
            ],
            temperature: 0.7
          })
        });

        if (!response.ok) {
             throw new Error("OpenAI Error: " + response.status);
        }

        const data = await response.json();
        const aiText = data.choices[0].message.content;

        // 3. Write the reply into the email draft
        item.body.setSelectedDataAsync(aiText, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            statusLabel.innerHTML = "Error: " + asyncResult.error.message;
          } else {
            statusLabel.innerHTML = "Draft inserted!";
          }
        });

      } catch (error) {
        statusLabel.innerHTML = "Failed: " + error.message;
      }
      // -----------------------------------------------------

    } else {
      statusLabel.innerHTML = "Failed to read email.";
    }
  });
}
