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
      
      statusLabel.innerHTML = "Thinking (Simulating AI)...";

      // --- SIMULATED AI PART (We will replace this later) ---
      // We pretend the AI read the email and wrote this text:
      const fakeAIResponse = `
        Hi there,
        
        Thanks for your email regarding: "${emailBody.substring(0, 30)}..."
        
        I have received your message and will get back to you shortly.
        
        Best regards,
        [Your Name]
      `;
      // -----------------------------------------------------

      // 3. Write the reply into the email draft
      // We use setSelectedDataAsync to insert text where the cursor is
      item.body.setSelectedDataAsync(fakeAIResponse, { coercionType: Office.CoercionType.Text }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            statusLabel.innerHTML = "Error: " + asyncResult.error.message;
        } else {
            statusLabel.innerHTML = "Draft inserted!";
        }
      });
    } else {
      statusLabel.innerHTML = "Failed to read email.";
    }
  });
}