Office.onReady(() => {
  const item = Office.context.mailbox.item;

  if (!item) {
    document.getElementById("status").innerText = "Open any email to get started.";
    return;
  }

  document.getElementById("reportBtn").onclick = () => {
    // Forward the current mail to your backend recipient
    const recipient = "username310310@gmail.com"; // configure here
    item.forwardAsync(
      { toRecipients: [recipient] },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById("status").innerText = "Email forwarded successfully.";
        } else {
          document.getElementById("status").innerText = "Failed to forward email.";
        }
      }
    );
  };

  document.getElementById("cancelBtn").onclick = () => {
    document.getElementById("status").innerText = "Report cancelled.";
  };
});
