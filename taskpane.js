Office.onReady(() => {
  const item = Office.context.mailbox.item;
  const reportBtn = document.getElementById("reportBtn");
  const cancelBtn = document.getElementById("cancelBtn");
  const status = document.getElementById("status");

  if (!item) {
    status.innerText = "Open any email to get started.";
    return;
  }

  let actionTaken = false; // prevents multiple clicks

  reportBtn.onclick = () => {
    if (actionTaken) return;
    actionTaken = true;

    // Disable both buttons
    reportBtn.disabled = true;
    cancelBtn.disabled = true;

    const recipient = "username310310@gmail.com";

    item.forwardAsync(
      { toRecipients: [recipient] },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          status.innerText = "Email reported successfully.";
        } else {
          status.innerText = "Failed to report email.";
        }

        // Close task pane after a short delay
        setTimeout(() => {
          Office.context.ui.closeContainer();
        }, 800);
      }
    );
  };

  cancelBtn.onclick = () => {
    if (actionTaken) return;
    actionTaken = true;

    // Disable both buttons
    reportBtn.disabled = true;
    cancelBtn.disabled = true;

    status.innerText = "Report cancelled.";

    // Close task pane immediately
    Office.context.ui.closeContainer();
  };
});
