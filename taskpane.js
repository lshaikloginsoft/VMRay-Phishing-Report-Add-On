let actionTaken = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  const item = Office.context.mailbox.item;
  const reportBtn = document.getElementById("reportBtn");
  const cancelBtn = document.getElementById("cancelBtn");
  const status = document.getElementById("status");

  if (!item) {
    status.innerText = "Please open an email to report phishing.";
    return;
  }

  function disableButtons() {
    reportBtn.disabled = true;
    cancelBtn.disabled = true;
  }

  reportBtn.onclick = () => {
    if (actionTaken) return;
    actionTaken = true;
    disableButtons();

    status.innerText = "Reporting email…";

    item.forwardAsync(
      {
        toRecipients: ["username310310@gmail.com"],
        subject: "[Phishing Report]",
        htmlBody: "<p>This email was reported as phishing.</p>"
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          status.innerText = "Email reported successfully.";
        } else {
          status.innerText = "Failed to report email.";
        }

        setTimeout(safeClose, 1200);
      }
    );
  };

  cancelBtn.onclick = () => {
    if (actionTaken) return;
    actionTaken = true;
    disableButtons();

    status.innerText = "Report cancelled.";
    setTimeout(safeClose, 500);
  };
});

/* Safe close */
function safeClose() {
  try {
    Office.context.ui.closeContainer();
  } catch (e) {
    console.error("Close failed:", e);
  }
}

window.onerror = function () {
  return true;
};
