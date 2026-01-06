Office.onReady(() => {
  const item = Office.context.mailbox.item;

  if (!item) {
    // No email is open
    showMessage("Open any email to get started.");
  } else {
    // Email is open → show buttons
    document.getElementById("buttons").style.display = "block";
    document.getElementById("reportBtn").onclick = reportEmail;
    document.getElementById("cancelBtn").onclick = cancelReport;
  }
});

function reportEmail() {
  const item = Office.context.mailbox.item;
  if (!item) {
    showMessage("Open any email to get started.");
    return;
  }

  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailContent = `
        Subject: ${item.subject}
        From: ${item.from.emailAddress}
        Body: ${result.value}
      `;
      forwardToSecurity(emailContent);
    } else {
      showMessage("Failed to read email content.");
    }
  });
}

function forwardToSecurity(content) {
  const recipient = "security@yourdomain.com";
  Office.context.mailbox.item.forwardAsync({
    toRecipients: [recipient],
    htmlBody: content
  }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      showMessage("Email reported successfully.");
    } else {
      showMessage("Failed to report email.");
    }
  });
}

function cancelReport() {
  showMessage("Report cancelled.");
}

function showMessage(msg) {
  document.getElementById("status").innerText = msg;
}
