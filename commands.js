Office.onReady(() => {
  console.log("VMRay Outlook Add-in ready");
});

function reportPhishing(event) {
  const item = Office.context.mailbox.item;

  // Example: get email subject and sender
  const report = {
    subject: item.subject,
    from: item.from?.emailAddress?.address || "unknown"
  };

  console.log("Reporting phishing:", report);

  // TODO: send `report` to your backend API
  // Example: fetch("https://your-backend/report", { method: "POST", body: JSON.stringify(report) })

  // Must call this to let Outlook know the function is done
  event.completed();
}
