let actionTaken = false;

Office.onReady((info) => {
  console.log("Office.onReady fired:", info);

  if (info.host !== Office.HostType.Outlook) {
    console.warn("Not running in Outlook, exiting.");
    return;
  }

  const item = Office.context.mailbox.item;
  const reportBtn = document.getElementById("reportBtn");
  const cancelBtn = document.getElementById("cancelBtn");
  const status = document.getElementById("status");

  if (!item) {
    status.innerText = "Please open an email to report phishing.";
    console.error("No item found in context.");
    return;
  }

  function disableButtons() {
    reportBtn.disabled = true;
    cancelBtn.disabled = true;
    console.log("Buttons disabled.");
  }

  reportBtn.onclick = () => {
    if (actionTaken) {
      console.log("Action already taken, ignoring click.");
      return;
    }
    actionTaken = true;
    disableButtons();

    status.innerText = "Reporting email…";
    console.log("Report button clicked, starting forwardMail.");

    forwardMailGraph(item, "username310310@gmail.com")
      .then(() => {
        status.innerText = "Email reported successfully.";
        console.log("Forward succeeded.");
        setTimeout(safeClose, 1200);
      })
      .catch((err) => {
        status.innerText = "Failed to report email.";
        console.error("Forward failed:", err);
        setTimeout(safeClose, 1200);
      });
  };

  cancelBtn.onclick = () => {
    if (actionTaken) {
      console.log("Action already taken, ignoring cancel.");
      return;
    }
    actionTaken = true;
    disableButtons();

    status.innerText = "Reporting cancelled.";
    console.log("Cancel button clicked.");
    setTimeout(safeClose, 500);
  };
});

/* Forward mail using Microsoft Graph */
async function forwardMailGraph(item, recipientEmail) {
  console.log("forwardMailGraph called with recipient:", recipientEmail);

  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get callback token:", result.error);
        return reject(result.error);
      }

      const accessToken = result.value;
      console.log("Got callback token:", accessToken);

      // REST ID of the item
      const restId = Office.context.mailbox.convertToRestId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      console.log("REST ID:", restId);

      const graphEndpoint = `https://graph.microsoft.com/v1.0/me/messages/${restId}/forward`;
      console.log("Graph endpoint:", graphEndpoint);

      try {
        const response = await fetch(graphEndpoint, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            comment: "Forwarded by Outlook add-in",
            toRecipients: [
              {
                emailAddress: {
                  address: recipientEmail,
                },
              },
            ],
          }),
        });

        if (response.ok) {
          console.log("Graph forward succeeded:", await response.text());
          resolve();
        } else {
          const errorText = await response.text();
          console.error("Graph forward failed:", errorText);
          reject(errorText);
        }
      } catch (err) {
        console.error("Graph fetch error:", err);
        reject(err);
      }
    });
  });
}

/* Safe close */
function safeClose() {
  try {
    Office.context.ui.closeContainer();
    console.log("Container closed successfully.");
  } catch (e) {
    console.error("Close failed:", e);
  }
}

window.onerror = function (msg, url, line, col, error) {
  console.error("Global error:", msg, "at", url, ":", line, ":", col, error);
  return true;
};
