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
    console.log("Report button clicked, starting forwardMailWithFallback.");

    forwardMailWithFallback(item, "phishing-report@vmray.com", (success) => {
      if (success) {
        status.innerText = "Email reported successfully.";
        console.log("Forward succeeded.");
      } else {
        status.innerText = "Failed to report email.";
        console.error("Forward failed.");
      }
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

/* Graph-first with EWS fallback */
async function forwardMailWithFallback(item, recipientEmail, statusCallback) {
  console.log("forwardMailWithFallback called with recipient:", recipientEmail);

  // Try Outlook REST API first (since callback token is scoped for REST)
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const accessToken = result.value;
      console.log("Got REST token");

      const restId = Office.context.mailbox.convertToRestId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      console.log("REST ID:", restId);

      const outlookRestEndpoint = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${restId}/forward`;

      try {
        const response = await fetch(outlookRestEndpoint, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            Comment: "Forwarded by Outlook add-in",
            ToRecipients: [{ EmailAddress: { Address: recipientEmail } }],
          }),
        });

        if (response.ok) {
          console.log("Outlook REST forward succeeded");
          statusCallback(true);
          return;
        } else {
          console.error("Outlook REST forward failed:", await response.text());
        }
      } catch (err) {
        console.error("Outlook REST fetch error:", err);
      }
    } else {
      console.error("Token error:", result.error);
    }

    // Fallback to EWS if REST fails
    console.log("Falling back to EWS SOAP forwarding...");
    forwardMailEws(item, recipientEmail, statusCallback);
  });
}


/* EWS SOAP fallback */
function forwardMailEws(item, recipientEmail, callback) {
  console.log("forwardMailEws called with recipient:", recipientEmail);

  const ewsId = Office.context.mailbox.convertToEwsId(item.itemId);
  if (!ewsId) {
    console.error("EWS ID conversion failed.");
    callback(false);
    return;
  }

  const ewsRequest = `
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                   xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                   xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <soap:Body>
        <CreateItem MessageDisposition="SendAndSaveCopy"
                    xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
          <Items>
            <ForwardItem>
              <ReferenceItemId Id="${ewsId}" />
              <ToRecipients>
                <t:Mailbox>
                  <t:EmailAddress>${recipientEmail}</t:EmailAddress>
                </t:Mailbox>
              </ToRecipients>
            </ForwardItem>
          </Items>
        </CreateItem>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("EWS forward succeeded");
      callback(true);
    } else {
      console.error("EWS forward failed:", asyncResult.error);
      callback(false);
    }
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
