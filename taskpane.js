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

    // Call  forward method
    forwardMail(item, "username310310@gmail.com", (success) => {
      if (success) {
        status.innerText = "Email reported successfully.";
      } else {
        status.innerText = "Failed to report email.";
      }
      setTimeout(safeClose, 1200);
    });
  };

  cancelBtn.onclick = () => {
    if (actionTaken) return;
    actionTaken = true;
    disableButtons();

    status.innerText = "Reporting cancelled.";
    setTimeout(safeClose, 500);
  };
});

/* Custom forward method using EWS */
function forwardMail(item, recipientEmail, callback) {
  const itemId = item.itemId;

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
              <ReferenceItemId Id="${itemId}" />
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
      console.log("Mail forwarded successfully!");
      callback(true);
    } else {
      console.error("Error forwarding mail: " + asyncResult.error.message);
      callback(false);
    }
  });
}

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
