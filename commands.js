// Function triggered by the ribbon button
function openRedirectDialog(event) {
  const dialogOptions = { height: 30, width: 20, displayInIframe: true };

  Office.context.ui.displayDialogAsync(
    'https://192.168.0.131:3000/dialog.html', 
    dialogOptions,
    function (asyncResult) {
      const dialog = asyncResult.value;
      
      // Listen for messages sent back from the popup
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const targetEmail = arg.message;
        redirectMail(targetEmail);
        dialog.close();
        event.completed();
      });
    }
  );
}

// Function to actually forward/redirect the mail
function redirectMail(recipient) {
  const item = Office.context.mailbox.item;

  // Use the 'forward' method to simulate redirection
  item.displayReplyAllForm(
    {
      'toRecipients': [{ emailAddress: recipient }],
      'htmlBody': "This message has been redirected."
    }
  );
}