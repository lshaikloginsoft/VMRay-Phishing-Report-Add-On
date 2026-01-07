function reportPhishing(event) {
  // Get current email
  Office.context.mailbox.item.getAllInternetHeadersAsync((result) => {
    // Send headers/content to your backend
    console.log(result.value);
  });

  event.completed();
}
