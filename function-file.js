function onMessageSendHandler(event) {
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    let recipients = asyncResult.value;
    let externalRecipients = recipients.filter(r => {
      const email = r.emailAddress.toLowerCase();
      return !(email.endsWith("@mossi.co.id") || email.endsWith("@ptsci.co.id"));
    });

    if (externalRecipients.length > 0) {
      const names = externalRecipients.map(r => r.displayName + " (" + r.emailAddress + ")").join("\n");

      Office.context.ui.displayDialogAsync("https://yourdomain.com/index.html", { height: 30, width: 40 }, function (result) {
        let dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          if (arg.message === "send") {
            dialog.close();
            event.completed({ allowEvent: true });
          } else {
            dialog.close();
            event.completed({ allowEvent: false });
          }
        });
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}
