Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    // Siap digunakan
  }
});

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  // Ambil semua penerima
  const toRecipients = item.to || [];
  const ccRecipients = item.cc || [];
  const bccRecipients = item.bcc || [];

  // Gabungkan semua penerima
  const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];

  // Ambil hanya alamat email-nya
  const allEmails = allRecipients.map(r => r.emailAddress.toLowerCase());

  // Cek jika ada email bukan dari domain mossi.co.id
  const externalEmails = allEmails.filter(email => !email.endsWith("@mossi.co.id"));

  if (externalEmails.length > 0) {
    const message = "⚠️ Warning: You are about to send an email to external recipients:\n\n" +
      externalEmails.join("\n") +
      "\n\nAre you sure you want to proceed?";

    Office.context.ui.displayDialogAsync(
      'https://abro212.github.io/outlook-addin-warning/confirm.html',
      { height: 40, width: 50, displayInIframe: true },
      function (result) {
        const dialog = result.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
          if (args.message === "yes") {
            dialog.close();
            event.completed({ allowEvent: true }); // Lanjutkan kirim
          } else {
            dialog.close();
            event.completed({ allowEvent: false }); // Batalkan kirim
          }
        });
      }
    );
  } else {
    // Tidak ada eksternal, langsung kirim
    event.completed({ allowEvent: true });
  }
}

// Perlu agar Outlook mengenali fungsi ini
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
