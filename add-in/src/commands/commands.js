import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

Office.onReady(() => {
  // Office.js is ready
});

Office.actions.associate("action", action);

function action(event) {
  console.log("Action button clicked");

  const item = Office.context.mailbox.item;

  // Get the email body
  item.body.getAsync("text", function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const body = asyncResult.value;

      const subject = item.subject || "No Subject";

      // Try to get 'from' and 'to' addresses
      const from = item.from?.emailAddress?.name || item.from?.displayName || "Unknown Sender";
      const toRecipients = item.to && item.to.map(r => r.emailAddress?.name || r.displayName).join(", ") || "Unknown Recipient";
      const date = item.dateTimeCreated || new Date().toLocaleString();

      // Generate PDF
      const doc = new jsPDF();
      doc.setFontSize(14);
      doc.text("Email Details", 14, 20);

      autoTable(doc, {
        startY: 30,
        styles: { fontSize: 11, cellPadding: 3 },
        columnStyles: { 0: { fontStyle: 'bold', cellWidth: 30 }, 1: { cellWidth: 160 } },
        body: [
          ["Subject", subject],
          ["From", from],
          ["To", toRecipients],
          ["Date", date]
        ],
        theme: "grid",
        showHead: 'never'
      });

      doc.setFontSize(12);
      doc.text("Body:", 14, doc.lastAutoTable.finalY + 10);
      doc.setFontSize(11);

      // Split long text into lines
      const splitBody = doc.splitTextToSize(body, 180);
      doc.text(splitBody, 14, doc.lastAutoTable.finalY + 18);

      // Save the PDF file
      const fileName = subject.replace(/[^a-z0-9]/gi, '_').substring(0, 50) + ".pdf";
      doc.save(fileName);

      // Show notification in Outlook
      Office.context.mailbox.item.notificationMessages.addAsync("saveMsg", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Email saved as PDF successfully!",
        icon: "Icon.16x16",
        persistent: false
      });

    } else {
      console.error("Failed to get email body:", asyncResult.error.message);
    }

    event.completed();
  });
}
