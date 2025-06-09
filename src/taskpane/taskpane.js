// Assuming you have jsPDF library loaded in your HTML
// <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>

Office.onReady(() => {
  document.getElementById("exportBtn").onclick = exportEmailAndAttachments;
  document.getElementById("pickFolder").onclick = pickFolder;
});

async function exportEmailAndAttachments() {
  const item = Office.context.mailbox.item;
  const subject = sanitizeFilename(item.subject || "Untitled");
  const from = item.from && item.from.emailAddress ? item.from.emailAddress : "unknown";

  let bodyText = "Loading...";

  await new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        bodyText = result.value;
      }
      resolve();
    });
  });

  const emailContent = `
    Subject: ${subject}
    From: ${from}

    ${bodyText}
  `;

  // Convert email to PDF
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const lines = doc.splitTextToSize(emailContent, 180);
  doc.text(lines, 10, 10);
  doc.save(`${subject}.pdf`);

  // Download attachments

  const attachments = item.attachments;
  attachments.forEach((attachment) => {
    item.getAttachmentContentAsync(attachment.id, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const { content, contentFormat, name } = result.value;
        const base64 = content;
        const filename = sanitizeFilename(attachment.name || `attachment_${attachment.id}`);
        
        const saveFolderPath = localStorage.getItem("saveFolderPath");
        if (saveFolderPath) {
          saveFile(base64, saveFolderPath, filename);
        } else {
          downloadBase64File(base64, filename);
        }
      } else {
        console.error("Attachment content error:", result.error.message);
      }
    });
  });
}

function sanitizeFilename(name) {
  return name.replace(/[^a-z0-9_\-\.]/gi, "_");
}

function downloadBase64File(base64Data, filename) {
  const byteCharacters = atob(base64Data);
  const byteArrays = [];

  for (let i = 0; i < byteCharacters.length; i += 512) {
    const slice = byteCharacters.slice(i, i + 512);
    const byteNumbers = new Array(slice.length);
    for (let j = 0; j < slice.length; j++) {
      byteNumbers[j] = slice.charCodeAt(j);
    }
    const byteArray = new Uint8Array(byteNumbers);
    byteArrays.push(byteArray);
  }

  const blob = new Blob(byteArrays, { type: "application/octet-stream" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
}

async function pickFolder() {
  // try {
  const response = await fetch("http://localhost:3001/select-folder");
  const result = await response.text();

  // Office.context.mailbox.item.notificationMessages.addAsync("saveMsg", {
  //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //   message: result,
  //   icon: "Icon.16x16",
  //   persistent: false
  // });

  if (result && result.length > 0) {
    localStorage.setItem("saveFolderPath", result);
  } else {
    alert("No folder was selected.");
  }
  // } catch (err) {
  //   console.error("Error selecting folder:", err);
  //   alert("Unable to connect to the folder picker helper app.");
  // }
}

async function saveFile(base64Data, folderPath, fileName) {
  // try {
  await fetch("http://localhost:3001/save-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fileName, folderPath, base64Data }),
  });
  // } catch (err) {
  //   console.error("Error:", err);
  //   alert("Unable to connect to the folder picker helper app.");
  // }
}
