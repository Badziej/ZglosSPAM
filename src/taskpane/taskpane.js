Office.onReady(function (info) {
  // Office is ready
  if (info.host === Office.HostType.Outlook) {
    // Initialize the "Report SPAM" button
    const button = document.getElementById("reportSpamButton");
    button.onclick = reportSpam;
  }
});

// Function to handle the "Report SPAM" button click
async function reportSpam() {
  try {
    // Get the selected message
    const item = Office.context.mailbox.item;
    
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      // Create a new email message
      const newMessage = Office.context.mailbox.createItem(Office.MailboxEnums.ItemType.Message);

      // Add the original email as an attachment (EML file)
      const attachment = {
        type: Office.MailboxEnums.AttachmentType.File,
        name: `${item.subject}.eml`,
        url: await getEMLAttachmentUrl(item)
      };
      newMessage.attachments.add(attachment);

      // Set up the new message properties
      newMessage.subject = "[SPAM Report] " + item.subject;
      newMessage.to = ["casb@ratels.pl"];
      newMessage.body = "Please find the attached message for spam report.";

      // Send the message
      await newMessage.sendAsync();
      console.log("SPAM report email sent.");

      // Move the original message to Trash
      await moveToTrash(item);
    } else {
      console.error("No email message selected.");
    }
  } catch (error) {
    console.error("Error reporting spam: ", error);
  }
}

// Function to get the EML attachment URL of the selected message
async function getEMLAttachmentUrl(item) {
  return new Promise((resolve, reject) => {
    item.getAttachmentAsync({ asyncContext: item.itemId }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

// Function to move the email to Trash
async function moveToTrash(item) {
  return new Promise((resolve, reject) => {
    item.moveAsync(Office.MailboxEnums.FolderType.Trash, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve("Message moved to trash.");
      } else {
        reject(result.error);
      }
    });
  });
}
