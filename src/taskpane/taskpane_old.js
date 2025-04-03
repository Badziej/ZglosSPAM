/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = reportSpam;
  }
});

export async function reportSpam() {
  const item = Office.context.mailbox.item;
  if (!item) {
    Office.context.mailbox.displayMessageForm({
      subject: "Error",
      body: "No email selected."
    });
    return;
  }

  item.loadCustomPropertiesAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to load custom properties.");
      return;
    }
    
    let props = result.value;
    let alreadyReported = props.get("requestSPAM");

    if (alreadyReported) {
      Office.context.mailbox.displayMessageForm({
        subject: "Spam Already Reported",
        body: "This email has already been reported."
      });
      return;
    }

    item.getAttachmentsAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        let attachment = result.value.length > 0 ? result.value[0] : null;
        
        Office.context.mailbox.displayNewMessageForm({
          to: ["casb@ratels.pl"],
          subject: "[ZS] " + item.subject,
          attachments: attachment ? [{ name: "ReportedEmail.eml", url: attachment.url }] : []
        });

        // Mark as reported
        props.set("requestSPAM", "true");
        props.saveAsync();
      }
    });
  });
}
