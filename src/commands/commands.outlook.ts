/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

console.log("[Add-in] Background script loaded into memory.");

export function setNotificationInOutlook(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  event.completed();
}

export function onMessageRecipientsChanged(event: Office.AddinCommands.Event) {
  console.log("[Add-in] EVENT TRIGGERED: onMessageRecipientsChanged");
  
  const item = Office.context.mailbox.item;
  
  if (!item) {
    console.warn("[Add-in] WARNING: Office.context.mailbox.item is undefined.");
    event.completed();
    return;
  }

  if (!item.to) {
    console.warn("[Add-in] WARNING: item.to is undefined. Ensure this is a compose window.");
    event.completed();
    return;
  }

  console.log("[Add-in] Fetching recipient list...");
  
  item.to.getAsync((asyncResult: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;
      console.log(`[Add-in] Successfully fetched ${recipients.length} recipient(s):`, JSON.stringify(recipients));

      const targetEmail = "shridhar.p@inspirozit.com"; // Putting my own email for testing
      console.log(`[Add-in] Looking for target email: ${targetEmail}`);

      const containsTarget = recipients.some((recipient) => {
        const isMatch = recipient.emailAddress.toLowerCase() === targetEmail.toLowerCase();
        console.log(`[Add-in] Comparing: [${recipient.emailAddress}] to [${targetEmail}] -> Match: ${isMatch}`);
        return isMatch;
      });

      if (containsTarget) {
        console.log("[Add-in] Target email found! Preparing to inject HTML block...");
        
        const htmlBlock = `
          <p>Here is the standard <b>Manseuto Ventures</b> client block:</p>
          <ul>
            <li></li>
            <li></li>
          </ul>
        `;

        item.body.prependAsync(
          htmlBlock,
          { coercionType: Office.CoercionType.Html },
          (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("[Add-in] SUCCESS: HTML block prepended successfully.");
            } else {
              console.error("[Add-in] ERROR: Failed to prepend HTML block.", setResult.error);
            }
            event.completed(); 
          }
        );
        return;
      } else {
        console.log("[Add-in] Target email not found in the current recipient list.");
      }
    } else {
      console.error("[Add-in] ERROR: Failed to get recipients.", asyncResult.error);
    }
    
    console.log("[Add-in] Event processing finished without injection.");
    event.completed();
  });
}

// Register the add-in commands with the Office host application.
Office.onReady(async () => {
  console.log("[Add-in] Office.onReady fired. Associating actions...");
  Office.actions.associate("onMessageRecipientsChanged", onMessageRecipientsChanged);
});