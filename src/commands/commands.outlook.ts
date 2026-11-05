/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Log that the background script has been loaded into memory
console.log("[Add-in] Background script loaded into memory.");

// Keep the existing notification function intact
export function setNotificationInOutlook(event: Office.AddinCommands.Event) {
  // Define the notification message details
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Replace the current notification with the new message
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Signal that the event has completed
  event.completed();
}

// Define the HTML template exactly as formatted in the reference file
function getHtmlTemplate(): string {
  return "<p><strong>RFP INTAKE</strong></p>" +
    "<p><strong>RFP (Y/N):</strong> </p>" +
    "<p><strong>Account:</strong> </p>" +
    "<p><strong>Deliverable Needed (formats):</strong> </p>" +
    "<p><strong>Brand being activated (Inc. / FC / Both):</strong> </p>" +
    "<p><strong>Due Date:</strong> </p>" +
    "<p><strong>Goals &amp; Objectives:</strong> </p>" +
    "<p><strong>Budget:</strong> </p>" +
    "<p><strong>General ideas or thoughts on tactics that are appropriate:</strong> </p>" +
    "<p><strong>Splits should be included:</strong> </p>" +
    "<p><strong>Your planner in RFP requests:</strong> </p>" +
    "<br/>";
}

// Define the Plain Text template exactly as formatted in the reference file
function getTextTemplate(): string {
  return "RFP INTAKE\n\n" +
    "RFP (Y/N): \n\n" +
    "Account: \n\n" +
    "Deliverable Needed (formats): \n\n" +
    "Brand being activated (Inc. / FC / Both): \n\n" +
    "Due Date: \n\n" +
    "Goals & Objectives: \n\n" +
    "Budget: \n\n" +
    "General ideas or thoughts on tactics that are appropriate: \n\n" +
    "Splits should be included: \n\n" +
    "Your planner in RFP requests: \n\n";
}

// Main event handler triggered when the To, Cc, or Bcc fields change
export function onMessageRecipientsChanged(event: Office.AddinCommands.Event) {
  // Log that the event was triggered
  console.log("[Add-in] EVENT TRIGGERED: onMessageRecipientsChanged");
  
  // Retrieve the current mail item from the Office context
  const item = Office.context.mailbox.item;
  
  // Verify that the mail item exists
  if (!item) {
    console.warn("[Add-in] WARNING: Office.context.mailbox.item is undefined.");
    event.completed(); // Free up the Outlook UI thread
    return;
  }

  // Verify that the 'to' field exists (ensures we are in compose mode)
  if (!item.to) {
    console.warn("[Add-in] WARNING: item.to is undefined. Ensure this is a compose window.");
    event.completed(); // Free up the Outlook UI thread
    return;
  }

  // Log the start of the recipient fetching process
  console.log("[Add-in] Fetching recipient list...");
  
  // Asynchronously retrieve the array of recipients in the 'To' field
  item.to.getAsync((asyncResult: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
    // Check if the API successfully returned the recipients
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      // Extract the recipient array
      const recipients = asyncResult.value;
      // Log the total count and the raw JSON of the recipients
      console.log(`[Add-in] Successfully fetched ${recipients.length} recipient(s):`, JSON.stringify(recipients));

      // Define the target email
      const targetEmail = "marketing@mansueto.com"; 
      console.log(`[Add-in] Looking for target email: ${targetEmail}`);

      // Iterate through recipients to see if the target email is present
      const containsTarget = recipients.some((recipient) => {
        // Perform a case-insensitive comparison
        const isMatch = recipient.emailAddress.toLowerCase() === targetEmail.toLowerCase();
        // Log each comparison step for debugging
        console.log(`[Add-in] Comparing: [${recipient.emailAddress}] to [${targetEmail}] -> Match: ${isMatch}`);
        return isMatch;
      });

      // If a match is found
      if (containsTarget) {
        console.log("[Add-in] Target email found! Checking if template already exists...");
        
        // Step 1: Asynchronously determine if the email is Plain Text or HTML
        item.body.getTypeAsync((typeResult: Office.AsyncResult<Office.CoercionType>) => {
          // Check if the API successfully determined the body type
          if (typeResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("[Add-in] ERROR: unable to determine body type.", typeResult.error);
            event.completed(); // Close the event on failure
            return;
          }

          // Extract the body type (Office.CoercionType.Html or Office.CoercionType.Text)
          const bodyType = typeResult.value;
          console.log(`[Add-in] Body type determined as: ${bodyType}`);

          // Step 2: Asynchronously read the current contents of the email body
          item.body.getAsync(bodyType, (bodyResult: Office.AsyncResult<string>) => {
            // Check if the API successfully read the body content
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
              console.error("[Add-in] ERROR: unable to read message body.", bodyResult.error);
              event.completed(); // Close the event on failure
              return;
            }

            // Extract the body content, defaulting to an empty string if null
            const currentBody = bodyResult.value || "";
            
            // Step 3: Check if "RFP INTAKE" already exists to prevent duplicate injections
            if (currentBody.toUpperCase().indexOf("RFP INTAKE") !== -1) {
              // Log that it was found and abort the injection
              console.log("[Add-in] Duplicate check: 'RFP INTAKE' already exists in the body. Aborting injection.");
              event.completed(); // Close the event
              return;
            }

            console.log("[Add-in] No existing template found. Preparing to inject...");

            // Select the appropriate string template based on the detected body type
            const template = (bodyType === Office.CoercionType.Html) ? getHtmlTemplate() : getTextTemplate();

            // Step 4: Inject the template at the very top of the email body
            item.body.prependAsync(
              template,
              { coercionType: bodyType }, // Explicitly declare the data type being injected
              (setResult: Office.AsyncResult<void>) => {
                // Verify the injection was successful
                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("[Add-in] SUCCESS: Template injected successfully.");
                } else {
                  console.error("[Add-in] ERROR: Failed to inject template.", setResult.error);
                }
                // Conclude the event lifecycle
                event.completed(); 
              }
            );
          });
        });
        
        // Return here so the asynchronous API calls aren't interrupted by the fallback event.completed()
        return; 
      } else {
        console.log("[Add-in] Target email not found in the current recipient list.");
      }
    } else {
      console.error("[Add-in] ERROR: Failed to get recipients.", asyncResult.error);
    }
    
    // Fallback completion if no target is found or an earlier API call fails
    console.log("[Add-in] Event processing finished without injection.");
    event.completed();
  });
}

// Global initialization hook: Register the command handler with the Office host
Office.onReady(async () => {
  console.log("[Add-in] Office.onReady fired. Associating actions...");
  Office.actions.associate("onMessageRecipientsChanged", onMessageRecipientsChanged);
});