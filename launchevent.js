// launchevent.js

function onMessageSendHandler(event) {
  // check whether the onsend event is firing or not
  console.log("onMessageSendHandler receiving the event !");
  event.completed({
    allowEvent: false,
    errorMessage: "onMessageSendHandler receiving the event !",
    debugMessage: "onMessageSendHandler receiving the event !",
  });
  const targetEmail = "coder.abhi02@gmail.com".toLowerCase();

  // 1. Check 'To' recipients
  Office.context.mailbox.item.to.getAsync((toResult) => {
    const toEmails = toResult.value.map((r) => r.emailAddress.toLowerCase());

    // 2. Check 'Cc' recipients
    Office.context.mailbox.item.cc.getAsync((ccResult) => {
      const ccEmails = ccResult.value.map((r) => r.emailAddress.toLowerCase());
      const hasTargetEmail =
        toEmails.includes(targetEmail) || ccEmails.includes(targetEmail);

      if (!hasTargetEmail) {
        return event.completed({ allowEvent: true }); // Let it send
      }

      // 3. Check for attachments
      Office.context.mailbox.item.getAttachmentsAsync((attResult) => {
        const attachments = attResult.value.filter((a) => !a.isInline); // Ignore inline images/signatures

        if (attachments.length === 0) {
          return event.completed({ allowEvent: true }); // No attachments, let it send
        }

        // 4. Check if they have been categorized
        Office.context.mailbox.item.loadCustomPropertiesAsync((propResult) => {
          const customProps = propResult.value;
          const categorizedCount =
            customProps.get("categorizedAttachmentCount") || 0;

          // If the number of categorized attachments matches the current attachments, allow send
          if (categorizedCount === attachments.length) {
            event.completed({ allowEvent: true });
          } else {
            // Block the send and alert the user
            event.completed({
              allowEvent: false,
              errorMessage:
                "You have uncategorized attachments. Please open the 'Categorize Attachments' add-in to categorize them before sending.",
            });
          }
        });
      });
    });
  });
}

// Office requires you to associate the function name with the background event
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
