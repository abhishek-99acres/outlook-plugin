"use strict";

// ── CONFIG ────────────────────────────────────────────────────────────────────

const TRIGGER_RECIPIENT_MAP = {
  "abhishek.a3@99acres.com": "Abhishek Anand",
  "sonia.m@99acres.com": "Sonia M",
  "coder.abhi02@gmail.com": "Abhishek Kumar",
  "rkd02122@gmail.com": "Abhishek Kumar",
  "finance@contoso.com": "Finance",
  "legal@contoso.com": "Legal",
  "hr@contoso.com": "HR",
  "compliance@contoso.com": "Compliance",
  "audit@contoso.com": "Audit",
};

const KNOWN_PREFIXES = [
  "Legal_",
  "Finance_",
  "HR_",
  "Compliance_",
  "Contract_",
  "Invoice_",
  "Report_",
  "Presentation_",
  "Reference_",
  "General_",
];

// ── SEND HANDLER ──────────────────────────────────────────────────────────────

function onMessageSendHandler(event) {
  try {
    var attachments = Office.context.mailbox.item.attachments || [];

    // No attachments — allow send
    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // Find files without a category prefix
    var bad = attachments.filter(function (a) {
      return !KNOWN_PREFIXES.some(function (p) {
        return a.name.startsWith(p);
      });
    });

    // All categorized — allow send
    if (bad.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // Block send — per Microsoft docs:
    //   cancelLabel → customizes the "Don't Send" button text (max 20 chars)
    //   commandId   → clicking that button opens the taskpane
    //                 value must match the Control id in manifest = "msgComposeOpenPaneButton"
    event.completed({
      allowEvent: false,
      errorMessage:
        bad.length +
        " attachment(s) not categorized:\n" +
        bad
          .map(function (a) {
            return "  \u2022 " + a.name;
          })
          .join("\n") +
        "\n\nClick 'Categorize Now' to label them.",
      cancelLabel: "Categorize Now", // button text (max 20 chars)
      commandId: "msgComposeOpenPaneButton", // opens taskpane on click
    });
  } catch (e) {
    console.error("[AttachCat] send error:", e);
    event.completed({ allowEvent: true });
  }
}

// ── COMPOSE EVENTS ────────────────────────────────────────────────────────────

function onNewMessageComposeHandler(event) {
  // console.log("***********************\n");
  // console.log("Composing the message handler!\n");
  // console.log("***********************");

  event.completed({
    allowEvent: false, // Cancel the send action
    errorMessage: "This email cannot be sent due to policy reasons.",
  });

  try {
    updateNotification();
  } catch (e) {}
  event.completed();
}

function onMessageAttachmentsChangedHandler(event) {
  try {
    updateNotification();
    var attachments = Office.context.mailbox.item.attachments || [];
    var hasUncategorized = attachments.some(function (a) {
      return !KNOWN_PREFIXES.some(function (p) {
        return a.name.startsWith(p);
      });
    });
    if (hasUncategorized && Office.addin && Office.addin.showAsTaskpane) {
      Office.addin.showAsTaskpane();
    }
  } catch (e) {}
  event.completed();
}

function onMessageRecipientsChangedHandler(event) {
  try {
    updateNotification();
  } catch (e) {}
  event.completed();
}

// ── NOTIFICATION BAR ──────────────────────────────────────────────────────────

function updateNotification() {
  var attachments = Office.context.mailbox.item.attachments || [];
  var bad = attachments.filter(function (a) {
    return !KNOWN_PREFIXES.some(function (p) {
      return a.name.startsWith(p);
    });
  });
  var msg =
    attachments.length === 0
      ? "Attach a file — categorizer opens automatically."
      : bad.length > 0
        ? "\u26a0\ufe0f " +
          bad.length +
          " attachment(s) need a category before sending."
        : "\u2713 All " +
          attachments.length +
          " attachment(s) categorized. Ready to send.";

  Office.context.mailbox.item.notificationMessages.replaceAsync("attachCat", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: "Icon.16x16",
    persistent: true,
  });
}

// ── REGISTER ─────────────────────────────────────────────────────────────────
// Per Microsoft docs — for classic Outlook on Windows (JSRuntime),
// Office.onReady() does NOT run. Use Office.actions.associate at top level.
// For Outlook Web/Mac (WebViewRuntime / shared runtime), both work.
// Solution: call associate both ways to cover all clients.

Office.actions.associate(
  "onNewMessageComposeHandler",
  onNewMessageComposeHandler,
);
Office.actions.associate(
  "onMessageAttachmentsChangedHandler",
  onMessageAttachmentsChangedHandler,
);
Office.actions.associate(
  "onMessageRecipientsChangedHandler",
  onMessageRecipientsChangedHandler,
);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
