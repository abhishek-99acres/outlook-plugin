// taskpane.js

let currentAttachments = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("attachments-container").innerHTML = "Loading...";
    loadAttachments();
  }
});

function loadAttachments() {
  Office.context.mailbox.item.getAttachmentsAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      document.getElementById("attachments-container").innerHTML =
        "Error loading attachments.";
      return;
    }

    // Filter out inline images (like email signatures)
    currentAttachments = result.value.filter((a) => !a.isInline);
    const container = document.getElementById("attachments-container");

    if (currentAttachments.length === 0) {
      container.innerHTML = "No attachments found.";
      document.getElementById("save-btn").style.display = "none";
      return;
    }

    container.innerHTML = ""; // Clear loading text

    // Generate UI for each attachment
    currentAttachments.forEach((att, index) => {
      const div = document.createElement("div");
      div.className = "attachment-item";

      div.innerHTML = `
                <div class="attachment-name">${att.name}</div>
                <select id="cat-${index}">
                    <option value="">-- Select Category --</option>
                    <option value="Invoice">Invoice</option>
                    <option value="Contract">Contract</option>
                    <option value="Report">Report</option>
                    <option value="Other">Other</option>
                </select>
            `;
      container.appendChild(div);
    });

    document.getElementById("save-btn").style.display = "block";
  });
}

function saveCategories() {
  let allSelected = true;
  let categoryData = {};

  // Validate that every attachment has a category selected
  currentAttachments.forEach((att, index) => {
    const selectedValue = document.getElementById(`cat-${index}`).value;
    if (!selectedValue) {
      allSelected = false;
    }
    categoryData[att.name] = selectedValue;
  });

  if (!allSelected) {
    document.getElementById("status").style.color = "red";
    document.getElementById("status").innerText =
      "Please select a category for all attachments.";
    return;
  }

  // Save the data to the email's custom properties
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    const customProps = result.value;

    // Save the count of attachments categorized (used by the background script to verify)
    customProps.set("categorizedAttachmentCount", currentAttachments.length);

    // Optional: Save the actual JSON mapping if you need to read it later on the backend
    customProps.set("attachmentCategories", JSON.stringify(categoryData));

    customProps.saveAsync((saveResult) => {
      if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("status").style.color = "green";
        document.getElementById("status").innerText =
          "Saved successfully! You can now send the email.";
      } else {
        document.getElementById("status").style.color = "red";
        document.getElementById("status").innerText =
          "Error saving categories.";
      }
    });
  });
}
