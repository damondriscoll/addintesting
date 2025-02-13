Office.onReady();

function forwardAsAttachment(event) {
  try {
      var item = Office.context.mailbox.item;

      if (!item) {
          console.error("No email selected.");
          event.completed();
          return;
      }

      item.forwardAsAttachmentAsync(
          {
              toRecipients: ["ddriscoll@perrknight.com"], // Change to actual email
              subject: "Forwarded Email",
              body: "Here is the forwarded email."
          },
          function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Email forwarded successfully.");
              } else {
                  console.error("Error forwarding email: " + asyncResult.error.message);
              }
              event.completed();
          }
      );
  } catch (error) {
      console.error("Exception in forwardAsAttachment: ", error);
      event.completed();
  }
}

/**
 * IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name
 * specified in the manifest to its JavaScript counterpart.
 */
Office.actions.associate("forwardAsAttachment", forwardAsAttachment);
