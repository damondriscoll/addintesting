/**
 * This function is triggered when the user clicks the custom "Forward as Attachment" button.
 */
function forwardAsAttachment() {
    // Ensure the Office.js library is initialized
    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        // The currently selected item (email in Read mode)
        const item = Office.context.mailbox.item;
        
        // The name of the current user (the original recipient of this email)
        const userName = Office.context.mailbox.userProfile.displayName;
        
        // Create a new message with the current item attached
        const messageOptions = {
          toRecipients: ['ddriscoll@perrknight.com'],
          subject: 'Forwarded email as attachment',
          htmlBody: `<p>Hello,</p>
                     <p>This email was originally received by: <strong>${userName}</strong></p>
                     <p>See attached original email.</p>`,
          attachments: [{
            // Use the built-in Outlook enum for item attachments
            type: Office.MailboxEnums.AttachmentType.Item,
            itemId: item.itemId
          }]
        };
        
        // Opens a new Outlook message form with the original email attached
        Office.context.mailbox.displayNewMessageForm(messageOptions);
      }
    });
  }
  
  // Make sure to expose the function if needed in the global scope
  window.forwardAsAttachment = forwardAsAttachment;