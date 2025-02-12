function forwardAsAttachment() {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        // The currently selected item (the email being read)
        const item = Office.context.mailbox.item;
  
        // The display name of the current user (original recipient)
        const userName = Office.context.mailbox.userProfile.displayName;
  
        // Prepare the new message
        const messageOptions = {
          toRecipients: ['ddriscoll@perrknight.com'],
          subject: 'Forwarded email as attachment',
          htmlBody: `
            <p>Hello,</p>
            <p>This email was originally received by: <strong>${userName}</strong></p>
            <p>See attached original email.</p>
          `,
          attachments: [
            {
              // Mark as an item attachment
              type: Office.MailboxEnums.AttachmentType.Item,
              itemId: item.itemId
            }
          ]
        };
  
        // Open the new mail form with the item attached
        Office.context.mailbox.displayNewMessageForm(messageOptions);
      }
    });
  }
  
  // Expose globally for the manifest
  window.forwardAsAttachment = forwardAsAttachment;