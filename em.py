import win32com.client
import os

# Initialize Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the inbox folder (Folder number 6 is the inbox)
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

# Loop through all emails in the inbox
for message in messages:
    # Check the subject of the email
    if message.SenderName == "Microsoft account team":
        print(f"Found email: {message.SenderName}")

        # Loop through attachments
        attachments = message.Attachments
        for attachment in attachments:
            # Save attachment to the desired location
            attachment.SaveAsFile(os.path.join("C:/Users/USER/Documents/", attachment.FileName))
            print(f"Attachment {attachment.FileName} saved.")

