import win32com.client
import os

# Initialize Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the inbox folder (Folder number 6 is the inbox)
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

# Loop through all emails in the inbox
filtered_messages = messages.Restrict("[SenderName] = 'Kwashie Andoh'")
for message in filtered_messages:
    print(f"Found email: {message.SenderName}")

# Loop through attachments
if message.Attachments.Count > 0:
    for attachment in message.Attachments:
        attachment.SaveAsFile(os.path.join("C:/Users/USER/Documents/", attachment.FileName))

