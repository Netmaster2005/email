import imaplib
import email
import os

# Your Outlook credentials
useName = "kwablahklawrence@outlook.com"  # Replace with your Outlook email
passWord = "ZKMPC-59P68-65BKQ-RJQPT-EN3L9"          # Replace with your Outlook password or app password
imap_url = "outlook.office365.com"

# Connect to the IMAP server
try:
    mail = imaplib.IMAP4_SSL(imap_url)
    mail.login(useName, passWord)
    print("Logged in successfully!")
except Exception as e:
    print(f"Failed to log in: {e}")
    exit()

# Select the mailbox you want to use
mail.select("Inbox")  # Or any other folder you want to access

# Search for emails (e.g., all emails with attachments)
status, email_ids = mail.search(None, 'ALL')  # You can filter emails using criteria like 'FROM', 'SUBJECT', etc.

# Convert email IDs to a list
email_id_list = email_ids[0].split()
print(f"Found {len(email_id_list)} emails.")

# Directory to save attachments
attachment_dir = "attachments"
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)

# Loop through emails to fetch and download attachments
for email_id in email_id_list:
    # Fetch the email
    status, data = mail.fetch(email_id, '(RFC822)')
    if status != "OK":
        print(f"Failed to fetch email ID {email_id}")
        continue

    # Parse the email
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    # Extract email details
    subject = msg["subject"]
    sender = msg["from"]
    print(f"Processing email from {sender} with subject '{subject}'")

    # Process email parts to find attachments
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue

        # Get the attachment filename
        filename = part.get_filename()
        if filename:
            filepath = os.path.join(attachment_dir, filename)
            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            print(f"Downloaded attachment: {filename}")

# Logout from the mail server
mail.logout()
