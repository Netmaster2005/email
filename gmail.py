import imaplib
import email
import os
import sys

# Usage: python script_name.py <email> <password>
# Example: python script_name.py myemail@gmail.com mypassword

if __name__ == "__main__":
    email_account = sys.argv[1]
    password = sys.argv[2]
    
imap_url = 'imap.gmail.com'
my_mail = imaplib.IMAP4_SSL(imap_url)
my_mail.login(email_account, password)

my_mail.select('Inbox')

status, data = my_mail.search(None, 'ALL')
mail_ids = data[0]  
id_list = mail_ids.split()
first_email_id = int(id_list[0])
latest_email_id = int(id_list[-1])

for i in range(latest_email_id, first_email_id - 1, -1):
    status, msg_data = my_mail.fetch(str(i), '(RFC822)')
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            # Basic email details
            email_subject = msg['subject']
            email_from = msg['from']
            email_date = msg['Date']

            print(f"From: {email_from}")
            print(f"Subject: {email_subject}")
            print(f"Date: {email_date}\n")

            # Check for attachments
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()
                if bool(filename):
                    save_dir = "C:/Users/USER/Documents/"
                    os.makedirs(save_dir, exist_ok=True)
                    file_path = os.path.join(save_dir, filename)

                    if not os.path.isfile(file_path):
                        with open(file_path, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        print(f"Attachment saved: {file_path}")

my_mail.close()
my_mail.logout()
