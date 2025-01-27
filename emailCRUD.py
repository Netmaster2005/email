import imaplib
import email
import os

from credentials import useName,passWord

imap_url ='imap.gmail.com'
my_mail = imaplib.IMAP4_SSL(imap_url)
my_mail.login(useName, passWord)

my_mail.select('Inbox')

data = my_mail.search(None, 'All')
mail_ids = data[1]  
id_list = mail_ids[0].split(b' ')  
first_email_id = int(id_list[0])
latest_email_id = int(id_list[-1])

for i in range(latest_email_id,first_email_id, -1):
     data = my_mail.fetch(str(i), '(RFC822)' )
     for response_part in data:
        arr = response_part[0]
        if isinstance(arr, tuple):
            msg = email.message_from_bytes(arr[1])   
            email_subject = msg['subject']
            index = (msg['from']).find('<')
            email_from = (msg['from'])[0:index]
            email_date =  msg['Date']

            if 'Google' in email_from:
                print('From : ' + email_from)
                print('Subject : ' + email_subject)
                print('Date: ' + email_date + '\n')

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                fileName = part.get_filename()
                if bool(fileName):
                    filePath = os.path.join('C:/Users/tejas/Documents/', fileName)
                    if not os.path.isfile(filePath) :
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
            
            if 'Google' in email_from:
               my_mail.store(str(i), "+FLAGS", "\\Deleted")

# my_mail.expunge()
my_mail.close()
my_mail.logout()