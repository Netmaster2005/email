from imaplib import IMAP4_SSL
import imaplib
import email
import win32com.client 
import pythoncom 
import re
import os
from django.views.decorators.csrf import csrf_exempt
from rest_framework.decorators import api_view
from django.http import JsonResponse
from email.policy import default
import traceback
from .serializers import EmailLoginSerializer
from rest_framework.views import APIView
from rest_framework.response import Response


@csrf_exempt
@api_view(['POST'])
def fetch_or_process_emails(request):
    """
    Django view to fetch exsmails from Gmail or process emails from Outlook.
    Accepts email and password via POST request.
    """
    # Validate the input using the serializer
    serializer = EmailLoginSerializer(data=request.data)
    if not serializer.is_valid():
        return Response(serializer.errors, status=400)

    # Extract email and password from the validated data
    email_account = serializer.validated_data['email']
    password = serializer.validated_data['password']
    
    domain = email_account.split('@')[-1].lower()
    # Determine the email domain
    if domain == "outlook.com":
         # Call the process_outlook_emails function for Outlook
        return process_outlook_emails()
    elif domain == "gmail.com":
         # Call the fetch_emails function for Gmail
         return fetch_emails(email_account, password)
    else:
        return Response({
            'status': 'error',
            'message': 'Unsupported email domain. Only Gmail and Outlook are supported.',
        }, status=400)

# @api_view(['POST'])
def fetch_emails(email_account, password):
        
        imap_url = 'imap.gmail.com'

        try:
            my_mail = imaplib.IMAP4_SSL(imap_url)
            my_mail.login(email_account, password)

            my_mail.select('Inbox')

            status, data = my_mail.search(None, 'ALL')
            mail_ids = data[0]  
            id_list = mail_ids.split()
            first_email_id = int(id_list[0])
            latest_email_id = int(id_list[-1])

            emails = []

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
        
            return JsonResponse({
                'status': 'success',
                'emails': emails,
                'message': 'Emails fetched successfully.',
            })

        except imaplib.IMAP4.error as e:
            # Handle IMAP errors (e.g., login failure)
            return JsonResponse({
                'status': 'error',
                'message': f'IMAP error: {str(e)}',
            }, status=500)

        except Exception as e:
            # Handle other exceptions
            return JsonResponse({
                'status': 'error',
                'message': f'An error occurred: {str(e)}',
            }, status=500)
        
def process_outlook_emails():
    """
    Process emails from Outlook using COM.
    """
    # Initialize COM
    pythoncom.CoInitialize()  

    try:
        # Initialize Outlook application and get the MAPI namespace
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Access the Inbox folder (Folder number 6 corresponds to Inbox)
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        
        # Sort emails by received time (newest first)
        messages.Sort("[ReceivedTime]", True)

        # Filter emails by sender; adjust the filter criteria as needed.
        # filter_criteria = "[SenderName] = 'Kwashie Andoh'"
        # filtered_messages = messages.Restrict(filter_criteria)
        
        # Get the number of emails matching the criteria
        total_emails = messages.Count
        print(f"Found {total_emails} email(s).")
        
        email_list = []
        
        for message in messages:
            email_list.append({
                "subject": message.Subject,
                "sender": message.SenderName,
                "body": message.Body[:100],  # Only take first 100 characters
                "received_time": str(message.ReceivedTime)  # Convert datetime to string
            })

        for i in range(1, total_emails + 1):
            message = messages.Item(i)
            email_details = {
                'from': message.SenderName,
                'subject': message.Subject,
                'date': message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S'),
                'attachments': [],
            }
        # Loop through the filtered emails using 1-based indexing
        # for i in range(1, total_emails + 1):
        #     message = messages.Item(i)
        #     print(f"\nProcessing email {i} from: {message.SenderName} - Subject: {message.Subject}")

            # Check if the email has attachments
            if message.Attachments.Count > 0:
                print(f"Processing {message.Attachments.Count} attachment(s)...")

                # Loop through attachments using 1-based indexing
                for j in range(1, message.Attachments.Count + 1):
                    attachment = message.Attachments.Item(j)
                    
                    # Sanitize the filename to remove invalid characters
                    safe_filename = re.sub(r'[<>:"/\\|?*]', '_', attachment.FileName)
                    
                    # Define the save path (create the directory if it doesn't exist)
                    save_dir = "C:/Users/USER/Documents/"
                    os.makedirs(save_dir, exist_ok=True)
                    save_path = os.path.join(save_dir, safe_filename)
                    
                    # Save the attachment
                    attachment.SaveAsFile(save_path)
                    print(f"Attachment '{safe_filename}' saved to {save_path}.")

        return Response({
            'status': 'success',
            'emails': email_list,
            'message': 'Emails fetched successfully.',
        })   
     

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Release COM resources
        pythoncom.CoUninitialize()

      
            
            

       
                