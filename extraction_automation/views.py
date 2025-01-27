from imaplib import IMAP4_SSL
import imaplib
import email
from email.policy import default
import traceback
from .serializers import UserSerializer
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.status import (
    HTTP_200_OK,
    HTTP_204_NO_CONTENT,
    HTTP_500_INTERNAL_SERVER_ERROR,
    HTTP_400_BAD_REQUEST,
    HTTP_201_CREATED,
)

class RegisterUser(APIView):
    def post(self, request):
        serializer = UserSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response({"message": "User registered successfully"}, status=HTTP_201_CREATED)
        return Response(serializer.errors, status=HTTP_400_BAD_REQUEST)
    

class RetrieveEmails(APIView):
    def post(self, request):
        # Get email and password from the request
        email = request.data.get('email')
        password = request.data.get('password')
        imap_server = request.data.get('imap_server', 'imap.gmail.com')  # Default to Gmail

        if not email or not password:
            return Response(
                {"error": "Email and password are required."},
                status=HTTP_400_BAD_REQUEST
            )

        try:
            # Connect to the IMAP server
            mail = imaplib.IMAP4_SSL('imap.gmail.com')

# Login to the server
            mail.login('email', 'password')

            # Select the mailbox (inbox in this case)
            mail.select('inbox')

            # Search for emails
            status, data = mail.search(None, 'ALL')

            # Get the list of email IDs
            email_ids = data[0].split()

            # Loop through the email IDs and fetch the email data
            for email_id in email_ids:
                status, data = mail.fetch(email_id, '(RFC822)')
                raw_email = data[0][1]
                print(raw_email)
                
                return Response({"emails": data}, status=HTTP_200_OK)

        except Exception as e:
            traceback.print_exc()
            return Response(
                {"error": "Failed to retrieve emails. Check your credentials and IMAP server."},
                status=HTTP_400_BAD_REQUEST
            )    

# class ProcessEmailAPIView(APIView):
#     def get(self):
#         try:
#             imap_host = 'imap.gmail.com'
#             imap_user = 'kwablahlawrence@gmail.com'
#             imap_pass = 'Lawrence@2024'

#             # if not imap_user or not imap_pass:
#             #     return Response(
#             #         {"error": "Email username and password are required."},
#             #         status=HTTP_400_BAD_REQUEST
#             #     )

#             mail = imaplib.IMAP4_SSL(imap_host)

# # Login to the server
#             mail.login(imap_user, imap_pass)

#             # Select the mailbox (inbox in this case)
#             mail.select('inbox')

#             # Search for emails
#             key = 'FROM'
#             value = 'bnsreenu@hotmail.com'
#             _, data = mail.search(None, key, value) 

#             # Get the list of email IDs
#             email_ids = data[0].split()

#             msgs = []

#         # Loop through the email IDs and fetch the email data
#             for num in email_ids:
#                 typ, data = mail.fetch(num, '(RFC822)') #RFC822 returns whole message (BODY fetches just body)
#                 msgs.append(data)

#             for msg in msgs[::-1]:
#                 for response_part in msg:
#                     if type(response_part) is tuple:
#                         my_msg=email.message_from_bytes((response_part[1]))
#                         print("_________________________________________")
#                         print ("subj:", my_msg['subject'])
#                         print ("from:", my_msg['from'])
#                         print ("body:")
#                         for part in my_msg.walk():  
#                             #print(part.get_content_type())
#                             if part.get_content_type() == 'text/plain':
#                                 print (part.get_payload())

#             return Response(
#                 {"message": "Emails processed successfully."},
#                 status=HTTP_200_OK
#             )

#         except Exception as e:
#             return Response(
#                 {"error": str(e)},
#                 status=HTTP_500_INTERNAL_SERVER_ERROR,
#             )
