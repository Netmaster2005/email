import win32com.client
import os
import pythoncom
import re

#  useName = "myronadjei407@gmail.com"  # Replace with your Outlook email
# passWord = "bpmpvuyqrvdapojn"
def process_outlook_emails():
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
        print(f"Found {total_emails} email(s) from 'Kwashie Andoh'.")

        # Loop through the filtered emails using 1-based indexing
        for i in range(1, total_emails + 1):
            message = messages.Item(i)
            print(f"\nProcessing email {i} from: {message.SenderName} - Subject: {message.Subject}")

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
            else:
                print("No attachments found in this email.")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Release COM resources
        pythoncom.CoUninitialize()

# Run the function
process_outlook_emails()
