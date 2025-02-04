import win32com.client
import os
import pythoncom

def process_outlook_emails():
    # Initialize COM
    pythoncom.CoInitialize()

    try:
        # Initialize Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the inbox folder (Folder number 6 is the inbox)
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        # Sort emails by received time (newest first)
        messages.Sort("[ReceivedTime]", True)

        # Filter emails by sender and subject (e.g., emails containing "invoice" in the subject)
        filter_criteria = "[SenderName] = 'Kwashie Andoh'"
        filtered_messages = messages.Restrict(filter_criteria)

        # Loop through filtered emails
        for message in filtered_messages:
            print(f"Found email from: {message.SenderName} - Subject: {message.Subject}")

            # Check if the email has attachments
            if message.Attachments.Count > 0:
                print(f"Processing {message.Attachments.Count} attachment(s)...")

                # Loop through attachments
                for attachment in message.Attachments:
                    try:
                        # Define the save path for the attachment
                        save_path = os.path.join("C:/Users/USER/Documents/", attachment.FileName)

                        # Save the attachment to the specified location
                        attachment.SaveAsFile(save_path)
                        print(f"Attachment '{attachment.FileName}' saved to {save_path}.")
                    except Exception as e:
                        print(f"Failed to save '{attachment.FileName}': {e}")
            else:
                print("No attachments found in this email.")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Release COM
        pythoncom.CoUninitialize()

# Run the function
process_outlook_emails()