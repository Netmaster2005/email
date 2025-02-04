import subprocess
import getpass

def run_email_script(email_address, password):
    # Extract the domain from the email address.
    domain = email_address.split('@')[-1].lower()

    # Depending on the domain, run the corresponding script.
    if domain == "gmail.com":
        print("Running Gmail script...")
        subprocess.run(["python", "gmail.py", email_address, password])
    elif domain == "outlook.com":
        print("Running Outlook script...")
        subprocess.run(["python", "outlook.py"])
    else:
        print("Unsupported email domain:", domain)

if __name__ == "__main__":
    # Prompt the user for their email and password.
    user_email = input("Enter your email address: ")
    user_password = getpass.getpass("Enter your password: ")
    
    run_email_script(user_email, user_password)
