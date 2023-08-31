import smtplib, csv, time, os
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Removes temp file if manually aborted
if os.path.exists("temp.csv"):
    os.remove("temp.csv")

# Define the scopes you want to authorize
scopes = ["https://www.googleapis.com/auth/drive"]

# Set up the OAuth 2.0 flow
flow = InstalledAppFlow.from_client_secrets_file(
    "PATH-TO-CREDENTIALS", scopes=scopes)

# Run the flow and authorize the user
creds = flow.run_local_server(port=0, authorization_prompt_message="[+] Authorize your client via your web browser",
                               success_message="[+] Authorization successful")

# Set up the API client
drive_service = build("drive", "v3", credentials=creds)

# Define the file ID of the CSV file you want to access
file_id = input("\n[+] Enter the ID of the target file: ")

choice = input("[+] Type the number of the email template you would like to use:\
                     \n[+] 1. TEMPLATE 1\
                     \n[+] 2. TEMPLATE 2\
                     \n\n[+] Selection: ")


if int(choice) == 1:
    # Create the HTML email message object
    with open("PATH-TO-TEMPLATE-1", "r") as f:
        html_content = f.read()

    # Subject line is factored in spam detection
    subject = "SUBJECT"
elif int(choice) == 2:
    # Create the HTML email message object
    with open("PATH-TO-TEMPLATE-2", "r") as f:
        html_content = f.read()

    # Subject line is factored in spam detection
    subject = "SUBJECT"
else:
    # Create the HTML email message object
    with open("PAT-TO-DEFAULT-TEMPLATE", "r") as f:
        html_content = f.read()

    # Subject line is factored in spam detection
    subject = "SUBJECT"


try:
    # Use the Drive API to get the file metadata
    file_metadata = drive_service.files().get(fileId=file_id).execute()

    # Show name of file
    print(f"[+] Name of file being used: {file_metadata['name']}")

    # Download the CSV file content
    csv_content = drive_service.files().export(fileId=file_id, mimeType="text/csv").execute()

    # Save the CSV file content to disk
    print("[+] Creating temporary file")
    with open("temp.csv", "wb") as f:
        f.write(csv_content)

    # Read the email addresses from the CSV file
    with open("temp.csv", "r") as csvfile:
        reader = csv.reader(csvfile)
        
        # Skip the header row
        next(reader)
        
        for count, row in enumerate(reader, start=1):
            # Set the email to send to
            email = row[0]

            # Skip emails starting with 'info'
            if email.startswith("info") or email.startswith("support") or email.startswith("noreply"):
                print(f"[+] Skipping email #{count}")
                
                # Take long break if 200 emails have been sent
                if count == 200:
                    print("[+] 200 Emails sent, taking a long break")
                    time_break = datetime.now() + timedelta(minutes=20)
                    print(format(time_break.time(), "[+] Resuming at %H:%M:%S"))
                    time.sleep(1200)
                # Take a break every 20 emails
                elif count % 20 == 0:
                    time_break = datetime.now() + timedelta(minutes=10)
                    print("[+] Taking a 10 min break")
                    print(format(time_break.time(), "[+] Resuming at %H:%M:%S"))
                    time.sleep(600)
            else:
                # Set the email message parameters
                sender = "SENDER-EMAIL"
                recipient = email
                
                # Set the MIME data
                message = MIMEText(html_content, "html")
                message["From"] = sender
                message["To"] = recipient
                message["Subject"] = subject

                # Set your SMTP server settings
                smtp_server = "SMTP-SERVER"
                smtp_port = 0000
                smtp_username = "SMTP-USER"
                smtp_password = "SMTP-PASS"

                # Catches any email sending errors
                try:
                    # Send the email using the SMTP server
                    with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                        smtp.starttls()
                        smtp.login(smtp_username, smtp_password)
                        smtp.sendmail(sender, recipient, message.as_string())
                        print(f"[+] Email #{count} sent successfully to: {email}")
                except Exception:
                    print(f"[+] Encountered error: {Exception}, skipping email")

                
                # Take long break if 200 emails have been sent
                if count == 200:
                    print("[+] 200 Emails sent, taking a long break")
                    time_break = datetime.now() + timedelta(minutes=20)
                    print(format(time_break.time(), "[+] Resuming at %H:%M:%S"))
                    time.sleep(1200)
                # Take a break every 20 emails
                elif count % 20 == 0:
                    time_break = datetime.now() + timedelta(minutes=10)
                    print("[+] Taking a 10 min break")
                    print(format(time_break.time(), "[+] Resuming at %H:%M:%S"))
                    time.sleep(600)
                else:
                    # May or may not need to sleep in between transactions!!
                    time.sleep(2)

except HttpError as error:
    print("An error occurred: %s" % error)

os.remove("temp.csv")
print("\n[+] Removing temporary file")
print("[+] Operation complete")
