import openpyxl
import smtplib
import os
import dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

dotenv.load_dotenv()

def read_email_addresses(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    email_addresses = []
    for row in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
        if row[0]:
            email_addresses.append(row[0])
    return email_addresses

def send_email_with_attachments(sender_email, sender_password, recipient_email, subject, body, directory):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))
        if os.path.isdir(directory):
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                if os.path.isfile(file_path):
                    with open(file_path, 'rb') as file:
                        part = MIMEApplication(file.read(), Name=filename)
                        part['Content-Disposition'] = f'attachment; filename="{filename}"'
                        msg.attach(part)
        else:
            print(f"Directory {directory} does not exist or is not valid.")

        with smtplib.SMTP('smtp.gmail.com', 587) as server:  # Adjust server for your email provider
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)

        print(f"Email with attachments sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

def main():
    file_path_for_emails = input("Enter file path for emails: ")
    file_path = file_path_for_emails
    sender_email = os.getenv("EMAIL")
    sender_password = os.getenv("PASSWORD")
    subject = "Your Subject Here"

    email_addresses = read_email_addresses(file_path)
    for email in email_addresses:
        print(f"Sending email to {email}...")
        directory = input("Enter the directory for attachments: ").strip()
        name = input("Enter the name of the recipient: ")
        body = f"Здравейте {name}"
        send_email_with_attachments(sender_email, sender_password, email, subject, body, directory)

if __name__ == "__main__":
    main()
