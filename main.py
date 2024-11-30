import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import openpyxl

load_dotenv()


def read_email_addresses(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    email_addresses = []
    for row in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
        if row[0]:
            email_addresses.append(row[0])
    return email_addresses


def send_email_with_signature_and_attachments(
        sender_email, sender_password, recipient_email, subject, body, signature_image_path, directory
):
    try:

        msg = MIMEMultipart("mixed")
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        related_part = MIMEMultipart("related")
        html_content = f"""
        <html>
        <body>
            <p>{body}</p>
            <br>
            <img src="cid:signature_image" alt="Signature" style="width:20%; height:auto;">
        </body>
        </html>
        """
        related_part.attach(MIMEText(html_content, "html"))

        if os.path.isfile(signature_image_path):
            with open(signature_image_path, "rb") as img_file:
                img = MIMEImage(img_file.read())
                img.add_header("Content-ID", "<signature_image>")
                img.add_header("Content-Disposition", "inline", filename=os.path.basename(signature_image_path))
                related_part.attach(img)
        else:
            print(f"Signature image {signature_image_path} not found.")

        msg.attach(related_part)

        if os.path.isdir(directory):
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                if os.path.isfile(file_path):
                    with open(file_path, "rb") as file:
                        content = file.read()

                        attachment = MIMEApplication(content, Name=filename)
                        attachment.add_header(
                            "Content-Disposition", f'attachment; filename="{filename}"'
                        )
                        msg.attach(attachment)

        else:
            print(f"Directory {directory} does not exist or is not valid.")

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print("Successfully sent email.")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")


def main():
    file_path_for_emails = input("Enter file path for emails: ")
    sender_email = os.getenv("EMAIL")
    sender_password = os.getenv("PASSWORD")
    subject = "Информационно писмо - Гимназия по информатика"
    signature_image_path = input("Enter file path for signature image: ").strip()

    if not os.path.isfile(signature_image_path):
        print(f"Error: Signature image '{signature_image_path}' not found.")
        return

    email_addresses = read_email_addresses(file_path_for_emails)
    for email in email_addresses:
        print(f"Sending email to {email}...")
        directory = input("Enter the directory for attachments: ").strip()

        if not os.path.isdir(directory):
            print(f"Error: Directory '{directory}' not found.")
            continue

        body = """Здравейте,
<br><br>
Приложено Ви изпращаме справка за резултатите от работата в клас на Вашето дете и сканирани писмени работи.
<br><br>
С уважение!"""

        send_email_with_signature_and_attachments(
            sender_email, sender_password, email, subject, body, signature_image_path, directory
        )


if __name__ == "__main__":
    main()
