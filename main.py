import os
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import openpyxl

load_dotenv()

def read_email_addresses(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        return [
            row[0] for row in sheet.iter_rows(min_col=1, max_col=1, values_only=True) if row[0]
        ]
    except Exception as e:
        print(f"Error reading email addresses: {e}")
        return []

def send_email_with_signature_and_attachments(sender_email, sender_password, recipient_email, subject, body, signature_image_path, directory):
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

        msg.attach(related_part)

        if os.path.isdir(directory):
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                if os.path.isfile(file_path):
                    with open(file_path, "rb") as file:
                        attachment = MIMEApplication(file.read(), Name=filename)
                        attachment.add_header("Content-Disposition", f'attachment; filename="{filename}"')
                        msg.attach(attachment)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

def process_emails(file_path, signature_image_path, sender_email, sender_password, subject, progress_label, root):
    email_addresses = read_email_addresses(file_path)
    if not email_addresses:
        messagebox.showerror("Error", "No email addresses found in the file.")
        return

    progress_label.config(text="Processing emails...")

    for email in email_addresses:
        directory = filedialog.askdirectory(title=f"Select attachments directory for {email}")
        if not directory:
            messagebox.showwarning("Warning", f"No directory selected for {email}. Skipping.")
            continue

        body = """Здравейте,
        <br><br>
        Приложено Ви изпращаме справка за резултатите от работата в клас на Вашето дете и сканирани писмени работи.
        <br><br>
        С уважение!"""
        send_email_with_signature_and_attachments(
            sender_email, sender_password, email, subject, body, signature_image_path, directory
        )

    progress_label.config(text="Emails are sent!")

    root.after(5000, root.destroy)

def main():
    root = tk.Tk()
    root.title("Email Sender")

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    window_width = int(screen_width * 0.5)
    window_height = int(screen_height * 0.5)

    root.geometry(f"{window_width}x{window_height}+{int((screen_width - window_width) / 2)}+{int((screen_height - window_height) / 2)}")

    tk.Label(root, text='Absolute path to the table with emails').grid(row=0, column=0, pady=5, sticky="nsew")
    tk.Label(root, text='Absolute path to the signature image').grid(row=1, column=0, pady=5, sticky="nsew")

    path_for_emails = tk.StringVar()
    path_for_image = tk.StringVar()

    tk.Entry(root, textvariable=path_for_emails).grid(row=0, column=1, pady=5, sticky="nsew")
    tk.Entry(root, textvariable=path_for_image).grid(row=1, column=1, pady=5, sticky="nsew")

    progress_label = tk.Label(root, text="")
    progress_label.grid(row=3, column=0, columnspan=2, pady=10, sticky="nsew")

    def on_submit():
        file_path = path_for_emails.get()
        signature_path = path_for_image.get()
        sender_email = os.getenv("EMAIL")
        sender_password = os.getenv("PASSWORD")
        subject = "Информационно писмо - Гимназия по информатика"

        if not file_path or not signature_path:
            messagebox.showerror("Error", "Please fill in all fields.")
            return

        if not os.path.isfile(signature_path):
            messagebox.showerror("Error", f"Signature image '{signature_path}' not found.")
            return

        process_emails(file_path, signature_path, sender_email, sender_password, subject, progress_label, root)

    tk.Button(root, text="Submit", command=on_submit).grid(row=2, column=1, pady=10, sticky="nsew")

    root.mainloop()

if __name__ == "__main__":
    main()
