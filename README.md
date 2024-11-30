# Bulk Email Sender with Attachments

This Python script automates sending emails with attachments to multiple recipients. It reads recipient email addresses from an Excel file and allows you to attach files from a specified directory for each email.

## Features

- Reads recipient email addresses from an Excel file.
- Sends emails with a custom subject and message body.
- Attaches all files from a specified directory to each email.
- Uses environment variables for email authentication.

## Requirements

- Python 3.6 or higher
- Required Python packages:
  - `openpyxl` (for reading Excel files)
  - `python-dotenv` (for managing environment variables)
  - `smtplib` (for sending emails)
- An Excel file with email addresses in the first column.
- A `.env` file for securely storing email credentials.

## Installation

1. Clone the repository or download the script.

2. Install the required Python packages:

   ```bash
   pip install openpyxl python-dotenv
   ```

3. Create a `.env` file in the project directory with the following content:

   ```plaintext
   EMAIL="your-email@example.com"
   PASSWORD="your-password"
   ```

4. Ensure you have an Excel file with email addresses in the first column.

## Usage

1. Run the script:

   ```bash
   python main.py
   ```

2. Follow the prompts:
   - Enter the file path to the Excel file containing email addresses.
   - Enter the directory path containing the attachments.
   - Provide the recipient's name for personalized emails.

3. The script will send emails to all listed addresses with the specified attachments.

## Example Workflow

1. Prepare an Excel file (`emails.xlsx`) with the following structure:

   | Email Address      |
   |--------------------|
   | recipient1@example.com |
   | recipient2@example.com |
   | recipient3@example.com |

2. Place your attachments in a directory, e.g., `attachments/`.

3. Run the script and provide inputs as prompted:
   - File path: `emails.xlsx`
   - Attachments directory: `attachments/`
   - Recipient's name: `John Doe`

The script will send personalized emails with attachments to each recipient in the Excel file.

## Notes

- Ensure you enable **Less Secure App Access** or use an app-specific password for Gmail. For other email providers, configure the SMTP settings accordingly.
- Make sure the attachments directory exists and contains the files you want to send.
- If you have 2FA enabled make App password and use that as your password in the script.

## License

This project is licensed under the GNU GENERAL PUBLIC LICENSE Version 3 License. See the `LICENSE` file for details.
