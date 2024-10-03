# Email-Sender-With-Attachment
This Python script allows you to send personalized emails to a list of recipients from an Excel file. The email body is fetched from a text file, and you can also attach a file (e.g., a resume) with each email. It uses Gmail's SMTP server for sending emails.

## Features:
- Fetches email addresses from an Excel file.
- Reads the body of the email from a text file.
- Sends the email with an attachment (e.g., a resume).
- Handles multiple recipients.
- Simple input via command line for email credentials and attachment.

## Prerequisites:
Make sure you have the following Python packages installed:
- `pandas`
- `openpyxl`

Install them via pip:
```bash
pip install pandas openpyxl
