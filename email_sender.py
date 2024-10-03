import smtplib
import pandas as pd
from email.message import EmailMessage
import os

def get_emails(file_path):
    """
    Reads email addresses from an Excel file.

    Args:
        file_path (str): Path to the Excel file.

    Returns:
        pd.Series: Series of email addresses.
        str: Error message if an issue occurs.
    """
    try:
        df = pd.read_excel(file_path)
        column_data = df['Email'].dropna()  # Drop rows where email is missing
        return column_data
    except FileNotFoundError:
        return f"The file {file_path} was not found."
    except Exception as e:
        return f"An error occurred while reading the Excel file: {e}"

def get_data(file_path):
    """
    Reads the content of a text file.

    Args:
        file_path (str): Path to the text file.

    Returns:
        str: Content of the text file.
        str: Error message if an issue occurs.
    """
    try:
        with open(file_path, 'r') as f:
            content = f.read()
        return content
    except FileNotFoundError:
        return f"The file {file_path} was not found."
    except Exception as e:
        return f"An error occurred while reading the text file: {e}"

def send_email_to_all(input_excel, input_text_file, subject, sender_email, app_password, smtp_server, smtp_port, attachment_path):
    """
    Sends an email to all addresses listed in the Excel file, attaching a file and using a specified subject.

    Args:
        input_excel (str): Path to the Excel file containing email addresses.
        input_text_file (str): Path to the text file containing the email body.
        subject (str): Subject of the email.
        sender_email (str): Sender's email address.
        app_password (str): App password for sender email login.
        smtp_server (str): SMTP server address (e.g., 'smtp.gmail.com').
        smtp_port (int): SMTP port (e.g., 587 for Gmail).
        attachment_path (str): Path to the file to be attached.
    """
    emails = get_emails(input_excel)
    email_body = get_data(input_text_file)

    if isinstance(emails, pd.Series) and isinstance(email_body, str):
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, app_password)

                for recipient_email in emails:
                    try:
                        msg = EmailMessage()
                        msg['Subject'] = subject
                        msg['From'] = sender_email
                        msg['To'] = recipient_email
                        msg.set_content(email_body)

                        with open(attachment_path, 'rb') as file:
                            file_data = file.read()
                            file_name = os.path.basename(attachment_path)
                            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

                        server.send_message(msg)
                        print(f"Email sent to {recipient_email} with attachment '{file_name}'")

                    except Exception as email_error:
                        print(f"Failed to send email to {recipient_email}: {email_error}")

        except Exception as e:
            print(f"An error occurred while sending emails: {e}")
    else:
        print(f"Failed to fetch valid emails or email body. Emails: {emails}, Email Body: {email_body}")

if __name__ == '__main__':
    # Hardcoded paths and details
    input_excel = 'hr_emails.xlsx'        # Excel file with email addresses
    input_text_file = 'input.txt'     # Text file containing the body of the email
    subject = 'Applying for the position of QA Engineer'  # Email subject
    smtp_server = 'smtp.gmail.com'    # SMTP server for Gmail
    smtp_port = 587                   # SMTP port for Gmail

    # Prompting for user input
    sender_email = input("Enter the sender's email address: ")
    app_password = input("Enter the sender's app password: ")
    attachment_path = input("Enter the path to the attachment file (RESUME.docx): ")

    # Call the function to send the emails with attachment
    send_email_to_all(input_excel, input_text_file, subject, sender_email, app_password, smtp_server, smtp_port, attachment_path)
