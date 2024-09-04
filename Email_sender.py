### Uchu Ni! - Bulk Email Sender
### Made by: M3shour 
### Lisense: Apache 2.0
### Description: This script is used to send bulk emails to a list of recipients. The script reads the recipient email addresses and names from an Excel file and sends personalized emails to each recipient. The script uses the smtplib library to send emails via an SMTP server and the imaplib library to save the sent emails to the IMAP "Sent" folder. The script also supports sending attachments with the emails. The email content is read from an HTML template file, which can be personalized with the recipient's name. The script logs the email sending status to a log file.

import imaplib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
from tqdm import tqdm
import logging
import time
import argparse
import os

def get_emails_and_names_from_excel(file_path, sheet_name, email_column, name_column=None):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    if name_column is None:
        return df[email_column].tolist(), [None] * len(df)
    return df[email_column].tolist(), df[name_column].tolist()

def send_bulk_emails(sender_email, sender_password, subject, recipients, imap_server, smtp_server, attachment_path, append_attachment, email_template):

    logging.basicConfig(filename='email_sending.log', level=logging.INFO, format='%(asctime)s - %(message)s')
    # Set up the server
    server = smtplib.SMTP(smtp_server, 587)
    server.starttls()
    server.login(sender_email, sender_password)
    imap_conn = imaplib.IMAP4_SSL(imap_server, 993)
    imap_conn.login(sender_email, sender_password)
    with open(email_template, 'r', encoding='utf-8') as file:
        html_template = file.read()
    for recipient_email, recipient_name in tqdm(recipients, desc="Sending emails", unit="email", unit_scale=True, bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]"):
        # Create the email content
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Personalize the message
        message = html_template.replace("{{recipient_name}}", recipient_name)
        msg.attach(MIMEText(message, 'html'))
        # Attach the PDF
        if attachment_path:
            with open(attachment_path, 'rb') as pdf_file:
                pdf_attachment = MIMEApplication(pdf_file.read(), _subtype='pdf')
                pdf_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                msg.attach(pdf_attachment)
        # Send the email
        try:
            server.send_message(msg)
            print(f"Email sent to {recipient_name} ({recipient_email})")
            logging.info(f"Email sent to {recipient_name} ({recipient_email})")
            if not append_attachment:
            # Sent to the IMAP server
                msg_without_pdf = MIMEMultipart()
                msg_without_pdf['From'] = msg['From']
                msg_without_pdf['To'] = msg['To']
                msg_without_pdf['Subject'] = msg['Subject']
                msg_without_pdf.attach(MIMEText(message, 'html'))
                # Save the sent message to the IMAP "Sent" folder without the attachment
                imap_conn.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), msg_without_pdf.as_bytes())
            else:
                # Save the sent message to the IMAP "Sent" folder with the attachment
                imap_conn.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())

        except Exception as e:
            logging.error(f"Failed to send email to {recipient_name} ({recipient_email}): {str(e)}")
            print(f"Failed to send email to {recipient_name} ({recipient_email}): {str(e)}")

    # Close the server connection
    server.quit()
    imap_conn.logout()
if __name__ == "__main__":
    # Example usage
    parser = argparse.ArgumentParser(description='Send bulk emails.')
    parser.add_argument('--sender_email', type=str, help='Sender email address', required=True)
    parser.add_argument('--sender_password', type=str, help='Sender email password', required=True)
    parser.add_argument('--subject', type=str, help='Email subject', required=True)
    parser.add_argument('--excel_file_path', type=str, help='Path to the Excel file', required=True)
    parser.add_argument('--sheet_name', type=str, help='Sheet name in the Excel file', required=True)
    parser.add_argument('--email-column', type=str, help='Column name in the Excel file containing the email addresses', required=True)
    parser.add_argument('--name-column', type=str, help='Column name in the Excel file containing the recipient names')
    parser.add_argument('--imap-server', type=str, help='IMAP server address', required=True)
    parser.add_argument('--smtp-server', type=str, help='SMTP server address', required=True)
    parser.add_argument('--attachment_path', type=str, help='Path to the attachment file')
    parser.add_argument('--append-attachment', type=bool, help='Append attachment in the IMAP "Sent" folder - default is True')
    parser.add_argument('--email-template', type=str, help='Path to the email template html file', required=True)
    args = parser.parse_args()
    
    # Use command-line arguments if provided, otherwise use hardcoded values
    sender_email = args.sender_email
    sender_password = args.sender_password
    subject = args.subject
    excel_file_path = args.excel_file_path
    sheet_name = args.sheet_name
    email_column = args.email_column
    name_column = args.name_column
    imap_server = args.imap_server
    smtp_server = args.smtp_server
    attachment_path = args.attachment_path
    append_attachment = args.append_attachment or True
    email_template = args.email_template
    # Get recipient emails and names from the Excel file

    recipient_emails, recipient_names = get_emails_and_names_from_excel(excel_file_path, sheet_name, email_column, name_column)
    recipients = list(zip(recipient_emails, recipient_names))
    
    send_bulk_emails(sender_email, sender_password, subject, recipients, imap_server,smtp_server, attachment_path, append_attachment, email_template)