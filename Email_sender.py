### Uchu Ni! - Bulk Email Sender
### Made by: M3shour 
### License: Apache 2.0
### Description: This script is used to send bulk emails to a list of recipients. The script reads the recipient email addresses and names from an Excel file and sends personalized emails to each recipient. The script uses the smtplib library to send emails via an SMTP server and the imaplib library to save the sent emails to the IMAP "Sent" folder. The script also supports sending attachments with the emails. The email content is read from an HTML template file, which can be personalized with the recipient's name. The script logs the email sending status to a log file.

import imaplib
import smtplib
import re
import time
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
import pandas as pd
from tqdm import tqdm
import logging
import argparse
import os
from typing import List, Tuple, Optional, Dict, Any

# Constants
DEFAULT_SMTP_PORT = 587
DEFAULT_IMAP_PORT = 993
DEFAULT_SENT_FOLDER = 'Sent'
DEFAULT_LOG_FILE = 'email_sending.log'
DEFAULT_DELAY = 0
MAX_RETRIES = 3
INITIAL_RETRY_DELAY = 1
BACKOFF_FACTOR = 2

# Email validation regex pattern
EMAIL_REGEX = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')


def validate_email(email: str) -> bool:
    """Validate email address using regex pattern."""
    if not email or not isinstance(email, str):
        return False
    return EMAIL_REGEX.match(email.strip()) is not None


def validate_file_path(file_path: str, description: str) -> None:
    """Validate that a file exists and is readable."""
    if not file_path or not os.path.exists(file_path):
        raise ValueError(f"{description} does not exist: {file_path}")
    if not os.path.isfile(file_path):
        raise ValueError(f"{description} is not a file: {file_path}")
    if not os.access(file_path, os.R_OK):
        raise ValueError(f"{description} is not readable: {file_path}")


def get_credentials_from_env() -> Dict[str, Optional[str]]:
    """Get email credentials from environment variables as fallback."""
    return {
        'sender_email': os.getenv('SENDER_EMAIL'),
        'sender_password': os.getenv('SENDER_PASSWORD')
    }


def create_email_message(sender_email: str, recipient_email: str, recipient_name: Optional[str], 
                         subject: str, html_content: str, attachment_path: Optional[str] = None,
                         cc_emails: Optional[List[str]] = None, bcc_emails: Optional[List[str]] = None) -> MIMEMultipart:
    """Create a complete email message with optional attachments and CC/BCC."""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    
    # Add CC and BCC if provided
    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)
    if bcc_emails:
        msg['Bcc'] = ', '.join(bcc_emails)
    
    msg['Subject'] = subject
    
    # Handle recipient name gracefully
    display_name = recipient_name if recipient_name and recipient_name.strip() else recipient_email
    msg['To'] = formataddr((recipient_name or '', recipient_email)) if recipient_name else recipient_email
    
    # Personalize the message
    personalized_content = html_content.replace("{{recipient_name}}", recipient_name or "there")
    msg.attach(MIMEText(personalized_content, 'html'))
    
    # Attach file if provided
    if attachment_path:
        with open(attachment_path, 'rb') as attachment_file:
            attachment = MIMEApplication(attachment_file.read(), _subtype='pdf')
            attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
            msg.attach(attachment)
    
    return msg


def send_email_with_retry(server: smtplib.SMTP, msg: MIMEMultipart, recipient_email: str, 
                          recipient_name: Optional[str], max_retries: int = MAX_RETRIES) -> bool:
    """Send email with exponential backoff retry logic."""
    for attempt in range(max_retries):
        try:
            server.send_message(msg)
            return True
        except (smtplib.SMTPException, smtplib.SMTPServerDisconnected, ConnectionError, TimeoutError) as e:
            if attempt == max_retries - 1:
                raise e
            
            delay = INITIAL_RETRY_DELAY * (BACKOFF_FACTOR ** attempt) + random.uniform(0, 1)
            logging.warning(f"Attempt {attempt + 1} failed for {recipient_email}. Retrying in {delay:.2f}s: {str(e)}")
            time.sleep(delay)
    
    return False


def get_emails_and_names_from_excel(file_path: str, sheet_name: str, email_column: str, 
                                   name_column: Optional[str] = None, cc_column: Optional[str] = None,
                                   bcc_column: Optional[str] = None) -> Tuple[List[Tuple[str, Optional[str], List[str], List[str]]], List[str]]:
    """Extract recipient data from Excel file with validation."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        raise ValueError(f"Failed to read Excel file: {str(e)}")
    
    # Validate required columns exist
    if email_column not in df.columns:
        raise ValueError(f"Email column '{email_column}' not found in Excel file")
    
    recipients = []
    invalid_emails = []
    
    for index, row in df.iterrows():
        email = str(row[email_column]).strip() if pd.notna(row[email_column]) else ''
        
        # Validate email
        if not validate_email(email):
            invalid_emails.append(email)
            continue
        
        # Get name (handle None/missing gracefully)
        name = str(row[name_column]).strip() if name_column and pd.notna(row[name_column]) and row[name_column] else None
        
        # Get CC emails
        cc_emails = []
        if cc_column and cc_column in df.columns and pd.notna(row[cc_column]):
            cc_list = str(row[cc_column]).split(',')
            cc_emails = [cc.strip() for cc in cc_list if validate_email(cc.strip())]
        
        # Get BCC emails
        bcc_emails = []
        if bcc_column and bcc_column in df.columns and pd.notna(row[bcc_column]):
            bcc_list = str(row[bcc_column]).split(',')
            bcc_emails = [bcc.strip() for bcc in bcc_list if validate_email(bcc.strip())]
        
        recipients.append((email, name, cc_emails, bcc_emails))
    
    return recipients, invalid_emails


def send_bulk_emails(sender_email: str, sender_password: str, subject: str, 
                    recipients: List[Tuple[str, Optional[str], List[str], List[str]]], 
                    imap_server: str, smtp_server: str, attachment_path: Optional[str] = None, 
                    append_attachment: bool = True, email_template: str = '', 
                    log_file: str = DEFAULT_LOG_FILE, delay_between_sends: float = DEFAULT_DELAY,
                    dry_run: bool = False, smtp_port: int = DEFAULT_SMTP_PORT, 
                    imap_port: int = DEFAULT_IMAP_PORT, sent_folder: str = DEFAULT_SENT_FOLDER) -> Dict[str, int]:
    """
    Send bulk emails with enhanced error handling and features.
    
    Returns:
        Dict with counts: {'success': int, 'failed': int, 'skipped': int}
    """
    # Configure logging once at startup
    logging.basicConfig(filename=log_file, level=logging.INFO, 
                       format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)
    
    stats = {'success': 0, 'failed': 0, 'skipped': 0}
    
    # Validate files before starting
    validate_file_path(email_template, "Email template file")
    if attachment_path:
        validate_file_path(attachment_path, "Attachment file")
    
    # Read email template
    try:
        with open(email_template, 'r', encoding='utf-8') as file:
            html_template = file.read()
    except Exception as e:
        logger.error(f"Failed to read email template: {str(e)}")
        raise ValueError(f"Failed to read email template: {str(e)}")
    
    if dry_run:
        logger.info("DRY RUN MODE: No emails will be sent")
        print("DRY RUN MODE: No emails will be sent")
    
    # Set up connections
    server = None
    imap_conn = None
    
    try:
        if not dry_run:
            # SMTP connection with timeout
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            server.starttls()
            server.login(sender_email, sender_password)
            
            # IMAP connection with timeout
            imap_conn = imaplib.IMAP4_SSL(imap_server, imap_port)
            imap_conn.login(sender_email, sender_password)
        
        # Send emails
        with tqdm(recipients, desc="Sending emails", unit="email", 
                 bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]") as pbar:
            
            for recipient_email, recipient_name, cc_emails, bcc_emails in pbar:
                try:
                    # Update progress bar with current stats
                    pbar.set_postfix({
                        'Success': stats['success'], 
                        'Failed': stats['failed'], 
                        'Skipped': stats['skipped']
                    })
                    
                    # Create email message
                    msg = create_email_message(
                        sender_email, recipient_email, recipient_name, subject,
                        html_template, attachment_path, cc_emails, bcc_emails
                    )
                    
                    if dry_run:
                        print(f"[DRY RUN] Would send email to {recipient_name or 'Unknown'} ({recipient_email})")
                        if cc_emails:
                            print(f"[DRY RUN] CC: {', '.join(cc_emails)}")
                        if bcc_emails:
                            print(f"[DRY RUN] BCC: {', '.join(bcc_emails)}")
                        stats['success'] += 1
                        logger.info(f"[DRY RUN] Email prepared for {recipient_name} ({recipient_email})")
                    else:
                        # Send email with retry logic
                        send_email_with_retry(server, msg, recipient_email, recipient_name)
                        stats['success'] += 1
                        logger.info(f"Email sent to {recipient_name} ({recipient_email})")
                        print(f"Email sent to {recipient_name or 'Unknown'} ({recipient_email})")
                        
                        # Save to IMAP Sent folder
                        try:
                            if not append_attachment and attachment_path:
                                # Create message without attachment for IMAP
                                msg_without_pdf = create_email_message(
                                    sender_email, recipient_email, recipient_name, subject,
                                    html_template, None, cc_emails, bcc_emails
                                )
                                imap_conn.append(sent_folder, '\\Seen', 
                                               imaplib.Time2Internaldate(time.time()), 
                                               msg_without_pdf.as_bytes())
                            else:
                                imap_conn.append(sent_folder, '\\Seen', 
                                               imaplib.Time2Internaldate(time.time()), 
                                               msg.as_bytes())
                        except Exception as e:
                            logger.warning(f"Failed to save to Sent folder for {recipient_email}: {str(e)}")
                    
                    # Rate limiting
                    if delay_between_sends > 0:
                        time.sleep(delay_between_sends)
                        
                except Exception as e:
                    stats['failed'] += 1
                    error_msg = f"Failed to send email to {recipient_name} ({recipient_email}): {str(e)}"
                    logger.error(error_msg)
                    print(error_msg)
        
        logger.info(f"Email sending completed. Success: {stats['success']}, Failed: {stats['failed']}, Skipped: {stats['skipped']}")
        print(f"\nEmail sending completed. Success: {stats['success']}, Failed: {stats['failed']}, Skipped: {stats['skipped']}")
        
    except Exception as e:
        logger.error(f"Fatal error in send_bulk_emails: {str(e)}")
        raise
    finally:
        # Clean up connections
        if server:
            try:
                server.quit()
            except:
                pass
        if imap_conn:
            try:
                imap_conn.logout()
            except:
                pass
    
    return stats


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Send bulk emails with enhanced features and robustness.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic usage
  python3 Email_sender.py --sender_email user@example.com --sender_password pass --subject "Hello" --excel_file_path list.xlsx --sheet_name Sheet1 --email-column Email --name-column Name --imap-server imap.example.com --smtp-server smtp.example.com --email-template template.html

  # With CC/BCC and rate limiting
  python3 Email_sender.py --sender_email user@example.com --sender_password pass --subject "Hello" --excel_file_path list.xlsx --sheet_name Sheet1 --email-column Email --name-column Name --cc-column CC --bcc-column BCC --delay 2 --imap-server imap.example.com --smtp-server smtp.example.com --email-template template.html

  # Dry run to preview
  python3 Email_sender.py --sender_email user@example.com --sender_password pass --subject "Hello" --excel_file_path list.xlsx --sheet_name Sheet1 --email-column Email --name-column Name --imap-server imap.example.com --smtp-server smtp.example.com --email-template template.html --dry-run
        """
    )
    
    # Required arguments
    parser.add_argument('--sender_email', type=str, help='Sender email address (or set SENDER_EMAIL env var)')
    parser.add_argument('--sender_password', type=str, help='Sender email password (or set SENDER_PASSWORD env var)')
    parser.add_argument('--subject', type=str, help='Email subject', required=True)
    parser.add_argument('--excel_file_path', type=str, help='Path to the Excel file', required=True)
    parser.add_argument('--sheet_name', type=str, help='Sheet name in the Excel file', required=True)
    parser.add_argument('--email-column', type=str, help='Column name containing email addresses', required=True)
    parser.add_argument('--imap-server', type=str, help='IMAP server address', required=True)
    parser.add_argument('--smtp-server', type=str, help='SMTP server address', required=True)
    parser.add_argument('--email-template', type=str, help='Path to HTML email template', required=True)
    
    # Optional arguments
    parser.add_argument('--name-column', type=str, help='Column name containing recipient names')
    parser.add_argument('--cc-column', type=str, help='Column name containing CC email addresses (comma-separated)')
    parser.add_argument('--bcc-column', type=str, help='Column name containing BCC email addresses (comma-separated)')
    parser.add_argument('--attachment_path', type=str, help='Path to attachment file')
    parser.add_argument('--append-attachment', action='store_true', help='Save attachment in IMAP Sent folder (default: False)')
    parser.add_argument('--log-file', type=str, default=DEFAULT_LOG_FILE, help=f'Log file path (default: {DEFAULT_LOG_FILE})')
    parser.add_argument('--delay', type=float, default=DEFAULT_DELAY, help=f'Delay between sends in seconds (default: {DEFAULT_DELAY})')
    parser.add_argument('--dry-run', action='store_true', help='Preview emails without sending')
    parser.add_argument('--smtp-port', type=int, default=DEFAULT_SMTP_PORT, help=f'SMTP port (default: {DEFAULT_SMTP_PORT})')
    parser.add_argument('--imap-port', type=int, default=DEFAULT_IMAP_PORT, help=f'IMAP port (default: {DEFAULT_IMAP_PORT})')
    parser.add_argument('--sent-folder', type=str, default=DEFAULT_SENT_FOLDER, help=f'IMAP Sent folder name (default: {DEFAULT_SENT_FOLDER})')
    
    args = parser.parse_args()
    
    # Get credentials from environment variables as fallback
    env_creds = get_credentials_from_env()
    sender_email = args.sender_email or env_creds['sender_email']
    sender_password = args.sender_password or env_creds['sender_password']
    
    if not sender_email:
        parser.error("sender_email is required (either via --sender_email argument or SENDER_EMAIL environment variable)")
    if not sender_password:
        parser.error("sender_password is required (either via --sender_password argument or SENDER_PASSWORD environment variable)")
    
    # Validate required files exist
    try:
        validate_file_path(args.excel_file_path, "Excel file")
        validate_file_path(args.email_template, "Email template")
        if args.attachment_path:
            validate_file_path(args.attachment_path, "Attachment file")
    except ValueError as e:
        parser.error(str(e))
    
    try:
        # Get recipient data from Excel
        recipients, invalid_emails = get_emails_and_names_from_excel(
            args.excel_file_path, args.sheet_name, args.email_column,
            args.name_column, args.cc_column, args.bcc_column
        )
        
        if invalid_emails:
            print(f"Warning: Found {len(invalid_emails)} invalid email addresses and skipped them:")
            for email in invalid_emails[:10]:  # Show first 10
                print(f"  - {email}")
            if len(invalid_emails) > 10:
                print(f"  ... and {len(invalid_emails) - 10} more")
        
        if not recipients:
            parser.error("No valid recipients found in the Excel file")
        
        print(f"Found {len(recipients)} valid recipients to process")
        
        # Send emails
        stats = send_bulk_emails(
            sender_email=sender_email,
            sender_password=sender_password,
            subject=args.subject,
            recipients=recipients,
            imap_server=args.imap_server,
            smtp_server=args.smtp_server,
            attachment_path=args.attachment_path,
            append_attachment=args.append_attachment,
            email_template=args.email_template,
            log_file=args.log_file,
            delay_between_sends=args.delay,
            dry_run=args.dry_run,
            smtp_port=args.smtp_port,
            imap_port=args.imap_port,
            sent_folder=args.sent_folder
        )
        
        # Final summary
        if args.dry_run:
            print(f"\nDRY RUN COMPLETED: Would process {stats['success']} emails")
        else:
            if stats['failed'] > 0:
                print(f"\nWARNING: {stats['failed']} emails failed to send. Check the log file for details.")
            else:
                print(f"\nSUCCESS: All {stats['success']} emails sent successfully!")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        exit(1)
