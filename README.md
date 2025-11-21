# Uchu-ni
A Python script to pull names and emails from an excel sheet and send an email to every email in the list with a customized message that uses their name.

## Available Interfaces

### 1. Text User Interface (TUI) ⭐ New!
For interactive use with a friendly interface. Simply run:
```bash
python3 email_tui.py
```

The TUI provides:
- 🎨 Beautiful, colorful interface
- 📋 Step-by-step configuration wizard
- 👀 Live preview of recipients
- ✅ Interactive progress tracking
- 🔧 Preset server configurations (Gmail, Outlook, Yahoo)
- 📊 Excel file browser with column selection
- 🎯 Built-in dry run mode
- 📝 Template preview
- ⚙️ Advanced configuration options

### 2. Command Line Interface (CLI)
For automation and scripting use cases. See CLI Usage section below.

## Enhanced Features
This enhanced version includes robustness improvements and new features:
- **Email validation** with regex pattern checking
- **Retry logic** with exponential backoff for failed sends
- **CC/BCC support** from Excel columns
- **Dry run mode** to preview emails without sending
- **Rate limiting** to avoid server throttling
- **Environment variable support** for credentials
- **Enhanced error handling** and logging
- **File validation** before processing
- **Configurable ports** and settings
- **Progress tracking** with success/failure counts
- **Interactive TUI** for easy configuration

# Installation
```bash
git clone https://github.com/M3shour/Uchu-ni.git
cd Uchu-ni
pip3 install -r requirements.txt
```

# Usage 
```bash
usage: Email_sender.py [-h] [--sender_email SENDER_EMAIL]
                       [--sender_password SENDER_PASSWORD] --subject SUBJECT
                       --excel_file_path EXCEL_FILE_PATH --sheet_name SHEET_NAME
                       --email-column EMAIL_COLUMN --imap-server IMAP_SERVER
                       --smtp-server SMTP_SERVER --email-template EMAIL_TEMPLATE
                       [--name-column NAME_COLUMN] [--cc-column CC_COLUMN]
                       [--bcc-column BCC_COLUMN] [--attachment_path ATTACHMENT_PATH]
                       [--append-attachment] [--log-file LOG_FILE] [--delay DELAY]
                       [--dry-run] [--smtp-port SMTP_PORT] [--imap-port IMAP_PORT]
                       [--sent-folder SENT_FOLDER]
```

# Parameters

## Required Parameters
1. **--sender_email**: Sender email address (or set `SENDER_EMAIL` environment variable)
2. **--sender_password**: Sender email password (or set `SENDER_PASSWORD` environment variable)
3. **--subject**: Email subject
4. **--excel_file_path**: Path to the Excel file containing recipient data
5. **--sheet_name**: Sheet name in the Excel file
6. **--email-column**: Column name containing email addresses
7. **--imap-server**: IMAP server address for saving sent emails
8. **--smtp-server**: SMTP server address for sending emails
9. **--email-template**: Path to HTML email template file

## Optional Parameters
10. **--name-column**: Column name containing recipient names (for personalization)
11. **--cc-column**: Column name containing CC email addresses (comma-separated)
12. **--bcc-column**: Column name containing BCC email addresses (comma-separated)
13. **--attachment_path**: Path to attachment file (PDF supported)
14. **--append-attachment**: Save attachment in IMAP Sent folder (default: False)
15. **--log-file**: Log file path (default: email_sending.log)
16. **--delay**: Delay between sends in seconds (default: 0)
17. **--dry-run**: Preview emails without sending
18. **--smtp-port**: SMTP port (default: 587)
19. **--imap-port**: IMAP port (default: 993)
20. **--sent-folder**: IMAP Sent folder name (default: Sent)

# Examples

## TUI Usage (Recommended for most users)
```bash
python3 email_tui.py
```
Follow the interactive wizard to:
1. Enter your email credentials (or use environment variables)
2. Choose server preset (Gmail, Outlook, Yahoo) or enter custom
3. Select Excel file and configure columns
4. Set up email content and attachments
5. Configure advanced options (rate limiting, ports, etc.)
6. Preview recipients and send with dry-run option

## CLI Usage
```bash
python3 Email_sender.py --sender_email "user@example.com" --sender_password "password" --subject "Hello World" --excel_file_path "Maillist.xlsx" --sheet_name "Sheet1" --email-column "Emails" --name-column "Names" --imap-server "imaps.example.com" --smtp-server "smtp.example.com" --email-template "email_template.html"
```

## With CC/BCC and Rate Limiting
```bash
python3 Email_sender.py --sender_email "user@example.com" --sender_password "password" --subject "Hello World" --excel_file_path "Maillist.xlsx" --sheet_name "Sheet1" --email-column "Emails" --name-column "Names" --cc-column "CC" --bcc-column "BCC" --delay 2 --imap-server "imaps.example.com" --smtp-server "smtp.example.com" --email-template "email_template.html"
```

## Dry Run to Preview
```bash
python3 Email_sender.py --sender_email "user@example.com" --sender_password "password" --subject "Hello World" --excel_file_path "Maillist.xlsx" --sheet_name "Sheet1" --email-column "Emails" --name-column "Names" --imap-server "imaps.example.com" --smtp-server "smtp.example.com" --email-template "email_template.html" --dry-run
```

## Using Environment Variables
```bash
export SENDER_EMAIL="user@example.com"
export SENDER_PASSWORD="password"
python3 Email_sender.py --subject "Hello World" --excel_file_path "Maillist.xlsx" --sheet_name "Sheet1" --email-column "Emails" --name-column "Names" --imap-server "imaps.example.com" --smtp-server "smtp.example.com" --email-template "email_template.html"
```

## With Attachment
```bash
python3 Email_sender.py --sender_email "user@example.com" --sender_password "password" --subject "Hello World" --excel_file_path "Maillist.xlsx" --sheet_name "Sheet1" --email-column "Emails" --name-column "Names" --attachment_path "document.pdf" --append-attachment --imap-server "imaps.example.com" --smtp-server "smtp.example.com" --email-template "email_template.html"
```

# Excel File Format

Your Excel file should contain columns for:
- **Email** (required): Email addresses of recipients
- **Name** (optional): Names for personalization (used in `{{recipient_name}}` placeholder)
- **CC** (optional): CC recipients (comma-separated)
- **BCC** (optional): BCC recipients (comma-separated)

Example Excel structure:
| Email | Name | CC | BCC |
|-------|------|-----|-----|
| john@example.com | John Doe | jane@example.com | admin@example.com |
| jane@example.com | Jane Smith | | |

# Email Template

Your HTML template should use the `{{recipient_name}}` placeholder for personalization:

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>
<body>
    <h1>Hello {{recipient_name}}!</h1>
    <p>This is your personalized message.</p>
    <p>Best regards,<br>Sender</p>
</body>
</html>
```

If no recipient name is available, "there" will be used as fallback.

# Error Handling

The script includes comprehensive error handling:
- Invalid email addresses are skipped with warnings
- Failed email sends are retried with exponential backoff
- Connection timeouts and network errors are handled gracefully
- Detailed logging provides troubleshooting information
- File validation prevents processing of missing/corrupt files

# Logging

All operations are logged to the specified log file (default: `email_sending.log`). Log entries include:
- Successful email sends
- Failed email attempts with error details
- Retry attempts
- File validation errors
- Connection issues

# Security Considerations

- Use environment variables for sensitive credentials instead of command-line arguments
- The script validates email addresses to prevent injection attacks
- File paths are validated to prevent directory traversal
- Connection timeouts prevent hanging operations

***Made by M3shour***
