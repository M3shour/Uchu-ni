# Uchu-ni
a Python script to pull names and emails from an excel sheet and send an email to every email in the list with a customized message that uses their name.
# Installation
```bash
git clone https://github.com/M3shour/Uchu-ni.git
cd Uchu-ni
pip3 -m requirements.txt
```
# Usage 
```bash
usage: Email_sender.py [-h] --sender_email SENDER_EMAIL --sender_password SENDER_PASSWORD --subject SUBJECT --excel_file_path EXCEL_FILE_PATH
                       --sheet_name SHEET_NAME --email-column EMAIL_COLUMN [--name-column NAME_COLUMN] --imap-server IMAP_SERVER
                       --smtp-server SMTP_SERVER [--attachment_path ATTACHMENT_PATH] [--append-attachment APPEND_ATTACHMENT] --email-template
                       EMAIL_TEMPLATE
```

# Parameters

1. ***--sender_email:*** The email you intend to send from - required
2. ***--sender_password:*** The password of the email above - required
3. ***--subject:*** subject of your email - required
4. ***--excel_file_path:*** The path of the excel file that contains the emails - required
5. ***--sheet_name:*** the name of the excel sheet - required
6. ***--email-column:*** the name of the column that contains the emails of the recipients  - required
7. ***--name-column:*** The name of the column that contains the names of the recipients
8. ***--imap-server:*** The address of the imap server to connect to - required
9. ***--smtp-server:*** The address of the imap server to connect to - required
10. ***--attachment_path:*** path to attachment to send with the emails 
11. ***--append-attachment:*** choose if you want to append the attachment to the imap server or not(useful for large attachments)
12. ***--email-template:*** path to HTML file that contains the structure and content of your emails required

***Made by M3shour***
