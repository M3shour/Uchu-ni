#!/usr/bin/env python3
"""
Uchu-ni Bulk Email Sender - TUI Edition
A text-based user interface for the enhanced bulk email sender.
"""

import os
from typing import Dict, List, Tuple, Optional, Any
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.prompt import Prompt, Confirm, IntPrompt
from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn
from rich import box as rich_box
from rich.text import Text

# Import from the main email sender module
from Email_sender import (
    validate_email, validate_file_path, get_emails_and_names_from_excel,
    send_bulk_emails, DEFAULT_SMTP_PORT, DEFAULT_IMAP_PORT, 
    DEFAULT_SENT_FOLDER, DEFAULT_LOG_FILE, DEFAULT_DELAY
)

console = Console()

class EmailSenderTUI:
    def __init__(self):
        self.config: Dict[str, Any] = {}
    
    def display_welcome(self):
        """Display welcome screen."""
        welcome_text = """
[bold blue]🚀 Uchu-ni Bulk Email Sender - TUI Edition[/bold blue]

[green]Enhanced Features:[/green]
• Email validation with regex checking
• Retry logic with exponential backoff  
• CC/BCC support from Excel columns
• Dry run mode for preview
• Rate limiting to avoid throttling
• Environment variable support
• Enhanced error handling & logging

[dim]Press Enter to continue...[/dim]
        """
        console.print(Panel(welcome_text.strip(), title="Welcome", border_style="blue"))
        input()
    
    def get_file_input(self, prompt_text: str, file_description: str) -> str:
        """Get file input with validation."""
        while True:
            file_path = Prompt.ask(prompt_text)
            if not file_path:
                console.print("[red]Error: File path cannot be empty[/red]")
                continue
                
            try:
                validate_file_path(file_path, file_description)
                return file_path
            except ValueError as e:
                console.print(f"[red]Error: {e}[/red]")
    
    def get_email_input(self, prompt_text: str, env_var: Optional[str] = None) -> str:
        """Get email input with validation and environment variable support."""
        # Check environment variable first
        if env_var:
            env_value = os.getenv(env_var)
            if env_value:
                if Confirm.ask(f"Found {env_var} in environment: '{env_value}'. Use it?"):
                    return env_value
        
        while True:
            email = Prompt.ask(prompt_text)
            if validate_email(email):
                return email
            console.print("[red]Error: Invalid email address format[/red]")
    
    def get_server_config(self) -> Dict[str, str]:
        """Get SMTP/IMAP server configuration."""
        console.print("\n[bold blue]📧 Server Configuration[/bold blue]")
        
        # Check for common server presets
        presets = {
            "gmail": {"smtp": "smtp.gmail.com", "imap": "imap.gmail.com"},
            "outlook": {"smtp": "smtp-mail.outlook.com", "imap": "outlook.office365.com"},
            "yahoo": {"smtp": "smtp.mail.yahoo.com", "imap": "imap.mail.yahoo.com"},
            "custom": {"smtp": "", "imap": ""}
        }
        
        console.print("[yellow]Common server presets:[/yellow]")
        for key, value in presets.items():
            if key != "custom":
                console.print(f"  {key.title()}: SMTP={value['smtp']}, IMAP={value['imap']}")
        
        preset_choice = Prompt.ask(
            "\nChoose preset or enter 'custom'",
            choices=list(presets.keys()),
            default="custom"
        )
        
        if preset_choice != "custom":
            return presets[preset_choice]
        
        # Custom server configuration
        while True:
            smtp_server = Prompt.ask("SMTP server address")
            imap_server = Prompt.ask("IMAP server address")
            
            if smtp_server and imap_server:
                return {"smtp": smtp_server, "imap": imap_server}
            
            console.print("[red]Error: Both server addresses are required[/red]")
    
    def get_recipient_preview(self, file_path: str, sheet_name: str, 
                            email_col: str, name_col: Optional[str] = None,
                            cc_col: Optional[str] = None, bcc_col: Optional[str] = None) -> Tuple[List, List]:
        """Get preview of recipients from Excel file."""
        try:
            recipients, invalid_emails = get_emails_and_names_from_excel(
                file_path, sheet_name, email_col, name_col, cc_col, bcc_col
            )
            return recipients, invalid_emails
        except Exception as e:
            console.print(f"[red]Error reading Excel file: {e}[/red]")
            return [], []
    
    def display_recipient_preview(self, recipients: List, invalid_emails: List):
        """Display preview of recipients in a table."""
        console.print("\n[bold blue]👥 Recipient Preview[/bold blue]")
        
        # Valid recipients table
        if recipients:
            table = Table(title="Valid Recipients", box=rich_box.ROUNDED)
            table.add_column("Email", style="cyan")
            table.add_column("Name", style="green")
            table.add_column("CC", style="yellow")
            table.add_column("BCC", style="magenta")
            
            # Show first 10 recipients
            for email, name, cc, bcc in recipients[:10]:
                cc_str = ", ".join(cc) if cc else "None"
                bcc_str = ", ".join(bcc) if bcc else "None"
                table.add_row(email, name or "None", cc_str, bcc_str)
            
            if len(recipients) > 10:
                table.add_row("...", "...", "...", "...", f"... and {len(recipients) - 10} more")
            
            console.print(table)
            console.print(f"[green]✓ Total valid recipients: {len(recipients)}[/green]")
        else:
            console.print("[red]No valid recipients found![/red]")
        
        # Invalid emails
        if invalid_emails:
            console.print(f"\n[red]⚠ Invalid emails found ({len(invalid_emails)}):[/red]")
            for email in invalid_emails[:5]:
                console.print(f"  • {email}")
            if len(invalid_emails) > 5:
                console.print(f"  ... and {len(invalid_emails) - 5} more")
    
    def get_excel_config(self) -> Dict[str, str]:
        """Get Excel file configuration."""
        console.print("\n[bold blue]📊 Excel File Configuration[/bold blue]")
        
        # Get file path
        file_path = self.get_file_input("Excel file path", "Excel file")
        
        # Get sheet name
        try:
            import pandas as pd
            df = pd.read_excel(file_path, sheet_name=None)
            sheets = list(df.keys())
            
            if len(sheets) == 1:
                sheet_name = sheets[0]
                console.print(f"[green]Found single sheet: '{sheet_name}'[/green]")
            else:
                console.print(f"\n[yellow]Available sheets:[/yellow]")
                for i, sheet in enumerate(sheets, 1):
                    console.print(f"  {i}. {sheet}")
                
                sheet_choice = IntPrompt.ask(
                    "Choose sheet number",
                    choices=list(range(1, len(sheets) + 1)),
                    default=1
                )
                sheet_name = sheets[sheet_choice - 1]
        except Exception as e:
            console.print(f"[red]Error reading Excel file: {e}[/red]")
            sheet_name = Prompt.ask("Sheet name")
        
        # Get column names
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            columns = list(df.columns)
            
            console.print(f"\n[yellow]Available columns in '{sheet_name}':[/yellow]")
            for i, col in enumerate(columns, 1):
                console.print(f"  {i}. {col}")
            
            # Email column (required)
            email_choice = IntPrompt.ask(
                "Choose email column number",
                choices=list(range(1, len(columns) + 1))
            )
            email_col = columns[email_choice - 1]
            
            # Name column (optional)
            name_col = None
            if Confirm.ask("Do you have a name column?"):
                name_choice = IntPrompt.ask(
                    "Choose name column number",
                    choices=list(range(1, len(columns) + 1))
                )
                name_col = columns[name_choice - 1]
            
            # CC column (optional)
            cc_col = None
            if Confirm.ask("Do you have a CC column?"):
                cc_choice = IntPrompt.ask(
                    "Choose CC column number",
                    choices=list(range(1, len(columns) + 1))
                )
                cc_col = columns[cc_choice - 1]
            
            # BCC column (optional)
            bcc_col = None
            if Confirm.ask("Do you have a BCC column?"):
                bcc_choice = IntPrompt.ask(
                    "Choose BCC column number",
                    choices=list(range(1, len(columns) + 1))
                )
                bcc_col = columns[bcc_choice - 1]
                
        except Exception as e:
            console.print(f"[red]Error reading columns: {e}[/red]")
            email_col = Prompt.ask("Email column name")
            name_col = Prompt.ask("Name column name (optional)", default="")
            cc_col = Prompt.ask("CC column name (optional)", default="")
            bcc_col = Prompt.ask("BCC column name (optional)", default="")
        
        config = {
            "file_path": file_path,
            "sheet_name": sheet_name,
            "email_col": email_col,
            "name_col": name_col,
            "cc_col": cc_col,
            "bcc_col": bcc_col
        }
        
        # Show preview
        recipients, invalid_emails = self.get_recipient_preview(**config)
        self.display_recipient_preview(recipients, invalid_emails)
        
        if not recipients:
            console.print("[red]No valid recipients found. Please check your Excel file configuration.[/red]")
            return self.get_excel_config()
        
        return config
    
    def get_email_config(self) -> Dict[str, str]:
        """Get email content configuration."""
        console.print("\n[bold blue]📝 Email Content Configuration[/bold blue]")
        
        subject = Prompt.ask("Email subject")
        
        template_path = self.get_file_input("Email template file path", "Email template")
        
        # Preview template
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_content = f.read()
            
            console.print("\n[yellow]Template preview (first 300 chars):[/yellow]")
            console.print(Panel(template_content[:300] + "..." if len(template_content) > 300 else template_content))
            
            if "{{recipient_name}}" not in template_content:
                console.print("[yellow]⚠ Template doesn't contain {{recipient_name}} placeholder[/yellow]")
        except Exception as e:
            console.print(f"[red]Error reading template: {e}[/red]")
        
        # Attachment (optional)
        attachment_path = None
        if Confirm.ask("Do you want to add an attachment?"):
            attachment_path = self.get_file_input("Attachment file path", "Attachment")
        
        return {
            "subject": subject,
            "template_path": template_path,
            "attachment_path": attachment_path
        }
    
    def get_advanced_config(self) -> Dict[str, Any]:
        """Get advanced configuration options."""
        console.print("\n[bold blue]⚙ Advanced Configuration[/bold blue]")
        
        config = {}
        
        # Rate limiting
        if Confirm.ask("Do you want to add delay between sends?", default=False):
            config["delay"] = IntPrompt.ask("Delay between sends (seconds)", default=1)
        else:
            config["delay"] = DEFAULT_DELAY
        
        # IMAP settings
        config["smtp_port"] = IntPrompt.ask("SMTP port", default=DEFAULT_SMTP_PORT)
        config["imap_port"] = IntPrompt.ask("IMAP port", default=DEFAULT_IMAP_PORT)
        config["sent_folder"] = Prompt.ask("IMAP Sent folder name", default=DEFAULT_SENT_FOLDER)
        
        # Attachment handling
        if self.config.get("email", {}).get("attachment_path"):
            config["append_attachment"] = Confirm.ask(
                "Save attachment in IMAP Sent folder?", default=False
            )
        else:
            config["append_attachment"] = False
        
        # Logging
        if Confirm.ask("Customize log file path?", default=False):
            config["log_file"] = Prompt.ask("Log file path", default=DEFAULT_LOG_FILE)
        else:
            config["log_file"] = DEFAULT_LOG_FILE
        
        return config
    
    def confirm_and_send(self):
        """Show configuration summary and send emails."""
        console.print("\n[bold blue]📋 Configuration Summary[/bold blue]")
        
        # Create summary table
        table = Table(box=rich_box.ROUNDED)
        table.add_column("Setting", style="cyan")
        table.add_column("Value", style="white")
        
        # Sender info
        table.add_row("Sender Email", self.config["sender_email"])
        table.add_row("SMTP Server", self.config["smtp_server"])
        table.add_row("IMAP Server", self.config["imap_server"])
        
        # Excel config
        excel_config = self.config["excel"]
        table.add_row("Excel File", excel_config["file_path"])
        table.add_row("Sheet Name", excel_config["sheet_name"])
        table.add_row("Email Column", excel_config["email_col"])
        if excel_config["name_col"]:
            table.add_row("Name Column", excel_config["name_col"])
        if excel_config["cc_col"]:
            table.add_row("CC Column", excel_config["cc_col"])
        if excel_config["bcc_col"]:
            table.add_row("BCC Column", excel_config["bcc_col"])
        
        # Email config
        email_config = self.config["email"]
        table.add_row("Subject", email_config["subject"])
        table.add_row("Template", email_config["template_path"])
        if email_config["attachment_path"]:
            table.add_row("Attachment", email_config["attachment_path"])
        
        # Advanced config
        adv_config = self.config["advanced"]
        table.add_row("Delay (s)", str(adv_config["delay"]))
        if email_config["attachment_path"]:
            table.add_row("Save Attachment", "Yes" if adv_config["append_attachment"] else "No")
        
        console.print(table)
        
        # Confirm recipients count
        recipients, invalid_emails = self.get_recipient_preview(**self.config["excel"])
        console.print(f"\n[green]Ready to send {len(recipients)} emails[/green]")
        if invalid_emails:
            console.print(f"[yellow]{len(invalid_emails)} invalid emails will be skipped[/yellow]")
        
        # Final confirmation
        if not Confirm.ask("\n[bold red]Proceed with sending emails?[/bold red]"):
            console.print("[yellow]Operation cancelled.[/yellow]")
            return
        
        # Ask for dry run
        dry_run = False
        if Confirm.ask("Do you want to do a dry run first?", default=True):
            dry_run = True
        
        # Send emails
        self.send_emails_interactive(dry_run)
    
    def send_emails_interactive(self, dry_run: bool = False):
        """Send emails with interactive progress display."""
        console.print(f"\n[bold blue]{'🔍 DRY RUN MODE' if dry_run else '🚀 SENDING EMAILS'}[/bold blue]")
        
        try:
            # Get recipients
            recipients, invalid_emails = get_emails_and_names_from_excel(
                **self.config["excel"]
            )
            
            if not recipients:
                console.print("[red]No valid recipients to send to![/red]")
                return
            
            # Prepare parameters
            excel_config = self.config["excel"]
            email_config = self.config["email"]
            adv_config = self.config["advanced"]
            
            # Show progress with rich progress bar
            with Progress(
                "[progress.description]{task.description}",
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeRemainingColumn(),
                console=console
            ) as progress:
                
                task = progress.add_task(
                    f"{'Preparing' if dry_run else 'Sending'} emails...",
                    total=len(recipients)
                )
                
                stats = send_bulk_emails(
                    sender_email=self.config["sender_email"],
                    sender_password=self.config["sender_password"],
                    subject=email_config["subject"],
                    recipients=recipients,
                    imap_server=self.config["imap_server"],
                    smtp_server=self.config["smtp_server"],
                    attachment_path=email_config["attachment_path"],
                    append_attachment=adv_config["append_attachment"],
                    email_template=email_config["template_path"],
                    log_file=adv_config["log_file"],
                    delay_between_sends=adv_config["delay"],
                    dry_run=dry_run,
                    smtp_port=adv_config["smtp_port"],
                    imap_port=adv_config["imap_port"],
                    sent_folder=adv_config["sent_folder"]
                )
                
                progress.update(task, completed=len(recipients))
            
            # Show results
            console.print("\n[bold green]✅ {'Dry Run' if dry_run else 'Sending'} Completed![/bold green]")
            
            results_table = Table(box=rich_box.ROUNDED)
            results_table.add_column("Status", style="white")
            results_table.add_column("Count", justify="right", style="white")
            
            results_table.add_row("✅ Success", f"[green]{stats['success']}[/green]")
            results_table.add_row("❌ Failed", f"[red]{stats['failed']}[/red]")
            results_table.add_row("⏭️ Skipped", f"[yellow]{stats['skipped']}[/yellow]")
            
            console.print(results_table)
            
            if dry_run:
                if Confirm.ask("\nDry run successful. Send actual emails now?"):
                    self.send_emails_interactive(dry_run=False)
            elif stats['failed'] > 0:
                console.print(f"\n[yellow]⚠ {stats['failed']} emails failed to send. Check the log file for details.[/yellow]")
            else:
                console.print("\n[bold green]🎉 All emails sent successfully![/bold green]")
                
        except Exception as e:
            console.print(f"\n[red]❌ Error: {e}[/red]")
    
    def run(self):
        """Main TUI flow."""
        try:
            self.display_welcome()
            
            # Get sender credentials
            console.print("\n[bold blue]🔐 Sender Credentials[/bold blue]")
            self.config["sender_email"] = self.get_email_input(
                "Sender email address", "SENDER_EMAIL"
            )
            self.config["sender_password"] = Prompt.ask("Sender password", password=True)
            
            # Get server configuration
            servers = self.get_server_config()
            self.config["smtp_server"] = servers["smtp"]
            self.config["imap_server"] = servers["imap"]
            
            # Get Excel configuration
            self.config["excel"] = self.get_excel_config()
            
            # Get email configuration
            self.config["email"] = self.get_email_config()
            
            # Get advanced configuration
            self.config["advanced"] = self.get_advanced_config()
            
            # Confirm and send
            self.confirm_and_send()
            
        except KeyboardInterrupt:
            console.print("\n[yellow]Operation cancelled by user.[/yellow]")
        except Exception as e:
            console.print(f"\n[red]An error occurred: {e}[/red]")


def main():
    """Main entry point for TUI."""
    tui = EmailSenderTUI()
    tui.run()


if __name__ == "__main__":
    main()
