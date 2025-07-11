import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import getpass
import time
import ssl
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import logging
from logging.handlers import RotatingFileHandler
import traceback
import os
import base64
import json
import mailbox
import re

class EmailCampaignApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Campaign")
        self.root.geometry("1200x800")
        
        self.gmail_mode = tk.BooleanVar(value=False)
        self.previous_smtp_server = ""  # Add this line to store previous SMTP server
        self.previous_smtp_port = ""    # Add this line to store previous SMTP port

        # Setup logging
        self.setup_logging()
        
        # Configuration
        self.config = {
            'excel_file': "",
            'html_template': "",
            'test_mode': False,
            'log_file': "logs/email_campaign.log",
            'failure_report': "reports/failed_contacts.xlsx"
        }
        
        # Tracking failed sends
        self.failed_contacts = []
        
        # UI Setup
        self.setup_ui()
        self.load_config()
        
        if not self.smtp_entry.get():
            self.smtp_entry.insert(0, "smtp.gmail.com")
        if not self.port_entry.get():
            self.port_entry.insert(9, "465")

        self.log("Application initialized", "INFO")

    def setup_logging(self):
        """Configure logging to file and console"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                RotatingFileHandler(
                    "email_campaign.log",
                    maxBytes=5*1024*1024,  # 5MB
                    backupCount=3
                ),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger()

    def log(self, message, level="INFO"):
        """Log message with level and update UI log"""
        level = level.upper()
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = f"[{timestamp}] {level}: {message}"
        
        # Log to file/console
        if level == "INFO":
            self.logger.info(message)
        elif level == "WARNING":
            self.logger.warning(message)
        elif level == "ERROR":
            self.logger.error(message)
        elif level == "DEBUG":
            self.logger.debug(message)
        
        # Update UI log if available
        if hasattr(self, 'log_text'):
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, log_entry + "\n")
            
            # Color coding
            if level == "ERROR":
                self.log_text.tag_add("error", "end-2l linestart", "end-1c")
                self.log_text.tag_config("error", foreground="red")
            elif level == "WARNING":
                self.log_text.tag_add("warning", "end-2l linestart", "end-1c")
                self.log_text.tag_config("warning", foreground="orange")
            
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        
        # Update status bar for errors
        if level == "ERROR" and hasattr(self, 'status'):
            self.status.config(text=f"Error: {message[:60]}...")
    
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left panel - Configuration
        config_frame = ttk.LabelFrame(main_frame, text="Email Campaign Configuration", padding="10")
        config_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # Configure grid weights for config_frame
        config_frame.columnconfigure(1, weight=1)  # Make the middle column expandable
        
        # File Selection Section
        file_frame = ttk.LabelFrame(config_frame, text="1. Select Files", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)
        
        # Configure grid weights for file_frame
        file_frame.columnconfigure(1, weight=1)  # Make entry fields expandable
        
        ttk.Label(file_frame, text="Excel Contacts File:").grid(row=0, column=0, sticky="w", padx=5)
        self.excel_entry = ttk.Entry(file_frame)
        self.excel_entry.grid(row=0, column=1, sticky="we", padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_excel_file).grid(row=0, column=2, padx=5)
        
        ttk.Label(file_frame, text="HTML Template:").grid(row=1, column=0, sticky="w", padx=5)
        self.html_entry = ttk.Entry(file_frame)
        self.html_entry.grid(row=1, column=1, sticky="we", padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_html_file).grid(row=1, column=2, padx=5)
        
        # Sender Configuration Section
        sender_frame = ttk.LabelFrame(config_frame, text="2. Sender Configuration", padding="5")
        sender_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)

        sender_frame.columnconfigure(1, weight=1)

        ttk.Label(sender_frame, text="Your Name:").grid(row=0, column=0, sticky="w", padx=5)
        self.sender_name_entry = ttk.Entry(sender_frame)
        self.sender_name_entry.grid(row=0, column=1, sticky="we", padx=5, columnspan=2)
        
        ttk.Label(sender_frame, text="Your Email:").grid(row=1, column=0, sticky="w", padx=5)
        self.email_entry = ttk.Entry(sender_frame)
        self.email_entry.grid(row=1, column=1, sticky="we", padx=5, columnspan=2)
        
        ttk.Label(sender_frame, text="Password:").grid(row=2, column=0, sticky="w", padx=5)
        self.pass_entry = ttk.Entry(sender_frame, show="*")
        self.pass_entry.grid(row=2, column=1, sticky="we", padx=5, columnspan=2)

        self.app_pass_label = ttk.Label(sender_frame, text="App Password:")
        self.app_pass_entry = ttk.Entry(sender_frame, show="*")

        # Row 4 - Email Subject (new field)
        ttk.Label(sender_frame, text="Email Subject:").grid(row=4, column=0, sticky="w", padx=5)
        self.subject_entry = ttk.Entry(sender_frame)
        self.subject_entry.grid(row=4, column=1, sticky="we", padx=5, columnspan=2)
        self.subject_entry.insert(0, "About our services")  # Default subject
        
        # SMTP Configuration Section
        smtp_frame = ttk.LabelFrame(config_frame, text="3. Server Configuration", padding="5")
        smtp_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        
        # Gmail mode toggle
        self.gmail_mode = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            smtp_frame, 
            text="Use Gmail", 
            variable=self.gmail_mode,
            command=self.toggle_gmail_mode
        ).grid(row=0, column=0, sticky="w", padx=5)
        
        # Gmail help button (added here)
        self.gmail_help_btn = ttk.Button(
            smtp_frame,
            text="Gmail Setup Help",
            command=self.show_gmail_help
        )
        self.gmail_help_btn.grid(row=0, column=1, padx=5)
        self.gmail_help_btn.grid_remove()  # Initially hidden
        
        # Server settings
        ttk.Label(smtp_frame, text="SMTP Server:").grid(row=1, column=0, sticky="w", padx=5)
        self.smtp_entry = ttk.Entry(smtp_frame)
        self.smtp_entry.grid(row=1, column=1, sticky="we", padx=5, columnspan=2)
        
        ttk.Label(smtp_frame, text="SMTP Port:").grid(row=2, column=0, sticky="w", padx=5)
        self.port_entry = ttk.Entry(smtp_frame, width=8)
        self.port_entry.grid(row=2, column=1, sticky="w", padx=5)
        
        # Options Section
        options_frame = ttk.LabelFrame(config_frame, text="4. Campaign Options", padding="5")
        options_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5)
        
        self.remember_me = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, 
            text="Remember my credentials", 
            variable=self.remember_me
        ).grid(row=0, column=0, sticky="w", padx=5)
        
        self.test_mode = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame, 
            text="Test Mode (send to yourself)", 
            variable=self.test_mode
        ).grid(row=0, column=1, sticky="w", padx=5)
        
        # Action Buttons
        button_frame = ttk.Frame(config_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Load Contacts", command=self.load_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Preview Email", command=self.preview_email).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Send Campaign", command=self.start_campaign).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Import Failed", command=self.import_failed_contacts).pack(side=tk.LEFT, padx=5)
        
        # Right panel - Data Table and Log
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        # Configure notebook to expand properly
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Data Table Tab
        table_frame = ttk.Frame(notebook)
        notebook.add(table_frame, text="Contact List")
        
        # Configure table frame grid
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Treeview with improved resizing
        self.tree = ttk.Treeview(table_frame)
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbars - now properly attached to grid
        ysb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        ysb.grid(row=0, column=1, sticky="ns")
        xsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        xsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        
        # Log Tab
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="Activity Log")
        
        # Configure log frame grid
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # ScrolledText with proper expansion
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.config(state=tk.DISABLED)
        
        # Status bar
        self.status = ttk.Label(main_frame, text="Ready", relief=tk.SUNKEN)
        self.status.grid(row=1, column=0, columnspan=2, sticky="ew")
        
        # Configure grid weights for main container
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=3)  # Give more space to the notebook
        main_frame.rowconfigure(0, weight=1)  # Allow vertical expansion

    def toggle_gmail_mode(self):
        if self.gmail_mode.get():
            # Save current SMTP settings before switching to Gmail mode
            self.previous_smtp_server = self.smtp_entry.get()
            self.previous_smtp_port = self.port_entry.get()
            
            # Show App Password field, hide regular password
            self.app_pass_label.grid(row=3, column=0, sticky="w", padx=5)
            self.app_pass_entry.grid(row=3, column=1, sticky="we", padx=5, columnspan=2)
            self.pass_entry.grid_remove()
            self.gmail_help_btn.grid() # Button Visible

            self.subject_entry.grid(row=5, column=1, sticky="we", padx=5, columnspan=2)
            self.sender_frame.grid_rowconfigure(5, weight=1)
            # Set Gmail defaults
            self.smtp_entry.delete(0, tk.END)
            self.smtp_entry.insert(0, "smtp.gmail.com")
            self.port_entry.delete(0, tk.END)
            self.port_entry.insert(0, "465")
        else:
            # Show regular password, hide App Password
            self.pass_entry.grid()
            self.app_pass_label.grid_remove()
            self.app_pass_entry.grid_remove()
            self.gmail_help_btn.grid_remove() # Button Invisible
            self.subject_entry.grid(row=4, column=1, sticky="we", padx=5, columnspan=2)
            
            # Restore previous SMTP settings
            if self.previous_smtp_server:  # Only restore if we have a saved value
                self.smtp_entry.delete(0, tk.END)
                self.smtp_entry.insert(0, self.previous_smtp_server)
                self.port_entry.delete(0, tk.END)
                self.port_entry.insert(0, self.previous_smtp_port)

    def show_gmail_help(self):
        """Show Gmail configuration help"""
        help_msg = """Gmail Configuration Requirements:

1. For Gmail accounts WITH 2FA:
   - Create an App Password:
     Google Account > Security > App Passwords
   - Use this password in the app

2. For Gmail accounts WITHOUT 2FA:
   - Enable "Less Secure Apps":
     https://myaccount.google.com/lesssecureapps

3. You may need to unlock captcha:
   - Visit before sending:
     https://accounts.google.com/DisplayUnlockCaptcha

4. Sending Limits:
   - 500 emails per day
   - 100 recipients per message"""
        messagebox.showinfo("Gmail Setup Help", help_msg)

    def browse_excel_file(self):
        """Browse for Excel file and store absolute path"""
        initial_dir = os.path.dirname(self.excel_entry.get()) if self.excel_entry.get() else os.getcwd()
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if filename:
            abs_path = os.path.abspath(filename)
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, self._get_relative_path(abs_path))
            self.log(f"Selected Excel file: {abs_path}", "INFO")

    def browse_html_file(self):
        """Browse for HTML template and store absolute path"""
        initial_dir = os.path.dirname(self.html_entry.get()) if self.html_entry.get() else os.getcwd()
        filename = filedialog.askopenfilename(
            title="Select HTML Template",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if filename:
            abs_path = os.path.abspath(filename)
            self.html_entry.delete(0, tk.END)
            self.html_entry.insert(0, self._get_relative_path(abs_path))
            self.log(f"Selected HTML template: {abs_path}", "INFO")

    def load_data(self):
        """Load and display contact data from Excel"""
        excel_file = self.excel_entry.get()
        if not excel_file:
            error_msg = "Please select an Excel file first"
            self.log(error_msg, "ERROR")
            messagebox.showerror("Error", error_msg)
            return
            
        try:
            self.log(f"Loading contacts from {excel_file}", "INFO")
            self.df = pd.read_excel(excel_file)
            self.display_data()
            self.status.config(text=f"Loaded {len(self.df)} contacts")
            self.log(f"Successfully loaded {len(self.df)} contacts", "INFO")
        except Exception as e:
            error_msg = f"Failed to load Excel file: {str(e)}"
            self.log(error_msg, "ERROR")
            self.log(traceback.format_exc(), "DEBUG")
            messagebox.showerror("Error", error_msg)

    def display_data(self):
        """Display data in treeview"""
        # Clear existing data
        for i in self.tree.get_children():
            self.tree.delete(i)
            
        # Set up columns
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"
        
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.W)
        
        # Add data
        for _, row in self.df.iterrows():
            self.tree.insert("", tk.END, values=list(row))
        
        self.log("Contact data displayed in table", "INFO")

    def preview_email(self):
        """Show formatted email preview"""
        html_file = self.html_entry.get()
        if not html_file:
            error_msg = "Please select an HTML template first"
            self.log(error_msg, "ERROR")
            messagebox.showerror("Error", error_msg)
            return
        
        try:
            with open(html_file, 'r', encoding='utf-8') as f:
                content = f.read()
        
            preview = tk.Toplevel(self.root)
            preview.title("Email Preview")
            preview.geometry("800x600")
            
            # Simple text preview since tkinterhtml isn't available
            text_widget = scrolledtext.ScrolledText(preview, wrap=tk.WORD)
            text_widget.pack(fill=tk.BOTH, expand=True)
            text_widget.insert(tk.END, content)
            text_widget.config(state=tk.DISABLED)
        
            self.log("Email preview displayed", "INFO")
        except Exception as e:
            error_msg = f"Failed to load HTML template: {str(e)}"
            self.log(error_msg, "ERROR")
            self.log(traceback.format_exc(), "DEBUG")
            messagebox.showerror("Error", error_msg)

    def start_campaign(self):
        """Start email campaign with failure tracking"""
        self.failed_contacts = []  # Reset failed contacts list
        self.log("Starting email campaign", "INFO")
        
        # Validate inputs
        if not all([self.excel_entry.get(), self.html_entry.get(),
                   self.email_entry.get(), self.pass_entry.get()]):
            error_msg = "Please fill all required fields"
            self.log(error_msg, "ERROR")
            messagebox.showerror("Error", error_msg)
            return
        
        # Save config if "Remember Me" is checked
        if self.remember_me.get():
            self.save_config(
                email=self.email_entry.get(),
                password=self.pass_entry.get(),
                sender_name=self.sender_name_entry.get(),
                smtp_server=self.smtp_entry.get(),
                smtp_port=int(self.port_entry.get())
            )

        # Confirm before sending
        if not messagebox.askyesno("Confirm", "Start email campaign?"):
            self.log("Campaign canceled by user", "INFO")
            return
            
        # Get configuration
        self.config.update({
            'excel_file': self.excel_entry.get(),
            'html_template': self.html_entry.get(),
            'smtp_server': self.smtp_entry.get(),
            'smtp_port': int(self.port_entry.get()),
            'test_mode': self.test_mode.get(),
            'sender_email': self.email_entry.get(),
            'password': self.pass_entry.get(),
            'sender_name': self.sender_name_entry.get()
        })
        
        # Load data
        try:
            self.log(f"Loading contacts from {self.config['excel_file']}", "INFO")
            contacts = load_contacts(self.config['excel_file'])
            self.log(f"Found {len(contacts)} valid contacts", "INFO")
            
            self.log(f"Loading HTML template from {self.config['html_template']}", "INFO")
            with open(self.config['html_template'], 'r', encoding='utf-8') as f:
                html_template = f.read()
        except Exception as e:
            error_msg = f"Failed to load files: {str(e)}"
            self.log(error_msg, "ERROR")
            self.log(traceback.format_exc(), "DEBUG")
            messagebox.showerror("Error", error_msg)
            return
            
        # Start sending
        success_count = 0
        total = len(contacts)
        self.log(f"Starting to send {total} emails", "INFO")
        
        progress = tk.Toplevel(self.root)
        progress.title("Sending Progress")
        progress.geometry("600x400")
        
        ttk.Label(progress, text="Sending emails...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress, maximum=total, mode='determinate')
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        status_label = ttk.Label(progress, text="0/0 sent")
        status_label.pack()
        
        log_frame = ttk.LabelFrame(progress, text="Sending Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        progress_log = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        progress_log.pack(fill=tk.BOTH, expand=True)
        
        with open(self.config['html_template'], 'r', encoding = 'utf-8') as f:
            html_template = f.read()
        
        for i, contact in enumerate(contacts, 1):

            personalized_html = html_template.replace('{{name}}', contact['Nombre Contacto']).replace('{{company}}', contact.get('Nombre Empresa', '')).replace('{{sender_name}}', self.config['sender_name'])
            # Update UI
            progress_bar['value'] = i
            status_label.config(text=f"{i}/{total} sent")
            progress.update()
            
            # Prepare email
            company = contact.get('Nombre Empresa', '')
            subject = f"{self.config['sender_name']} - Servicios que entregan valor"
            recipient = self.config['sender_email'] if self.config['test_mode'] else contact['Email Contacto']
            
            log_msg = f"Sending to {recipient} ({contact['Nombre Contacto']} at {company})"
            self.log(log_msg, "INFO")
            progress_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {log_msg}\n")
            progress_log.see(tk.END)
            
            # Send email with detailed error handling
            try:
                result, error_detail = self.send_email(
                    self.config['sender_email'], self.config['sender_name'],
                    recipient, contact['Nombre Contacto'],
                    self.subject_entry.get(), 
                    personalized_html,
                    self.config['smtp_server'], 
                    int(self.config['smtp_port'])
                )
                
                if result:
                    success_msg = f"Successfully sent to {recipient}"
                    success_count += 1
                    self.log(success_msg, "INFO")
                    progress_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {success_msg}\n")
                else:
                    error_msg = f"Failed to send to {recipient}: {error_detail}"
                    self.log(error_msg, "ERROR")
                    progress_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] ERROR: {error_msg}\n", "error")
                    
                    # Add to failed contacts list
                    failed_contact = contact.copy()
                    failed_contact['Error'] = error_detail
                    failed_contact['Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.failed_contacts.append(failed_contact)
                    
            except Exception as e:
                error_msg = f"Unexpected error sending to {recipient}: {str(e)}"
                self.log(error_msg, "ERROR")
                self.log(traceback.format_exc(), "DEBUG")
                progress_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] CRITICAL: {error_msg}\n", "error")
                
                # Add to failed contacts list
                failed_contact = contact.copy()
                failed_contact['Error'] = str(e)
                failed_contact['Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                self.failed_contacts.append(failed_contact)
            
            progress_log.see(tk.END)
            progress.update()
            
            # Delay between emails
            time.sleep(3)
        
        # Save failure report if needed
        if self.failed_contacts:
            self.save_failure_report()
        
        # Show results
        progress.destroy()
        final_msg = f"Campaign finished - {success_count} succeeded, {total - success_count} failed"
        if self.failed_contacts:
            final_msg += f"\nFailure report saved to: {self.config['failure_report']}"
        self.log(final_msg, "INFO")
        messagebox.showinfo("Complete", final_msg)

    def save_failure_report(self):
        """Save failed contacts to Excel file with error details"""
        try:
            df_failed = pd.DataFrame(self.failed_contacts)
            
            # Ensure all original columns are included, even if empty
            original_df = pd.read_excel(self.config['excel_file'])
            for col in original_df.columns:
                if col not in df_failed.columns:
                    df_failed[col] = ""
            
            # Reorder columns to match original + error info
            cols = list(original_df.columns) + ['Error', 'Timestamp']
            df_failed = df_failed[cols]
            
            # Save to Excel
            df_failed.to_excel(self.config['failure_report'], index=False)
            self.log(f"Saved failure report with {len(df_failed)} contacts to {self.config['failure_report']}", "INFO")
        except Exception as e:
            error_msg = f"Failed to save failure report: {str(e)}"
            self.log(error_msg, "ERROR")
            self.log(traceback.format_exc(), "DEBUG")

    def _encode_config(self, data_dict: dict) -> str:
        """Obfuscate config data with base64"""
        if not data_dict:
            raise ValueError("Config data cannot be empty")
        json_str = json.dumps(data_dict)
        return base64.b64encode(json_str.encode()).decode()
    
    def _decode_config(self, encoded_str: str) -> dict:
        """Decode base64-obfuscated config"""
        if not encoded_str:
            raise ValueError("Encoded string cannot be empty")
        try:
            decoded = base64.b64decode(encoded_str.encode()).decode()
            return json.loads(decoded)
        except Exception as e:
            raise ValueError(f"Decoding failed: {str(e)}")
            
    def _get_relative_path(self, path):
        """Convert absolute path to relative if it's under current directory"""
        if not path:
            return ""
        try:
            rel_path = os.path.relpath(path)
            return rel_path if len(rel_path) < len(path) else path
        except ValueError:
            return path

    def _get_absolute_path(self, path):
        """Convert path to absolute, handling both relative and absolute paths"""
        if not path:
            return ""
        if os.path.isabs(path):
            return path
        return os.path.abspath(path)

    def save_config(self, email: str, password: str, sender_name: str, smtp_server: str, smtp_port: int):
        """Save encrypted config to file including last used paths"""
        config_file = ".creds"
        try:
            if not email or not password:
                self.log("Empty credentials provided, not saving", "WARNING")
                return
                
            config_data = {
                "email": self.email_entry.get(),
                "password": self.pass_entry.get(),
                "app_password": self.app_pass_entry.get() if self.gmail_mode.get() else "",
                "sender_name": self.sender_name_entry.get(),
                "smtp_server": self.smtp_entry.get(),
                "smtp_port": self.port_entry.get(),
                "previous_smtp_server": self.previous_smtp_server,  # Add this line
                "previous_smtp_port": self.previous_smtp_port, 
                "is_gmail": self.gmail_mode.get(),
                "last_excel": self._get_absolute_path(self.excel_entry.get()),
                "last_html": self._get_absolute_path(self.html_entry.get())
            }
            
            encoded = self._encode_config(config_data)
            with open(config_file, "w") as f:
                f.write(encoded)
            if os.name != 'nt':
                os.chmod(config_file, 0o600)
            self.log("Config saved securely", "INFO")
        except Exception as e:
            self.log(f"Error saving config: {str(e)}", "ERROR")
            messagebox.showerror("Save Error", f"Could not save config:\n{str(e)}")

    def load_config(self):
        """Load and decode config including last used files"""
        config_file = ".creds"
        try:
            if not os.path.exists(config_file):
                self.log("No config file found (first run?)", "INFO")
                return
                
            with open(config_file, "r") as f:
                encoded = f.read().strip()
                if not encoded:
                    self.log("Config file is empty", "WARNING")
                    return
                    
            config_data = self._decode_config(encoded)
            
            # Load previous SMTP settings
            self.previous_smtp_server = config_data.get("previous_smtp_server", "")
            self.previous_smtp_port = config_data.get("previous_smtp_port", "")
            
            # Only populate fields if they're empty
            if not self.email_entry.get():
                self.email_entry.delete(0, tk.END)
                self.email_entry.insert(0, config_data.get("email", ""))
            if not self.pass_entry.get():
                self.pass_entry.delete(0, tk.END)
                self.pass_entry.insert(0, config_data.get("password", ""))
            if config_data.get("is_gmail", False):
                self.gmail_mode.set(True)
                self.toggle_gmail_mode()
                self.app_pass_entry.insert(0, config_data.get("app_password", ""))
            if not self.smtp_entry.get():
                self.smtp_entry.delete(0, tk.END)
                self.smtp_entry.insert(0, config_data.get("smtp_server", ""))
            if not self.port_entry.get():
                self.port_entry.delete(0, tk.END)
                port = config_data.get("smtp_port", "465")
                self.port_entry.insert(0, str(port) if port else "465")
            if not self.sender_name_entry.get():
                self.sender_name_entry.delete(0, tk.END)
                self.sender_name_entry.insert(0, config_data.get("sender_name", ""))
                
            # Load last used files (show relative paths)
            if "last_excel" in config_data:
                abs_path = config_data["last_excel"]
                if os.path.exists(abs_path):
                    self.excel_entry.delete(0, tk.END)
                    self.excel_entry.insert(0, self._get_relative_path(abs_path))
                    
            if "last_html" in config_data:
                abs_path = config_data["last_html"]
                if os.path.exists(abs_path):
                    self.html_entry.delete(0, tk.END)
                    self.html_entry.insert(0, self._get_relative_path(abs_path))
                    
            self.log("Successfully loaded config", "INFO")
        except Exception as e:
            self.log(f"Error loading config: {str(e)}", "ERROR")

    def import_failed_contacts(self):
        """Import failed contacts from MBOX bounce messages"""
        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("Import Failed Contacts from MBOX")
            dialog.geometry("800x600")
            
            ttk.Label(dialog, text="Select MBOX file containing bounce messages:").pack(pady=10)
            
            def browse_mbox():
                filepath = filedialog.askopenfilename(
                    title="Select MBOX File",
                    filetypes=[("MBOX files", "*.mbox"), ("All files", "*.*")]
                )
                if filepath:
                    process_mbox_file(filepath)
            
            def process_mbox_file(filepath):
                try:
                    rejected_emails = set()
                    
                    # Open the mbox file
                    mbox = mailbox.mbox(filepath)
                    
                    self.log(f"Processing {len(mbox)} messages from MBOX", "INFO")
                    
                    for message in mbox:
                        text_content = self._extract_text_content(message)
                        found = self._parse_bounce_content(text_content)
                        rejected_emails.update(found)
                    
                    self._display_results(rejected_emails, dialog)
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to process MBOX file: {str(e)}")
                    self.log(f"MBOX processing error: {str(e)}", "ERROR")
            
            ttk.Button(dialog, text="Browse MBOX File", command=browse_mbox).pack(pady=10)
            
            # Result display area
            self.result_text = scrolledtext.ScrolledText(dialog, height=20)
            self.result_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            self.result_text.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize: {str(e)}")

    def _extract_text_content(self, msg):
        """Extract text content from email message"""
        text_content = ""
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            text_content += payload.decode(errors="ignore") + "\n"
            else:
                payload = msg.get_payload(decode=True)
                if payload:
                    text_content = payload.decode(errors="ignore")
        except Exception as e:
            self.log(f"Error extracting content: {str(e)}", "WARNING")
        return text_content

    def _parse_bounce_content(self, text_content):
        """Parse bounce message content for failed addresses"""
        found = set()
        
        if not text_content:
            return found
        
        # Patterns to match in bounce messages
        patterns = [
            r'failed:\s*\n\n\[([^\]]+)\]',  # [address] after "failed:"
            r'RCPT TO\s*[<:]+([^\s>:]+)',   # RCPT TO:<address>
            r'Original-Recipient:\s*rfc822;\s*([^\s]+)',  # Original-Recipient
            r'Final-Recipient:\s*rfc822;\s*([^\s]+)',     # Final-Recipient
            r'To:\s*([^\s<]+@[^\s>]+)',                  # To: address
            r'<([^>]+@[^>]+)>'                           # <address>
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text_content, re.IGNORECASE)
            for addr in matches:
                clean_addr = addr.strip('<>:').lower()
                if '@' in clean_addr and not any(x in clean_addr for x in ['mailer-daemon', 'postmaster']):
                    found.add(clean_addr)
        
        return found

    def _display_results(self, rejected_emails, parent):
        """Display results in the text area"""
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        
        if rejected_emails:
            self.result_text.insert(tk.END, f"Found {len(rejected_emails)} rejected addresses:\n\n")
            for email in sorted(rejected_emails):
                self.result_text.insert(tk.END, f"{email}\n")
            
            # Add button to save results
            save_btn = ttk.Button(parent, text="Save to Excel", 
                                command=lambda: self._save_results_to_excel(rejected_emails))
            save_btn.pack(pady=10)
        else:
            self.result_text.insert(tk.END, "No rejected addresses found in the messages")
        
        self.result_text.config(state=tk.DISABLED)

    def _save_results_to_excel(self, rejected_emails):
        """Save found addresses to Excel file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Rejected Addresses"
        )
        if filename:
            df = pd.DataFrame(sorted(rejected_emails), columns=["Rejected Emails"])
            df.to_excel(filename, index=False)
            messagebox.showinfo("Success", f"Saved {len(rejected_emails)} addresses to {filename}")

    def send_email(self, sender_email, sender_name, recipient_email, 
                recipient_name, subject, html_content, 
                smtp_server, smtp_port):
        """Universal email sender that handles both Gmail and regular SMTP"""
        
        password = self.app_pass_entry.get() if self.gmail_mode.get() else self.pass_entry.get()
        
        try:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = formataddr((sender_name, sender_email))
            msg['To'] = recipient_email
            msg.attach(MIMEText(html_content, 'html'))

            context = ssl.create_default_context()
                
            # Gmail mode handling
            if self.gmail_mode.get():
                if smtp_port == 465:
                    with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
                        server.login(sender_email, password)
                        server.send_message(msg)
                elif smtp_port == 587:
                    with smtplib.SMTP(smtp_server, smtp_port) as server:
                        server.starttls(context=context)
                        server.login(sender_email, password)
                        server.send_message(msg)
                else:
                    return False, "Gmail requires port 465 (SSL) or 587 (TLS)"            
            # Regular SMTP handling
            else:
                with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
                    server.login(sender_email, password)
                    server.send_message(msg)            
            return True, None
                
        except smtplib.SMTPAuthenticationError as e:
            if self.gmail_mode.get():
                return False, ("Gmail authentication failed. Possible causes:\n"
                            "1. Need App Password (if 2FA enabled)\n"
                            "2. 'Less secure apps' not enabled\n"
                            "3. Unlock captcha required\n"
                            f"Technical details: {str(e)}")
            else:
                return False, f"SMTP Authentication Failed: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
        
def load_contacts(file_path):

    """Load and validate contacts with logging"""
    try:
        df = pd.read_excel(file_path)
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        valid_contacts = df[df['Email Contacto'].str.contains('@', na=False)]
        invalid_count = len(df) - len(valid_contacts)
        
        if invalid_count > 0:
            logging.warning(f"Filtered out {invalid_count} invalid contacts")
        
        return valid_contacts.to_dict('records')
    except Exception as e:
        logging.error(f"Error loading contacts: {str(e)}")
        raise




if __name__ == "__main__":
    root = tk.Tk()
    app = EmailCampaignApp(root)
    root.mainloop()