import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import requests
import json
import smtplib
import threading
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from yaml_config_manager import YAMLConfigManager as ConfigManager

class EmailGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Fresh Start Cleaning - Email Generator")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Initialize config manager
        self.config = ConfigManager()
        
        # Data storage
        self.prospects_data = []
        self.generated_emails = []
        self.current_email_index = 0
        
        # Create main interface
        self.create_widgets()
        
        # Check configuration on startup
        self.check_initial_setup()

    def create_widgets(self):
        """Create the main GUI widgets"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Email Generation
        self.email_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.email_tab, text="Generate Emails")
        self.create_email_tab()
        
        # Tab 2: Configuration
        self.config_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.config_tab, text="Configuration")
        self.create_config_tab()
        
        # Tab 3: Results
        self.results_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.results_tab, text="Results & History")
        self.create_results_tab()

    def create_email_tab(self):
        """Create the email generation tab"""
        # File upload section
        upload_frame = ttk.LabelFrame(self.email_tab, text="Step 1: Upload Prospects CSV", padding=10)
        upload_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(upload_frame, text="Select CSV File", command=self.upload_csv).pack(side=tk.LEFT)
        self.file_label = ttk.Label(upload_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(upload_frame, text="Download CSV Template", command=self.download_template).pack(side=tk.RIGHT)
        
        # Prospects preview
        preview_frame = ttk.LabelFrame(self.email_tab, text="Loaded Prospects", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Treeview for prospects
        columns = ('Company', 'Industry', 'Contact', 'Email', 'Location')
        self.prospects_tree = ttk.Treeview(preview_frame, columns=columns, show='headings', height=6)
        
        for col in columns:
            self.prospects_tree.heading(col, text=col)
            self.prospects_tree.column(col, width=150)
        
        scrollbar_prospects = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.prospects_tree.yview)
        self.prospects_tree.configure(yscrollcommand=scrollbar_prospects.set)
        
        self.prospects_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_prospects.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Generate emails section
        generate_frame = ttk.LabelFrame(self.email_tab, text="Step 2: Generate Emails", padding=10)
        generate_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.generate_btn = ttk.Button(generate_frame, text="Generate All Emails", 
                                     command=self.generate_emails, state=tk.DISABLED)
        self.generate_btn.pack(side=tk.LEFT)
        
        self.progress = ttk.Progressbar(generate_frame, mode='determinate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        
        self.status_label = ttk.Label(generate_frame, text="Upload CSV to begin")
        self.status_label.pack(side=tk.RIGHT)
        
        # Email review section
        review_frame = ttk.LabelFrame(self.email_tab, text="Step 3: Review & Send Emails", padding=10)
        review_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Navigation
        nav_frame = ttk.Frame(review_frame)
        nav_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(nav_frame, text="◀ Previous", command=self.prev_email).pack(side=tk.LEFT)
        self.email_counter = ttk.Label(nav_frame, text="No emails generated")
        self.email_counter.pack(side=tk.LEFT, padx=10)
        ttk.Button(nav_frame, text="Next ▶", command=self.next_email).pack(side=tk.LEFT)
        
        ttk.Button(nav_frame, text="Send Current Email", command=self.send_current_email).pack(side=tk.RIGHT)
        ttk.Button(nav_frame, text="Send All Emails", command=self.send_all_emails).pack(side=tk.RIGHT, padx=(0, 10))
        
        # Email editor
        editor_frame = ttk.Frame(review_frame)
        editor_frame.pack(fill=tk.BOTH, expand=True)
        
        # Subject line
        ttk.Label(editor_frame, text="Subject:").pack(anchor=tk.W)
        self.subject_entry = ttk.Entry(editor_frame, font=('Arial', 10))
        self.subject_entry.pack(fill=tk.X, pady=(0, 10))
        
        # Email body
        ttk.Label(editor_frame, text="Email Body:").pack(anchor=tk.W)
        self.email_text = scrolledtext.ScrolledText(editor_frame, height=15, font=('Arial', 10))
        self.email_text.pack(fill=tk.BOTH, expand=True)

    def create_config_tab(self):
        """Create the configuration tab"""
        # Email configuration
        email_frame = ttk.LabelFrame(self.config_tab, text="Email Configuration", padding=10)
        email_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Email settings
        ttk.Label(email_frame, text="From Email:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.email_entry = ttk.Entry(email_frame, width=40)
        self.email_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        ttk.Label(email_frame, text="App Password:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.password_entry = ttk.Entry(email_frame, width=40, show="*")
        self.password_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        ttk.Button(email_frame, text="Test Email Connection", command=self.test_email_connection).grid(row=2, column=0, pady=10)
        ttk.Button(email_frame, text="Save Email Config", command=self.save_email_config).grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=10)
        
        # Company information
        company_frame = ttk.LabelFrame(self.config_tab, text="Company Information", padding=10)
        company_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Company details
        ttk.Label(company_frame, text="Company Name:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.company_name_entry = ttk.Entry(company_frame, width=40)
        self.company_name_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        ttk.Label(company_frame, text="Phone:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.phone_entry = ttk.Entry(company_frame, width=40)
        self.phone_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        ttk.Label(company_frame, text="Website:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.website_entry = ttk.Entry(company_frame, width=40)
        self.website_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # Services
        ttk.Label(company_frame, text="Services (one per line):").grid(row=3, column=0, sticky=tk.NW, pady=2)
        self.services_text = scrolledtext.ScrolledText(company_frame, height=8, width=50)
        self.services_text.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        ttk.Button(company_frame, text="Save Company Info", command=self.save_company_config).grid(row=4, column=1, sticky=tk.W, padx=(10, 0), pady=10)
        
        # Load existing configuration
        self.load_config_values()

    def create_results_tab(self):
        """Create the results and history tab"""
        # Sent emails history
        history_frame = ttk.LabelFrame(self.results_tab, text="Email History", padding=10)
        history_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # History treeview
        history_columns = ('Timestamp', 'Company', 'Email', 'Subject', 'Status')
        self.history_tree = ttk.Treeview(history_frame, columns=history_columns, show='headings')
        
        for col in history_columns:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, width=150)
        
        scrollbar_history = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar_history.set)
        
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_history.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Export options
        export_frame = ttk.Frame(self.results_tab)
        export_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(export_frame, text="Export History to CSV", command=self.export_history).pack(side=tk.LEFT)
        ttk.Button(export_frame, text="Clear History", command=self.clear_history).pack(side=tk.LEFT, padx=(10, 0))

    def check_initial_setup(self):
        """Check if initial setup is complete"""
        if not self.config.is_email_configured():
            messagebox.showinfo("Setup Required", 
                              "Please configure your email settings in the Configuration tab before generating emails.")
            self.notebook.select(1)  # Switch to config tab

    def upload_csv(self):
        """Upload and process CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select Prospects CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Read CSV
                df = pd.read_csv(file_path)
                
                # Validate required columns
                required_columns = ['Company Name', 'Email']
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    messagebox.showerror("Invalid CSV", 
                                       f"Missing required columns: {', '.join(missing_columns)}")
                    return
                
                # Store data
                self.prospects_data = df.to_dict('records')
                self.file_label.config(text=f"Loaded: {len(self.prospects_data)} prospects")
                
                # Update treeview
                self.update_prospects_tree()
                
                # Enable generate button
                self.generate_btn.config(state=tk.NORMAL)
                self.status_label.config(text="Ready to generate emails")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {str(e)}")

    def download_template(self):
        """Download CSV template"""
        template_data = {
            'Company Name': ['Turner Industries', 'Landis Construction', 'Meta Data Center'],
            'Industry': ['Industrial Construction', 'Commercial Construction', 'Technology'],
            'Contact Name': ['Facilities Manager', 'Project Manager', 'Construction Manager'],
            'Email': ['facilities@turner.com', 'projects@landis.com', 'construction@meta.com'],
            'Company Size': ['10,000+', '100-500', '5000+'],
            'Location': ['Baton Rouge, LA', 'New Orleans, LA', 'Richland Parish, LA'],
            'Notes': ['Large industrial construction company', 'Historic renovations', '$10B AI data center']
        }
        
        file_path = filedialog.asksaveasfilename(
            title="Save CSV Template",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        
        if file_path:
            df = pd.DataFrame(template_data)
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Template saved to {file_path}")

    def update_prospects_tree(self):
        """Update the prospects treeview"""
        # Clear existing items
        for item in self.prospects_tree.get_children():
            self.prospects_tree.delete(item)
        
        # Add new items
        for prospect in self.prospects_data:
            values = (
                prospect.get('Company Name', ''),
                prospect.get('Industry', ''),
                prospect.get('Contact Name', ''),
                prospect.get('Email', ''),
                prospect.get('Location', '')
            )
            self.prospects_tree.insert('', 'end', values=values)

    def generate_emails(self):
        """Generate emails for all prospects"""
        if not self.prospects_data:
            messagebox.showwarning("No Data", "Please upload a CSV file first.")
            return
        
        if not self.config.is_email_configured():
            messagebox.showwarning("Email Not Configured", "Please configure email settings first.")
            self.notebook.select(1)  # Switch to config tab
            return
        
        # Disable button and show progress
        self.generate_btn.config(state=tk.DISABLED)
        self.progress.config(maximum=len(self.prospects_data))
        self.status_label.config(text="Generating emails...")
        
        # Start generation in thread to prevent GUI freezing
        thread = threading.Thread(target=self._generate_emails_thread)
        thread.daemon = True
        thread.start()

    def _generate_emails_thread(self):
        """Generate emails in separate thread"""
        self.generated_emails = []
        ollama_config = self.config.get_ollama_config()
        
        for i, prospect in enumerate(self.prospects_data):
            try:
                # Update progress in main thread
                self.root.after(0, lambda i=i: self.progress.config(value=i))
                self.root.after(0, lambda: self.status_label.config(text=f"Generating email {i+1}/{len(self.prospects_data)}..."))
                
                # Generate email
                email_content = self._generate_single_email(prospect, ollama_config)
                subject, body = self._parse_email_content(email_content)
                
                email_data = {
                    'prospect': prospect,
                    'subject': subject,
                    'body': body,
                    'generated_at': datetime.now().isoformat(),
                    'sent': False
                }
                
                self.generated_emails.append(email_data)
                
            except Exception as e:
                print(f"Error generating email for {prospect.get('Company Name', 'Unknown')}: {e}")
        
        # Update UI in main thread
        self.root.after(0, self._generation_complete)

    def _generate_single_email(self, prospect, ollama_config):
        """Generate a single email using Ollama"""
        prompt = self._create_email_prompt(prospect)
        
        payload = {
            "model": ollama_config['model'],
            "prompt": prompt,
            "stream": False
        }
        
        response = requests.post(
            ollama_config['url'], 
            json=payload, 
            timeout=ollama_config['timeout']
        )
        
        if response.status_code == 200:
            return response.json()["response"]
        else:
            raise Exception(f"Ollama API error: {response.status_code}")

    def _create_email_prompt(self, prospect):
        """Create email generation prompt"""
        company_info = self.config.get_company_info()
        
        return f"""
Write a professional, personalized email for a cleaning company to send to a potential business client.

CLEANING COMPANY DETAILS:
- Company: {company_info['name']}
- Website: {company_info['website']}
- Location: {company_info['location']}
- Phone: {company_info['phone']}
- Services: {', '.join(company_info['services'])}
- Experience: {company_info['years_experience']} years
- Certifications: {', '.join(company_info['certifications'])}

PROSPECT DETAILS:
- Company Name: {prospect.get('Company Name', 'N/A')}
- Industry: {prospect.get('Industry', 'N/A')}
- Contact Name: {prospect.get('Contact Name', 'Facilities Manager')}
- Email: {prospect.get('Email', 'N/A')}
- Company Size: {prospect.get('Company Size', 'N/A')}
- Location: {prospect.get('Location', 'Louisiana')}
- Notes: {prospect.get('Notes', 'N/A')}

FORMAT REQUIREMENTS:
1. Start with: SUBJECT: [compelling subject line]
2. Then: EMAIL BODY: [the email content]
3. Professional but friendly tone
4. Personalized opening showing research
5. Clear value proposition
6. Specific services for their industry
7. Call to action for meeting/quote
8. Professional signature
9. Keep email body under 150 words

Generate the complete email with subject and body clearly separated.
"""

    def _parse_email_content(self, content):
        """Parse generated email content"""
        lines = content.strip().split('\n')
        subject = ""
        body = ""
        
        for i, line in enumerate(lines):
            if line.upper().startswith('SUBJECT:'):
                subject = line.replace('SUBJECT:', '').strip()
                body_lines = lines[i+1:]
                if body_lines and body_lines[0].upper().startswith('EMAIL BODY:'):
                    body_lines = body_lines[1:]
                body = '\n'.join(body_lines).strip()
                break
        
        if not subject and lines:
            subject = lines[0].strip()
            body = '\n'.join(lines[1:]).strip()
        
        if not subject:
            subject = "Professional Cleaning Services - Fresh Start Cleaning Co."
        
        return subject, body

    def _generation_complete(self):
        """Called when email generation is complete"""
        self.progress.config(value=len(self.prospects_data))
        self.status_label.config(text=f"Generated {len(self.generated_emails)} emails")
        self.generate_btn.config(state=tk.NORMAL)
        
        if self.generated_emails:
            self.current_email_index = 0
            self.display_current_email()
            messagebox.showinfo("Success", f"Generated {len(self.generated_emails)} emails!")

    def display_current_email(self):
        """Display the current email in the editor"""
        if not self.generated_emails:
            return
        
        email = self.generated_emails[self.current_email_index]
        
        # Update counter
        self.email_counter.config(text=f"Email {self.current_email_index + 1} of {len(self.generated_emails)}")
        
        # Update editor
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, email['subject'])
        
        self.email_text.delete(1.0, tk.END)
        self.email_text.insert(1.0, email['body'])

    def prev_email(self):
        """Navigate to previous email"""
        if self.generated_emails and self.current_email_index > 0:
            self.save_current_email_edits()
            self.current_email_index -= 1
            self.display_current_email()

    def next_email(self):
        """Navigate to next email"""
        if self.generated_emails and self.current_email_index < len(self.generated_emails) - 1:
            self.save_current_email_edits()
            self.current_email_index += 1
            self.display_current_email()

    def save_current_email_edits(self):
        """Save edits to current email"""
        if self.generated_emails:
            email = self.generated_emails[self.current_email_index]
            email['subject'] = self.subject_entry.get()
            email['body'] = self.email_text.get(1.0, tk.END).strip()

    def send_current_email(self):
        """Send the current email"""
        if not self.generated_emails:
            messagebox.showwarning("No Emails", "No emails to send.")
            return
        
        self.save_current_email_edits()
        email = self.generated_emails[self.current_email_index]
        
        if messagebox.askyesno("Confirm Send", 
                              f"Send email to {email['prospect']['Company Name']} ({email['prospect']['Email']})?"):
            success = self._send_email(email)
            if success:
                email['sent'] = True
                email['sent_at'] = datetime.now().isoformat()
                self.update_history()
                messagebox.showinfo("Success", "Email sent successfully!")
            else:
                messagebox.showerror("Error", "Failed to send email.")

    def send_all_emails(self):
        """Send all generated emails"""
        if not self.generated_emails:
            messagebox.showwarning("No Emails", "No emails to send.")
            return
        
        unsent_count = sum(1 for email in self.generated_emails if not email['sent'])
        
        if messagebox.askyesno("Confirm Send All", 
                              f"Send {unsent_count} unsent emails?"):
            
            sent_count = 0
            for email in self.generated_emails:
                if not email['sent']:
                    if self._send_email(email):
                        email['sent'] = True
                        email['sent_at'] = datetime.now().isoformat()
                        sent_count += 1
            
            self.update_history()
            messagebox.showinfo("Complete", f"Sent {sent_count} out of {unsent_count} emails.")

    def _send_email(self, email_data):
        """Send a single email"""
        try:
            email_config = self.config.get_email_config()
            company_info = self.config.get_company_info()
            
            msg = MIMEMultipart()
            msg['From'] = f"{company_info['name']} <{email_config['from_email']}>"
            msg['To'] = email_data['prospect']['Email']
            msg['Subject'] = email_data['subject']
            
            msg.attach(MIMEText(email_data['body'], 'plain'))
            
            server = smtplib.SMTP(email_config['smtp_server'], email_config['smtp_port'])
            server.starttls()
            server.login(email_config['from_email'], email_config['from_password'])
            
            text = msg.as_string()
            server.sendmail(email_config['from_email'], email_data['prospect']['Email'], text)
            server.quit()
            
            return True
            
        except Exception as e:
            print(f"Error sending email: {e}")
            return False

    def load_config_values(self):
        """Load configuration values into UI"""
        email_config = self.config.get_email_config()
        company_info = self.config.get_company_info()
        
        self.email_entry.insert(0, email_config.get('from_email', ''))
        self.password_entry.insert(0, email_config.get('from_password', ''))
        
        self.company_name_entry.insert(0, company_info.get('name', ''))
        self.phone_entry.insert(0, company_info.get('phone', ''))
        self.website_entry.insert(0, company_info.get('website', ''))
        
        services_text = '\n'.join(company_info.get('services', []))
        self.services_text.insert(1.0, services_text)

    def save_email_config(self):
        """Save email configuration"""
        self.config.set('email', 'from_email', self.email_entry.get())
        self.config.set('email', 'from_password', self.password_entry.get())
        
        if self.config.save_config():
            messagebox.showinfo("Success", "Email configuration saved!")
        else:
            messagebox.showerror("Error", "Failed to save email configuration.")

    def save_company_config(self):
        """Save company configuration"""
        services_list = [line.strip() for line in self.services_text.get(1.0, tk.END).split('\n') if line.strip()]
        
        self.config.set('company', 'name', self.company_name_entry.get())
        self.config.set('company', 'phone', self.phone_entry.get())
        self.config.set('company', 'website', self.website_entry.get())
        self.config.set('company', 'services', services_list)
        
        if self.config.save_config():
            messagebox.showinfo("Success", "Company configuration saved!")
        else:
            messagebox.showerror("Error", "Failed to save company configuration.")

    def test_email_connection(self):
        """Test email connection"""
        email = self.email_entry.get()
        password = self.password_entry.get()
        
        if not email or not password:
            messagebox.showwarning("Missing Info", "Please enter email and password.")
            return
        
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email, password)
            server.quit()
            messagebox.showinfo("Success", "Email connection successful!")
        except Exception as e:
            messagebox.showerror("Connection Failed", f"Failed to connect: {str(e)}")

    def update_history(self):
        """Update the email history display"""
        # Clear existing items
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # Add sent emails
        for email in self.generated_emails:
            if email.get('sent', False):
                values = (
                    email.get('sent_at', email.get('generated_at', '')),
                    email['prospect'].get('Company Name', ''),
                    email['prospect'].get('Email', ''),
                    email['subject'][:50] + '...' if len(email['subject']) > 50 else email['subject'],
                    'Sent'
                )
                self.history_tree.insert('', 'end', values=values)

    def export_history(self):
        """Export email history to CSV"""
        sent_emails = [email for email in self.generated_emails if email.get('sent', False)]
        
        if not sent_emails:
            messagebox.showinfo("No Data", "No sent emails to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Email History",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        
        if file_path:
            try:
                export_data = []
                for email in sent_emails:
                    export_data.append({
                        'Sent Date': email.get('sent_at', ''),
                        'Company Name': email['prospect'].get('Company Name', ''),
                        'Contact Email': email['prospect'].get('Email', ''),
                        'Industry': email['prospect'].get('Industry', ''),
                        'Subject': email['subject'],
                        'Body': email['body']
                    })
                
                df = pd.DataFrame(export_data)
                df.to_csv(file_path, index=False)
                messagebox.showinfo("Success", f"History exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")

    def clear_history(self):
        """Clear email history"""
        if messagebox.askyesno("Confirm Clear", "Clear all email history?"):
            self.generated_emails = []
            self.update_history()
            self.email_counter.config(text="No emails generated")
            self.subject_entry.delete(0, tk.END)
            self.email_text.delete(1.0, tk.END)

def main():
    """Main application entry point"""
    root = tk.Tk()
    app = EmailGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()