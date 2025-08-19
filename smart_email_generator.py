import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import requests
import smtplib
import threading
import os
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from yaml_config_manager import YAMLConfigManager

class BoilerplateEmailGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Fresh Start Cleaning - Smart Email Generator v3.0")
        self.root.geometry("1000x750")
        self.root.minsize(900, 650)
        
        # Initialize config manager
        try:
            self.config = YAMLConfigManager()
        except Exception as e:
            messagebox.showerror("Configuration Error", 
                               f"Error loading config: {e}\n\nPlease check your config.yaml file.")
            self.root.quit()
            return
        
        # Data storage
        self.prospects_data = []
        self.generated_emails = []
        self.current_email_index = 0
        
        # Email boilerplate templates
        self.email_boilerplate = self._load_email_boilerplate()
        
        # Create main interface
        self.create_widgets()
        
        # Check configuration on startup
        self.check_initial_setup()

    def _load_email_boilerplate(self):
        """Load email boilerplate that AI will customize"""
        return {
            'base_template': '''Dear {contact_name},

{opening_line}

Fresh Start Cleaning Co. specializes in {industry_services} for {industry_type} like {company_name}. {industry_specific_benefits}

Our comprehensive services include:
{service_list}

{value_proposition}

{call_to_action}

Best regards,
Fresh Start Cleaning Co.
{phone}
{website}''',
            
            'industry_mappings': {
                'education': {
                    'industry_type': 'educational institutions',
                    'industry_services': 'campus cleaning and sanitization',
                    'service_list': '''• Campus-wide cleaning and sanitization
• Classroom and laboratory cleaning
• Student health and safety protocols
• Flexible scheduling around academic calendars''',
                    'value_proposition': 'A clean learning environment directly impacts student health and academic performance.'
                },
                'construction': {
                    'industry_type': 'construction companies',
                    'industry_services': 'post-construction cleanup and site maintenance',
                    'service_list': '''• Post-construction debris removal
• Deep cleaning of newly completed facilities
• Construction site maintenance cleaning
• Safety compliance cleaning protocols''',
                    'value_proposition': 'We understand the demanding timelines and quality standards of construction projects.'
                },
                'technology': {
                    'industry_type': 'technology companies',
                    'industry_services': 'professional office cleaning',
                    'service_list': '''• Regular office cleaning and sanitization
• Server room and equipment area cleaning
• Professional workspace maintenance
• Flexible scheduling around business operations''',
                    'value_proposition': 'A clean, organized office environment directly impacts productivity and creates a positive impression.'
                },
                'manufacturing': {
                    'industry_type': 'manufacturing facilities',
                    'industry_services': 'industrial cleaning and maintenance',
                    'service_list': '''• Manufacturing floor deep cleaning
• Equipment and machinery cleaning
• Safety compliance cleaning protocols
• Hazardous material cleanup''',
                    'value_proposition': 'Safety and compliance are paramount in manufacturing environments.'
                },
                'office': {
                    'industry_type': 'professional offices',
                    'industry_services': 'comprehensive janitorial services',
                    'service_list': '''• Daily janitorial and maintenance cleaning
• Restroom cleaning and sanitization
• Break room and kitchen area cleaning
• Trash removal and recycling''',
                    'value_proposition': 'We understand the importance of maintaining a clean, professional environment.'
                },
                'default': {
                    'industry_type': 'businesses',
                    'industry_services': 'professional cleaning services',
                    'service_list': '''• Commercial cleaning and janitorial services
• Specialized industry cleaning solutions
• Flexible scheduling and competitive pricing
• Emergency cleanup services''',
                    'value_proposition': 'Our local team provides reliable, professional service tailored to your specific requirements.'
                }
            }
        }

    def create_widgets(self):
        # Create main menu
        self.create_menu()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Email Generation
        self.email_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.email_tab, text="📧 Generate Emails")
        self.create_email_tab()
        
        # Tab 2: Configuration
        self.config_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.config_tab, text="⚙️ Configuration")
        self.create_config_tab()
        
        # Tab 3: Results
        self.results_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.results_tab, text="📊 Results")
        self.create_results_tab()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Load CSV", command=self.upload_csv)
        file_menu.add_command(label="Create Test CSV", command=self.create_test_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

    def create_email_tab(self):
        # Info section
        info_frame = ttk.LabelFrame(self.email_tab, text="How It Works", padding=10)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        info_text = "🚀 Smart Generation: Uses AI to customize email boilerplates - much faster than full generation!"
        ttk.Label(info_frame, text=info_text, foreground="blue").pack()
        
        # File upload section
        upload_frame = ttk.LabelFrame(self.email_tab, text="Step 1: Upload Prospects CSV", padding=10)
        upload_frame.pack(fill=tk.X, padx=10, pady=5)
        
        button_frame = ttk.Frame(upload_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="📁 Select CSV File", command=self.upload_csv).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="⚡ Create Test CSV", command=self.create_test_csv).pack(side=tk.LEFT, padx=(10, 0))
        self.file_label = ttk.Label(button_frame, text="No file selected", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Prospects preview
        preview_frame = ttk.LabelFrame(self.email_tab, text="Loaded Prospects", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Treeview for prospects
        columns = ('Company', 'Industry', 'Contact', 'Email', 'Location')
        self.prospects_tree = ttk.Treeview(preview_frame, columns=columns, show='headings', height=6)
        
        for col in columns:
            self.prospects_tree.heading(col, text=col)
            self.prospects_tree.column(col, width=150)
        
        # Scrollbar for prospects
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar_prospects = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.prospects_tree.yview)
        self.prospects_tree.configure(yscrollcommand=scrollbar_prospects.set)
        
        self.prospects_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_prospects.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Generate emails section
        generate_frame = ttk.LabelFrame(self.email_tab, text="Step 2: Generate Smart Emails", padding=10)
        generate_frame.pack(fill=tk.X, padx=10, pady=5)
        
        gen_button_frame = ttk.Frame(generate_frame)
        gen_button_frame.pack(fill=tk.X)
        
        self.generate_btn = ttk.Button(gen_button_frame, text="🤖 Generate Smart Emails", 
                                     command=self.generate_emails, state=tk.DISABLED)
        self.generate_btn.pack(side=tk.LEFT)
        
        # Progress bar with detailed label
        progress_frame = ttk.Frame(gen_button_frame)
        progress_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 10))
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.pack(fill=tk.X)
        
        self.progress_label = ttk.Label(progress_frame, text="", font=('Arial', 8))
        self.progress_label.pack()
        
        self.status_label = ttk.Label(gen_button_frame, text="Upload CSV to begin")
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
        
        ttk.Button(nav_frame, text="📧 Send Current", command=self.send_current_email).pack(side=tk.RIGHT)
        ttk.Button(nav_frame, text="📬 Send All", command=self.send_all_emails).pack(side=tk.RIGHT, padx=(0, 10))
        
        # Email editor
        editor_frame = ttk.Frame(review_frame)
        editor_frame.pack(fill=tk.BOTH, expand=True)
        
        # Subject line
        ttk.Label(editor_frame, text="Subject:").pack(anchor=tk.W)
        self.subject_entry = ttk.Entry(editor_frame, font=('Arial', 10))
        self.subject_entry.pack(fill=tk.X, pady=(0, 10))
        
        # Email body
        ttk.Label(editor_frame, text="Email Body:").pack(anchor=tk.W)
        self.email_text = scrolledtext.ScrolledText(editor_frame, height=12, font=('Arial', 10))
        self.email_text.pack(fill=tk.BOTH, expand=True)

    def create_config_tab(self):
        # Simple config display
        config_frame = ttk.LabelFrame(self.config_tab, text="Configuration Status", padding=20)
        config_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Show config status
        status_text = scrolledtext.ScrolledText(config_frame, height=20, width=80)
        status_text.pack(fill=tk.BOTH, expand=True)
        
        # Display config summary
        try:
            config_info = f"""📋 Smart Email Generator Configuration:
{'='*60}

📧 Email Configuration:
   Email: {self.config.get('email', 'from_email')}
   Status: {'✅ Ready to Send' if self.config.is_email_configured() else '❌ Edit config.yaml'}

🏢 Company Information:
   Name: {self.config.get('company', 'name')}
   Website: {self.config.get('company', 'website')}
   Phone: {self.config.get('company', 'phone')}

🤖 AI Configuration:
   Model: {self.config.get('ollama', 'model')}
   URL: {self.config.get('ollama', 'url')}
   Mode: Smart Boilerplate Customization

🚀 How Smart Generation Works:
   1. Uses pre-written email boilerplates
   2. AI customizes only specific parts:
      • Opening line
      • Industry-specific benefits  
      • Call to action
   3. Much faster than full AI generation
   4. More reliable and consistent results

💡 Benefits:
   • 10x faster than full AI generation
   • Consistent professional quality
   • Fallback protection if AI fails
   • Industry-specific customization

📝 To modify settings, edit config.yaml file directly.
{'='*60}
Status: {self.config.get_config_status()}
"""
            
            status_text.insert(1.0, config_info)
            status_text.config(state=tk.DISABLED)
            
        except Exception as e:
            status_text.insert(1.0, f"Error loading configuration: {e}")

    def create_results_tab(self):
        # Results display
        results_frame = ttk.LabelFrame(self.results_tab, text="Email Generation Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Results treeview
        columns = ('Company', 'Email', 'Method', 'Status', 'Generated At')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show='headings')
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=120)
        
        scrollbar_results = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar_results.set)
        
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_results.pack(side=tk.RIGHT, fill=tk.Y)

    def check_initial_setup(self):
        """Check if initial setup is complete"""
        if not self.config.is_email_configured():
            messagebox.showinfo("Setup Required", 
                              "Please edit config.yaml with your email credentials before sending emails.")

    def create_test_csv(self):
        """Create a quick test CSV file"""
        test_data = {
            'Company Name': ['Macalester College', 'Louisiana Construction Co', 'Tech Solutions LLC', 'Manufacturing Corp', 'Office Complex'],
            'Industry': ['Education', 'Construction', 'Technology', 'Manufacturing', 'Professional Services'],
            'Contact Name': ['William Acosta', 'Project Manager', 'IT Director', 'Operations Manager', 'Building Manager'],
            'Email': ['wacostal@macalester.edu'] * 5,
            'Company Size': ['2000+', '50-100', '25-50', '200-500', '100+ tenants'],
            'Location': ['Minnesota', 'Louisiana', 'Louisiana', 'Louisiana', 'Louisiana'],
            'Notes': ['Liberal arts college', 'Commercial construction', 'Tech startup', 'Manufacturing plant', 'Multi-tenant building']
        }
        
        df = pd.DataFrame(test_data)
        file_path = 'test_prospects.csv'
        df.to_csv(file_path, index=False)
        
        messagebox.showinfo("Success", f"✅ Created {file_path}")
        
        # Auto-load the test file
        self.prospects_data = df.to_dict('records')
        self.file_label.config(text=f"✅ Loaded: {len(self.prospects_data)} test prospects", foreground="green")
        self.update_prospects_tree()
        self.generate_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Ready to generate smart emails")

    def new_project(self):
        """Start a new project"""
        if messagebox.askyesno("New Project", "Clear all current data?"):
            self.prospects_data = []
            self.generated_emails = []
            self.current_email_index = 0
            self.update_prospects_tree()
            self.file_label.config(text="No file selected", foreground="gray")
            self.generate_btn.config(state=tk.DISABLED)
            self.status_label.config(text="Upload CSV to begin")

    def upload_csv(self):
        """Upload and process CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select Prospects CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                df = pd.read_csv(file_path)
                
                # Validate required columns
                required_columns = ['Company Name', 'Email']
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    messagebox.showerror("Invalid CSV", 
                                       f"Missing required columns: {', '.join(missing_columns)}")
                    return
                
                # Filter out rows with empty required fields
                df = df.dropna(subset=required_columns)
                
                self.prospects_data = df.to_dict('records')
                self.file_label.config(text=f"✅ Loaded: {len(self.prospects_data)} prospects", foreground="green")
                
                self.update_prospects_tree()
                self.generate_btn.config(state=tk.NORMAL)
                self.status_label.config(text="Ready to generate smart emails")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {str(e)}")

    def update_prospects_tree(self):
        """Update the prospects treeview"""
        for item in self.prospects_tree.get_children():
            self.prospects_tree.delete(item)
        
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
        """Generate emails using smart boilerplate customization"""
        if not self.prospects_data:
            messagebox.showwarning("No Data", "Please upload a CSV file first.")
            return
        
        self.generate_btn.config(state=tk.DISABLED)
        self.progress.config(maximum=len(self.prospects_data), value=0)
        self.status_label.config(text="🤖 Smart generation in progress...")
        
        thread = threading.Thread(target=self._generate_emails_smart)
        thread.daemon = True
        thread.start()

    def _generate_emails_smart(self):
        """Generate emails using AI-customized boilerplates"""
        self.generated_emails = []
        
        for i, prospect in enumerate(self.prospects_data):
            try:
                # Update progress
                self.root.after(0, lambda i=i: self.progress.config(value=i))
                self.root.after(0, lambda i=i, p=prospect: self.progress_label.config(
                    text=f"Customizing email {i+1}/{len(self.prospects_data)}: {p.get('Company Name', 'Unknown')[:20]}..."))
                
                # Generate smart email
                subject, body, method = self._generate_smart_email(prospect)
                
                email_data = {
                    'prospect': prospect,
                    'subject': subject,
                    'body': body,
                    'generated_at': datetime.now().isoformat(),
                    'sent': False,
                    'method': method
                }
                
                self.generated_emails.append(email_data)
                
                # Small delay for visual progress
                time.sleep(0.2)
                
            except Exception as e:
                print(f"Error generating smart email for {prospect.get('Company Name', 'Unknown')}: {e}")
                # Fallback to basic template
                subject, body = self._generate_fallback_email(prospect)
                email_data = {
                    'prospect': prospect,
                    'subject': subject,
                    'body': body,
                    'generated_at': datetime.now().isoformat(),
                    'sent': False,
                    'method': 'fallback'
                }
                self.generated_emails.append(email_data)
        
        self.root.after(0, self._generation_complete)

    def _generate_smart_email(self, prospect):
        """Generate email using AI to customize boilerplate"""
        company_name = prospect.get('Company Name', 'Your Company')
        contact_name = prospect.get('Contact Name', 'Facilities Manager')
        industry = prospect.get('Industry', '').lower()
        notes = prospect.get('Notes', '')
        company_info = self.config.get_company_info()
        
        # Get appropriate boilerplate
        industry_key = self._map_industry(industry)
        boilerplate_data = self.email_boilerplate['industry_mappings'][industry_key]
        
        # Use AI to customize specific parts
        try:
            customizations = self._get_ai_customizations(prospect)
            method = 'smart_ai'
        except Exception as e:
            print(f"AI customization failed, using smart defaults: {e}")
            customizations = self._get_smart_defaults(prospect)
            method = 'smart_default'
        
        # Build email from boilerplate + AI customizations
        subject = f"Professional {boilerplate_data['industry_services'].title()} for {company_name}"
        
        body = self.email_boilerplate['base_template'].format(
            contact_name=contact_name,
            opening_line=customizations['opening_line'],
            industry_services=boilerplate_data['industry_services'],
            industry_type=boilerplate_data['industry_type'],
            company_name=company_name,
            industry_specific_benefits=customizations['industry_benefits'],
            service_list=boilerplate_data['service_list'],
            value_proposition=boilerplate_data['value_proposition'],
            call_to_action=customizations['call_to_action'],
            phone=company_info.get('phone', ''),
            website=company_info.get('website', '')
        )
        
        return subject, body, method

    def _get_ai_customizations(self, prospect):
        """Use AI to customize specific parts of the email"""
        company_name = prospect.get('Company Name', '')
        industry = prospect.get('Industry', '')
        notes = prospect.get('Notes', '')
        
        # Short AI prompt for customization only
        prompt = f"""Customize these 3 parts for an email to {company_name} ({industry}):

1. Opening line (1 sentence, friendly but professional)
2. Industry benefit (1 sentence about why cleaning matters for their industry)
3. Call to action (1 sentence asking for a meeting)

Company: {company_name}
Industry: {industry}
Notes: {notes}

Format:
OPENING: [opening line]
BENEFIT: [industry benefit]
ACTION: [call to action]

Keep each under 20 words."""

        ollama_config = self.config.get_ollama_config()
        
        payload = {
            "model": ollama_config['model'],
            "prompt": prompt,
            "stream": False
        }
        
        response = requests.post(
            ollama_config['url'], 
            json=payload, 
            timeout=20  # Short timeout for customization
        )
        
        if response.status_code == 200:
            ai_response = response.json()["response"]
            return self._parse_ai_customizations(ai_response, prospect)
        else:
            raise Exception(f"AI API error: {response.status_code}")

    def _parse_ai_customizations(self, ai_response, prospect):
        """Parse AI response into customization parts"""
        lines = ai_response.strip().split('\n')
        
        customizations = {
            'opening_line': f"I hope this message finds you well.",
            'industry_benefits': "This will help maintain a professional environment.",
            'call_to_action': f"Could we schedule a brief 15-minute call to discuss your cleaning needs?"
        }
        
        for line in lines:
            if line.upper().startswith('OPENING:'):
                customizations['opening_line'] = line.replace('OPENING:', '').strip()
            elif line.upper().startswith('BENEFIT:'):
                customizations['industry_benefits'] = line.replace('BENEFIT:', '').strip()
            elif line.upper().startswith('ACTION:'):
                customizations['call_to_action'] = line.replace('ACTION:', '').strip()
        
        return customizations

    def _get_smart_defaults(self, prospect):
        """Generate smart defaults without AI"""
        company_name = prospect.get('Company Name', 'Your Company')
        industry = prospect.get('Industry', '').lower()
        
        # Smart defaults based on industry
        if 'education' in industry:
            opening = f"I hope the academic year is going well at {company_name}."
            benefit = "A clean campus environment is essential for student health and learning."
        elif 'construction' in industry:
            opening = f"I've been following {company_name}'s impressive projects across Louisiana."
            benefit = "Post-construction cleanup is crucial for project completion and safety."
        elif 'tech' in industry or 'software' in industry:
            opening = f"I hope your team at {company_name} is having a productive week."
            benefit = "A clean workspace directly impacts productivity and team morale."
        elif 'manufacturing' in industry:
            opening = f"I understand {company_name} maintains high operational standards."
            benefit = "Industrial cleaning is essential for safety compliance and efficiency."
        else:
            opening = f"I hope business is going well at {company_name}."
            benefit = "Professional cleaning services help maintain your business image."
        
        return {
            'opening_line': opening,
            'industry_benefits': benefit,
            'call_to_action': f"Would you be available for a brief 15-minute conversation to discuss {company_name}'s cleaning needs?"
        }

    def _map_industry(self, industry):
        """Map industry to boilerplate category"""
        industry = industry.lower()
        
        if any(word in industry for word in ['education', 'school', 'college', 'university']):
            return 'education'
        elif any(word in industry for word in ['construction', 'building', 'contractor']):
            return 'construction'
        elif any(word in industry for word in ['technology', 'tech', 'software', 'it']):
            return 'technology'
        elif any(word in industry for word in ['manufacturing', 'industrial', 'factory', 'plant']):
            return 'manufacturing'
        elif any(word in industry for word in ['office', 'professional', 'services']):
            return 'office'
        else:
            return 'default'

    def _generate_fallback_email(self, prospect):
        """Generate basic email when AI fails"""
        company_name = prospect.get('Company Name', 'Your Company')
        contact_name = prospect.get('Contact Name', 'Facilities Manager')
        company_info = self.config.get_company_info()
        
        subject = f"Professional Cleaning Services for {company_name}"
        
        body = f"""Dear {contact_name},

I hope this message finds you well. Fresh Start Cleaning Co. would like to provide professional cleaning services for {company_name}.

Our services include:
• Commercial cleaning and janitorial services
• Professional maintenance and sanitization
• Flexible scheduling and competitive pricing
• Licensed, bonded, and insured service

With over 10+ years of experience serving Louisiana businesses, we provide reliable, professional cleaning solutions.

Could we schedule a brief call to discuss your cleaning needs?

Best regards,
Fresh Start Cleaning Co.
{company_info.get('phone', '')}
{company_info.get('website', '')}"""
        
        return subject, body

    def _generation_complete(self):
        """Called when email generation is complete"""
        self.progress.config(value=len(self.prospects_data))
        self.progress_label.config(text="✅ Complete!")
        self.status_label.config(text=f"✅ Generated {len(self.generated_emails)} smart emails")
        self.generate_btn.config(state=tk.NORMAL)
        
        if self.generated_emails:
            self.current_email_index = 0
            self.display_current_email()
            self.update_results_tree()
            
            # Show generation summary
            ai_count = sum(1 for email in self.generated_emails if email['method'] == 'smart_ai')
            default_count = sum(1 for email in self.generated_emails if email['method'] == 'smart_default')
            fallback_count = sum(1 for email in self.generated_emails if email['method'] == 'fallback')
            
            summary = f"✅ Generated {len(self.generated_emails)} emails!\n\n"
            summary += f"🤖 AI Customized: {ai_count}\n"
            summary += f"🧠 Smart Defaults: {default_count}\n"
            summary += f"📝 Fallback: {fallback_count}"
            
            messagebox.showinfo("Generation Complete", summary)

    def update_results_tree(self):
        """Update the results tree with generation info"""
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        for email in self.generated_emails:
            method_display = {
                'smart_ai': '🤖 AI Smart',
                'smart_default': '🧠 Smart Default', 
                'fallback': '📝 Fallback'
            }.get(email['method'], email['method'])
            
            values = (
                email['prospect'].get('Company Name', '')[:20],
                email['prospect'].get('Email', '')[:25],
                method_display,
                'Generated' if not email['sent'] else 'Sent',
                email['generated_at'][:16]  # Just date and time
            )
            self.results_tree.insert('', 'end', values=values)

    def display_current_email(self):
        """Display the current email in the editor"""
        if not self.generated_emails:
            return
        
        email = self.generated_emails[self.current_email_index]
        company_name = email['prospect'].get('Company Name', 'Unknown')
        method = email['method']
        
        method_icons = {
            'smart_ai': '🤖',
            'smart_default': '🧠', 
            'fallback': '📝'
        }
        
        icon = method_icons.get(method, '📧')
        self.email_counter.config(text=f"{icon} Email {self.current_email_index + 1} of {len(self.generated_emails)} - {company_name}")
        
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
        
        if not self.config.is_email_configured():
            messagebox.showwarning("Email Not Configured", 
                                 "Please edit config.yaml with your email credentials first.")
            return
        
        self.save_current_email_edits()
        email = self.generated_emails[self.current_email_index]
        
        company_name = email['prospect'].get('Company Name', 'Unknown')
        recipient_email = email['prospect'].get('Email', '')
        
        if messagebox.askyesno("Confirm Send", 
                              f"Send email to:\n\nCompany: {company_name}\nEmail: {recipient_email}\nSubject: {email['subject'][:50]}..."):
            
            # Show sending progress
            progress_window = self._create_progress_window("Sending Email", f"Sending to {company_name}...")
            
            def send_thread():
                success = self._send_email(email)
                progress_window.after(0, lambda: progress_window.destroy())
                
                if success:
                    email['sent'] = True
                    email['sent_at'] = datetime.now().isoformat()
                    self.root.after(0, lambda: self.update_results_tree())
                    self.root.after(0, lambda: messagebox.showinfo("Success", f"✅ Email sent successfully to {company_name}!"))
                else:
                    self.root.after(0, lambda: messagebox.showerror("Error", "❌ Failed to send email. Check your email configuration."))
            
            thread = threading.Thread(target=send_thread)
            thread.daemon = True
            thread.start()

    def send_all_emails(self):
        """Send all generated emails"""
        if not self.generated_emails:
            messagebox.showwarning("No Emails", "No emails to send.")
            return
        
        if not self.config.is_email_configured():
            messagebox.showwarning("Email Not Configured", 
                                 "Please edit config.yaml with your email credentials first.")
            return
        
        # Count unsent emails
        unsent_emails = [email for email in self.generated_emails if not email['sent']]
        
        if not unsent_emails:
            messagebox.showinfo("No Emails", "All emails have already been sent.")
            return
        
        if messagebox.askyesno("Confirm Send All", 
                              f"Send {len(unsent_emails)} unsent emails?\n\nThis will send emails to all prospects at once."):
            
            # Create progress dialog
            progress_window = self._create_progress_window("Sending Emails", "Sending batch emails...")
            send_progress = ttk.Progressbar(progress_window, maximum=len(unsent_emails), mode='determinate')
            send_progress.pack(fill=tk.X, padx=20, pady=10)
            
            status_label = ttk.Label(progress_window, text="")
            status_label.pack(pady=5)
            
            # Send emails in thread
            def send_thread():
                sent_count = 0
                for i, email in enumerate(unsent_emails):
                    progress_window.after(0, lambda i=i: send_progress.config(value=i))
                    progress_window.after(0, lambda email=email: status_label.config(
                        text=f"Sending to {email['prospect'].get('Company Name', 'Unknown')}..."))
                    
                    if self._send_email(email):
                        email['sent'] = True
                        email['sent_at'] = datetime.now().isoformat()
                        sent_count += 1
                    
                    # Small delay between emails
                    time.sleep(1)
                
                # Update UI and close progress
                progress_window.after(0, lambda: progress_window.destroy())
                self.root.after(0, lambda: self.update_results_tree())
                self.root.after(0, lambda: messagebox.showinfo("Batch Send Complete", 
                    f"✅ Sent {sent_count} out of {len(unsent_emails)} emails successfully."))
            
            thread = threading.Thread(target=send_thread)
            thread.daemon = True
            thread.start()

    def _create_progress_window(self, title, message):
        """Create a progress window"""
        progress_window = tk.Toplevel(self.root)
        progress_window.title(title)
        progress_window.geometry("400x120")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Center the window
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
        y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
        progress_window.geometry(f"+{x}+{y}")
        
        ttk.Label(progress_window, text=message).pack(pady=20)
        
        return progress_window

    def _send_email(self, email_data):
        """Send a single email"""
        try:
            email_config = self.config.get_email_config()
            company_info = self.config.get_company_info()
            
            msg = MIMEMultipart()
            msg['From'] = f"{company_info['name']} <{email_config['from_email']}>"
            msg['To'] = email_data['prospect']['Email']
            msg['Subject'] = email_data['subject']
            
            # Add signature to body if not already present
            body = email_data['body']
            if company_info['name'] not in body:
                signature = f"\n\nBest regards,\n{company_info['name']}\n{company_info['phone']}\n{company_info['website']}"
                body += signature
            
            msg.attach(MIMEText(body, 'plain'))
            
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

def main():
    """Main application entry point"""
    root = tk.Tk()
    
    # Set better fonts and styling
    try:
        root.tk.call('tk', 'scaling', 1.0)
    except:
        pass
    
    app = BoilerplateEmailGenerator(root)
    
    # Center window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()