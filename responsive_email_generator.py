import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import requests
import smtplib
import threading
import queue
import time
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from yaml_config_manager import YAMLConfigManager

class HybridEmailGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Fresh Start Cleaning - Hybrid AI Email Generator")
        self.root.geometry("1100x800")
        self.root.minsize(1000, 700)

        # Configuration
        self.config = self._load_config()
        self.session = requests.Session()
        self.session.headers.update({"Content-Type": "application/json"})
        
        # AI settings optimized for slow Ollama
        self.ai_timeout_fast = 15  # Quick timeout for first attempt
        self.ai_timeout_slow = 60  # Longer timeout for retry
        self.max_workers = 1  # Single worker to avoid overwhelming slow Ollama
        
        # Application state
        self.prospects = []
        self.emails = []
        self.current_idx = 0
        self.is_generating = False
        self.cancel_event = threading.Event()
        self.result_queue = queue.Queue()
        
        # Industry-specific boilerplates for AI customization
        self.industry_data = self._load_industry_data()
        
        self._build_ui()
        self._start_ui_updater()

    def _load_config(self):
        try:
            return YAMLConfigManager()
        except Exception as e:
            messagebox.showerror("Config Error", f"Error loading config: {e}")
            raise

    def _load_industry_data(self):
        """Industry-specific data for AI customization and fallbacks"""
        return {
            'education': {
                'services': ['Campus-wide cleaning', 'Classroom sanitization', 'Laboratory cleaning', 'Student facility maintenance'],
                'benefits': 'clean learning environments impact student health and academic performance',
                'pain_points': 'maintaining health standards across large campus facilities',
                'fallback_opening': "Hope the academic year is going well at {company_name}.",
                'fallback_benefit': "Campus cleanliness is essential for student health and learning success.",
                'fallback_action': "Could we schedule a call to discuss your campus cleaning needs?"
            },
            'construction': {
                'services': ['Post-construction cleanup', 'Site maintenance', 'Debris removal', 'Safety compliance cleaning'],
                'benefits': 'proper cleanup is crucial for project completion and safety standards',
                'pain_points': 'meeting tight deadlines while maintaining quality cleanup standards',
                'fallback_opening': "I've been following {company_name}'s impressive construction projects.",
                'fallback_benefit': "Post-construction cleanup is crucial for project completion and safety.",
                'fallback_action': "Would you be available to discuss your cleanup requirements?"
            },
            'technology': {
                'services': ['Office cleaning', 'Server room maintenance', 'Equipment area cleaning', 'Workspace sanitization'],
                'benefits': 'clean workspaces directly impact productivity and professional image',
                'pain_points': 'maintaining professional environments that support productivity',
                'fallback_opening': "Hope your team at {company_name} is having a productive week.",
                'fallback_benefit': "Clean workspaces directly impact productivity and team morale.",
                'fallback_action': "Could we schedule a brief call about your office cleaning needs?"
            },
            'manufacturing': {
                'services': ['Industrial floor cleaning', 'Equipment maintenance', 'Safety compliance', 'Hazardous material cleanup'],
                'benefits': 'industrial cleaning is essential for safety compliance and operational efficiency',
                'pain_points': 'maintaining safety standards while keeping operations running',
                'fallback_opening': "I understand {company_name} maintains high operational standards.",
                'fallback_benefit': "Industrial cleaning is essential for safety and operational efficiency.",
                'fallback_action': "Would you be interested in discussing your facility cleaning needs?"
            },
            'residential': {
                'services': ['House cleaning', 'Deep cleaning', 'Move-in/out cleaning', 'Regular maintenance cleaning'],
                'benefits': 'professional cleaning saves time and ensures a healthy living environment',
                'pain_points': 'maintaining a clean home while managing busy schedules',
                'fallback_opening': "Hope you and your family are doing well.",
                'fallback_benefit': "Professional cleaning saves time and ensures a healthy home environment.",
                'fallback_action': "Would you be interested in learning about our residential cleaning services?"
            },
            'office': {
                'services': ['Daily janitorial', 'Restroom maintenance', 'Break room cleaning', 'Trash removal'],
                'benefits': 'professional environments enhance employee satisfaction and client impressions',
                'pain_points': 'maintaining professional appearance for employees and clients',
                'fallback_opening': "Hope business is going well at {company_name}.",
                'fallback_benefit': "Professional cleaning helps maintain your business image and employee satisfaction.",
                'fallback_action': "Could we schedule a call to discuss your office cleaning needs?"
            },
            'default': {
                'services': ['Commercial cleaning', 'Professional maintenance', 'Customized solutions', 'Reliable service'],
                'benefits': 'professional cleaning maintains business standards and creates positive impressions',
                'pain_points': 'maintaining professional standards while focusing on core business',
                'fallback_opening': "Hope business is going well at {company_name}.",
                'fallback_benefit': "Professional cleaning helps maintain business standards.",
                'fallback_action': "Could we schedule a brief call to discuss your cleaning needs?"
            }
        }

    def _build_ui(self):
        # Menu
        self._create_menu()
        
        # Main notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tabs
        self._create_generation_tab()
        self._create_config_tab()
        self._create_results_tab()
        
        # Status bar
        self.status_bar = ttk.Label(self.root, text="🚀 Ready - Hybrid AI + Smart Fallbacks", anchor="w")
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self._new_project)
        file_menu.add_command(label="Load CSV", command=self._load_csv)
        file_menu.add_command(label="Create Test CSV", command=self._create_test_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

    def _create_generation_tab(self):
        self.gen_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.gen_tab, text="📧 Generate Emails")
        
        # Progress section
        progress_frame = ttk.Frame(self.gen_tab)
        progress_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(fill=tk.X, side=tk.LEFT, expand=True)
        self.progress_label = ttk.Label(progress_frame, text="")
        self.progress_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # File operations
        file_frame = ttk.LabelFrame(self.gen_tab, text="Step 1: Load Prospects", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        file_buttons = ttk.Frame(file_frame)
        file_buttons.pack(fill=tk.X)
        
        ttk.Button(file_buttons, text="📁 Load CSV", command=self._load_csv).pack(side=tk.LEFT)
        ttk.Button(file_buttons, text="⚡ Create Test CSV", command=self._create_test_csv).pack(side=tk.LEFT, padx=(10, 0))
        self.file_status = ttk.Label(file_buttons, text="No file loaded", foreground="gray")
        self.file_status.pack(side=tk.LEFT, padx=(10, 0))
        
        # Preview
        preview_frame = ttk.LabelFrame(self.gen_tab, text="Prospects Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("Company", "Industry", "Contact", "Email", "Location")
        self.prospects_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=6)
        for col in columns:
            self.prospects_tree.heading(col, text=col)
            self.prospects_tree.column(col, width=150)
        
        scrollbar1 = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.prospects_tree.yview)
        self.prospects_tree.configure(yscrollcommand=scrollbar1.set)
        self.prospects_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar1.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Generation controls
        gen_frame = ttk.LabelFrame(self.gen_tab, text="Step 2: Generate Hybrid AI Emails", padding=10)
        gen_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Mode info
        info_frame = ttk.Frame(gen_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        info_label = ttk.Label(info_frame, text="🤖 Hybrid Mode: Tries AI first (15s timeout), falls back to smart templates if slow", 
                              foreground="blue", font=("Arial", 9))
        info_label.pack()
        
        # Controls
        controls = ttk.Frame(gen_frame)
        controls.pack(fill=tk.X)
        
        self.generate_btn = ttk.Button(controls, text="🚀 Generate Hybrid Emails", 
                                     command=self._start_generation, state=tk.DISABLED)
        self.generate_btn.pack(side=tk.LEFT)
        
        self.cancel_btn = ttk.Button(controls, text="⛔ Cancel", 
                                   command=self._cancel_generation, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        self.gen_status = ttk.Label(controls, text="Load CSV file to begin")
        self.gen_status.pack(side=tk.RIGHT)
        
        # Review section
        review_frame = ttk.LabelFrame(self.gen_tab, text="Step 3: Review & Send", padding=10)
        review_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Navigation
        nav_frame = ttk.Frame(review_frame)
        nav_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(nav_frame, text="◀ Previous", command=self._prev_email).pack(side=tk.LEFT)
        self.email_counter = ttk.Label(nav_frame, text="No emails generated")
        self.email_counter.pack(side=tk.LEFT, padx=10)
        ttk.Button(nav_frame, text="Next ▶", command=self._next_email).pack(side=tk.LEFT)
        
        ttk.Button(nav_frame, text="📧 Send Current", command=self._send_current).pack(side=tk.RIGHT)
        ttk.Button(nav_frame, text="📬 Send All", command=self._send_all).pack(side=tk.RIGHT, padx=(0, 10))
        
        # Editor
        editor_frame = ttk.Frame(review_frame)
        editor_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(editor_frame, text="Subject:").pack(anchor="w")
        self.subject_entry = ttk.Entry(editor_frame, font=("Arial", 10))
        self.subject_entry.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(editor_frame, text="Email Body:").pack(anchor="w")
        self.email_text = scrolledtext.ScrolledText(editor_frame, height=12, font=("Arial", 10))
        self.email_text.pack(fill=tk.BOTH, expand=True)

    def _create_config_tab(self):
        self.config_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.config_tab, text="⚙️ Configuration")
        
        config_frame = ttk.LabelFrame(self.config_tab, text="Hybrid AI Configuration", padding=10)
        config_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        config_text = scrolledtext.ScrolledText(config_frame, height=25, wrap=tk.WORD)
        config_text.pack(fill=tk.BOTH, expand=True)
        
        # Test AI and show config
        ai_status = self._test_ai_connection()
        
        config_info = f"""📧 Email Configuration:
Email: {self.config.get('email', 'from_email')}
Status: {'✅ Ready' if self.config.is_email_configured() else '❌ Configure in config.yaml'}

🏢 Company Information:
Name: {self.config.get('company', 'name')}
Website: {self.config.get('company', 'website')}
Phone: {self.config.get('company', 'phone')}

🤖 AI Configuration:
Model: {self.config.get('ollama', 'model')}
URL: {self.config.get('ollama', 'url')}
Status: {ai_status}

⚡ Hybrid Generation Settings:
• Fast AI Timeout: {self.ai_timeout_fast}s (quick attempt)
• Slow AI Timeout: {self.ai_timeout_slow}s (retry)
• Max Workers: {self.max_workers} (prevents overwhelming slow AI)
• Fallback: Smart industry-specific templates

🏠 Supported Industries:
• Education (schools, colleges, universities)
• Construction (commercial, residential projects)
• Technology (offices, startups, tech companies)
• Manufacturing (industrial facilities, plants)
• Residential (homes, apartments, condos)
• Office (professional services, businesses)
• Default (all other business types)

🔄 How Hybrid Mode Works:
1. Try AI generation with {self.ai_timeout_fast}s timeout
2. If AI times out → Use smart template fallback
3. If AI fails → Retry once with {self.ai_timeout_slow}s timeout
4. If still fails → Use template with industry customization

💡 This ensures you always get emails, even if AI is slow!

Overall Status: {self.config.get_config_status()}
"""
        
        config_text.insert(1.0, config_info)
        config_text.config(state=tk.DISABLED)

    def _create_results_tab(self):
        self.results_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.results_tab, text="📊 Results")
        
        results_frame = ttk.LabelFrame(self.results_tab, text="Generation Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("Company", "Email", "Method", "Status", "Time")
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=150)
        
        scrollbar2 = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar2.set)
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar2.pack(side=tk.RIGHT, fill=tk.Y)

    def _start_ui_updater(self):
        """Start the UI update loop"""
        self._process_ui_events()

    def _process_ui_events(self):
        """Process events from background threads"""
        try:
            while True:
                event = self.result_queue.get_nowait()
                self._handle_ui_event(event)
        except queue.Empty:
            pass
        
        self.root.after(100, self._process_ui_events)  # Update every 100ms

    def _handle_ui_event(self, event):
        """Handle UI events from background threads"""
        event_type = event.get("type")
        
        if event_type == "progress":
            current, total = event["current"], event["total"]
            self.progress.config(maximum=total, value=current)
            self.progress_label.config(text=f"{current}/{total}")
            if "message" in event:
                self.status_bar.config(text=event["message"])
        
        elif event_type == "generation_complete":
            self._handle_generation_complete(event["results"])
        
        elif event_type == "status":
            self.status_bar.config(text=event["message"])

    def _test_ai_connection(self):
        """Test AI connection with proper timeout handling"""
        try:
            tags_url = self.config.get('ollama', 'url').replace('/api/generate', '/api/tags')
            response = requests.get(tags_url, timeout=3)
            if response.status_code == 200:
                return "✅ Connected (fast)"
            else:
                return "⚠️ Connected (server error)"
        except requests.exceptions.Timeout:
            return "🐌 Connected (very slow - will use fallbacks)"
        except Exception as e:
            return f"❌ Not available ({str(e)[:30]}...)"

    def _load_csv(self):
        """Load prospects from CSV file with flexible field handling"""
        file_path = filedialog.askopenfilename(
            title="Select Prospects CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            df = pd.read_csv(file_path)
            
            # Check for required columns (more flexible)
            required = ["Company Name"]
            if "Company Name" not in df.columns:
                # Try common variations
                company_cols = [col for col in df.columns if any(word in col.lower() for word in ['company', 'business', 'name'])]
                if company_cols:
                    df = df.rename(columns={company_cols[0]: "Company Name"})
                else:
                    messagebox.showerror("Invalid CSV", "Need at least a company name column")
                    return
            
            # Handle missing Email column
            if "Email" not in df.columns:
                # Create placeholder emails that can be filled later
                df["Email"] = df["Company Name"].apply(lambda x: f"contact@{x.lower().replace(' ', '').replace('&', 'and')}.com")
                messagebox.showinfo("Email Column Missing", 
                                  "No email column found. Generated placeholder emails.\n" +
                                  "You can edit these in the generated emails or find real contacts later.")
            
            # Handle missing Contact Name
            if "Contact Name" not in df.columns:
                df["Contact Name"] = "Manager"  # Default contact name
            
            # Handle missing Industry - try to infer or set default
            if "Industry" not in df.columns:
                # Try to infer from company name or other fields
                def infer_industry(row):
                    company = str(row.get("Company Name", "")).lower()
                    if any(word in company for word in ['plumbing', 'hvac', 'ac', 'contractor']):
                        return 'Construction'
                    elif any(word in company for word in ['technology', 'tech', 'it', 'computer']):
                        return 'Technology'
                    elif any(word in company for word in ['consulting', 'consultant']):
                        return 'Professional Services'
                    elif any(word in company for word in ['security', 'exterminating', 'pest']):
                        return 'Services'
                    elif any(word in company for word in ['coffee', 'cafe', 'restaurant']):
                        return 'Food & Beverage'
                    elif any(word in company for word in ['coworking', 'workspace', 'office']):
                        return 'Office'
                    else:
                        return 'Business'
                
                df["Industry"] = df.apply(infer_industry, axis=1)
            
            # Handle missing Location - try to use City or set default
            if "Location" not in df.columns:
                if "City" in df.columns:
                    df["Location"] = df["City"] + ", LA"
                else:
                    df["Location"] = "Louisiana"
            
            # Handle missing Company Size
            if "Company Size" not in df.columns:
                df["Company Size"] = "Small Business"
            
            # Handle missing Notes
            if "Notes" not in df.columns:
                df["Notes"] = ""
            
            # Fill any NaN values
            df = df.fillna("")
            
            # Remove rows with empty company names
            df = df[df["Company Name"].str.strip() != ""]
            
            self.prospects = df.to_dict("records")
            
            # Show what was loaded
            missing_fields = []
            if "Email" not in df.columns or df["Email"].str.contains("@").sum() == 0:
                missing_fields.append("emails (generated placeholders)")
            if "Contact Name" not in df.columns:
                missing_fields.append("contact names (using 'Manager')")
            if "Industry" not in df.columns:
                missing_fields.append("industries (auto-detected)")
            
            status_text = f"✅ Loaded {len(self.prospects)} prospects"
            if missing_fields:
                status_text += f"\n📝 Auto-filled: {', '.join(missing_fields)}"
            
            self.file_status.config(text=status_text, foreground="green")
            self._refresh_prospects_tree()
            self.generate_btn.config(state=tk.NORMAL)
            self.gen_status.config(text="Ready to generate hybrid emails")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV: {e}")

    def _create_test_csv(self):
        """Create test CSV with some missing fields to demonstrate flexibility"""
        test_data = {
            "Company Name": [
                "A5 Star Plumbing Company",
                "Acadiana Security Plus", 
                "DDG Architectural Design",
                "Techneaux Technology Services",
                "Rêve Coffee Roasters",
                "Smith Family Home"
            ],
            "Industry": [
                "Plumbing / Contractors",
                "Security Services", 
                "Architectural Design",
                "",  # Missing industry
                "Food & Beverage",
                "Residential"
            ],
            "City": [
                "Youngsville",
                "Broussard",
                "Lafayette", 
                "Lafayette",
                "Lafayette",
                "Youngsville"
            ],
            "Phone": [
                "(337) 202-0246",
                "(337) 839-1880",
                "(337) 233-9914",
                "(800) 337-5313",
                "(337) 534-8336",
                "(337) 555-0123"
            ],
            # Note: No Email column to test auto-generation
            # Note: No Contact Name column to test defaults
            # Note: No Company Size column to test defaults
        }
        
        df = pd.DataFrame(test_data)
        file_path = "test_prospects_flexible.csv"
        df.to_csv(file_path, index=False)
        
        self.prospects = self._process_flexible_csv(df)
        self.file_status.config(text=f"✅ Created & loaded {len(self.prospects)} test prospects\n📝 Demonstrates missing field handling", foreground="green")
        self._refresh_prospects_tree()
        self.generate_btn.config(state=tk.NORMAL)
        self.gen_status.config(text="Ready to generate hybrid emails")
        
        messagebox.showinfo("Flexible Test CSV Created", 
                          f"Created {file_path} with missing fields to demonstrate flexibility!\n\n" +
                          "✅ Company names provided\n" +
                          "❌ Email column missing (auto-generated)\n" +
                          "❌ Contact names missing (using 'Manager')\n" +
                          "❌ Company size missing (using 'Small Business')\n" +
                          "✅ Industries mostly provided (auto-detected where missing)")

    def _process_flexible_csv(self, df):
        """Process CSV with missing fields"""
        # Handle missing Email column
        if "Email" not in df.columns:
            df["Email"] = df["Company Name"].apply(self._generate_placeholder_email)
        
        # Handle missing Contact Name
        if "Contact Name" not in df.columns:
            df["Contact Name"] = "Manager"
        
        # Handle missing Industry
        if "Industry" not in df.columns or df["Industry"].isna().any():
            df["Industry"] = df.apply(self._infer_industry, axis=1)
        
        # Handle missing Location
        if "Location" not in df.columns:
            if "City" in df.columns:
                df["Location"] = df["City"].apply(lambda x: f"{x}, LA" if x else "Louisiana")
            else:
                df["Location"] = "Louisiana"
        
        # Handle missing Company Size
        if "Company Size" not in df.columns:
            df["Company Size"] = "Small Business"
        
        # Handle missing Notes
        if "Notes" not in df.columns:
            df["Notes"] = ""
        
        # Fill NaN values
        df = df.fillna("")
        
        return df.to_dict("records")

    def _generate_placeholder_email(self, company_name):
        """Generate placeholder email from company name"""
        if not company_name:
            return "contact@company.com"
        
        # Clean up company name for email
        clean_name = company_name.lower()
        clean_name = clean_name.replace(' & ', 'and')
        clean_name = clean_name.replace('&', 'and')
        clean_name = ''.join(c for c in clean_name if c.isalnum())
        
        # Truncate if too long
        if len(clean_name) > 20:
            clean_name = clean_name[:20]
        
        return f"contact@{clean_name}.com"

    def _infer_industry(self, row):
        """Infer industry from company name and other fields"""
        company = str(row.get("Company Name", "")).lower()
        existing_industry = str(row.get("Industry", "")).strip()
        
        # If industry already provided and not empty, use it
        if existing_industry and existing_industry.lower() not in ['', 'nan', 'null']:
            return existing_industry
        
        # Infer from company name
        if any(word in company for word in ['plumbing', 'hvac', 'ac', 'contractor', 'construction']):
            return 'Construction'
        elif any(word in company for word in ['technology', 'tech', 'it', 'computer', 'software']):
            return 'Technology'
        elif any(word in company for word in ['consulting', 'consultant', 'advisory']):
            return 'Professional Services'
        elif any(word in company for word in ['security', 'exterminating', 'pest', 'cleaning']):
            return 'Services'
        elif any(word in company for word in ['coffee', 'cafe', 'restaurant', 'food', 'roaster']):
            return 'Food & Beverage'
        elif any(word in company for word in ['coworking', 'workspace', 'office', 'regus']):
            return 'Office'
        elif any(word in company for word in ['marketing', 'advertising', 'agency']):
            return 'Marketing'
        elif any(word in company for word in ['insurance', 'financial', 'trust']):
            return 'Financial Services'
        elif any(word in company for word in ['engineering', 'engineer', 'design']):
            return 'Engineering'
        else:
            return 'Business'

    def _refresh_prospects_tree(self):
        """Refresh the prospects tree view"""
        for item in self.prospects_tree.get_children():
            self.prospects_tree.delete(item)
        
        for prospect in self.prospects:
            values = (
                prospect.get("Company Name", ""),
                prospect.get("Industry", ""),
                prospect.get("Contact Name", ""),
                prospect.get("Email", ""),
                prospect.get("Location", "")
            )
            self.prospects_tree.insert("", "end", values=values)

    def _start_generation(self):
        """Start hybrid email generation"""
        if not self.prospects:
            messagebox.showwarning("No Data", "Please load a CSV file first.")
            return
        
        if self.is_generating:
            return
        
        self.is_generating = True
        self.cancel_event.clear()
        self.emails = []
        
        # Update UI
        self.generate_btn.config(state=tk.DISABLED, text="🤖 Generating...")
        self.cancel_btn.config(state=tk.NORMAL)
        self.progress.config(maximum=len(self.prospects), value=0)
        self.progress_label.config(text="0/0")
        self.status_bar.config(text="Starting hybrid generation...")
        
        # Start background generation
        thread = threading.Thread(target=self._generate_worker, daemon=True)
        thread.start()

    def _cancel_generation(self):
        """Cancel email generation"""
        self.cancel_event.set()
        self.status_bar.config(text="Cancelling generation...")

    def _generate_worker(self):
        """Background worker for email generation"""
        total = len(self.prospects)
        ai_success = 0
        fallback_used = 0
        failed = 0
        
        # Use single worker to avoid overwhelming slow Ollama
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks
            future_to_prospect = {
                executor.submit(self._generate_single_email, i, prospect): (i, prospect)
                for i, prospect in enumerate(self.prospects)
            }
            
            # Process completed tasks
            for future in as_completed(future_to_prospect):
                if self.cancel_event.is_set():
                    break
                
                i, prospect = future_to_prospect[future]
                try:
                    email_data = future.result()
                    self.emails.append(email_data)
                    
                    # Count methods
                    method = email_data["method"]
                    if method == "ai_fast" or method == "ai_slow":
                        ai_success += 1
                    elif method == "fallback":
                        fallback_used += 1
                    else:
                        failed += 1
                    
                    # Update progress
                    current = len(self.emails)
                    self.result_queue.put({
                        "type": "progress",
                        "current": current,
                        "total": total,
                        "message": f"Generated {current}/{total}: {email_data['method']} for {prospect.get('Company Name', 'Unknown')}"
                    })
                    
                except Exception as e:
                    print(f"Error processing {prospect.get('Company Name', 'Unknown')}: {e}")
                    failed += 1
        
        # Sort emails by original order
        self.emails.sort(key=lambda x: x["original_index"])
        
        # Send completion event
        self.result_queue.put({
            "type": "generation_complete",
            "results": {
                "ai_success": ai_success,
                "fallback_used": fallback_used,
                "failed": failed,
                "total": len(self.emails)
            }
        })

    def _generate_single_email(self, index, prospect):
        """Generate a single email using hybrid approach"""
        start_time = time.time()
        
        try:
            # Try AI generation first (fast timeout)
            subject, body = self._try_ai_generation(prospect, self.ai_timeout_fast)
            method = "ai_fast"
        except (requests.exceptions.Timeout, TimeoutError):
            try:
                # Retry with longer timeout
                subject, body = self._try_ai_generation(prospect, self.ai_timeout_slow)
                method = "ai_slow"
            except:
                # Fall back to smart template
                subject, body = self._generate_fallback_email(prospect)
                method = "fallback"
        except Exception:
            # Fall back to smart template
            subject, body = self._generate_fallback_email(prospect)
            method = "fallback"
        
        generation_time = time.time() - start_time
        
        return {
            "original_index": index,
            "prospect": prospect,
            "subject": subject,
            "body": body,
            "method": method,
            "generation_time": f"{generation_time:.1f}s",
            "generated_at": datetime.now().isoformat(),
            "sent": False
        }

    def _try_ai_generation(self, prospect, timeout):
        """Try AI generation with specified timeout"""
        company_name = prospect.get("Company Name", "Your Company")
        industry = prospect.get("Industry", "business").lower()
        contact_name = prospect.get("Contact Name", "Manager")
        
        # Get industry data for context
        industry_key = self._map_industry(industry)
        industry_info = self.industry_data[industry_key]
        
        # Create focused prompt for AI
        prompt = f"""Write 3 short customized lines for {company_name} ({industry}):

1. Personal opening line mentioning {company_name} (max 12 words)
2. Industry-specific benefit about why {industry_info['benefits']} (max 15 words)  
3. Call to action for meeting (max 10 words)

Company: {company_name}
Industry: {industry}
Contact: {contact_name}

Format exactly (no colons in the content):
OPEN: [opening line]
BENEFIT: [why cleaning matters for this industry]
ACTION: [meeting request]

Example:
OPEN: Hope your team at Acme Corp is having a productive week
BENEFIT: Clean workspaces boost productivity and create positive impressions
ACTION: Could we schedule a brief call about your needs"""

        # Make AI request
        ollama_config = self.config.get("ollama")
        payload = {
            "model": ollama_config["model"],
            "prompt": prompt,
            "stream": False
        }
        
        response = self.session.post(
            ollama_config["url"],
            json=payload,
            timeout=timeout
        )
        response.raise_for_status()
        
        ai_text = response.json().get("response", "").strip()
        opening, benefit, action = self._parse_ai_response(ai_text)
        
        # Generate email using AI customizations
        subject = f"Professional Cleaning Services for {company_name}"
        body = self._build_email_body(prospect, industry_info, opening, benefit, action)
        
        return subject, body

    def _parse_ai_response(self, ai_text):
        """Parse AI response into components"""
        opening = benefit = action = ""
        
        for line in ai_text.split('\n'):
            line = line.strip()
            if line.startswith('OPEN:'):
                opening = line[5:].strip()
            elif line.startswith('BENEFIT:'):
                benefit = line[7:].strip()
            elif line.startswith('ACTION:'):
                action = line[7:].strip()
        
        # Fallback if parsing fails
        if not all([opening, benefit, action]):
            raise ValueError("AI response incomplete")
        
        return opening, benefit, action

    def _generate_fallback_email(self, prospect):
        """Generate email using smart templates"""
        company_name = prospect.get("Company Name", "Your Company")
        industry = prospect.get("Industry", "business").lower()
        
        # Get industry-specific template
        industry_key = self._map_industry(industry)
        industry_info = self.industry_data[industry_key]
        
        # Use fallback content
        opening = industry_info['fallback_opening'].format(company_name=company_name)
        benefit = industry_info['fallback_benefit']
        action = industry_info['fallback_action']
        
        subject = f"Professional Cleaning Services for {company_name}"
        body = self._build_email_body(prospect, industry_info, opening, benefit, action)
        
        return subject, body

    def _build_email_body(self, prospect, industry_info, opening, benefit, action):
        """Build email body using customizations and industry data"""
        company_name = prospect.get("Company Name", "Your Company")
        contact_name = prospect.get("Contact Name", "Manager")
        company_config = self.config.get_company_info()
        
        # Build service list
        services = industry_info['services']
        service_list = '\n'.join(f"• {service}" for service in services[:4])
        
        # Clean up benefit text - remove colons and extra punctuation
        benefit_clean = benefit.strip()
        if benefit_clean.startswith(':'):
            benefit_clean = benefit_clean[1:].strip()
        if benefit_clean.endswith(':'):
            benefit_clean = benefit_clean[:-1].strip()
        
        # Ensure proper capitalization
        if benefit_clean and not benefit_clean[0].isupper():
            benefit_clean = benefit_clean[0].upper() + benefit_clean[1:]
        
        # Ensure proper ending punctuation
        if benefit_clean and not benefit_clean.endswith('.'):
            benefit_clean += '.'

        body = f"""Dear {contact_name},

{opening}

Fresh Start Cleaning Co. specializes in professional cleaning for businesses like {company_name}. {benefit_clean}

Our services include:
{service_list}

With over 10+ years of experience serving Louisiana businesses, we're Licensed, Bonded, and Insured. Our local team provides reliable, professional service tailored to your specific needs.

{action}

Best regards,
Fresh Start Cleaning Co.
{company_config.get('phone', '')}
{company_config.get('website', '')}"""

        return body

    def _map_industry(self, industry):
        """Map industry string to category"""
        industry = industry.lower()
        
        if any(word in industry for word in ['education', 'school', 'college', 'university', 'campus']):
            return 'education'
        elif any(word in industry for word in ['construction', 'building', 'contractor', 'builder']):
            return 'construction'
        elif any(word in industry for word in ['technology', 'tech', 'software', 'it', 'startup']):
            return 'technology'
        elif any(word in industry for word in ['manufacturing', 'industrial', 'factory', 'plant']):
            return 'manufacturing'
        elif any(word in industry for word in ['residential', 'home', 'house', 'apartment', 'family']):
            return 'residential'
        elif any(word in industry for word in ['office', 'professional', 'services', 'business']):
            return 'office'
        else:
            return 'default'

    def _handle_generation_complete(self, results):
        """Handle completion of email generation"""
        self.is_generating = False
        self.generate_btn.config(state=tk.NORMAL, text="🚀 Generate Hybrid Emails")
        self.cancel_btn.config(state=tk.DISABLED)
        
        ai_success = results["ai_success"]
        fallback_used = results["fallback_used"]
        failed = results["failed"]
        total = results["total"]
        
        # Update status
        self.gen_status.config(text=f"✅ AI: {ai_success} | 📝 Fallback: {fallback_used} | ❌ Failed: {failed}")
        self.progress_label.config(text="Complete!")
        
        if self.emails:
            self.current_idx = 0
            self._display_current_email()
            self._refresh_results_tree()
            
            # Show summary
            summary = f"Generated {total} emails!\n\n"
            summary += f"🤖 AI Generated: {ai_success}\n"
            summary += f"📝 Smart Fallbacks: {fallback_used}\n"
            if failed > 0:
                summary += f"❌ Failed: {failed}\n"
            summary += f"\nAll emails ready for review and sending!"
            
            messagebox.showinfo("Generation Complete", summary)
        
        self.status_bar.config(text=f"✅ Generated {total} emails - {ai_success} AI, {fallback_used} fallback")

    def _display_current_email(self):
        """Display current email in editor"""
        if not self.emails:
            return
        
        email = self.emails[self.current_idx]
        company_name = email["prospect"].get("Company Name", "Unknown")
        method = email["method"]
        time_taken = email["generation_time"]
        
        # Method icons
        method_icons = {
            "ai_fast": "🤖⚡",
            "ai_slow": "🤖🐌", 
            "fallback": "📝",
            "failed": "❌"
        }
        
        icon = method_icons.get(method, "📧")
        self.email_counter.config(text=f"{icon} Email {self.current_idx + 1} of {len(self.emails)} - {company_name} ({time_taken})")
        
        # Load content
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, email["subject"])
        
        self.email_text.delete(1.0, tk.END)
        self.email_text.insert(1.0, email["body"])

    def _prev_email(self):
        """Navigate to previous email"""
        if not self.emails:
            return
        self._save_current_edits()
        if self.current_idx > 0:
            self.current_idx -= 1
            self._display_current_email()

    def _next_email(self):
        """Navigate to next email"""
        if not self.emails:
            return
        self._save_current_edits()
        if self.current_idx < len(self.emails) - 1:
            self.current_idx += 1
            self._display_current_email()

    def _save_current_edits(self):
        """Save current email edits"""
        if self.emails:
            email = self.emails[self.current_idx]
            email["subject"] = self.subject_entry.get()
            email["body"] = self.email_text.get(1.0, tk.END).strip()

    def _refresh_results_tree(self):
        """Refresh the results tree"""
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        for email in self.emails:
            method_display = {
                "ai_fast": "🤖⚡ AI Fast",
                "ai_slow": "🤖🐌 AI Slow",
                "fallback": "📝 Smart Template",
                "failed": "❌ Failed"
            }.get(email["method"], email["method"])
            
            status = "✅ Sent" if email.get("sent") else "📝 Draft"
            
            values = (
                email["prospect"].get("Company Name", "")[:25],
                email["prospect"].get("Email", "")[:30],
                method_display,
                status,
                email["generation_time"]
            )
            self.results_tree.insert("", "end", values=values)

    def _send_current(self):
        """Send current email"""
        if not self.emails:
            messagebox.showwarning("No Emails", "No emails to send.")
            return
        
        if not self.config.is_email_configured():
            messagebox.showwarning("Email Not Configured", "Please configure email in config.yaml")
            return
        
        self._save_current_edits()
        email = self.emails[self.current_idx]
        
        company_name = email["prospect"].get("Company Name", "Unknown")
        recipient = email["prospect"].get("Email", "")
        
        if messagebox.askyesno("Confirm Send", f"Send email to {recipient} ({company_name})?"):
            success = self._send_email(email)
            if success:
                email["sent"] = True
                email["sent_at"] = datetime.now().isoformat()
                self._refresh_results_tree()
                messagebox.showinfo("Success", f"✅ Email sent to {company_name}!")
            else:
                messagebox.showerror("Error", "❌ Failed to send email. Check configuration.")

    def _send_all(self):
        """Send all unsent emails"""
        if not self.emails:
            messagebox.showwarning("No Emails", "No emails to send.")
            return
        
        if not self.config.is_email_configured():
            messagebox.showwarning("Email Not Configured", "Please configure email in config.yaml")
            return
        
        unsent = [email for email in self.emails if not email.get("sent")]
        if not unsent:
            messagebox.showinfo("Nothing to Send", "All emails have been sent.")
            return
        
        if messagebox.askyesno("Confirm Send All", f"Send {len(unsent)} unsent emails?"):
            sent_count = 0
            for email in unsent:
                if self._send_email(email):
                    email["sent"] = True
                    email["sent_at"] = datetime.now().isoformat()
                    sent_count += 1
                time.sleep(1)  # Small delay between sends
            
            self._refresh_results_tree()
            messagebox.showinfo("Batch Send Complete", f"✅ Sent {sent_count} out of {len(unsent)} emails.")

    def _send_email(self, email_data):
        """Send a single email via SMTP"""
        try:
            email_config = self.config.get_email_config()
            company_info = self.config.get_company_info()
            
            msg = MIMEMultipart()
            msg["From"] = f"{company_info['name']} <{email_config['from_email']}>"
            msg["To"] = email_data["prospect"]["Email"]
            msg["Subject"] = email_data["subject"]
            
            body = email_data["body"]
            msg.attach(MIMEText(body, "plain"))
            
            server = smtplib.SMTP(email_config["smtp_server"], email_config["smtp_port"])
            server.starttls()
            server.login(email_config["from_email"], email_config["from_password"])
            server.sendmail(email_config["from_email"], email_data["prospect"]["Email"], msg.as_string())
            server.quit()
            
            return True
        except Exception as e:
            print(f"SMTP Error: {e}")
            return False

    def _new_project(self):
        """Start a new project"""
        if messagebox.askyesno("New Project", "Clear all current data and start fresh?"):
            self.cancel_event.set()
            self.prospects = []
            self.emails = []
            self.current_idx = 0
            self.is_generating = False
            
            # Reset UI
            self.prospects_tree.delete(*self.prospects_tree.get_children())
            self.results_tree.delete(*self.results_tree.get_children())
            self.file_status.config(text="No file loaded", foreground="gray")
            self.generate_btn.config(state=tk.DISABLED, text="🚀 Generate Hybrid Emails")
            self.cancel_btn.config(state=tk.DISABLED)
            self.gen_status.config(text="Load CSV file to begin")
            self.email_counter.config(text="No emails generated")
            self.subject_entry.delete(0, tk.END)
            self.email_text.delete(1.0, tk.END)
            self.progress.config(value=0)
            self.progress_label.config(text="")
            self.status_bar.config(text="🚀 Ready for new project")

def main():
    """Main application entry point"""
    root = tk.Tk()
    
    # Window optimization
    try:
        root.tk.call('tk', 'scaling', 1.0)
    except:
        pass
    
    app = HybridEmailGenerator(root)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()