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
        self.ai_timeout_fast = 20   # Generous for fast models
        self.ai_timeout_slow = 45   # Retry timeout
        self.max_workers = 1        # Keep sequential
        self.debug_mode = True      # Always debug until working
        print(f"🤖 AI configured: Fast={self.ai_timeout_fast}s, Slow={self.ai_timeout_slow}s")

        
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
    
    def _parse_ai_response(self, ai_text):
        """Parse AI response and clean labels"""
        if hasattr(self, 'debug_mode') and self.debug_mode:
            print(f"   📝 Raw AI response: {repr(ai_text[:100])}...")
        
        opening = benefit = action = ""
        
        # Method 1: Try exact format first (OPEN:, BENEFIT:, ACTION:)
        for line in ai_text.split('\n'):
            line = line.strip()
            if line.startswith('OPEN:'):
                opening = line[5:].strip()  # Remove "OPEN:" label
            elif line.startswith('BENEFIT:'):
                benefit = line[7:].strip()  # Remove "BENEFIT:" label
            elif line.startswith('ACTION:'):
                action = line[7:].strip()   # Remove "ACTION:" label
        
        # Method 2: If exact format failed, try flexible extraction
        if not all([opening, benefit, action]):
            if hasattr(self, 'debug_mode') and self.debug_mode:
                print(f"   🔄 Trying flexible extraction...")
            
            # Split into lines and clean up
            lines = [line.strip() for line in ai_text.split('\n') if line.strip() and len(line.strip()) > 5]
            
            # Remove any numbered lines (1., 2., 3.) and labels
            clean_lines = []
            for line in lines:
                # Remove leading numbers, bullets, and labels
                line = line.strip('123456789.- ')
                
                # Remove any remaining labels that might be embedded
                for label in ['OPEN:', 'BENEFIT:', 'ACTION:']:
                    if line.startswith(label):
                        line = line[len(label):].strip()
                
                if line and len(line) > 5:
                    clean_lines.append(line)
            
            # Try to assign based on content keywords
            for line in clean_lines:
                lower_line = line.lower()
                
                # Opening: contains greetings, company name, or hope/hello
                if not opening and any(word in lower_line for word in ['hope', 'hello', 'greetings', 'hi', 'good', 'team']):
                    opening = line
                    continue
                
                # Benefit: contains cleaning/business benefits
                if not benefit and any(word in lower_line for word in ['clean', 'professional', 'productivity', 'maintain', 'environment', 'image', 'standards']):
                    benefit = line
                    continue
                
                # Action: contains meeting/call requests
                if not action and any(word in lower_line for word in ['call', 'meeting', 'schedule', 'discuss', 'available', 'talk', 'contact']):
                    action = line
                    continue
            
            # If still missing pieces, assign by position
            if not all([opening, benefit, action]) and len(clean_lines) >= 3:
                if not opening: opening = clean_lines[0]
                if not benefit: benefit = clean_lines[1] 
                if not action: action = clean_lines[2]
        
        # Method 3: Generate fallbacks for missing pieces
        if not opening:
            company_name = getattr(self, '_current_prospect_name', 'your company')
            opening = f"Hope business is going well at {company_name}"
        
        if not benefit:
            benefit = "Professional cleaning services help maintain business standards and create positive impressions"
        
        if not action:
            action = "Could we schedule a brief call to discuss your cleaning needs"
        
        # CRITICAL: Clean up any remaining labels and formatting
        def clean_text(text):
            # Remove any remaining labels
            for label in ['OPEN:', 'BENEFIT:', 'ACTION:', 'open:', 'benefit:', 'action:']:
                text = text.replace(label, '').strip()
            
            # Remove extra quotes, formatting
            text = text.strip(' "\'.,!?-*')
            
            # Ensure proper capitalization
            if text and not text[0].isupper():
                text = text[0].upper() + text[1:]
            
            return text
        
        opening = clean_text(opening)
        benefit = clean_text(benefit)
        action = clean_text(action)
        
        # Ensure proper punctuation
        if opening and not opening.endswith(('.', '!', '?')):
            opening += '.'
        if benefit and not benefit.endswith(('.', '!', '?')):
            benefit += '.'
        if action and not action.endswith(('.', '!', '?')):
            action += '?'
        
        if hasattr(self, 'debug_mode') and self.debug_mode:
            print(f"   ✅ CLEANED parsing results:")
            print(f"      OPEN: '{opening}'")
            print(f"      BENEFIT: '{benefit}'") 
            print(f"      ACTION: '{action}'")
        
        return opening, benefit, action

    def _load_config(self):
        try:
            return YAMLConfigManager()
        except Exception as e:
            messagebox.showerror("Config Error", f"Error loading config: {e}")
            raise
    def _warmup_model_if_needed(self):
        """Quick model warmup to avoid slow first request"""
        try:
            ollama_config = self.config.get("ollama")
            warmup_payload = {
                "model": ollama_config["model"], 
                "prompt": "ready",
                "stream": False,
                "options": {"num_predict": 1}
            }
            
            response = self.session.post(
                ollama_config["url"],
                json=warmup_payload, 
                timeout=10
            )
            
            if response.status_code == 200:
                print("🔥 Model warmed up and ready")
                return True
            else:
                print("⚠️ Model may be cold - first generation will be slower")
                return False
        except:
            print("⚠️ Warmup failed - first generation may be slower")
            return False


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
            'professional_services': {
                'services': ['Office cleaning', 'Reception area maintenance', 'Conference room cleaning', 'Professional space upkeep'],
                'benefits': 'professional environments create positive client impressions and boost productivity',
                'pain_points': 'maintaining professional appearance for client meetings and staff productivity',
                'fallback_opening': "Hope business is going well at {company_name}.",
                'fallback_benefit': "Professional cleaning helps maintain your business image and client impressions.",
                'fallback_action': "Could we schedule a brief call about your office cleaning needs?"
            },
            'food_beverage': {
                'services': ['Kitchen deep cleaning', 'Dining area maintenance', 'Health code compliance', 'Equipment sanitization'],
                'benefits': 'spotless facilities are essential for health compliance and customer satisfaction',
                'pain_points': 'maintaining health department standards while serving customers',
                'fallback_opening': "Hope your customers are enjoying {company_name}.",
                'fallback_benefit': "Professional cleaning is essential for health compliance and customer satisfaction.",
                'fallback_action': "Could we schedule a call to discuss your cleaning needs?"
            },
            'retail': {
                'services': ['Sales floor cleaning', 'Window cleaning', 'Restroom maintenance', 'Storage area organization'],
                'benefits': 'clean retail spaces create positive shopping experiences and drive sales',
                'pain_points': 'maintaining appealing spaces while serving customers throughout the day',
                'fallback_opening': "Hope your customers are having great experiences at {company_name}.",
                'fallback_benefit': "Clean retail spaces create positive shopping experiences.",
                'fallback_action': "Could we discuss your store cleaning needs?"
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
        file_menu.add_command(label="Load CSV/Excel", command=self._load_csv)
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
        
        ttk.Button(file_buttons, text="📁 Load CSV/Excel", command=self._load_csv).pack(side=tk.LEFT)
        ttk.Button(file_buttons, text="⚡ Create Test CSV", command=self._create_test_csv).pack(side=tk.LEFT, padx=(10, 0))
        self.file_status = ttk.Label(file_buttons, text="No file loaded", foreground="gray")
        self.file_status.pack(side=tk.LEFT, padx=(10, 0))
        
        # Preview
        preview_frame = ttk.LabelFrame(self.gen_tab, text="Prospects Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("Company", "Category", "City", "Email", "Website")
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
        
        self.gen_status = ttk.Label(controls, text="Load CSV/Excel file to begin")
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

📊 Expected Excel/CSV Fields:
• Company Name (required)
• Category (required - business category/industry)
• City (required - business location)
• Email (required - contact email)
• Website (optional - company website)

🏠 Supported Categories:
• Education, Construction, Technology, Manufacturing
• Residential, Office, Professional Services
• Food & Beverage, Retail, and more
• Auto-detected from category field

🔄 How Hybrid Mode Works:
1. Try AI generation with {self.ai_timeout_fast}s timeout
2. If AI times out → Use smart template fallback
3. If AI fails → Retry once with {self.ai_timeout_slow}s timeout
4. If still fails → Use template with category customization

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
        """Load prospects from CSV/Excel file with your specific field structure"""
        file_path = filedialog.askopenfilename(
            title="Select Prospects File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"), 
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            # Load file based on extension
            if file_path.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)
            
            # Validate required fields
            required_fields = ["Company Name", "Category", "City", "Email"]
            missing_fields = [field for field in required_fields if field not in df.columns]
            
            if missing_fields:
                messagebox.showerror(
                    "Missing Required Fields", 
                    f"Your file is missing these required columns:\n{', '.join(missing_fields)}\n\n"
                    f"Expected columns: Company Name, Category, City, Email, Website (optional)"
                )
                return
            
            # Handle optional Website field
            if "Website" not in df.columns:
                df["Website"] = ""
                print("📝 Website column not found - added empty Website field")
            
            # Clean and validate data
            # Remove rows with empty company names
            df = df[df["Company Name"].str.strip().astype(bool)]
            
            # Remove rows with empty emails
            df = df[df["Email"].str.strip().astype(bool)]
            
            # Fill NaN values
            df = df.fillna("")
            
            # Store as records
            self.prospects = df.to_dict("records")
            
            # Success message
            self.file_status.config(
                text=f"✅ Loaded {len(self.prospects)} prospects\n📊 Fields: {', '.join(df.columns)}", 
                foreground="green"
            )
            self._refresh_prospects_tree()
            self.generate_btn.config(state=tk.NORMAL)
            self.gen_status.config(text="Ready to generate hybrid emails")
            
            print(f"📂 Successfully loaded {len(self.prospects)} prospects from {file_path}")
            print(f"📊 Columns found: {list(df.columns)}")
            
        except Exception as e:
            messagebox.showerror("File Load Error", f"Failed to load file: {str(e)}")
            print(f"❌ Error loading file: {e}")

    def _create_test_csv(self):
        """Create test CSV with your exact field structure"""
        test_data = {
            "Company Name": [
                "A5 Star Plumbing Company",
                "Acadiana Security Plus", 
                "DDG Architectural Design",
                "Techneaux Technology Services",
                "Rêve Coffee Roasters",
                "Fresh Market Solutions"
            ],
            "Category": [
                "Construction",
                "Professional Services", 
                "Professional Services",
                "Technology",
                "Food & Beverage",
                "Retail"
            ],
            "City": [
                "Youngsville",
                "Broussard",
                "Lafayette", 
                "Lafayette",
                "Lafayette",
                "Youngsville"
            ],
            "Email": [
                "info@a5starplumbing.com",
                "contact@acadianasecurity.com",
                "hello@ddgarchitectural.com",
                "support@techneaux.com",
                "info@revecoffee.com",
                "contact@freshmarketsolutions.com"
            ],
            "Website": [
                "www.a5starplumbing.com",
                "www.acadianasecurity.com",
                "www.ddgarchitectural.com",
                "www.techneaux.com",
                "www.revecoffee.com",
                ""  # Test empty website
            ]
        }
        
        df = pd.DataFrame(test_data)
        file_path = "test_prospects.csv"
        df.to_csv(file_path, index=False)
        
        # Load the test data
        self.prospects = df.to_dict("records")
        self.file_status.config(
            text=f"✅ Created & loaded {len(self.prospects)} test prospects\n📊 Perfect field structure", 
            foreground="green"
        )
        self._refresh_prospects_tree()
        self.generate_btn.config(state=tk.NORMAL)
        self.gen_status.config(text="Ready to generate hybrid emails")
        
        messagebox.showinfo(
            "Test CSV Created", 
            f"Created {file_path} with your exact field structure!\n\n" +
            "✅ Company Name, Category, City, Email, Website\n" +
            "✅ Includes various business categories\n" +
            "✅ Ready for email generation testing"
        )

    def _refresh_prospects_tree(self):
        """Refresh the prospects tree view"""
        for item in self.prospects_tree.get_children():
            self.prospects_tree.delete(item)
        
        for prospect in self.prospects:
            values = (
                prospect.get("Company Name", "")[:30],
                prospect.get("Category", "")[:20],
                prospect.get("City", "")[:15],
                prospect.get("Email", "")[:30],
                prospect.get("Website", "")[:30]
            )
            self.prospects_tree.insert("", "end", values=values)
    
    def _pre_warm_model(self):
        """Pre-warm model to ensure it stays loaded"""
        try:
            ollama_config = self.config.get("ollama")
            
            # Send a warming request
            payload = {
                "model": ollama_config["model"],
                "prompt": "Ready for email generation",
                "stream": False,
                "options": {
                    "num_predict": 5,
                    "temperature": 0.7
                }
            }
            
            print(f"🔥 Pre-warming {ollama_config['model']}...")
            start_time = time.time()
            
            response = self.session.post(
                ollama_config["url"],
                json=payload,
                timeout=60
            )
            
            duration = time.time() - start_time
            
            if response.status_code == 200:
                print(f"✅ Model warmed in {duration:.1f}s - ready for {len(self.prospects)} emails")
                return True
            else:
                print(f"❌ Warmup failed: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"❌ Warmup error: {e}")
            return False
      
    def _start_generation(self):
        """Start generation with model warmup"""
        if not self.prospects:
            messagebox.showwarning("No Data", "Please load a CSV/Excel file first.")
            return
        
        if self.is_generating:
            return
        
        # Warmup AI model FIRST
        self.status_bar.config(text="🔥 Warming up AI model for ALL emails...")
        self.root.update()
        
        # Pre-warm the model
        success = self._pre_warm_model()
        if success:
            self.status_bar.config(text="✅ AI ready - generating ALL emails with AI...")
        else:
            self.status_bar.config(text="⚠️ AI warmup failed - may use some templates...")
        
        self.root.update()
        time.sleep(1)
        
        self.is_generating = True
        self.cancel_event.clear()
        self.emails = []
        
        # Update UI
        self.generate_btn.config(state=tk.DISABLED, text="🤖 AI Generating...")
        self.cancel_btn.config(state=tk.NORMAL)
        self.progress.config(maximum=len(self.prospects), value=0)
        self.progress_label.config(text="0/0")
        
        # Start background generation
        thread = threading.Thread(target=self._generate_worker_sequential, daemon=True)
        thread.start()
    
    def _generate_single_email_with_retry(self, index, prospect):
        """Generate single email with aggressive AI retry"""
        start_time = time.time()
        company_name = prospect.get("Company Name", "Unknown")
        
        # Try AI with multiple attempts
        for attempt in range(3):  # Up to 3 attempts per email
            try:
                timeout = self.ai_timeout_fast if attempt == 0 else self.ai_timeout_slow
                print(f"   🤖 AI attempt {attempt+1} (timeout: {timeout}s)")
                
                subject, body = self._try_ai_generation(prospect, timeout)
                method = "ai_fast" if attempt == 0 else "ai_slow"
                
                generation_time = time.time() - start_time
                print(f"   ✅ AI SUCCESS on attempt {attempt+1} ({generation_time:.1f}s)")
                
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
                
            except requests.exceptions.Timeout:
                print(f"   ⏰ AI timeout on attempt {attempt+1}")
                continue
            except Exception as e:
                print(f"   ❌ AI error on attempt {attempt+1}: {e}")
                continue
        
        # If all AI attempts failed, use template
        print(f"   📝 All AI attempts failed - using template")
        subject, body = self._generate_fallback_email(prospect)
        generation_time = time.time() - start_time
        
        return {
            "original_index": index,
            "prospect": prospect,
            "subject": subject,
            "body": body,
            "method": "fallback",
            "generation_time": f"{generation_time:.1f}s",
            "generated_at": datetime.now().isoformat(),
            "sent": False
        }
        
    def _generate_worker_sequential(self):
        """Sequential generation to keep model hot"""
        total = len(self.prospects)
        ai_success = 0
        fallback_used = 0
        failed = 0
        
        print(f"\n🚀 SEQUENTIAL AI GENERATION - Target: {total} AI emails")
        print("=" * 60)
        
        for i, prospect in enumerate(self.prospects):
            if self.cancel_event.is_set():
                break
            
            company_name = prospect.get("Company Name", "Unknown")
            print(f"\n📧 Email {i+1}/{total}: {company_name}")
            
            try:
                # Generate with debugging
                email_data = self._generate_single_email_with_retry(i, prospect)
                self.emails.append(email_data)
                
                # Count and report
                method = email_data["method"]
                if method in ["ai_fast", "ai_slow"]:
                    ai_success += 1
                    print(f"   🎉 AI SUCCESS ({method}) for {company_name}")
                elif method == "fallback":
                    fallback_used += 1
                    print(f"   📝 TEMPLATE used for {company_name}")
                else:
                    failed += 1
                    print(f"   ❌ FAILED for {company_name}")
                
                # Update progress
                current = len(self.emails)
                self.result_queue.put({
                    "type": "progress",
                    "current": current,
                    "total": total,
                    "message": f"AI: {ai_success}, Templates: {fallback_used} - {company_name}"
                })
                
                # Keep model warm between requests
                if i < total - 1:
                    print(f"   ⏸️ Keeping model warm...")
                    time.sleep(1)  # Brief pause
                
            except Exception as e:
                print(f"❌ Processing error for {company_name}: {e}")
                failed += 1
        
        print(f"\n🎯 FINAL RESULTS:")
        print(f"   🤖 AI Generated: {ai_success}/{total}")
        print(f"   📝 Templates: {fallback_used}/{total}")
        print(f"   ❌ Failed: {failed}/{total}")
        
        # Send completion
        self.result_queue.put({
            "type": "generation_complete",
            "results": {
                "ai_success": ai_success,
                "fallback_used": fallback_used,
                "failed": failed,
                "total": len(self.emails)
            }
        })
    
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
        category = prospect.get("Category", "business").lower()
        city = prospect.get("City", "Louisiana")
        website = prospect.get("Website", "")
        
        # Get industry data for context
        industry_key = self._map_category_to_industry(category)
        industry_info = self.industry_data[industry_key]
        
        # Create focused prompt for AI with all available data
        website_context = f"I noticed their website is {website} - " if website else ""
        
        prompt = f"""Write 3 short customized lines for {company_name} ({category}) in {city}:

1. Professional opening email line (max 12 words This is the first time Fresh Start Cleaning Louisiana,LLC reach out to the company and this is an opening email) 
2. Professional Industry-specific benefits {industry_info['benefits']} (max 15 words how Fresh Start Cleaning is useful to their business)  
3. Professional Call to action to schedule a call or zoom meeting(max 10 words)

Company: {company_name}
Category: {category}
City: {city}
{f"Website: {website}" if website else ""}

Format exactly (no colons in the content):
OPEN: [opening line]
BENEFIT: [why cleaning matters for this category]
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
        if hasattr(self, 'debug_mode') and self.debug_mode:
            print(f"   📝 Full AI response: {repr(ai_text)}")
        opening, benefit, action = self._parse_ai_response(ai_text)
        
        # Generate email using AI customizations
        subject = f"Professional Cleaning Services for {company_name}"
        body = self._build_email_body(prospect, industry_info, opening, benefit, action)
        
        return subject, body


    def _generate_fallback_email(self, prospect):
        """Generate email using smart templates"""
        company_name = prospect.get("Company Name", "Your Company")
        category = prospect.get("Category", "business").lower()
        
        # Get industry-specific template
        industry_key = self._map_category_to_industry(category)
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
        city = prospect.get("City", "Louisiana")
        website = prospect.get("Website", "")
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

        # Build location context
        location_context = f"We're based locally and serve businesses throughout {city} and the surrounding Louisiana area."
        
        # Build website reference if available
        # website_reference = f"I had a chance to look at {website} and " if website else ""

        body = f"""Dear {company_name},

{opening}

Fresh Start Cleaning Louisiana, LLC. specializes in professional cleaning for businesses like {company_name}. {benefit_clean}

Our services include:
{service_list}

{location_context} With over 5+ years of experience serving Louisiana businesses, we're licensed, bonded, and insured. Our local team provides reliable, professional service tailored to your specific needs.

{action}

Best regards,
Fresh Start Cleaning Louisiana, LLC.
{company_config.get('phone', '')}
{company_config.get('website', '')}"""

        return body

    def _map_category_to_industry(self, category):
        """Map category string to industry category"""
        category = category.lower()
        
        if any(word in category for word in ['education', 'preschool', 'school', 'academy', 'college', 'university', 'campus', 'steam']):
            return 'education'
        elif any(word in category for word in ['construction', 'building', 'contractor', 'builder', 'plumbing', 'hvac', 'realty']):
            return 'construction'
        elif any(word in category for word in ['technology', 'tech', 'software', 'it', 'startup', 'computer']):
            return 'technology'
        elif any(word in category for word in ['manufacturing', 'industrial', 'factory', 'plant']):
            return 'manufacturing'
        elif any(word in category for word in ['residential', 'home', 'house', 'apartment', 'family']):
            return 'residential'
        elif any(word in category for word in ['office', 'professional services', 'consulting', 'consultant']):
            return 'professional_services'
        elif any(word in category for word in ['food', 'beverage', 'restaurant', 'cafe', 'coffee', 'roaster']):
            return 'food_beverage'
        elif any(word in category for word in ['retail', 'store', 'shop', 'market']):
            return 'retail'
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
            self.gen_status.config(text="Load CSV/Excel file to begin")
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