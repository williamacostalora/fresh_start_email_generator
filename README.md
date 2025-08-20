# Fresh Start Cleaning - Hybrid AI Email Generator

A professional email automation tool built for Fresh Start Cleaning Co. in Louisiana. This application uses hybrid AI generation (Ollama + smart fallbacks) to create personalized cleaning service emails for potential business clients.

##  Features

### **Hybrid AI Generation**
- **Primary**: Ollama AI model for personalized emails (15s timeout)
- **Backup**: Extended AI retry with 60s timeout  
- **Fallback**: Smart industry-specific templates
- **Result**: Always generates professional emails, even when AI is slow

### **Industry Specialization**
-  **Education** (schools, colleges, universities)
-  **Construction** (commercial, residential projects)
-  **Technology** (offices, startups, tech companies)
-  **Manufacturing** (industrial facilities, plants)
-  **Residential** (homes, apartments, families)
-  **Office** (professional services, businesses)
- 📝**Default** (all other business types)

### **Professional Features**
- **CSV Upload**: Bulk prospect management
- **Email Editor**: Review and edit before sending
- **Progress Tracking**: Real-time generation status
- **Method Tracking**: See which emails used AI vs templates
- **SMTP Integration**: Send emails directly via Gmail
- **Results Export**: Track sent emails and response rates

##  Requirements

### **Software Dependencies**
```bash
Python 3.11+
Ollama AI (local installation)
Gmail account with app password
```

### **Python Packages**
```bash
pip install PyYAML pandas requests tkinter openpyxl
```

### **AI Model**
```bash
ollama pull mistral
```

##  Installation

### **1. Clone/Download Project**
```bash
git clone <repository-url>
cd fresh_start_email_generator
```

### **2. Install Ollama**
```bash
# macOS
brew install ollama

# Linux
curl -fsSL https://ollama.ai/install.sh | sh

# Windows - Download from https://ollama.ai
```

### **3. Install AI Model**
```bash
ollama serve
ollama pull mistral
```

### **4. Install Python Dependencies**
```bash
pip3.11 install PyYAML pandas requests openpyxl
```

### **5. Configure Email Credentials**
Create `config.yaml`:
```yaml
email:
  from_email: "your_email@gmail.com"
  from_password: "your_16_character_app_password"
  smtp_server: "smtp.gmail.com"
  smtp_port: 1234
  from_name: "Your company."

company:
  name: "Fresh Start Cleaning Co."
  website: "https://freshcleaningcolouisiana.com/"
  location: "Louisiana"
  phone: "(337) XXX-XXXX"
  # ... additional company settings

ollama:
  url: "http://localhost:11434/api/generate"
  model: "mistral"
  timeout: 180
```

##  Usage

### **Quick Start**
```bash
python3.11 responsive_email_generator.py
```

### **Step-by-Step Process**

1. **Configure Settings**
   - Edit `config.yaml` with your Gmail credentials
   - Update company information

2. **Load Prospects**
   - Click "Load CSV" or "Create Test CSV"
   - Required columns: `Company Name`, `Email`
   - Optional: `Industry`, `Contact Name`, `Location`, `Notes`

3. **Generate Emails**
   - Click "Generate Hybrid Emails"
   - Watch real-time progress with method tracking
   - AI attempts with smart fallbacks ensure 100% success

4. **Review & Edit**
   - Navigate through generated emails
   - Edit subject lines and content as needed
   - See generation method (AI Fast/Slow/Template)

5. **Send Emails**
   - Send individual emails or batch send all
   - Track delivery status and results

## CSV Format

Your prospects CSV should include:

| Column Name | Required | Example |
|-------------|----------|---------|
| Company Name | ✅ Yes | Turner Industries |
| Email | ✅ Yes | facilities@turner.com |
| Industry | No | Construction |
| Contact Name | No | Facilities Manager |
| Company Size | No | 1,000+ |
| Location | No | Baton Rouge, LA |
| Notes | No | Large industrial contractor |

### **Sample CSV Data**
The project includes Louisiana business prospects:
- Major construction companies (Turner Industries, Performance Contractors)
- Technology companies (IBM Louisiana, CGI)
- Healthcare facilities (Ochsner Health, OLOL)
- Educational institutions (LSU, Tulane)
- Manufacturing plants (ExxonMobil, BASF, Dow)
- Local Youngsville/Lafayette businesses

## ⚙ Configuration

### **Gmail App Password Setup**
1. Enable 2-Step Verification in Google Account
2. Go to Security → App passwords
3. Generate password for "Mail"
4. Use 16-character password in `config.yaml`

### **Ollama Optimization**
```yaml
ollama:
  url: "http://localhost:11434/api/generate"
  model: "mistral"  # or "phi3" for faster responses
  timeout: 180      # adjust based on your hardware
```

### **Generation Mode Settings**
- **Fast AI Timeout**: 15 seconds (quick attempt)
- **Slow AI Timeout**: 60 seconds (retry)
- **Max Workers**: 1 (prevents overwhelming slow AI)
- **Fallback**: Industry-specific smart templates

## Troubleshooting

### **Common Issues**

#### **"Cannot connect to Ollama"**
```bash
# Check if Ollama is running
ps aux | grep ollama

# Start Ollama
ollama serve

# Test model
ollama run mistral "Hello"
```

#### **"Email connection failed"**
- Use Gmail app password, not regular password
- Enable 2-step verification first
- Check email/password in config.yaml

#### **"AI generation slow/timing out"**
- This is expected! Hybrid mode handles it automatically
- Slow responses fall back to smart templates
- Consider using `phi3` model for faster responses

#### **"Import errors"**
```bash
# Check Python version
python3.11 --version

# Install missing packages
pip3.11 install PyYAML pandas requests openpyxl
```

##  Project Structure

```
fresh_start_email_generator/
├── responsive_email_generator.py    # Main application
├── yaml_config_manager.py       # Configuration management
├── config.yaml                  # Your credentials (add to .gitignore)
├── .gitignore                   # Git ignore file          
├── README.md                    # This file
```

##  Performance Optimization

### ** (Recommended Settings)**
```yaml
# config.yaml
ollama:
  url: # Your localhost url for ollama
  model: "mistral"    # Good quality
  timeout: 180        # Extended timeout
```

### **Generation Modes**
- **AI Only**: Pure AI generation (slow but highly customized)
- **Hybrid**: AI with smart fallbacks (recommended - best balance)
- **Templates**: Instant generation using smart templates

##  Security & Privacy

- **Local Processing**: All AI generation happens locally via Ollama
- **Secure Storage**: Passwords stored in local config.yaml file
- **No External APIs**: No data sent to external services (except Gmail SMTP)
- **Privacy First**: Your prospect data never leaves your computer

##  Results Tracking

The application tracks:
- **Generation Method**: AI Fast/Slow vs Template fallback
- **Generation Time**: How long each email took to create
- **Send Status**: Which emails were sent successfully
- **Email History**: Complete audit trail of all communications

## Contributing

This is a custom business application for Fresh Start Cleaning Co. For modifications or enhancements, contact the development team.

##  Support

For technical support or business inquiries:

**Fresh Start Cleaning Co.**
-  Website: https://freshcleaningcolouisiana.com/
-  Location: Louisiana
-  Email: [Contact information]

##  License

Copyright (c) 2025 Fresh Start Cleaning Co. All rights reserved.

This software is proprietary and confidential. Unauthorized copying, distribution, or use is strictly prohibited.

---

**Built with ❤️ for Fresh Start Cleaning Co. - Louisiana's premier cleaning service provider** 
