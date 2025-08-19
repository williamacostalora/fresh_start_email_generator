import yaml
import os
from typing import Dict, Any, Optional

class YAMLConfigManager:
    def __init__(self, config_file: str = "config.yaml"):
        self.config_file = config_file
        self.config = self.load_config()

    def load_config(self) -> Dict[str, Any]:
        """Load configuration from YAML file"""
        if not os.path.exists(self.config_file):
            raise FileNotFoundError(
                f"Configuration file '{self.config_file}' not found. "
                f"Please create it using the config.yaml template."
            )
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                if not config:
                    raise ValueError("Configuration file is empty or invalid")
                return config
        except yaml.YAMLError as e:
            raise ValueError(f"Error parsing YAML file: {e}")
        except Exception as e:
            raise Exception(f"Error loading configuration: {e}")

    def reload_config(self):
        """Reload configuration from file"""
        self.config = self.load_config()

    def get(self, section: str, key: Optional[str] = None) -> Any:
        """Get configuration value"""
        if section not in self.config:
            return {} if key is None else None
        
        if key is None:
            return self.config[section]
        
        return self.config[section].get(key)

    def get_email_config(self) -> Dict[str, Any]:
        """Get email configuration"""
        return self.get('email') or {}

    def get_company_info(self) -> Dict[str, Any]:
        """Get company information"""
        return self.get('company') or {}

    def get_ollama_config(self) -> Dict[str, Any]:
        """Get Ollama configuration"""
        return self.get('ollama') or {}

    def is_email_configured(self) -> bool:
        """Check if email is properly configured"""
        email_config = self.get_email_config()
        
        email = email_config.get('from_email', '')
        password = email_config.get('from_password', '')
        
        # Check if values are not default placeholders and are present
        return (
            bool(email and password) and 
            email != "your_email@gmail.com" and 
            password != "your_16_character_app_password" and
            "@" in email
        )

    def validate_config(self) -> tuple[bool, str]:
        """Validate configuration and return (is_valid, error_message)"""
        # Check if config was loaded
        if not self.config:
            return False, "Configuration is empty"

        # Check required sections
        required_sections = ['email', 'company', 'ollama']
        for section in required_sections:
            if section not in self.config:
                return False, f"Missing required section: {section}"

        # Check email configuration
        email_config = self.get_email_config()
        required_email_fields = ['from_email', 'from_password', 'smtp_server', 'smtp_port']
        for field in required_email_fields:
            if field not in email_config:
                return False, f"Missing email field: {field}"

        # Check if email is configured (not placeholder)
        if not self.is_email_configured():
            return False, "Email credentials not configured (still using placeholder values)"

        # Check company configuration
        company_config = self.get_company_info()
        required_company_fields = ['name', 'website', 'phone']
        for field in required_company_fields:
            if field not in company_config:
                return False, f"Missing company field: {field}"

        # Check Ollama configuration
        ollama_config = self.get_ollama_config()
        required_ollama_fields = ['url', 'model']
        for field in required_ollama_fields:
            if field not in ollama_config:
                return False, f"Missing Ollama field: {field}"

        return True, "Configuration is valid"

    def get_config_status(self) -> str:
        """Get human-readable configuration status"""
        is_valid, message = self.validate_config()
        
        if is_valid:
            return "✅ Configuration loaded and valid"
        else:
            return f"❌ Configuration error: {message}"

    def print_config_summary(self):
        """Print a summary of the current configuration (without sensitive data)"""
        print("\n📋 Configuration Summary:")
        print("=" * 40)
        
        # Email status
        email_config = self.get_email_config()
        if email_config:
            email = email_config.get('from_email', 'Not set')
            # Mask email for security
            masked_email = email[:3] + "*" * (len(email) - 6) + email[-3:] if len(email) > 6 else "***"
            print(f"📧 Email: {masked_email}")
            print(f"📧 SMTP: {email_config.get('smtp_server', 'Not set')}")
        
        # Company info
        company_config = self.get_company_info()
        if company_config:
            print(f"🏢 Company: {company_config.get('name', 'Not set')}")
            print(f"🌐 Website: {company_config.get('website', 'Not set')}")
            print(f"📱 Phone: {company_config.get('phone', 'Not set')}")
        
        # Ollama info
        ollama_config = self.get_ollama_config()
        if ollama_config:
            print(f"🤖 AI Model: {ollama_config.get('model', 'Not set')}")
            print(f"🔗 Ollama URL: {ollama_config.get('url', 'Not set')}")
        
        print("=" * 40)
        print(self.get_config_status())
        print()

# Example usage and testing
if __name__ == "__main__":
    try:
        config = YAMLConfigManager()
        config.print_config_summary()
        
        # Test email configuration
        if config.is_email_configured():
            print("✅ Email is configured and ready")
        else:
            print("❌ Email needs to be configured in config.yaml")
            
    except FileNotFoundError as e:
        print(f"❌ {e}")
        print("\n📝 Please create a config.yaml file using the provided template")
    except Exception as e:
        print(f"❌ Error: {e}")