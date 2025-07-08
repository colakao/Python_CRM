# Email Campaign Script Setup Guide

## System Requirements

### Minimum System Requirements
- **Operating System**: Windows 10/11, macOS 10.12+, or Linux (Ubuntu 18.04+)
- **RAM**: 4GB minimum, 8GB recommended
- **Storage**: 1GB free space
- **Network**: Stable internet connection for SMTP/IMAP operations

### Python Version
- **Python 3.8 or higher** (recommended: Python 3.10+)

## Step-by-Step Installation Guide

### Step 1: Install Python

#### Windows:
1. Go to [python.org](https://www.python.org/downloads/)
2. Download Python 3.10+ installer
3. **Important**: Check "Add Python to PATH" during installation
4. Run the installer as administrator
5. Verify installation:
   ```cmd
   python --version
   pip --version
   ```

#### macOS:
1. Option A - Using Homebrew (recommended):
   ```bash
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   brew install python
   ```
2. Option B - Download from python.org (same as Windows)

#### Linux (Ubuntu/Debian):
```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

### Step 2: Create Virtual Environment (Recommended)

```bash
# Create virtual environment
python -m venv email_campaign_env

# Activate it
# Windows:
email_campaign_env\Scripts\activate
# macOS/Linux:
source email_campaign_env/bin/activate
```

### Step 3: Install Required Libraries

#### Core Python Libraries (via pip):
```bash
pip install pandas>=1.5.0
pip install openpyxl>=3.0.0
pip install xlrd>=2.0.0
```

#### GUI and HTML Libraries:
```bash
pip install tkinterhtml
```

**Note**: `tkinter` comes pre-installed with Python, but if you encounter issues:
- **Linux**: `sudo apt-get install python3-tk`
- **macOS**: Usually included with Python installation

### Step 4: Verify Installation

Create a test script to verify all imports:

```python
# test_imports.py
try:
    import pandas as pd
    print("✓ pandas imported successfully")
    
    import smtplib
    print("✓ smtplib imported successfully")
    
    import imaplib
    print("✓ imaplib imported successfully")
    
    import tkinter as tk
    print("✓ tkinter imported successfully")
    
    from tkinterhtml import HtmlFrame
    print("✓ tkinterhtml imported successfully")
    
    import openpyxl
    print("✓ openpyxl imported successfully")
    
    print("\n✅ All required libraries are installed!")
    
except ImportError as e:
    print(f"❌ Import error: {e}")
```

Run: `python test_imports.py`

## Complete Requirements List

### Standard Library Modules (Built-in):
- `smtplib` - SMTP email sending
- `imaplib` - IMAP email reading
- `email` - Email message handling
- `ssl` - SSL/TLS security
- `tkinter` - GUI framework
- `logging` - Logging functionality
- `os` - Operating system interface
- `re` - Regular expressions
- `time` - Time operations
- `datetime` - Date/time handling
- `getpass` - Password input
- `traceback` - Error tracing
- `base64` - Base64 encoding
- `webbrowser` - Web browser operations
- `tempfile` - Temporary file operations
- `mailbox` - Mailbox file handling

### Third-Party Libraries:
- `pandas` - Data manipulation (Excel files)
- `openpyxl` - Excel file operations (.xlsx)
- `xlrd` - Excel file reading (.xls)
- `tkinterhtml` - HTML display in tkinter

## File Structure Setup

Create this folder structure:
```
email_campaign/
├── execute.py                 # Your main script
├── templates/
│   └── email_template.html   # HTML email template
├── data/
│   └── contacts.xlsx         # Contact list
├── logs/
│   └── (log files will be created here)
└── reports/
    └── (failure reports will be saved here)
```

## Pre-run Checklist

### Required Files:
1. **Excel file** with columns:
   - `Email Contacto` (required)
   - `Nombre Contacto` (required)
   - `Nombre Empresa` (optional)

2. **HTML template file** for email content

### SMTP Configuration:
- SMTP server address
- SMTP port (usually 465 for SSL, 587 for TLS)
- Email credentials (username/password)

### Permissions:
- **Windows**: Run as administrator if needed
- **macOS/Linux**: Ensure file permissions for credential storage

## Troubleshooting Common Issues

### tkinterhtml Installation Issues:
If `pip install tkinterhtml` fails:

**Windows**:
```cmd
pip install --upgrade pip setuptools wheel
pip install tkinterhtml
```

**macOS**:
```bash
xcode-select --install
pip install tkinterhtml
```

**Linux**:
```bash
sudo apt-get install python3-dev python3-tk
pip install tkinterhtml
```

### Alternative if tkinterhtml fails:
The script will show a warning and disable HTML preview, but core functionality will work.

### Excel File Issues:
- Ensure Excel files are not open in Excel when running the script
- Use UTF-8 encoding for special characters
- Verify column names match exactly

### SMTP Connection Issues:
- Check firewall settings
- Verify SMTP server settings
- Ensure "Less secure app access" is enabled (if using Gmail)
- Use app-specific passwords for Gmail/Outlook

## Security Notes

1. **Credential Storage**: The script stores credentials in a hidden `.creds` file with base64 encoding
2. **File Permissions**: On Unix systems, credential files get 600 permissions (owner read/write only)
3. **SSL/TLS**: Uses SSL context for secure SMTP connections
4. **Test Mode**: Always test with a small group first

## Running the Script

1. Place your script in the project folder
2. Activate virtual environment (if using)
3. Run: `python execute.py`
4. Use the GUI to configure and run campaigns

## Additional Recommendations

### For Production Use:
- Use environment variables for sensitive data
- Implement proper logging rotation
- Add email validation
- Consider rate limiting for large campaigns
- Use professional SMTP services (SendGrid, Mailgun, etc.)

### For Development:
- Use test email addresses
- Enable detailed logging
- Keep backups of contact lists
- Test HTML templates across email clients