# Microsoft Automation Emailing

<div align="center">

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)

**Professional Email Automation Tool for Business & Educational Use**

Automate your business email campaigns with Microsoft Graph API. Perfect for educational institutions, businesses, and organizations that need to send personalized emails with attachments and company signatures.

[Features](#-features) • [Quick Start](#-quick-start) • [Documentation](#-documentation) • [License](#-license)

</div>

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Use Cases](#-use-cases)
- [Prerequisites](#-prerequisites)
- [Quick Start](#-quick-start)
- [Detailed Setup Guide](#-detailed-setup-guide)
- [Configuration](#-configuration)
- [Usage Examples](#-usage-examples)
- [Company Signature Setup](#-company-signature-setup)
- [PDF Attachment Guide](#-pdf-attachment-guide)
- [Troubleshooting](#-troubleshooting)
- [Security & Best Practices](#-security--best-practices)
- [API Reference](#-api-reference)
- [License](#-license)
- [Contributing](#-contributing)

---

## 🎯 Overview

**Microsoft Automation Emailing** is a powerful Python application designed for businesses and educational institutions to automate email sending through Microsoft Graph API. This tool enables you to:

- Send personalized emails to multiple recipients
- Attach PDF documents (invoices, reports, certificates, etc.)
- Include professional company signatures with logos
- Automate business communications and educational notifications
- Track email delivery with detailed reporting

### Perfect For:
- **Businesses**: Marketing campaigns, invoice distribution, report sharing
- **Educational Institutions**: Course notifications, certificate distribution, announcements
- **Organizations**: Newsletter distribution, event invitations, document sharing

---

## ✨ Features

### Core Capabilities
- ✅ **One-by-One Sending**: Prevents rate limiting and ensures reliable delivery
- ✅ **HTML Email Templates**: Create beautiful, professional email designs
- ✅ **PDF Attachment Support**: Attach documents, invoices, certificates, reports
- ✅ **Company Signature**: Automatically embed company logos and signatures
- ✅ **Email Validation**: Built-in validation using Disify API (removes invalid/disposable emails)
- ✅ **Progress Tracking**: Resume interrupted sending sessions
- ✅ **Detailed Reporting**: Comprehensive success/failure reports
- ✅ **Error Handling**: Automatic retry logic for throttled requests
- ✅ **No User Interaction**: Fully automated using Application permissions

### Advanced Features
- 🔄 **Automatic Token Refresh**: Handles Microsoft Graph API token expiration
- 📊 **Real-time Progress**: Live updates during email sending
- 🛡️ **Duplicate Prevention**: Tracks sent emails to prevent duplicates
- ⚡ **Rate Limit Handling**: Intelligent backoff for throttled requests
- 📝 **Logging**: Comprehensive logging for debugging

---

## 🎓 Use Cases

### Business Applications
- **Invoice Distribution**: Automatically send invoices to clients
- **Marketing Campaigns**: Send promotional emails to customer lists
- **Report Sharing**: Distribute monthly/quarterly reports to stakeholders
- **Document Delivery**: Send contracts, proposals, and legal documents

### Educational Applications
- **Certificate Distribution**: Send course completion certificates to students
- **Course Notifications**: Notify students about new courses, assignments, or deadlines
- **Announcements**: Distribute institutional announcements to faculty/students
- **Grade Reports**: Send grade reports to students and parents

---

## 📦 Prerequisites

Before you begin, ensure you have:

1. **Python 3.7 or higher** installed
   ```bash
   python3 --version
   ```

2. **Microsoft Azure Account** with Entra ID (Azure AD)
   - Free tier available at [azure.microsoft.com](https://azure.microsoft.com)

3. **Microsoft 365 Account** (for sending emails)
   - Can be Business, Education, or Enterprise plan

4. **App Registration** in Azure Portal with:
   - `Mail.Send` (Application) permission
   - Admin consent granted

---

## 🚀 Quick Start

### Step 1: Clone or Download
```bash
git clone https://github.com/nabilW/Email-Sender.git
cd Email-Sender
```

### Step 2: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 3: Configure Environment
```bash
cp .env.example .env
# Edit .env with your credentials
```

### Step 4: Add Recipients
```bash
# Create recipients.txt with one email per line
echo "client1@example.com" > recipients.txt
echo "client2@example.com" >> recipients.txt
```

### Step 5: Run
```bash
python3 email_sender.py
```

**That's it!** Your emails will start sending automatically.

---

## 📖 Detailed Setup Guide

### 1. Azure App Registration Setup

#### Create App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Name: `Email Automation App` (or your preferred name)
5. Supported account types: **Single tenant**
6. Click **Register**

#### Configure API Permissions
1. In your app, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Search for `Mail.Send`
6. Select **Mail.Send** and click **Add permissions**
7. **Important**: Click **Grant admin consent** (requires admin privileges)

#### Create Client Secret
1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Description: `Email Automation Secret`
4. Expires: Choose appropriate duration (6 months, 12 months, or never)
5. Click **Add**
6. **Copy the secret value immediately** (you won't see it again!)

#### Get Your Credentials
1. **Tenant ID**: Go to **Overview** > Copy **Tenant ID**
2. **Client ID**: Go to **Overview** > Copy **Application (client) ID**
3. **Client Secret**: The value you copied from step above

### 2. Environment Configuration

Edit your `.env` file:

```env
# Microsoft Graph API Credentials (Required)
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-client-id-here
CLIENT_SECRET=your-client-secret-here
SENDER_EMAIL=your-email@yourdomain.com

# Email Template (Optional - can also be set in email_sender.py)
EMAIL_SUBJECT=Your Email Subject Here
EMAIL_BODY_HTML=<html><body><p>Your HTML email content here</p></body></html>
EMAIL_BODY_TEXT=Your plain text email content here
```

**Where to find these values:**
- **TENANT_ID**: Azure Portal > Entra ID > Overview > Tenant ID
- **CLIENT_ID**: Azure Portal > App registrations > Your App > Overview > Application (client) ID
- **CLIENT_SECRET**: The secret value you created (from Certificates & secrets)
- **SENDER_EMAIL**: Your Microsoft 365 email address (must be in your tenant)

### 3. Create Recipients File

Create `recipients.txt` with one email address per line:

```
client1@example.com
client2@example.com
client3@example.com
```

**Tips:**
- Empty lines are ignored
- Comments starting with `#` are ignored
- The script automatically validates and filters invalid emails

### 4. (Optional) Add PDF Attachment

1. Place your PDF file in the project directory
2. Update `attachment_path` in `email_sender.py` (line ~850):
   ```python
   attachment_path = Path("your-document.pdf")
   ```

**Supported formats:** PDF only (`.pdf`)

---

## ⚙️ Configuration

### Email Template Configuration

You can configure email templates in two ways:

#### Method 1: Environment Variables (.env file)
```env
EMAIL_SUBJECT=Welcome to Our Service
EMAIL_BODY_HTML=<html><body><h1>Welcome!</h1><p>Thank you for joining us.</p></body></html>
EMAIL_BODY_TEXT=Welcome! Thank you for joining us.
```

#### Method 2: Direct Code Modification
Edit `email_sender.py` (lines 48-88) to customize templates directly.

### Delay Between Emails

Default: 5 seconds (to avoid rate limiting)

To change, modify `delay_seconds` in `send_emails_one_by_one()` function:
```python
delay_seconds=10.0  # Wait 10 seconds between emails
```

### Email Validation

The script automatically validates emails using:
- Format validation (RFC 5322 compliant)
- Disify API (checks for disposable emails, invalid DNS)

To disable Disify validation:
```python
recipients, stats = load_recipients_from_file(recipients_file, use_disify=False)
```

---

## 💼 Company Signature Setup

### Adding Your Company Logo

1. **Place logo file** in the project directory with one of these names:
   - `logo.png`
   - `logo.jpg` or `logo.jpeg`
   - `logo_black.png`
   - `logo.gif`

2. **Supported formats:**
   - PNG (automatically converted to JPEG for better compatibility)
   - JPEG/JPG
   - GIF

3. **Logo will automatically be embedded** in your email signature

### Customizing Email Signature

Edit the email template in `.env` or `email_sender.py` to include your signature:

```html
<div class="signature">
    <img src="PLACEHOLDER_LOGO_URL" alt="Company Logo" />
    <p><strong>Your Name</strong></p>
    <p>Your Title</p>
    <p>Company Name</p>
    <p>Email: your@email.com</p>
    <p>Phone: +1 (555) 123-4567</p>
</div>
```

The script automatically replaces `PLACEHOLDER_LOGO_URL` with your embedded logo.

### Using Logo from URL

If you prefer to host your logo online:

1. Upload logo to your website/CDN
2. In `email_sender.py`, set:
   ```python
   logo_url = "https://yourdomain.com/logo.png"
   ```

---

## 📎 PDF Attachment Guide

### Adding PDF Attachments

1. **Place PDF file** in the project directory
2. **Update attachment path** in `email_sender.py`:
   ```python
   attachment_path = Path("invoice.pdf")  # Your PDF filename
   ```

### Common Use Cases

- **Invoices**: `invoice.pdf`, `invoice-2024.pdf`
- **Certificates**: `certificate.pdf`, `completion-certificate.pdf`
- **Reports**: `monthly-report.pdf`, `annual-report.pdf`
- **Documents**: `contract.pdf`, `proposal.pdf`

### Multiple Attachments

Currently, the script supports **one PDF attachment per email**. For multiple attachments, you can:
- Combine PDFs into a single file
- Modify the code to support multiple attachments (see [Contributing](#contributing))

---

## 📝 Usage Examples

### Example 1: Basic Email Sending

```bash
# 1. Configure .env with your credentials
# 2. Create recipients.txt with email addresses
# 3. Run the script
python3 email_sender.py
```

### Example 2: With PDF Attachment

```python
# In email_sender.py, set:
attachment_path = Path("monthly-report.pdf")
```

### Example 3: With Company Logo

```bash
# Place logo.png in project directory
# Script automatically detects and embeds it
python3 email_sender.py
```

### Example 4: Custom Email Template

```env
# In .env file:
EMAIL_SUBJECT=Important Announcement
EMAIL_BODY_HTML=<html><body><h1>Important</h1><p>This is your message.</p></body></html>
EMAIL_BODY_TEXT=Important: This is your message.
```

---

## 🔧 Troubleshooting

### Authentication Errors

**Error**: `Failed to acquire token`

**Solutions:**
- Verify `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET` are correct
- Check that client secret hasn't expired
- Ensure app registration has `Mail.Send` (Application) permission
- Verify admin consent is granted

### Email Sending Errors

**Error**: `401 Unauthorized` or `403 Forbidden`

**Solutions:**
- Verify `SENDER_EMAIL` matches an email in your tenant
- Check that app registration has admin consent for Mail.Send
- Ensure the email account exists and is active

### Rate Limiting

**Error**: `ApplicationThrottled` or `429 Too Many Requests`

**Solutions:**
- Increase delay between emails (default: 5 seconds)
- The script automatically retries with exponential backoff
- Consider sending in smaller batches

### Attachment Issues

**Error**: `Attachment file not found`

**Solutions:**
- Verify PDF file exists in project directory
- Check file path in `email_sender.py`
- Ensure file has `.pdf` extension
- Check file permissions

### Logo Not Appearing

**Solutions:**
- Verify logo file exists with correct name (`logo.png`, `logo.jpg`, etc.)
- Check file format (PNG, JPG, or GIF)
- Ensure logo file is in project root directory
- Check file size (very large files may cause issues)

---

## 🔒 Security & Best Practices

### Security Recommendations

1. **Never commit `.env` file**
   - Already excluded in `.gitignore`
   - Contains sensitive credentials

2. **Rotate client secrets regularly**
   - Set expiration dates
   - Create new secrets before old ones expire

3. **Use environment variables in production**
   - Don't hardcode credentials
   - Use secure secret management (Azure Key Vault, AWS Secrets Manager)

4. **Limit permissions**
   - Only grant `Mail.Send` permission (not full mailbox access)
   - Use application permissions (not delegated)

5. **Monitor usage**
   - Review `send_results.json` regularly
   - Check Azure Portal for unusual activity

### Best Practices

- **Test with small batches first**: Send to 5-10 recipients before large campaigns
- **Validate email lists**: Use the built-in validation to remove invalid emails
- **Monitor rate limits**: Adjust delays if you encounter throttling
- **Keep logs**: Review `email_send_log.txt` for debugging
- **Backup progress**: The script saves progress automatically

---

## 📚 API Reference

### Microsoft Graph API

**Endpoint:**
```
POST https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail
```

**Required Permission:**
- `Mail.Send` (Application)
- Admin consent required

**Resource App ID:**
```
00000003-0000-0000-c000-000000000000
```

**Permission ID:**
```
b633e1c5-b582-4048-a93e-9f11b44c7e96
```

### Main Functions

#### `GraphEmailSender`
Main class for sending emails via Microsoft Graph API.

**Methods:**
- `get_access_token()`: Acquires and refreshes access tokens
- `send_email()`: Sends a single email with retry logic
- `send_emails_one_by_one()`: Sends emails to multiple recipients

#### `load_recipients_from_file()`
Loads and validates email addresses from a text file.

**Parameters:**
- `file_path`: Path to recipients file
- `use_disify`: Enable/disable Disify API validation (default: True)

---

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

### Educational Use

This software is provided for **educational and business purposes**. Users are responsible for:

- Complying with email regulations (CAN-SPAM Act, GDPR, etc.)
- Obtaining proper consent from recipients
- Following Microsoft's Terms of Service
- Respecting recipient privacy and preferences

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### How to Contribute

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Areas for Contribution

- Support for multiple attachments
- Email template builder
- GUI interface
- Additional email validation services
- Enhanced error handling
- Performance optimizations

---

## 📞 Support

### Getting Help

- **Issues**: Open an issue on [GitHub](https://github.com/nabilW/Email-Sender/issues)
- **Documentation**: Check this README and code comments
- **Microsoft Graph API**: [Official Documentation](https://docs.microsoft.com/graph/api/resources/mail-api-overview)

### Common Questions

**Q: Can I send to Gmail/Yahoo/other email providers?**  
A: Yes! The script works with any email address, not just Microsoft accounts.

**Q: How many emails can I send per day?**  
A: Depends on your Microsoft 365 plan. Check your plan's limits in Azure Portal.

**Q: Can I schedule emails?**  
A: Use a task scheduler (cron, Windows Task Scheduler) to run the script at specific times.

**Q: Is this free?**  
A: The script is free and open source. Microsoft 365 subscription required for sending emails.

---

## 🎓 Educational Certificate

This project is designed for **educational and business automation purposes**. It demonstrates:

- Microsoft Graph API integration
- Email automation best practices
- Python application development
- API authentication and security
- Error handling and retry logic

**Perfect for:**
- Learning Microsoft Graph API
- Understanding email automation
- Business process automation
- Educational institution communications

---

## 📊 Project Status

**Current Version:** 1.0.0  
**Status:** ✅ Production Ready  
**Maintenance:** Active

---

<div align="center">

**Made with ❤️ for businesses and educational institutions**

[⭐ Star this repo](https://github.com/nabilW/Email-Sender) • [🐛 Report Bug](https://github.com/nabilW/Email-Sender/issues) • [💡 Request Feature](https://github.com/nabilW/Email-Sender/issues)

</div>
