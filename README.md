# Microsoft Automation Emailing

A Python application to send emails one by one using Microsoft Graph API with attachment support. This tool automates email sending through Microsoft's Graph API, perfect for marketing campaigns, notifications, and bulk email operations.

## Features

- ✅ Sends emails one by one (not bulk) to avoid rate limiting
- ✅ Supports HTML email templates
- ✅ PDF attachment support
- ✅ Configurable delay between emails
- ✅ Detailed sending results and error reporting
- ✅ Uses Microsoft Graph API with Application permissions (no user interaction required)

## Prerequisites

- Python 3.7 or higher
- Microsoft Azure account with Entra ID (Azure AD)
- App registration configured with Mail.Send (Application) permission
- Client ID, Client Secret, and Tenant ID from Azure Portal

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Environment Variables

Copy the example environment file:

```bash
cp .env.example .env
```

Edit `.env` and add your Microsoft Graph API credentials:

```
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-client-id-here
CLIENT_SECRET=your-client-secret-here
SENDER_EMAIL=your-email@yourdomain.com
EMAIL_SUBJECT=Your Email Subject
EMAIL_BODY_HTML=<html>Your HTML content here</html>
EMAIL_BODY_TEXT=Your plain text content here
```

**Where to find these values:**

- **TENANT_ID**: Azure Portal > Entra ID > Overview > Tenant ID
- **CLIENT_ID**: Azure Portal > Entra ID > App registrations > Your App > Overview > Application (client) ID
- **CLIENT_SECRET**: Azure Portal > Entra ID > App registrations > Your App > Certificates & secrets > Create a new secret
- **SENDER_EMAIL**: The email address you want to send from (must be in your tenant)
- **EMAIL_SUBJECT**: (Optional) Email subject line
- **EMAIL_BODY_HTML**: (Optional) HTML email body
- **EMAIL_BODY_TEXT**: (Optional) Plain text email body

### 3. Create Recipients File

Create a file named `recipients.txt` with one email address per line:

```
client1@example.com
client2@example.com
client3@example.com
```

### 4. (Optional) Add PDF Attachment

If you want to attach a PDF file, place it in the same directory and update the `attachment_path` variable in `email_sender.py` (default: `attachment.pdf`).

## Usage

Run the email sender:

```bash
python3 email_sender.py
```

The script will:
1. Load recipients from `recipients.txt`
2. Authenticate with Microsoft Graph API
3. Send emails one by one with a 1-second delay between each
4. Display progress and results
5. Save detailed results to `send_results.json`

## Configuration Options

You can modify these settings:

- **Delay between emails**: Change `delay_seconds` parameter in `send_emails_one_by_one()` (default: 5.0 seconds)
- **Email subject**: Set `EMAIL_SUBJECT` in `.env` or modify in `email_sender.py`
- **Email body**: Set `EMAIL_BODY_HTML` and `EMAIL_BODY_TEXT` in `.env` or modify in `email_sender.py`
- **Attachment**: Update `attachment_path` variable in `main()` function

## Output

The script provides:
- Real-time progress updates
- Success/failure status for each email
- Summary statistics
- Detailed results saved to `send_results.json`

## Troubleshooting

### Authentication Errors

- Verify your `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET` are correct
- Ensure the client secret hasn't expired
- Check that the app registration has `Mail.Send` (Application) permission and admin consent

### Email Sending Errors

- Verify `SENDER_EMAIL` matches an email address in your tenant
- Check that the app registration has admin consent for Mail.Send permission
- Ensure the recipient email addresses are valid

### Rate Limiting

If you encounter rate limiting errors, increase the `delay_seconds` value in the script.

## Microsoft Graph API Endpoint

The script uses:
```
POST https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail
```

This endpoint requires:
- **Permission**: `Mail.Send` (Application)
- **Admin Consent**: Required

## Security Notes

- Never commit `.env` file to version control
- Keep your `CLIENT_SECRET` secure
- Rotate client secrets regularly
- Use environment variables or secure secret management in production

## License

This script is provided as-is for sending emails via Microsoft Graph API.

