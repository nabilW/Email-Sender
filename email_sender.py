#!/usr/bin/env python3
"""
Microsoft Automation Emailing - Email Sender using Microsoft Graph API
Sends emails one by one with attachment support

Microsoft Graph API Permission Details:
- Resource App ID: 00000003-0000-0000-c000-000000000000 (Microsoft Graph)
- Permission ID: b633e1c5-b582-4048-a93e-9f11b44c7e96 (Mail.Send)
- Required Permission: Mail.Send (Application) - Admin consent required
- Endpoint: POST https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail
"""

import os
import json
import base64
import time
import io
import re
import random
from urllib.parse import quote
from typing import List, Optional, Tuple
from pathlib import Path
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Load environment variables from .env file
load_dotenv()

# Configuration
TENANT_ID = os.getenv("TENANT_ID", "")
CLIENT_ID = os.getenv("CLIENT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# Microsoft Graph API Permission Details
# Resource: https://graph.microsoft.com/Mail.Send
GRAPH_RESOURCE_APP_ID = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
MAIL_SEND_PERMISSION_ID = "b633e1c5-b582-4048-a93e-9f11b44c7e96"  # Mail.Send

# Email template - Customize these for your needs
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "Your Email Subject Here")

EMAIL_BODY_HTML = os.getenv("EMAIL_BODY_HTML", """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header { background-color: #1a5490; color: white; padding: 20px; text-align: center; }
        .content { padding: 20px; background-color: #f9f9f9; }
        .footer { background-color: #333; color: white; padding: 20px; text-align: center; font-size: 12px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Your Company Name</h1>
        </div>
        <div class="content">
            <p>Dear Recipient,</p>
            <p>This is a sample email template. Please customize the HTML content in your .env file or directly in the code.</p>
            <p>Best regards,<br>Your Team</p>
        </div>
        <div class="footer">
            <p>&copy; 2025 Your Company. All rights reserved.</p>
        </div>
    </div>
</body>
</html>
""")

EMAIL_BODY_TEXT = os.getenv("EMAIL_BODY_TEXT", """
Dear Recipient,

This is a sample email template. Please customize the text content in your .env file or directly in the code.

Best regards,
Your Team
""")


class GraphEmailSender:
    """
    Microsoft Graph API Email Sender
    
    Uses Microsoft Graph API with Mail.Send permission:
    - Resource: https://graph.microsoft.com/Mail.Send
    - Resource App ID: 00000003-0000-0000-c000-000000000000 (Microsoft Graph)
    - Permission ID: b633e1c5-b582-4048-a93e-9f11b44c7e96 (Mail.Send)
    - Required: Mail.Send (Application) permission with admin consent
    
    The app registration in Azure Portal must have this permission configured.
    """
    
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        # Using .default scope requests all application permissions granted to the app
        # This includes Mail.Send (Permission ID: b633e1c5-b582-4048-a93e-9f11b44c7e96)
        self.scope = ["https://graph.microsoft.com/.default"]
        self.app = ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=self.authority
        )
        self.access_token = None
        self.token_expires_at = None  # Track token expiration time
    
    def get_access_token(self, force_refresh: bool = False) -> str:
        """
        Get access token using client credentials flow with automatic refresh
        
        This method authenticates with Microsoft Graph API using the Mail.Send permission.
        The app registration must have Mail.Send (Application) permission configured:
        - Resource App ID: 00000003-0000-0000-c000-000000000000
        - Permission ID: b633e1c5-b582-4048-a93e-9f11b44c7e96
        
        Args:
            force_refresh: If True, force token refresh even if token exists
        """
        import time as time_module
        
        # Check if token exists and is still valid (refresh 5 minutes before expiration)
        if not force_refresh and self.access_token and self.token_expires_at:
            current_time = time_module.time()
            # Refresh if token expires in less than 5 minutes
            if current_time < (self.token_expires_at - 300):
                return self.access_token
        
        # Acquire new token
        result = self.app.acquire_token_for_client(scopes=self.scope)
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            # Calculate expiration time (tokens typically expire in 1 hour)
            expires_in = result.get("expires_in", 3600)  # Default to 1 hour if not provided
            self.token_expires_at = time_module.time() + expires_in
            return self.access_token
        else:
            error_msg = result.get("error_description", result.get("error", "Unknown error"))
            raise Exception(f"Failed to acquire token: {error_msg}")
    
    def get_logo_base64(self, logo_path: Optional[Path] = None, logo_url: Optional[str] = None) -> Optional[str]:
        """Get logo as base64 encoded string for embedding in email"""
        # Try to download from URL first
        if logo_url:
            try:
                response = requests.get(logo_url, timeout=10)
                if response.status_code == 200:
                    logo_content = response.content
                    logo_base64 = base64.b64encode(logo_content).decode("utf-8")
                    # Determine content type from URL or response headers
                    content_type = response.headers.get('Content-Type', 'image/gif')
                    if 'gif' in logo_url.lower() or 'gif' in content_type.lower():
                        content_type = "image/gif"
                    elif 'png' in logo_url.lower() or 'png' in content_type.lower():
                        content_type = "image/png"
                    elif 'jpg' in logo_url.lower() or 'jpeg' in logo_url.lower() or 'jpeg' in content_type.lower():
                        content_type = "image/jpeg"
                    else:
                        content_type = "image/gif"
                    return f"data:{content_type};base64,{logo_base64}"
            except Exception as e:
                print(f"Warning: Could not download logo from URL: {e}")
        
        # Try local file if URL didn't work
        if logo_path is None:
            # Try common logo file names
            possible_names = ["logo_black.png", "logo.png", "logo.jpg", "logo.jpeg", "logo.gif"]
            for name in possible_names:
                test_path = Path(name)
                if test_path.exists():
                    logo_path = test_path
                    break
        
        if logo_path and logo_path.exists():
            # Convert PNG to JPEG for better email client compatibility
            if logo_path.suffix.lower() == '.png' and PIL_AVAILABLE:
                try:
                    # Open PNG and convert to JPEG
                    img = Image.open(logo_path)
                    # Convert RGBA to RGB if necessary (remove alpha channel)
                    if img.mode in ('RGBA', 'LA', 'P'):
                        # Create white background
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'P':
                            img = img.convert('RGBA')
                        background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                        img = background
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                    
                    # Save to bytes as JPEG
                    jpeg_buffer = io.BytesIO()
                    img.save(jpeg_buffer, format='JPEG', quality=95, optimize=True)
                    logo_content = jpeg_buffer.getvalue()
                    content_type = "image/jpeg"
                    logo_base64 = base64.b64encode(logo_content).decode("utf-8")
                    return f"data:{content_type};base64,{logo_base64}"
                except Exception as e:
                    print(f"Warning: Could not convert PNG to JPEG: {e}. Using PNG instead.")
            
            # Use original file format if conversion failed or not PNG
            with open(logo_path, "rb") as f:
                logo_content = f.read()
                logo_base64 = base64.b64encode(logo_content).decode("utf-8")
                # Determine content type
                if logo_path.suffix.lower() in ['.jpg', '.jpeg']:
                    content_type = "image/jpeg"
                elif logo_path.suffix.lower() == '.png':
                    content_type = "image/png"
                elif logo_path.suffix.lower() == '.gif':
                    content_type = "image/gif"
                else:
                    content_type = "image/jpeg"  # Default to JPEG
                return f"data:{content_type};base64,{logo_base64}"
        return None
    
    def encode_attachment(self, file_path: Path) -> dict:
        """Encode file as base64 for attachment"""
        if not file_path.exists():
            raise FileNotFoundError(f"Attachment file not found: {file_path}")
        
        with open(file_path, "rb") as f:
            file_content = f.read()
            file_base64 = base64.b64encode(file_content).decode("utf-8")
        
        return {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": file_path.name,
            "contentType": "application/pdf",
            "contentBytes": file_base64
        }
    
    def send_email(
        self,
        recipient_email: str,
        subject: str,
        body_html: str,
        body_text: str,
        attachment_path: Optional[Path] = None,
        logo_path: Optional[Path] = None,
        logo_url: Optional[str] = None,
        max_retries: int = 3
    ) -> dict:
        """
        Send email to a single recipient with retry logic for throttling
        
        Args:
            recipient_email: Email address of recipient
            subject: Email subject
            body_html: HTML email body
            body_text: Plain text email body
            attachment_path: Optional path to PDF attachment
            logo_path: Optional path to logo file
            logo_url: Optional URL to logo
            max_retries: Maximum number of retry attempts for throttled requests
        """
        # Get access token (will auto-refresh if needed)
        access_token = self.get_access_token()
        
        # Prepare attachments list
        attachments = []
        
        # Add logo - use base64 embedding directly in HTML (most reliable method)
        logo_base64 = self.get_logo_base64(logo_path, logo_url)
        if logo_base64:
            # Extract base64 data and content type
            if logo_base64.startswith("data:"):
                parts = logo_base64.split(",", 1)
                content_type = parts[0].split(":")[1].split(";")[0]
                logo_data = parts[1]
            else:
                # If not in data URI format, determine from file extension
                if logo_path and logo_path.suffix.lower() == '.gif':
                    content_type = "image/gif"
                elif logo_path and logo_path.suffix.lower() in ['.jpg', '.jpeg']:
                    content_type = "image/jpeg"
                else:
                    content_type = "image/jpeg"  # Default to JPEG for converted PNG
                logo_data = logo_base64
            
            # Embed logo directly in HTML using base64 (most compatible method)
            base64_img_tag = f'<img src="data:{content_type};base64,{logo_data}" alt="Company Logo" style="max-width: 250px; height: auto; display: block; margin: 0 auto; border: 0;" />'
            
            # Replace logo placeholder with base64 embedded image (if exists in HTML)
            # You can customize this placeholder in your email template
            body_html = body_html.replace(
                '<img src="PLACEHOLDER_LOGO_URL" alt="Company Logo" />',
                base64_img_tag
            )
        
        # Prepare message with updated HTML
        message = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": body_html
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": recipient_email
                        }
                    }
                ]
            }
        }
        
        # Add PDF attachment if provided
        if attachment_path:
            attachment = self.encode_attachment(attachment_path)
            attachments.append(attachment)
        
        # Add all attachments to message
        if attachments:
            message["message"]["attachments"] = attachments
        
        # Send email using Microsoft Graph API Mail.Send endpoint
        # API: https://graph.microsoft.com/Mail.Send
        # Requires Mail.Send (Application) permission:
        # - Resource App ID: 00000003-0000-0000-c000-000000000000
        # - Permission ID: b633e1c5-b582-4048-a93e-9f11b44c7e96
        url = f"{GRAPH_API_ENDPOINT}/users/{SENDER_EMAIL}/sendMail"
        
        # Retry logic with exponential backoff for throttling
        for attempt in range(max_retries):
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            }
            
            try:
                response = requests.post(url, headers=headers, json=message, timeout=30)
                
                # Success
                if response.status_code == 202:
                    return {
                        "success": True,
                        "recipient": recipient_email,
                        "status_code": response.status_code,
                        "message": "Email sent successfully"
                    }
                
                # Handle token expiration
                elif response.status_code == 401:
                    error_data = response.json() if response.text else {}
                    if "InvalidAuthenticationToken" in str(error_data) or "expired" in str(error_data).lower():
                        # Token expired - refresh token but don't retry (max_retries=1)
                        access_token = self.get_access_token(force_refresh=True)
                        # Return error immediately (no retries)
                        return {
                            "success": False,
                            "recipient": recipient_email,
                            "status_code": response.status_code,
                            "error": "Token expired"
                        }
                
                # Handle throttling (429 or ApplicationThrottled)
                elif response.status_code == 429 or "ApplicationThrottled" in response.text:
                    error_data = response.json() if response.text else {}
                    
                    # Check for Retry-After header (tells us how long to wait)
                    retry_after = response.headers.get("Retry-After")
                    if retry_after:
                        try:
                            wait_seconds = int(retry_after)
                        except ValueError:
                            wait_seconds = None
                    else:
                        wait_seconds = None
                    
                    # If we have retries left, wait and retry
                    if attempt < max_retries - 1:
                        if wait_seconds:
                            # Use the Retry-After value from the header
                            wait_time = wait_seconds
                        else:
                            # Exponential backoff: 60, 120, 240 seconds for IncomingBytes limit
                            wait_time = min(60 * (2 ** attempt), 600)  # Max 10 minutes
                        
                        print(f"\n   ⚠ Throttled (IncomingBytes limit). Waiting {wait_time} seconds before retry...")
                        time.sleep(wait_time)
                        continue  # Retry the request
                    else:
                        # All retries exhausted
                        return {
                            "success": False,
                            "recipient": recipient_email,
                            "status_code": response.status_code,
                            "error": response.text
                        }
                
                # Other errors
                return {
                    "success": False,
                    "recipient": recipient_email,
                    "status_code": response.status_code,
                    "error": response.text
                }
                
            except requests.exceptions.RequestException as e:
                if attempt < max_retries - 1:
                    wait_time = min(2 ** attempt, 30)  # Exponential backoff, max 30 seconds
                    print(f"\n   ⚠ Request error: {str(e)}. Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                    continue
                else:
                    return {
                        "success": False,
                        "recipient": recipient_email,
                        "error": f"Request exception: {str(e)}"
                    }
        
        # All retries exhausted
        return {
            "success": False,
            "recipient": recipient_email,
            "error": f"Failed after {max_retries} attempts"
        }
    
    def send_emails_one_by_one(
        self,
        recipient_list: List[str],
        subject: str,
        body_html: str,
        body_text: str,
        attachment_path: Optional[Path] = None,
        logo_path: Optional[Path] = None,
        logo_url: Optional[str] = None,
        delay_seconds: float = 5.0,  # Default delay between emails (increased to reduce throttling)
        start_index: int = 0,
        progress_file: Optional[Path] = None
    ) -> List[dict]:
        """
        Send emails to multiple recipients ONE BY ONE (not in bulk)
        
        This method sends each email individually with a delay between sends
        to avoid rate limiting and ensure reliable delivery.
        
        Args:
            start_index: Index to start from (0-based). Useful for resuming after interruption.
            progress_file: Path to file that tracks progress (prevents duplicates on restart).
        """
        results = []
        
        # Load progress tracking file to prevent duplicates
        sent_emails = set()
        if progress_file and progress_file.exists():
            try:
                with open(progress_file, "r", encoding="utf-8") as f:
                    for line in f:
                        email = line.strip()
                        if email:
                            sent_emails.add(email)
                if sent_emails:
                    print(f"Loaded {len(sent_emails):,} already-sent emails from progress file (will skip duplicates)")
            except Exception as e:
                print(f"Warning: Could not load progress file: {e}")
        
        # Skip already sent emails if resuming
        if start_index > 0:
            recipient_list = recipient_list[start_index:]
            print(f"Resuming from email #{start_index + 1}...")
            print(f"Remaining emails to send: {len(recipient_list):,}")
        else:
            print(f"Starting to send {len(recipient_list)} emails ONE BY ONE...")
        
        print(f"Mode: Individual sending (not bulk)")
        print(f"Delay between emails: 5 seconds\n")
        
        # Open progress file for appending
        progress_fp = None
        if progress_file:
            try:
                progress_fp = open(progress_file, "a", encoding="utf-8")
            except Exception as e:
                print(f"Warning: Could not open progress file for writing: {e}")
        
        try:
            for i, recipient in enumerate(recipient_list, 1):
                # Skip if already sent (prevents duplicates)
                if recipient in sent_emails:
                    print(f"[SKIP] {recipient} already sent (skipping to prevent duplicate)")
                    continue
                
                # Calculate actual email number (for display)
                actual_email_num = start_index + i
                total_emails = start_index + len(recipient_list)
                print(f"[{actual_email_num}/{total_emails}] Sending to {recipient}...", end=" ", flush=True)
                
                try:
                    # Send email ONE BY ONE - each email is sent individually with retries for throttling
                    result = self.send_email(
                        recipient_email=recipient,
                        subject=subject,
                        body_html=body_html,
                        body_text=body_text,
                        attachment_path=attachment_path,
                        logo_path=logo_path,
                        logo_url=logo_url,
                        max_retries=5  # Retry up to 5 times for throttling errors
                    )
                    
                    if result["success"]:
                        print("✓ Success")
                        # Mark as sent in tracking set
                        sent_emails.add(recipient)
                        # Save to progress file immediately to prevent duplicates on restart
                        if progress_fp:
                            progress_fp.write(f"{recipient}\n")
                            progress_fp.flush()  # Ensure it's written immediately
                    else:
                        error_msg = result.get('error', 'Unknown error')
                        print(f"✗ Failed: {error_msg}")
                        
                        # If throttled, wait longer before next email to avoid hitting limit again
                        if "ApplicationThrottled" in str(error_msg) or "429" in str(result.get('status_code', '')):
                            extra_wait = 30  # Wait 30 extra seconds after throttling error
                            print(f"   ⚠ Throttling detected. Waiting {extra_wait} extra seconds before next email...")
                            time.sleep(extra_wait)
                    
                    results.append(result)
                    
                    # IMPORTANT: Fixed 5 second delay between emails to avoid rate limiting
                    # This ensures emails are sent one by one, not in bulk
                    if i < len(recipient_list):
                        delay_seconds = 5.0
                        print(f"   Waiting {delay_seconds} seconds before next email...")
                        time.sleep(delay_seconds)
                        
                except Exception as e:
                    print(f"✗ Error: {str(e)}")
                    results.append({
                        "success": False,
                        "recipient": recipient,
                        "error": str(e)
                    })
        finally:
            if progress_fp:
                progress_fp.close()
        
        return results


def is_valid_email(email: str) -> bool:
    """
    Validate email format using regex
    Returns True if email is valid, False otherwise
    """
    # RFC 5322 compliant email regex (simplified but comprehensive)
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def extract_email_from_line(line: str) -> Optional[str]:
    """
    Extract a valid email address from a line that might contain extra text
    Returns the first valid email found, or None
    """
    line = line.strip()
    if not line:
        return None
    
    # Try to find email pattern in the line
    # Look for pattern: word@domain.extension
    email_pattern = r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b'
    matches = re.findall(email_pattern, line)
    
    if matches:
        # Return the first valid email found
        for match in matches:
            if is_valid_email(match):
                return match.lower()  # Normalize to lowercase
    
    return None


def validate_emails_with_disify(email_list: List[str]) -> Tuple[List[str], dict]:
    """
    Validate emails using Disify API (free email validation service)
    Checks for: disposable emails, invalid DNS, invalid format
    
    API: https://www.disify.com/
    Limit: 10,000 emails per request
    """
    if not email_list:
        return [], {
            "total": 0,
            "invalid_format": 0,
            "invalid_dns": 0,
            "disposable": 0,
            "unique": 0,
            "valid": 0
        }
    
    # Disify API limit is 10,000 emails per request, but URL length limits us
    # Use smaller batches to avoid URL length issues (414 error)
    batch_size = 500  # Smaller batches to avoid URL length limits
    all_valid_emails = []
    total_stats = {
        "total": 0,
        "invalid_format": 0,
        "invalid_dns": 0,
        "disposable": 0,
        "unique": len(email_list),
        "valid": 0
    }
    
    # Process in batches if needed
    for i in range(0, len(email_list), batch_size):
        batch = email_list[i:i + batch_size]
        # Use comma-separated format as per API documentation
        batch_emails_str = ",".join(batch)
        
        try:
            # Call Disify API using GET method
            # URL encode the email list to handle special characters
            # Note: Keep commas unencoded as API expects comma-separated format
            encoded_emails = quote(batch_emails_str, safe=',@.')
            url = f"https://www.disify.com/api/email/{encoded_emails}/mass"
            response = requests.get(url, timeout=60)
            
            if response.status_code == 200:
                result = response.json()
                
                # Get valid emails using session
                session = result.get("session")
                if session:
                    view_url = f"https://www.disify.com/api/view/{session}"
                    valid_response = requests.get(view_url, timeout=30)
                    
                    if valid_response.status_code == 200:
                        valid_emails_text = valid_response.text.strip()
                        if valid_emails_text:
                            valid_emails = [e.strip() for e in valid_emails_text.split("\n") if e.strip()]
                            all_valid_emails.extend(valid_emails)
                
                # Accumulate statistics
                total_stats["total"] += result.get("total", 0)
                total_stats["invalid_format"] += result.get("invalid_format", 0)
                total_stats["invalid_dns"] += result.get("invalid_dns", 0)
                total_stats["disposable"] += result.get("disposable", 0)
                total_stats["valid"] += result.get("valid", 0)
                
                print(f"  Batch {i//batch_size + 1}: Validated {len(batch)} emails, {result.get('valid', 0)} valid")
            else:
                print(f"  Warning: Disify API returned status {response.status_code} for batch {i//batch_size + 1}")
                # Fallback: use format validation only
                for email in batch:
                    if is_valid_email(email):
                        all_valid_emails.append(email)
                        
        except Exception as e:
            print(f"  Warning: Error validating batch {i//batch_size + 1} with Disify: {e}")
            print(f"  Falling back to format validation only for this batch...")
            # Fallback: use format validation only
            for email in batch:
                if is_valid_email(email):
                    all_valid_emails.append(email)
    
    total_stats["valid"] = len(all_valid_emails)
    return all_valid_emails, total_stats


def filter_and_validate_emails(emails: List[str], use_disify: bool = True) -> Tuple[List[str], dict]:
    """
    Filter and validate email addresses:
    - Extract valid emails from lines
    - Remove duplicates
    - Validate format
    - Optionally validate with Disify API (disposable, DNS, etc.)
    - Return cleaned list and statistics
    """
    extracted_emails = []
    invalid_lines = []
    
    # Step 1: Extract emails from lines
    print("Step 1: Extracting emails from lines...")
    for i, line in enumerate(emails, 1):
        email = extract_email_from_line(line)
        if email:
            extracted_emails.append(email)
        else:
            invalid_lines.append((i, line[:50]))  # Store first 50 chars for reporting
    
    # Step 2: Remove duplicates while preserving order
    print("Step 2: Removing duplicates...")
    seen = set()
    unique_emails = []
    duplicates_count = 0
    
    for email in extracted_emails:
        if email not in seen:
            seen.add(email)
            unique_emails.append(email)
        else:
            duplicates_count += 1
    
    # Step 3: Basic format validation
    print("Step 3: Validating email format...")
    format_valid_emails = []
    invalid_emails = []
    
    for email in unique_emails:
        if is_valid_email(email):
            format_valid_emails.append(email)
        else:
            invalid_emails.append(email)
    
    # Step 4: Advanced validation with Disify API (if enabled)
    final_valid_emails = format_valid_emails
    disify_stats = {}
    
    if use_disify and format_valid_emails:
        print(f"Step 4: Validating with Disify API (checking disposable emails, DNS, etc.)...")
        print(f"  This may take a few minutes for large lists...")
        final_valid_emails, disify_stats = validate_emails_with_disify(format_valid_emails)
    else:
        disify_stats = {
            "total": len(format_valid_emails),
            "invalid_format": 0,
            "invalid_dns": 0,
            "disposable": 0,
            "unique": len(format_valid_emails),
            "valid": len(format_valid_emails)
        }
    
    stats = {
        "total_lines": len(emails),
        "extracted_emails": len(extracted_emails),
        "duplicates_removed": duplicates_count,
        "invalid_lines": len(invalid_lines),
        "invalid_emails": len(invalid_emails),
        "format_valid_emails": len(format_valid_emails),
        "disposable_emails": disify_stats.get("disposable", 0),
        "invalid_dns": disify_stats.get("invalid_dns", 0),
        "final_valid_emails": len(final_valid_emails)
    }
    
    return final_valid_emails, stats


def load_recipients_from_file(file_path: Path, use_disify: bool = True) -> Tuple[List[str], dict]:
    """
    Load recipient emails from a text file, validate and filter them
    Returns: (list of valid unique emails, statistics dictionary)
    
    Args:
        file_path: Path to recipients file
        use_disify: If True, use Disify API for advanced validation (disposable, DNS checks)
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Recipients file not found: {file_path}")
    
    raw_lines = []
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            raw_lines.append(line)
    
    # Filter and validate
    valid_emails, stats = filter_and_validate_emails(raw_lines, use_disify=use_disify)
    
    return valid_emails, stats


def main():
    """Main function"""
    # Check required environment variables
    required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print("Error: Missing required environment variables:")
        for var in missing_vars:
            print(f"  - {var}")
        print("\nPlease set these in your .env file or environment.")
        return
    
    # Initialize sender
    sender = GraphEmailSender(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    
    # Get recipients - check for failed_recipients.txt first (for retrying failed emails)
    recipients_file = Path("failed_recipients.txt")
    if not recipients_file.exists():
        recipients_file = Path("recipients.txt")
    
    if not recipients_file.exists():
        print("Error: recipients.txt or failed_recipients.txt file not found!")
        print("\nPlease create a recipients.txt file with one email address per line.")
        print("Example:")
        print("  client1@example.com")
        print("  client2@example.com")
        return
    
    if recipients_file.name == "failed_recipients.txt":
        print("⚠️  Using failed_recipients.txt - retrying failed emails only")
        print()
    
    # Load and validate recipients
    print(f"Loading and validating recipients from {recipients_file.name}...")
    print("Using Disify API for advanced validation (disposable emails, DNS checks)...\n")
    recipients, stats = load_recipients_from_file(recipients_file, use_disify=True)
    
    # Display validation statistics
    print("\n" + "="*60)
    print("EMAIL VALIDATION RESULTS")
    print("="*60)
    print(f"Total lines in file:           {stats['total_lines']:,}")
    print(f"Emails extracted:             {stats['extracted_emails']:,}")
    print(f"Duplicates removed:          {stats['duplicates_removed']:,}")
    print(f"Invalid lines skipped:       {stats['invalid_lines']:,}")
    print(f"Invalid format filtered:     {stats['invalid_emails']:,}")
    if stats.get('disposable_emails', 0) > 0:
        print(f"Disposable emails removed:    {stats['disposable_emails']:,}")
    if stats.get('invalid_dns', 0) > 0:
        print(f"Invalid DNS emails removed:   {stats['invalid_dns']:,}")
    print(f"Final valid unique emails:    {stats['final_valid_emails']:,}")
    print("="*60)
    
    if not recipients:
        print("\nError: No valid email addresses found in recipients.txt")
        print("Please check your recipients file and try again.")
        return
    
    # Ask for confirmation before sending
    print(f"\nReady to send emails to {len(recipients):,} valid recipients.")
    print("All emails have been validated and duplicates removed.")
    print("\nProceeding with email sending...\n")
    
    # Attachment path - customize this to your PDF file name
    attachment_path = Path("attachment.pdf")
    
    if not attachment_path.exists():
        print(f"Warning: Attachment file not found: {attachment_path}")
        print("Emails will be sent without attachment.\n")
        attachment_path = None
    
    # Logo - prioritize JPEG, but will convert PNG to JPEG automatically
    logo_path = None
    logo_url = None
    
    # Try to find local logo file (prioritize logo_black.png, PNG will be converted to JPEG automatically)
    possible_logo_names = ["logo_black.png", "logo.png", "logo.jpg", "logo.jpeg", "logo.gif"]
    for logo_name in possible_logo_names:
        test_logo = Path(logo_name)
        if test_logo.exists():
            logo_path = test_logo
            print(f"Found logo file: {logo_name}")
            if logo_name.endswith('.png'):
                print("(PNG will be converted to JPEG for better email compatibility)\n")
            elif logo_name.endswith('.gif'):
                print("(GIF will be embedded - animation may not work in all email clients)\n")
            else:
                print()
            break
    
    # If no local file found, you can set a logo URL here
    # logo_url = "https://your-logo-url.com/logo.png"
    
    if logo_path:
        print(f"Using logo from file: {logo_path}\n")
    elif logo_url:
        print(f"Using logo from URL: {logo_url}\n")
    else:
        print("Note: No logo file found. Logo will not appear in email signature.")
        print("To add logo, place a PNG or JPG file named 'logo.png' or 'logo.jpg' in the same directory.\n")
    
    # Send emails
    # Start from the beginning (index 0)
    start_from = 0  # Start from email #1
    
    # Progress tracking file to prevent duplicates on restart
    progress_file = Path("sent_emails_progress.txt")
    
    print("\n" + "="*60)
    print("RESUMING EMAIL SENDING")
    print("="*60)
    results = sender.send_emails_one_by_one(
        recipient_list=recipients,
        subject=EMAIL_SUBJECT,
        body_html=EMAIL_BODY_HTML,
        body_text=EMAIL_BODY_TEXT,
        attachment_path=attachment_path,
        logo_path=logo_path,
        logo_url=logo_url,
        start_index=start_from,
        progress_file=progress_file
    )
    
    # Use results directly (no priority recipients)
    all_results = results
    
    # Summary
    print("\n" + "="*50)
    print("SENDING SUMMARY")
    print("="*50)
    successful = sum(1 for r in all_results if r.get("success"))
    failed = len(all_results) - successful
    print(f"Total: {len(all_results)}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    
    if failed > 0:
        print("\nFailed recipients:")
        for result in all_results:
            if not result.get("success"):
                print(f"  - {result.get('recipient')}: {result.get('error', 'Unknown error')}")
    
    # Save results to file
    results_file = Path("send_results.json")
    with open(results_file, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)
    print(f"\nResults saved to: {results_file}")


if __name__ == "__main__":
    main()

