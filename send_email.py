import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email(to_email, code, email_type="verify", username=""):
    sender_email = "justaarnavjindal@gmail.com"
    sender_password = "jfep ivij usqy gsrl"  # Provided by the user
    
    if email_type == "reset":
        subject = "Reset your ProMatX Password"
        greeting = f"<p>Hello {username},</p>" if username else "<p>Hello,</p>"
        body_text = f"We received a request to reset your password for your ProMatX account. Please use the 6-digit verification code below to reset your password:"
        if username:
            body_text = f"We received a request to reset your password for your ProMatX account (username: <strong>{username}</strong>). Please use the 6-digit verification code below to reset your password:"
    else:
        subject = "Verify your ProMatX Account"
        greeting = "<p>Hello,</p>"
        body_text = "Thank you for registering an account with ProMatX. Please use the 6-digit verification code below to complete your registration and activate your account:"
    
    # Beautiful HTML email template matching the modern premium dark aesthetic of the calculator
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
                background-color: #0D0D0D;
                margin: 0;
                padding: 0;
                color: #E5E7EB;
            }}
            .container {{
                max-width: 580px;
                margin: 40px auto;
                background-color: #1C1C1C;
                border-radius: 12px;
                overflow: hidden;
                border: 1px solid rgba(229, 231, 235, 0.1);
                box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
            }}
            .header {{
                background-color: #1C1C1C;
                border-bottom: 3px solid #3B82F6;
                padding: 30px 20px;
                text-align: center;
            }}
            .header h1 {{
                margin: 0;
                font-size: 24px;
                font-weight: 700;
                letter-spacing: -0.01em;
                color: #3B82F6;
            }}
            .content {{
                padding: 40px;
                color: #E5E7EB;
                line-height: 1.6;
            }}
            .content p {{
                margin: 0 0 20px 0;
                font-size: 15px;
                font-weight: 400;
            }}
            .code-container {{
                text-align: center;
                margin: 35px 0;
                padding: 24px;
                background-color: rgba(59, 130, 246, 0.08);
                border-radius: 8px;
                border: 1px dashed rgba(59, 130, 246, 0.25);
            }}
            .verification-code {{
                font-family: 'Consolas', 'Courier New', Courier, monospace;
                font-size: 34px;
                font-weight: 700;
                letter-spacing: 0.25em;
                color: #F97316;
            }}
            .footer {{
                background-color: #0D0D0D;
                padding: 24px;
                text-align: center;
                font-size: 11px;
                color: #9CA3AF;
                border-top: 1px solid rgba(229, 231, 235, 0.08);
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>ProMatX</h1>
            </div>
            <div class="content">
                {greeting}
                <p>{body_text}</p>
                <div class="code-container">
                    <span class="verification-code">{code}</span>
                </div>
                <p>This code will expire shortly. If you did not request an account setup, you can safely ignore this email.</p>
                <p>Best regards,<br><strong>The ProMatX Team</strong></p>
            </div>
            <div class="footer">
                &copy; 2026 ProMatX &middot; Radiation Attenuation Calculator<br>
                Interactive Calculation and Graphical Shielding Analysis
            </div>
        </div>
    </body>
    </html>
    """
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = to_email
    
    msg.attach(MIMEText(html_content, "html"))
    
    try:
        # Connect to Gmail SMTP server using TLS
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
        print(f"Verification email successfully sent to {to_email}")
    except Exception as e:
        print(f"Error sending email: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python send_email.py <to_email> <code> [email_type] [username]")
        sys.exit(1)
        
    to_email = sys.argv[1]
    code = sys.argv[2]
    email_type = sys.argv[3] if len(sys.argv) > 3 else "verify"
    username = sys.argv[4] if len(sys.argv) > 4 else ""
    
    send_email(to_email, code, email_type, username)
