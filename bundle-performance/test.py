import smtplib
from email.message import EmailMessage
import ssl

SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587

USERNAME = "partnerhub@safaricom.et"  # Your full email address
PASSWORD = "j0jV55u206Cki#"   # Your working password

SENDER = "Mekin.Jemal@safaricom.et"
RECEIVER = "Mekin.Jemal@safaricom.et"

# Create a clean message structure using email.message
msg = EmailMessage()
msg["From"] = SENDER
msg["To"] = RECEIVER
msg["Subject"] = "SMTP Test from Python"
msg.set_content("Hello, this is a test email from Python matching JavaMail settings.")

try:
    print("Connecting to Office 365 SMTP server...")
    # Increase timeout to 60 seconds to match heavy corporate networks
    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60)
    
    # Optional: uncomment if debugging network handshake issues
    # server.set_debuglevel(1)
    
    server.ehlo()
    
    # Secure the connection using a default SSL context (similar to Java's starttls)
    context = ssl.create_default_context()
    server.starttls(context=context)
    server.ehlo()

    print("Authenticating...")
    server.login(USERNAME, PASSWORD)

    print("Sending email...")
    server.send_message(msg)

    print("Email sent successfully!")
    server.quit()

except Exception as e:
    print(f"Email sending failed: {e}")