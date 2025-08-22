from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import ssl


class EmailService:
    """Service for sending emails via SMTP."""

    def __init__(
        self,
        smtp_server: str,
        port: int,
        sender_email: str,
        sender_password: str
    ) -> None:
        """Initialize the EmailService with SMTP server details.

        Args:
            smtp_server: SMTP server address (e.g., "smtp.gmail.com").
            port: SMTP server port (e.g., 587 for TLS).
            sender_email: Email address used to send emails.
            sender_password: Password or app-specific token for the sender email.
        """
        self.smtp_server = smtp_server
        self.port = port
        self.sender_email = sender_email
        self.sender_password = sender_password

    def send_email(
        self,
        recipient_email: str,
        subject: str,
        html: str | None = None,
    ) -> None:
        """Send an email to a recipient with optional HTML content.

        Args:
            recipient_email: Recipient's email address.
            subject: Subject line of the email.
            html: Optional HTML content of the email.

        Raises:
            Exception: If sending the email fails.
        """
        message = MIMEMultipart("alternative")
        message["From"] = self.sender_email
        message["To"] = recipient_email
        message["Subject"] = subject

        if html:
            html_part = MIMEText(html, "html")
            message.attach(html_part)

        try:
            context = ssl.create_default_context()
            with smtplib.SMTP(self.smtp_server, self.port) as server:
                server.starttls(context=context)
                server.login(self.sender_email, self.sender_password)
                server.sendmail(self.sender_email, recipient_email, message.as_string())
            print("Email sent successfully")
        except Exception as e:
            print(f"Failed to send email: {e}")
