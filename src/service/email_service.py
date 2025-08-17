from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import ssl


class EmailService:
    def __init__(self,
                 smtp_server: str,
                 port: int,
                 sender_email: str,
                 sender_password: str) -> None:
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