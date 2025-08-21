from unittest.mock import MagicMock, patch

from mypy.checker import Mapping

from config import email_service


def test_send_email() ->  None:

    with patch("smtplib.SMTP") as mock_smtp:
        instance = mock_smtp.return_value.__enter__.return_value
        instance.sendmail = MagicMock()

        email_service.send_email("test@example.com", "Subject", "<b>HTML</b>")
        instance.sendmail.assert_called_once()

def test_send_email_with_except() ->  None:
    with patch("smtplib.SMTP") as mock_smtp:
        instance = mock_smtp.return_value.__enter__.return_value
        instance.sendmail.side_effect = Exception("SMTP fail")

        email_service.send_email("test@example.com", "Subject", "<b>HTML</b>")
        assert instance.sendmail.call_count == 1