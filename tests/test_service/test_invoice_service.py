from unittest.mock import patch, MagicMock
from src.model.invoice import InvoiceDict
from config import invoice_service



def test_create_invoice() -> None:
    data: InvoiceDict = {
        "client_name": "test_name",
        "client_email": "test_email",
        "client_tax_no": "test_tax_no",
        "item_price": 100,
        "item_name": "test_item_name",
        "item_quantity": 1
    }

    with patch("httpx.Client") as mock_client:
        mock_client_instance = mock_client.return_value

        fake_invoice = MagicMock()
        mock_client_instance.post = fake_invoice

        invoice_service.create_invoice(data)
        mock_client.assert_called()

def test_get_invoice() -> None:
    fake_data = {
        "invoices": [
            {"id": 1, "amount": 100},
            {"id": 2, "amount": 200},
        ]
    }

    with patch("httpx.Client") as mock_client:
        mock_client_instance = mock_client.return_value.__enter__.return_value
        fake_response = MagicMock()
        fake_response.json.return_value = fake_data
        fake_response.raise_for_status.return_value = None
        mock_client_instance.get.return_value = fake_response

        invoice_service.get_invoice()

    mock_client_instance.get.assert_called()

def test_update_invoice() -> None:
    fake_data = {
        "invoices":
            {"id": 1, "amount": 100},

    }

    with patch("httpx.Client") as mock_client:
        mock_client_instance = mock_client.return_value.__enter__.return_value
        fake_invoice = MagicMock()
        fake_invoice.json.return_value = fake_data
        fake_invoice.raise_for_status.return_value = None

        mock_client_instance.put.return_value = fake_invoice
        invoice_service.update_invoice(1)

    mock_client_instance.put.assert_called()





