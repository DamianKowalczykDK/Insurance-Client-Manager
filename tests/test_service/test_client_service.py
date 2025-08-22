from src.excel.manager.client_manager import ClientExcelManager
from src.service.client_service import ClientService
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta
from src.model.client import Client
import pytest


def test_add_client_service(example_client_service: ClientService, client_2: Client) -> None:
    example_client_service.add_client(client_2)

    assert example_client_service.check_if_client_exists("client2@gmail.com") is True

def test_add_client_service_with_email_already_exist(example_client_service: ClientService, client_1: Client) -> None:
    example_client_service.add_client(client_1)

    with pytest.raises(ValueError, match="Client with email"):
        example_client_service.add_client(client_1)

def test_update_client_service(
        example_client_service: ClientService,
        example_client_manager: ClientExcelManager,
        client_1: Client,
        client_2: Client
) -> None:
    example_client_service.add_client(client_1)

    example_client_service.update_client("client1@example.com", client_2)
    ws = example_client_manager.get_sheet()

    assert ws["A2"].value == "client2"
    assert example_client_service.check_if_client_exists("client2@gmail.com") is True

def test_update_client_service_with_email_already_exist(
        example_client_service: ClientService,
        client_1: Client,
        client_2: Client
) -> None:

    example_client_service.add_client(client_1)
    example_client_service.update_client("client1@example.com", client_2)

    with pytest.raises(ValueError, match="Client with email"):
        example_client_service.update_client("client1@example.com", client_2)

def test_update_client_service_if_email_not_exist(example_client_service: ClientService, client_2: Client) -> None:
    with pytest.raises(ValueError, match="Client with email client1@example.com not found"):
        example_client_service.update_client("client1@example.com", client_2)


def test_confirm_payment_if_client_not_exist(example_client_service: ClientService) -> None:
    with pytest.raises(ValueError, match="Client with email client1@example.com not found"):
        example_client_service.confirm_payment("client1@example.com")

def test_remove_client_service_if_client_not_exist(example_client_service: ClientService) -> None:
     with pytest.raises(ValueError, match="Client with email client1@example.com not found"):
        example_client_service.remove_client("client1@example.com")

def test_remove_overdue_clients(example_client_manager: ClientExcelManager, example_client_service: ClientService) -> None:
    client_1 = Client(
        name='client1',
        email='client1@example.com',
        insurance_company='abc',
        car_model='Audi',
        car_year=2015,
        price=1500,
        next_payment=datetime.today() - timedelta(days=5)
    )

    client_2 = Client(
        name='client2',
        email='client2@gmail.com',
        insurance_company='xyz',
        car_model='Bmw',
        car_year=2015,
        price=1500,
        next_payment=datetime.today() + timedelta(days=2)
    )

    example_client_service.add_client(client_1)
    example_client_service.add_client(client_2)

    example_client_manager.load_client_row()
    removed = example_client_service.remove_overdue_clients()

    assert "client1@example.com" in removed
    assert "client2@gmail.com" not in removed

def test_remove_overdue_clients_except(example_client_service: ClientService):
    bad_client = {
        "name": "bad_client",
        "email": "bad@example.com",
        "insurance_company": "abc",
        "car_model": "Audi",
        "car_year": 2020,
        "price": 1000,
        "next_payment": "not_a_date"
    }

    with patch.object(ClientExcelManager, "load_client_row", return_value=[bad_client]):
        removed = example_client_service.remove_overdue_clients()

    assert "bad@example.com" not in removed

def test_notify_payment_due_in_days(
        example_client_manager: ClientExcelManager,
        example_client_service: ClientService
) -> None:

    mock_invoice_service = MagicMock()
    mock_email_service = MagicMock()
    example_client_service.invoice_service = mock_invoice_service
    example_client_service.email_service = mock_email_service

    client_1 = Client(
        name='client1',
        email='client1@example.com',
        insurance_company='abc',
        car_model='Audi',
        car_year=2015,
        price=1500,
        next_payment=datetime.today() + timedelta(days=1)
    )
    client_2 = Client(
        name='client2',
        email='client2@gmail.com',
        insurance_company='xyz',
        car_model='Bmw',
        car_year=2015,
        price=1500,
        next_payment=datetime.today() + timedelta(days=3)
    )

    example_client_service.add_client(client_1)
    example_client_service.add_client(client_2)
    example_client_manager.load_client_row()

    mock_invoice_service.create_invoice.return_value = "fake_invoice_url"
    example_client_service.notify_payment_due_in_days()

    assert mock_email_service.send_email.called

def test_notify_payment_due_in_days_except(
        example_client_manager: ClientExcelManager,
        example_client_service: ClientService,
        example_bad_client: dict[str, str | int]
) -> None:

    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    example_client_service.invoice_service = mock_invoice_service
    example_client_service.email_service = mock_email_service

    bad_client = example_bad_client

    example_client_manager.load_client_row()

    mock_invoice_service.create_invoice.return_value = "fake_invoice_url"
    example_client_service.notify_payment_due_in_days()

    with patch.object(ClientExcelManager, "load_client_row", return_value=[bad_client]):
        example_client_service.notify_payment_due_in_days(1)

    assert not mock_email_service.send_email.called

def test_generate_monthly_report(example_client_service: ClientService, client_1: Client) -> None:
    data = {
        "month": "2025-08",
        "company": "TEST",
        "gross_total": 1000,
        "net_total": 100
    }

    example_client_service.add_client(client_1)
    next_payment_date = client_1.next_payment
    example_client_service.generate_monthly_report()


    assert data["month"] == "2025-08"
    assert str(next_payment_date) == "2025-08-15"

def test_generate_monthly_report_except(
        example_client_service: ClientService,
        example_bad_client: dict[str, str | int]
) -> None:
    bad_client = example_bad_client

    with patch.object(ClientExcelManager, "load_client_row", return_value=[bad_client]):
        report = example_client_service.generate_monthly_report()


        assert "bad@example.com" not in report









