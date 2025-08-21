from datetime import datetime, timedelta
from pathlib import Path
from re import match
from unittest.mock import MagicMock, patch

import pytest

from config import client_excel_manager, email_service, invoice_service
from src.excel.manager.client_manager import ClientExcelManager
from src.model.client import Client
from src.service.client_service import ClientService
from src.service.email_service import EmailService


def test_add_client_service(tmp_path: Path, client_2: Client) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))

    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)
    client.add_client(client_2)

    assert client.check_if_client_exists("client2@gmail.com") is True

def test_add_client_service_with_email_already_exist(tmp_path: Path, client_1: Client) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)
    client.add_client(client_1)

    with pytest.raises(ValueError, match="Client with email"):
        client.add_client(client_1)

def test_update_client_service(tmp_path: Path, client_1: Client, client_2: Client) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)
    client.add_client(client_1)

    client.update_client("client1@example.com", client_2)
    ws = manager.get_sheet()

    assert ws["A2"].value == "client2"
    assert client.check_if_client_exists("client2@gmail.com") is True

def test_update_client_service_with_email_already_exist(tmp_path: Path, client_1: Client, client_2: Client) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)
    client.add_client(client_1)
    client.update_client("client1@example.com", client_2)

    with pytest.raises(ValueError, match="Client with email"):
        client.update_client("client1@example.com", client_2)

def test_update_client_service_if_email_not_exist(tmp_path: Path, client_1: Client, client_2: Client) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)

    with pytest.raises(ValueError, match="Client with email client1@example.com not found"):
        client.update_client("client1@example.com", client_2)


def test_confirm_payment_if_client_not_exist(tmp_path: Path) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)

    with pytest.raises(ValueError, match="Client with email client1@example.com not found"):
        client.confirm_payment("client1@example.com")

def test_remove_client_service_if_client_not_exist(tmp_path: Path) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)

    with pytest.raises(ValueError, match="Client with email client1@example.com not found"):
        client.remove_client("client1@example.com")

def test_remove_overdue_clients(tmp_path: Path) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)
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

    client.add_client(client_1)
    client.add_client(client_2)

    manager.load_client_row()
    removed = client.remove_overdue_clients()

    assert "client1@example.com" in removed
    assert "client2@gmail.com" not in removed

def test_remove_overdue_clients_except(tmp_path: Path):
    # przygotowanie managera
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    service = ClientService(manager, mock_email_service, mock_invoice_service)

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
        removed = service.remove_overdue_clients()

    assert "bad@example.com" not in removed

def test_notify_payment_due_in_days(tmp_path: Path, client_1: Client) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)

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

    client.add_client(client_1)
    client.add_client(client_2)
    manager.load_client_row()

    mock_invoice_service.create_invoice.return_value = "fake_invoice_url"

    client.notify_payment_due_in_days()

    assert mock_email_service.send_email.called

def test_notify_payment_due_in_days_except(tmp_path: Path) -> None:
    file_path = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(file_path))
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client = ClientService(manager, mock_email_service, mock_invoice_service)

    bad_client = {
        "name": "bad_client",
        "email": "bad@example.com",
        "insurance_company": "abc",
        "car_model": "Audi",
        "car_year": 2020,
        "price": 1000,
        "next_payment": "not_a_date"
    }

    manager.load_client_row()

    mock_invoice_service.create_invoice.return_value = "fake_invoice_url"
    client.notify_payment_due_in_days()

    with patch.object(ClientExcelManager, "load_client_row", return_value=[bad_client]):
        client.notify_payment_due_in_days(1)


    assert not mock_email_service.send_email.called




