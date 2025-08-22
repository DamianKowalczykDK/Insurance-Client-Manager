from src.excel.manager.client_manager import ClientExcelManager
from src.excel.manager.base_manager import ExcelManager
from src.service.client_service import ClientService
from src.model.client import Client, ClientDict
from src.model.report import MonthlyReportDict
from src.model.invoice import InvoiceDict
from unittest.mock import MagicMock
from typing import Generator
from datetime import date
from pathlib import Path
import tempfile
import pytest
import os


@pytest.fixture
def example_bad_client() -> dict[str, str | int]:
    return  {
        "name": "bad_client",
        "email": "bad@example.com",
        "insurance_company": "abc",
        "car_model": "Audi",
        "car_year": 2020,
        "price": 1000,
        "next_payment": "not_a_date"
    }


@pytest.fixture
def client_1() -> Client:
    return Client(
        name="client1",
        email="client1@example.com",
        insurance_company="abc",
        car_model="Audi",
        car_year=2015,
        price=1500,
        next_payment=date(2025, 8, 15)
    )

@pytest.fixture
def client_2() -> Client:
    return Client(
        name="client2",
        email="client2@gmail.com",
        insurance_company="xyz",
        car_model="Bmw",
        car_year=2015,
        price=1500,
        next_payment=date(2025, 8, 15)
    )

@pytest.fixture
def client1_data() -> ClientDict:
    return {
        "name":"client1",
        "email":"client1@example.com",
        "insurance_company":"abc",
        "car_model":"Audi",
        "car_year":2015,
        "price":1500,
        "next_payment":"2025-08-15"
    }

@pytest.fixture
def client2_data() -> ClientDict:
    return {
        "name":"client2",
        "email":"client2@example.com",
        "insurance_company":"def",
        "car_model":"Audi",
        "car_year":2012,
        "price":1500,
        "next_payment":"2025-08-15"
    }


@pytest.fixture
def invoice_dict() -> InvoiceDict:
    return {
        "client_name":"client1",
        "client_email":"client1@example.com",
        "client_tax_no": "123",
        "item_name": "abc",
        "item_quantity": 1,
        "item_price": 1500,
    }

@pytest.fixture
def report_dict() -> MonthlyReportDict:
    return {
        "month": "May",
        "company": {"abc": 1},
        "gross_total": 500,
        "net_total": 100,
    }


@pytest.fixture
def example_base_manager() -> Generator[ExcelManager, None, None]:
    tmp_dir = tempfile.gettempdir()
    file_path = os.path.join(tmp_dir, "test.xlsx")
    if os.path.exists(file_path):
        os.remove(file_path)
    manager = ExcelManager(file_path)
    sheet_name = "test"
    manager.workbook.create_sheet(sheet_name)

    yield  manager

    if os.path.exists(file_path):
        os.remove(file_path)

@pytest.fixture
def example_client_manager(tmp_path: Path) -> ClientExcelManager:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    return manager


@pytest.fixture
def example_client_service(example_client_manager: ClientExcelManager) -> ClientService:
    mock_email_service = MagicMock()
    mock_invoice_service = MagicMock()
    client_service = ClientService(example_client_manager, mock_email_service, mock_invoice_service)
    return client_service
