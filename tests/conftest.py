from datetime import date, datetime

import pytest

from src.model.client import Client, ClientDict
from src.model.invoice import InvoiceDict
from src.model.report import MonthlyReportDict


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
