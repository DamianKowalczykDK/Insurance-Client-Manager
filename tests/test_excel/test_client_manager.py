from datetime import datetime, date, timedelta
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from src.excel.manager.client_manager import ClientExcelManager
from src.model.client import ClientDict
from tests.conftest import client1_data
from tests.test_excel.test_base_manager_and_style import italic_font_style, bold_font_style


def test_insert_and_load_client(tmp_path: Path, client1_data: ClientDict):
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))

    manager.insert_main_row(client1_data)
    clients = manager.load_client_row()

    assert len(clients) == 1
    assert clients[0]["name"] == "client1"
    assert clients[0]["insurance_company"] == "abc"

def test_get_next_main_table_row(tmp_path: Path, client1_data: ClientDict, client2_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)
    manager.insert_main_row(client2_data)

    row_idx = manager.get_next_main_table_row()

    assert row_idx == 4


def test_get_next_main_table_row_empty_sheet():
    manager = ClientExcelManager("dummy.xlsx")

    mock_ws = MagicMock()
    mock_ws.max_row = 5

    manager.get_sheet = MagicMock(return_value=mock_ws)

    next_row = manager.get_next_main_table_row()
    assert next_row == 6

def test_update_client_row(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)
    ws = manager.get_sheet()
    email_value = ws.cell(row=2, column=2).value

    client_updated: ClientDict = {
        "name": "client2",
        "email": "client1@example.com",
        "insurance_company": "abc",
        "car_model": "Audi",
        "car_year": 2015,
        "price": 1500,
        "next_payment": "2025-08-15"
    }

    updated_client = manager.update_client_row(2, "client1@example.com", data=client_updated)
    no_updated_client = manager.update_client_row(2, "client2@example.com", data=client_updated)

    assert updated_client is True
    assert email_value == "client1@example.com"
    assert no_updated_client is False

def test_shift_payment_date_with_str(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    manager.shift_payment_date(2, "client1@example.com",7, 30)
    ws = manager.get_sheet()
    next_payment = ws.cell(row=2, column=7).value

    assert next_payment == "2025-09-14"
    assert isinstance(next_payment, str)

def test_shift_payment_date_with_data_object(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    ws = manager.get_sheet()
    ws.cell(row=2, column=7).value = date(2025,8,15)

    manager.shift_payment_date(2, "client1@example.com", 7, 30)

    next_payment = ws.cell(row=2, column=7).value

    assert next_payment == "2025-09-14"

def test_shift_payment_date_invalid_type(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)
    ws = manager.get_sheet()
    ws.cell(row=2, column=7).value = 123.45

    result = manager.shift_payment_date(2, "client1@example.com", 7, 30)

    assert result is False

def test_shift_payment_date_value_not_found(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    result = manager.shift_payment_date(2, "test@example.com", 7, 30)

    assert result is False

def test_remove_client(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    remove_client = manager.remove_client_row(2, "client1@example.com")
    ws = manager.get_sheet()
    email = ws.cell(row=2, column=2).value

    assert email is None
    assert remove_client is True

def test_remove_client_if_not_data(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    remove_client = manager.remove_client_row(2, "test@example.com")

    assert remove_client is False

def test_overwrite_clients(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    manager.overwrite_clients([client1_data])

    ws = manager.get_sheet()
    email = ws.cell(row=2, column=2).value

    assert email == "client1@example.com"

def test_aplay_uppercase(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    manager.uppercase_columns = ["A"]
    ws = manager.get_sheet()
    ws.cell(row=2, column=1).value = "client1"

    manager._aplay_uppercase()
    name = ws.cell(row=2, column=1).value
    assert name == "CLIENT1"

def test_validate_headers_with_raise_value_error(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    manager.main_table_headers = []
    with pytest.raises(ValueError, match="Main table headers collection should have at least 1 item"):
        manager._validate_headers()

def test_validate_headers_with_raise_value_company_table_header(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    manager.company_table_headers = []

    with pytest.raises(ValueError, match="Insurance Company table headers collection should have 2 items"):
        manager._validate_headers()

def test_validate_headers_with_raise_value_summary_table_header(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)
    manager.summary_table_headers = []
    with pytest.raises(ValueError, match="Summary table headers collection should have 2 items"):
        manager._validate_headers()

def test_validate_column_ranges(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    manager.main_table_start_col = "A"
    manager.main_table_headers = ["H1", "H2"]
    manager.company_table_start_col = "B"
    manager.company_table_headers = ["C1", "C2"]

    with pytest.raises(ValueError, match="Incorrect columns ranges"):
        manager._validate_column_ranges()

def test_style_summary_table(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    manager.summary_table_start_col = "A"

    manager.header_style = italic_font_style()
    manager.row_style = bold_font_style()

    manager._style_summary_table()
    ws = manager.get_sheet()

    cell_header = ws["A1"]
    cell_row = ws["B2"]
    assert cell_header.font.italic == True
    assert cell_row.font.bold == True

def test_highlight_overdue_payment(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    ws = manager.get_sheet()
    yesterday = datetime.today().date() - timedelta(days=1)
    ws.cell(row=2, column=3).value = yesterday
    ws.cell(row=3, column=3).value = "invalid_data"

    manager.main_table_headers = ["NAME", "EMAIL", "NEXT_PAYMENT"]
    manager.overdue_style = italic_font_style()

    manager._highlight_overdue_payment()

    cell = ws.cell(row=2, column=3)
    cell_invalid = ws.cell(row=3, column=3)
    assert cell.font.italic == True
    assert cell_invalid.font.italic == False

def test_load_client_row_with_exceptions(tmp_path: Path, client1_data: ClientDict) -> None:
    filepath = tmp_path / "clients.xlsx"
    manager = ClientExcelManager(str(filepath))
    manager.insert_main_row(client1_data)

    ws = manager.get_sheet()
    ws.cell(row=2, column=5).value = "invalid_data"
    clients = manager.load_client_row()

    assert clients == []













