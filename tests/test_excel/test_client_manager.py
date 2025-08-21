from tests.test_excel.test_base_manager_and_style import italic_font_style, bold_font_style
from src.excel.manager.client_manager import ClientExcelManager
from datetime import datetime, date, timedelta
from src.model.client import ClientDict
from tests.conftest import client1_data
from unittest.mock import MagicMock, patch
import pytest


def test_insert_and_load_client(example_client_manager: ClientExcelManager, client1_data: ClientDict):
    example_client_manager.insert_main_row(client1_data)
    clients = example_client_manager.load_client_row()

    assert len(clients) == 1
    assert clients[0]["name"] == "client1"
    assert clients[0]["insurance_company"] == "abc"

def test_get_next_main_table_row(example_client_manager: ClientExcelManager, client1_data: ClientDict, client2_data: ClientDict) -> None:

    example_client_manager.insert_main_row(client1_data)
    example_client_manager.insert_main_row(client2_data)

    row_idx = example_client_manager.get_next_main_table_row()

    assert row_idx == 4

def test_get_next_main_table_row_empty_sheet(example_client_manager: ClientExcelManager):
    mock_ws = MagicMock()
    mock_ws.max_row = 5

    with patch.object(example_client_manager, "get_sheet", return_value=mock_ws):
        next_row = example_client_manager.get_next_main_table_row()
        assert next_row == 6

def test_update_client_row(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)
    ws = example_client_manager.get_sheet()
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

    updated_client = example_client_manager.update_client_row(2, "client1@example.com", data=client_updated)
    no_updated_client = example_client_manager.update_client_row(2, "client2@example.com", data=client_updated)

    assert updated_client is True
    assert email_value == "client1@example.com"
    assert no_updated_client is False

def test_shift_payment_date_with_str(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    example_client_manager.shift_payment_date(2, "client1@example.com",7, 30)
    ws = example_client_manager.get_sheet()
    next_payment = ws.cell(row=2, column=7).value

    assert next_payment == "2025-09-14"
    assert isinstance(next_payment, str)

def test_shift_payment_date_with_data_object(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    ws = example_client_manager.get_sheet()
    ws.cell(row=2, column=7).value = date(2025,8,15)

    example_client_manager.shift_payment_date(2, "client1@example.com", 7, 30)
    next_payment = ws.cell(row=2, column=7).value

    assert next_payment == "2025-09-14"

def test_shift_payment_date_invalid_type(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)
    ws = example_client_manager.get_sheet()
    ws.cell(row=2, column=7).value = 123.45

    result = example_client_manager.shift_payment_date(2, "client1@example.com", 7, 30)

    assert result is False

def test_shift_payment_date_value_not_found(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    result = example_client_manager.shift_payment_date(2, "test@example.com", 7, 30)

    assert result is False

def test_remove_client(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    remove_client = example_client_manager.remove_client_row(2, "client1@example.com")
    ws = example_client_manager.get_sheet()
    email = ws.cell(row=2, column=2).value

    assert email is None
    assert remove_client is True

def test_remove_client_if_not_data(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    remove_client = example_client_manager.remove_client_row(2, "test@example.com")

    assert remove_client is False

def test_overwrite_clients(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    example_client_manager.overwrite_clients([client1_data])

    ws = example_client_manager.get_sheet()
    email = ws.cell(row=2, column=2).value

    assert email == "client1@example.com"

def test_aplay_uppercase(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    example_client_manager.uppercase_columns = ["A"]
    ws = example_client_manager.get_sheet()
    ws.cell(row=2, column=1).value = "client1"

    example_client_manager._aplay_uppercase()
    name = ws.cell(row=2, column=1).value
    assert name == "CLIENT1"

def test_validate_headers_with_raise_value_error(
        example_client_manager: ClientExcelManager,
        client1_data: ClientDict
)-> None:
    example_client_manager.insert_main_row(client1_data)

    example_client_manager.main_table_headers = []
    with pytest.raises(ValueError, match="Main table headers collection should have at least 1 item"):
        example_client_manager._validate_headers()

def test_validate_headers_with_raise_value_company_table_header(
        example_client_manager: ClientExcelManager,
        client1_data: ClientDict
) -> None:
    example_client_manager.insert_main_row(client1_data)

    example_client_manager.company_table_headers = []

    with pytest.raises(ValueError, match="Insurance Company table headers collection should have 2 items"):
        example_client_manager._validate_headers()

def test_validate_headers_with_raise_value_summary_table_header(
        example_client_manager: ClientExcelManager,
        client1_data: ClientDict
) -> None:
    example_client_manager.insert_main_row(client1_data)
    example_client_manager.summary_table_headers = []
    with pytest.raises(ValueError, match="Summary table headers collection should have 2 items"):
        example_client_manager._validate_headers()

def test_validate_column_ranges(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    example_client_manager.main_table_start_col = "A"
    example_client_manager.main_table_headers = ["H1", "H2"]
    example_client_manager.company_table_start_col = "B"
    example_client_manager.company_table_headers = ["C1", "C2"]

    with pytest.raises(ValueError, match="Incorrect columns ranges"):
        example_client_manager._validate_column_ranges()

def test_style_summary_table(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    example_client_manager.summary_table_start_col = "A"

    example_client_manager.header_style = italic_font_style()
    example_client_manager.row_style = bold_font_style()

    example_client_manager._style_summary_table()
    ws = example_client_manager.get_sheet()

    cell_header = ws["A1"]
    cell_row = ws["B2"]
    assert cell_header.font.italic == True
    assert cell_row.font.bold == True

def test_highlight_overdue_payment(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    ws = example_client_manager.get_sheet()
    yesterday = datetime.today().date() - timedelta(days=1)
    ws.cell(row=2, column=3).value = yesterday
    ws.cell(row=3, column=3).value = "invalid_data"

    example_client_manager.main_table_headers = ["NAME", "EMAIL", "NEXT_PAYMENT"]
    example_client_manager.overdue_style = italic_font_style()

    example_client_manager._highlight_overdue_payment()

    cell = ws.cell(row=2, column=3)
    cell_invalid = ws.cell(row=3, column=3)
    assert cell.font.italic == True
    assert cell_invalid.font.italic == False

def test_load_client_row_with_exceptions(example_client_manager: ClientExcelManager, client1_data: ClientDict) -> None:
    example_client_manager.insert_main_row(client1_data)

    ws = example_client_manager.get_sheet()
    ws.cell(row=2, column=5).value = "invalid_data"
    clients = example_client_manager.load_client_row()

    assert clients == []













