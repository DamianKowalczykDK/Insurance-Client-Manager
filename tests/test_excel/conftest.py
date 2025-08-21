import tempfile
from datetime import date

import pytest
import os
from typing import Generator
from src.excel.manager.base_manager import ExcelManager
from src.excel.manager.client_manager import ClientExcelManager
from src.model.client import Client, ClientDict


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
def example_client_manager() -> Generator[ClientExcelManager, None, None]:
    tmp_dir = tempfile.gettempdir()
    file_path = os.path.join(tmp_dir, "test.xlsx")
    if os.path.exists(file_path):
        os.remove(file_path)
    manager = ClientExcelManager(file_path)

    yield  manager

    if os.path.exists(file_path):
        os.remove(file_path)

