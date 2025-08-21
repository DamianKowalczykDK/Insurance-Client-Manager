from src.model.client import Client, ClientDict
from src.model.report import MonthlyReportDict
from src.model.invoice import InvoiceDict


def test_model_client(client_1: Client, client1_data: ClientDict) -> None:
    result = client_1.to_dict()
    assert result == client1_data
    assert result["name"] == "client1"

def test_invoice_dict(invoice_dict: InvoiceDict) -> None:
    assert invoice_dict["client_name"] == "client1"

def test_report_dict(report_dict: MonthlyReportDict) -> None:
    assert report_dict["gross_total"] == 500

