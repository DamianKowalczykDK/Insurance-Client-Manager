from datetime import datetime, date
from config import create_client_service, client_service, invoice_service
from src.model.client import Client
from src.model.invoice import InvoiceDict
from src.service.invoice_service import InvoiceService


def main() -> None:
    client1 = Client(
        name='Piotr Nowak',
        email='piotrnowak@example.com ',
        insurance_company="PZU",
        car_model="Audi A5",
        car_year=2013,
        price=2500,
        next_payment=date(2025,8,18)
    )

    client2 = Client(
        name='Tadeusz Nowak',
        email='tadeusznowak@example.pl',
        insurance_company="WARTA",
        car_model="Opel Corsa",
        car_year=2009,
        price=700,
        next_payment=date(2025, 8, 16)
    )

    client3 = Client(
        name='Damian Nowak',
        email='damiannowak@example.pl',
        insurance_company="LINK4",
        car_model="Seat Leon",
        car_year=2011,
        price=1500,
        next_payment=date(2025, 8, 12)
    )


    # client_service.add_client(client1)
    # client_service.add_client(client2)
    # client_service.add_client(client3)

    # client_service.confirm_payment("alonamelnyk@example.pl")

    # client_service.notify_payment_due_in_days()


    # invoice_service.get_invoice()
    # invoice_service.update_invoice(406516632)
    # print(client_service.generate_monthly_report())

if __name__ == '__main__':
    main()
