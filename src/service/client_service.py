from src.excel.manager.client_manager import ClientExcelManager
from src.service.invoice_service import InvoiceService
from src.service.email_service import EmailService
from datetime import datetime, timedelta
from src.model.client import Client


class ClientService:
    def __init__(
            self,
            client_excel_manager: ClientExcelManager,
            email_service: EmailService,
            invoice_service: InvoiceService
    ) -> None:
        self.client_excel_manager = client_excel_manager
        self.email_service = email_service
        self.invoice_service = invoice_service

    def add_client(self, client: Client) -> None:
        if self.check_if_client_exists(client.email):
            raise ValueError(f"Client with email {client.email} already exists")
        self.client_excel_manager.insert_main_row(client.to_dict())

    def update_client(self, email: str, update_client: Client) -> None:
        if email != update_client.email and self.check_if_client_exists(update_client.email):
            raise ValueError(f"Client with email {update_client.email} already exists")

        if not self.client_excel_manager.update_client_row(2, email, update_client.to_dict()):
            raise ValueError(f"Client with email {email} not found")

    def confirm_payment(self, email: str, days: int = 30) -> None:
        if not self.client_excel_manager.shift_payment_date(2, email, 7, days):
            raise ValueError(f"Client with email {email} not found")

    def remove_client(self, email: str) -> None:
        if not self.client_excel_manager.remove_client_row(2, email):
            raise ValueError(f"Client with email {email} not found")

    def check_if_client_exists(self, email: str) -> bool:
        clients = self.client_excel_manager.load_client_row()
        return any(c["email"] == email for c in clients)


    def notify_payment_due_in_days(self, days_ahead: int = 1) -> None:
        target_days = (datetime.today() + timedelta(days=days_ahead)).date()
        clients = self.client_excel_manager.load_client_row()

        for client in clients:
            try:
                payment_date_str = client["next_payment"].split()[0]
                payment_date = datetime.strptime(payment_date_str, "%Y-%m-%d").date()
            except Exception:
                continue

            if payment_date == target_days:
                invoice_url = self.invoice_service.create_invoice({
                        "client_name": client["name"],
                        "client_email": client["email"],
                        "client_tax_no": "123-456-78-90",
                        "item_name": f"Polisa ubezpieczeniowa za auto marki {client['car_model']}",
                        "item_quantity": 1,
                        "item_price": client["price"],
                    })
                self.email_service.send_email(
                    recipient_email=client["email"],
                    subject=f"Payment for insurance policy",
                    html=f"""
                    <html>
                        <body>
                            <p>Hello {client['name']},</p>
                            <p>We would like to remind you that the payment deadline for your insurance policy for the
                            {client["car_model"]} is on <b>{payment_date}</b>.</p>
                            <p>Please make sure to complete the payment on time.</p>
                            <p>Your invoice <a href={invoice_url}>Link</a></p>
                        </body>
                    </html>
                    """
                )
                print(f"Email with reminder send to {client['email']} date: {payment_date}")

    def remove_overdue_clients(self, overdue_days: int = 3) -> list[str]:
        today = datetime.today().date()
        all_clients = self.client_excel_manager.load_client_row()

        remaining_clients = []
        removed_clients = []
        for client in all_clients:
            try:
                payment_date_str = client["next_payment"].split()[0]
                payment_date = datetime.strptime(payment_date_str, "%Y-%m-%d").date()
            except Exception:
                remaining_clients.append(client)
                continue

            if (today - payment_date).days < overdue_days:
                remaining_clients.append(client)
            else:
                removed_clients.append(client["email"])

        self.client_excel_manager.overwrite_clients(remaining_clients)
        return removed_clients


    # def generate_monthly_report(self) -> MonthlyReportDict:
    #     students = self.client_excel_manager.load_client_row()
    #     current_month = datetime.today().strftime("%Y-%m")
    #     ratio = self.client_excel_manager.ratio
    #
    #     course_count: dict[str, int] = defaultdict(int)
    #     gross_total = 0
    #
    #     for student in students:
    #         try:
    #             payment_date_str = student["next_payment"].split()[0]
    #             payment_date = datetime.strptime(payment_date_str, "%Y-%m-%d").date()
    #         except Exception:
    #             continue
    #
    #         if payment_date.strftime("%Y-%m") == current_month:
    #             course = student["course"]
    #             price = student["price"]
    #             course_count[course] += 1
    #             gross_total += price
    #
    #     net_total = round(gross_total * ratio)
    #
    #     return {
    #         "month": current_month,
    #         "courses": course_count,
    #         "gross_total": gross_total,
    #         "net_total": net_total
    #     }




