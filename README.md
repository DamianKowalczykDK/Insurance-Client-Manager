# Insurance Client Manager

A Python system for managing insurance clients, invoices, and payment notifications with Excel backend.  

---

## Features

- Manage clients in Excel (`Clients.xlsx`)  
- Track upcoming and overdue payments  
- Generate and send invoices via Fakturownia API  
- Styled Excel tables for better readability  
- Email reminders for clients  
- Monthly reports per insurance company  
- Automated background scheduler for notifications & cleanup  

---

## Quick Start
### Run the application
```bash
poetry run python main.py
poetry run python .\main_scheduler.py
```
1. **Clone the repo**
```bash
git clone https://github.com/DamianKowalczykDK/Insurance-Client-Manager
cd insurance-client-manager
```
2. **Create a virtual environment & install dependencies** 
```bash
poetry install --with dev
```
3. **.env**
```text
SMTP_SERVER=smtp.example.com
SMTP_PORT=587
SENDER_EMAIL=your_email@example.com
SENDER_PASSWORD=your_password
INVOICE_API_TOKEN=your_invoice_api_token
INVOICE_DOMAIN=your_invoice_subdomain
```
4. **Initialize services**
```python
from src.config import create_client_service
client_service = create_client_service()
```
5. **Add a client**
```python
from src.model.client import Client
from datetime import date

client_service.add_client(Client(
    name="John Doe",
    email="john@example.com",
    insurance_company="ACME Insurance",
    car_model="Toyota Corolla",
    car_year=2020,
    price=1500,
    next_payment=date(2025, 7, 10)
))
```
6. **Send reminders and cleanup**
```python
client_service.notify_payment_due_in_days(days_ahead=1)
removed = client_service.remove_overdue_clients(overdue_days=3)
print("Removed clients:", removed)
```
7. **Generate a monthly report**
```python
report = client_service.generate_monthly_report()
print(report)
```
8. **Start background scheduler**
```python
from src.jobs.scheduler import create_scheduler
scheduler = create_scheduler(days_ahead=1, overdue_days=3)
scheduler.start()
```
9. **Excel Styling**

- Header: Bold, green background
- Row: Calibri, light green background
- Overdue Payment: Light red background
- Styles can be customized via CellStyle dictionaries in config.py.

10. **Tests & Coverage**

- coverage 100% ✅
- poetry run pytest --cov=src --cov-report=html
- View HTML coverage report online:
https://damiankowalczykdk.github.io/Insurance-Client-Manager/htmlcov/index.html

11. **License**
- License © 2025 Damian Kowalczyk