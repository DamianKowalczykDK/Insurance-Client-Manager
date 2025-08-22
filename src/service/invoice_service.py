from src.model.invoice import InvoiceDict
import httpx
import json


class InvoiceService:
    """Service for creating, retrieving, and updating invoices via the Fakturownia API."""

    def __init__(self, api_token: str, domain: str):
        """Initialize the InvoiceService with API credentials.

        Args:
            api_token: API token for authentication.
            domain: Subdomain of the Fakturownia account (used in the API URL).
        """
        self.api_token = api_token
        self.api_url = f"https://{domain}.fakturownia.pl/invoices.json"
        self.headers = {'Content-Type': 'application/json'}

    def create_invoice(self, data: InvoiceDict) -> str:
        """Create a new invoice for a client.

        Args:
            data: Dictionary containing invoice details (InvoiceDict).

        Returns:
            str: URL to view the created invoice.

        Raises:
            httpx.HTTPStatusError: If the API request fails.
        """
        payload = {
            "api_token": self.api_token,
            "invoice": {
                "kind": "vat",
                "number": None,
                "sell_date": "2025-06-16",
                "issue_date": "2025-06-16",
                "payment_to": "2025-06-23",
                "seller_name": "Damian Kowalczyk",
                "buyer_name": data["client_name"],
                "buyer_tax_no": "6272616681",
                "positions": [
                    {"name": data["item_name"], "tax": 23, "total_price_gross": data["item_price"], "quantity": 1}
                ]
            }
        }

        with httpx.Client(timeout=10.0) as client:
            response = client.post(
                self.api_url,
                headers=self.headers,
                json=payload
            )

        response.raise_for_status()
        result = response.json()
        invoice_url = result.get("view_url", "N/A")
        return invoice_url

    def get_invoice(self, number: int = 5) -> None:
        """Retrieve a list of invoices from the API.

        Args:
            number: Number of latest invoices to fetch.

        Raises:
            httpx.HTTPStatusError: If the API request fails.
        """
        url = self.api_url
        params: dict[str, str | int] = {
            "api_token": self.api_token,
            "sort": "desc",
            "page": 1,
            "per_page": number,
        }

        with httpx.Client(timeout=10.0) as client:
            response = client.get(url, params=params)
            response.raise_for_status()
            print(json.dumps(response.json(), indent=4, ensure_ascii=False))

    def update_invoice(self, invoice_id: int) -> None:
        """Update an existing invoice by its ID.

        Args:
            invoice_id: ID of the invoice to update.

        Raises:
            httpx.HTTPStatusError: If the API request fails.
        """
        url = f"https://damiankowalczyk.fakturownia.pl/invoices/{invoice_id}.json"
        params = {
            "api_token": self.api_token,
        }

        data = {
            "invoice": {
                "seller_name": "Damian Kowalczyk",
                "buyer_name": "Arkadiusz Piotr.",
                "positions": [{
                    "total_price_gross": 1000.0,
                    "tax": 23
                }]}
        }

        with httpx.Client(timeout=10.0) as client:
            response = client.put(url, params=params, json=data)
            response.raise_for_status()
            print(json.dumps(response.json(), indent=4, ensure_ascii=False))
