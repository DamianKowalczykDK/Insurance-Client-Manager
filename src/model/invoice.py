from typing import TypedDict

class InvoiceDict(TypedDict):
    client_name: str
    client_email: str
    client_tax_no: str
    item_name: str
    item_quantity: int
    item_price: int