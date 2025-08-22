from typing import TypedDict


class InvoiceDict(TypedDict):
    """Typed dictionary representation of an invoice.

    Attributes:
        client_name: Full name of the client.
        client_email: Email address of the client.
        client_tax_no: Tax identification number of the client.
        item_name: Name of the billed item or service.
        item_quantity: Quantity of the item or service.
        item_price: Price per item or service unit (in the smallest currency unit, e.g. cents).
    """
    client_name: str
    client_email: str
    client_tax_no: str
    item_name: str
    item_quantity: int
    item_price: int
