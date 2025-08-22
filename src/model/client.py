from dataclasses import dataclass
from typing import TypedDict
from datetime import date


class ClientDict(TypedDict):
    """Typed dictionary representation of a client.

    Attributes:
        name: Full name of the client.
        email: Email address of the client.
        insurance_company: Name of the insurance company.
        car_model: Model of the client's car.
        car_year: Year of manufacture of the car.
        price: Insurance price for the client.
        next_payment: Next payment date in ISO format (YYYY-MM-DD).
    """
    name: str
    email: str
    insurance_company: str
    car_model: str
    car_year: int
    price: int
    next_payment: str


@dataclass
class Client:
    """Dataclass representing a client and their insurance details.

    Attributes:
        name: Full name of the client.
        email: Email address of the client.
        insurance_company: Name of the insurance company.
        car_model: Model of the client's car.
        car_year: Year of manufacture of the car.
        price: Insurance price for the client.
        next_payment: Next payment date as a datetime.date object.
    """
    name: str
    email: str
    insurance_company: str
    car_model: str
    car_year: int
    price: int
    next_payment: date

    def to_dict(self) -> ClientDict:
        """Convert the Client instance into a dictionary representation.

        Returns:
            ClientDict: A dictionary with string and int values.
                The next_payment field is formatted as YYYY-MM-DD.
        """
        return {
            "name": self.name,
            "email": self.email,
            "insurance_company": self.insurance_company,
            "car_model": self.car_model,
            "car_year": self.car_year,
            "price": self.price,
            "next_payment": self.next_payment.strftime("%Y-%m-%d"),
        }
