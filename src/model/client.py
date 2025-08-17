from dataclasses import dataclass
from typing import TypedDict
from datetime import date

class ClientDict(TypedDict):
    name: str
    email: str
    insurance_company: str
    car_model: str
    car_year: int
    price: int
    next_payment: str

@dataclass
class Client:
    name: str
    email: str
    insurance_company: str
    car_model: str
    car_year: int
    price: int
    next_payment: date

    def to_dict(self) -> ClientDict:
        return {
            "name": self.name,
            "email": self.email,
            "insurance_company": self.insurance_company,
            "car_model": self.car_model,
            "car_year": self.car_year,
            "price": self.price,
            "next_payment": self.next_payment.strftime("%Y-%m-%d"),
        }