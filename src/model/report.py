from typing import TypedDict

class MonthlyReportDict(TypedDict):
    month: str
    company: dict[str, int]
    gross_total: int
    net_total: int