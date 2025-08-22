from typing import TypedDict

class MonthlyReportDict(TypedDict):
    """Typed dictionary representation of a monthly report.

    Attributes:
        month: Month of the report (e.g., "2025-08").
        company: Dictionary mapping company names to integer values (e.g., number of clients).
        gross_total: Total gross amount for the month.
        net_total: Total net amount for the month.
    """
    month: str
    company: dict[str, int]
    gross_total: int
    net_total: int
