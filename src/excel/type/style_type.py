from openpyxl.styles import Font, Alignment, Border, PatternFill, Side
from typing import TypedDict, Literal
from openpyxl.cell.cell import Cell


class BorderStyle(TypedDict, total=False):
    style: Literal['dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'medium', 'mediumDashDot',
    'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin']
    color: str


class CellStyle(TypedDict, total=False):
    font: Font
    fill: PatternFill
    alignment: Alignment
    border_sides: dict[str, BorderStyle]


class SheetStyle(TypedDict, total=False):
    header: CellStyle
    row: CellStyle


def apply_style(cell: Cell, style: CellStyle) -> None:
    if "font" in style:
        cell.font = style["font"]
    if "fill" in style:
        cell.fill = style["fill"]
    if "alignment" in style:
        cell.alignment = style["alignment"]
    if "border_sides" in style:
        sides: dict[str, Side] = {}
        for direction, border_style in style["border_sides"].items():
            sides[direction] = Side(
                style=border_style.get("style"),
                color=border_style.get("color", "000000")
            )

        cell.border = Border(
            left=sides.get("left", Side()),
            right=sides.get("right", Side()),
            top=sides.get("top", Side()),
            bottom=sides.get("bottom", Side())
        )
