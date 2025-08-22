from openpyxl.styles import Font, Alignment, Border, PatternFill, Side
from typing import TypedDict, Literal
from openpyxl.cell.cell import Cell


class BorderStyle(TypedDict, total=False):
    """Typed dictionary defining a border style for Excel cells.

    Attributes:
        style: Border line style. Possible values include
            'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double',
            'hair', 'medium', 'mediumDashDot', 'mediumDashDotDot',
            'mediumDashed', 'slantDashDot', 'thick', 'thin'.
        color: Hex color code for the border (e.g., "000000" for black).
    """
    style: Literal[
        'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair',
        'medium', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed',
        'slantDashDot', 'thick', 'thin'
    ]
    color: str


class CellStyle(TypedDict, total=False):
    """Typed dictionary defining a style for an Excel cell.

    Attributes:
        font: Font style applied to the cell.
        fill: Fill pattern or background color for the cell.
        alignment: Text alignment inside the cell.
        border_sides: Dictionary mapping border sides
            ('left', 'right', 'top', 'bottom') to a BorderStyle.
    """
    font: Font
    fill: PatternFill
    alignment: Alignment
    border_sides: dict[str, BorderStyle]


class SheetStyle(TypedDict, total=False):
    """Typed dictionary defining styles for an entire Excel sheet.

    Attributes:
        header: CellStyle applied to header cells.
        row: CellStyle applied to regular row cells.
    """
    header: CellStyle
    row: CellStyle


def apply_style(cell: Cell, style: CellStyle) -> None:
    """Apply a given style to an Excel cell.

    Args:
        cell: The Excel cell to apply the style to.
        style: Dictionary of style attributes to apply. May include
            font, fill, alignment, and border_sides.
    """
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
