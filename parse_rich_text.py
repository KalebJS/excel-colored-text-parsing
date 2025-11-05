#!/usr/bin/env python3
"""
Parse Excel cells with rich text formatting and extract color information.
"""
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from typing import Any, Literal

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell as OpenpyxlCell
from openpyxl.cell.rich_text import CellRichText
from openpyxl.styles.fonts import Font
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

# Type alias for cell values based on openpyxl's internal types
CellValue = (
    float
    | Decimal
    | str
    | CellRichText
    | datetime
    | date
    | time
    | timedelta
    | Literal[True]
    | Any  # For formula types and other edge cases
    | None
)


class Segment(BaseModel):
    """Represents a text segment with RGB color information."""
    r: int
    g: int
    b: int
    text: str
    
    @property
    def is_black(self) -> bool:
        """Check if color is black (with some leeway)."""
        # Allow up to 30 in each channel to be considered "black"
        return self.r <= 30 and self.g <= 30 and self.b <= 30
    
    @property
    def is_red(self) -> bool:
        """Check if color is red (with some leeway)."""
        # Red should be high (>200), green and blue should be low (<80)
        return self.r > 200 and self.g < 80 and self.b < 80
    
    @property
    def is_blue(self) -> bool:
        """Check if color is blue (with some leeway)."""
        # Blue should be high (>200), red and green should be low (<80)
        return self.r < 80 and self.g < 80 and self.b > 200


class Cell(BaseModel):
    """Represents an Excel cell with rich text segments."""
    cell_number: str
    color_groups: list[Segment]


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert hex color string to RGB tuple."""
    # Remove any leading '#' if present
    hex_color = hex_color.lstrip('#')
    
    # Ensure we have 6 characters
    if len(hex_color) != 6:
        return (0, 0, 0)
    
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    
    return (r, g, b)


def get_color_hex(font: Font | None, default_color: str = '000000') -> str:
    """Extract color hex from font object."""
    if not font or not font.color:
        return default_color
    
    color_obj = font.color
    
    # RGB format (ARGB - first 2 chars are alpha channel)
    if color_obj.rgb:
        rgb_value = str(color_obj.rgb)
        # Strip alpha channel (first 2 chars) to get RGB
        return rgb_value[2:] if len(rgb_value) > 6 else rgb_value
    
    # Theme or indexed colors would need additional lookup
    # For now, return default
    return default_color


def parse_rich_text_cell(cell_value: CellValue, cell_ref: str, default_color: str = '000000') -> Cell:
    """
    Parse a cell with rich text and return a Cell object with color segments.
    """
    segments: list[Segment] = []
    
    if cell_value is None:
        return Cell(cell_number=cell_ref, color_groups=[])
    
    if not isinstance(cell_value, CellRichText):
        # Plain text cell - no rich formatting
        r, g, b = hex_to_rgb(default_color)
        segment = Segment(r=r, g=g, b=b, text=str(cell_value))
        return Cell(cell_number=cell_ref, color_groups=[segment])
    
    # Process rich text segments
    for item in cell_value:
        if isinstance(item, str):
            # Unformatted text segment
            r, g, b = hex_to_rgb(default_color)
            segments.append(Segment(r=r, g=g, b=b, text=item))
        else:
            # Formatted text segment with font styling (TextBlock)
            text = item.text
            color_hex = get_color_hex(item.font, default_color)
            r, g, b = hex_to_rgb(color_hex)
            segments.append(Segment(r=r, g=g, b=b, text=text))
    
    return Cell(cell_number=cell_ref, color_groups=segments)


def main() -> list[Cell]:
    # Load workbook with rich text support enabled
    print("Loading Book.xlsx with rich text support...")
    wb = load_workbook('Book.xlsx', rich_text=True)
    ws = wb.active
    
    if not isinstance(ws, Worksheet):
        raise ValueError("Active sheet is not a valid worksheet")
    
    print(f"\nWorksheet: {ws.title}")
    print("=" * 80)
    
    cells: list[Cell] = []
    
    # Iterate through all rows and cells
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell, OpenpyxlCell):
                continue
            if cell.value:
                # Parse the cell into a Cell object
                cell_ref = f"{cell.column_letter}{cell.row}"
                cell_obj = parse_rich_text_cell(cell.value, cell_ref)
                cells.append(cell_obj)
                
                # Print cell information
                print(f"\nCell {cell_obj.cell_number}:")
                print(f"  Total segments: {len(cell_obj.color_groups)}")
                
                for i, segment in enumerate(cell_obj.color_groups, 1):
                    print(f"\n  Segment {i}:")
                    print(f"    RGB: ({segment.r}, {segment.g}, {segment.b})")
                    print(f"    Text: {segment.text[:50]}{'...' if len(segment.text) > 50 else ''}")
                    print(f"    is_black: {segment.is_black}")
                    print(f"    is_red: {segment.is_red}")
                    print(f"    is_blue: {segment.is_blue}")
                
                print()
    
    print("=" * 80)
    print(f"\nTotal cells processed: {len(cells)}")
    
    return cells


if __name__ == "__main__":
    cells = main()
