#!/usr/bin/env python3
"""
Parse Excel cells with rich text formatting using LibreOffice UNO API.
This provides character-level color access for legacy Excel files.

Requirements:
- LibreOffice installed
- uno Python package (comes with LibreOffice)

Usage:
1. Start LibreOffice in headless mode with socket:
   soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"

2. Run this script:
   python tech_libreoffice.py
"""

import os
import time
from typing import Any

from pydantic import BaseModel

# Import UNO - this comes with LibreOffice (NOT from PyPI)
try:
    import uno
    from com.sun.star.beans import PropertyValue
except ImportError as e:
    print("=" * 80)
    print("ERROR: UNO not found. LibreOffice's Python bridge is required.")
    print("=" * 80)
    print("\nLibreOffice must be installed, and you need to use its Python:")
    print("\nOn Ubuntu/Debian:")
    print("  sudo apt install libreoffice python3-uno")
    print("\nOn macOS:")
    print("  brew install libreoffice")
    print("  # Then use LibreOffice's Python:")
    print("  /Applications/LibreOffice.app/Contents/Resources/python this_script.py")
    print("\nOn Arch Linux:")
    print("  sudo pacman -S libreoffice-fresh")
    print("\nAlternatively, add LibreOffice's Python to PYTHONPATH:")
    print("  export PYTHONPATH=/path/to/libreoffice/python:$PYTHONPATH")
    print("=" * 80)
    raise SystemExit(1) from e


class Segment(BaseModel):
    """Represents a text segment with RGB color information."""
    r: int
    g: int
    b: int
    text: str
    is_default_color: bool = False
    
    @property
    def is_black(self) -> bool:
        """Check if color is black (with some leeway)."""
        return self.r <= 30 and self.g <= 30 and self.b <= 30
    
    @property
    def is_red(self) -> bool:
        """Check if color is red (with some leeway)."""
        return self.r > 200 and self.g < 80 and self.b < 80
    
    @property
    def is_blue(self) -> bool:
        """Check if color is blue (with some leeway)."""
        return self.r < 80 and self.g < 80 and self.b > 200


class Cell(BaseModel):
    """Represents an Excel cell with rich text segments."""
    cell_number: str
    color_groups: list[Segment]


def rgb_from_long(color_long: int) -> tuple[int, int, int]:
    """Convert LibreOffice color (long integer) to RGB tuple."""
    # LibreOffice stores colors as BGR in a long integer
    b = (color_long >> 16) & 0xFF
    g = (color_long >> 8) & 0xFF
    r = color_long & 0xFF
    return (r, g, b)


def connect_to_libreoffice(host: str = "localhost", port: int = 2002) -> Any:
    """
    Connect to a running LibreOffice instance.
    
    Returns:
        Component context for LibreOffice
    """
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    
    url = f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
    
    try:
        context = resolver.resolve(url)
        return context
    except Exception as e:
        raise ConnectionError(
            f"Could not connect to LibreOffice at {host}:{port}. "
            "Make sure LibreOffice is running with:\n"
            f"soffice --headless --accept=\"socket,host={host},port={port};urp;StarOffice.ComponentContext\""
        ) from e


def start_libreoffice_headless(port: int = 2002, timeout: int = 10) -> bool:
    """
    Start LibreOffice in headless mode with socket connection.
    
    Returns:
        True if LibreOffice started successfully, False otherwise
    """
    print(f"  Starting LibreOffice on port {port}...")
    
    # Try different soffice locations
    soffice_paths = [
        'soffice',  # In PATH
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',  # macOS
        '/usr/bin/soffice',  # Linux
        '/usr/local/bin/soffice',  # Linux alternative
    ]
    
    soffice_cmd = None
    for path in soffice_paths:
        if os.path.exists(path) or path == 'soffice':
            soffice_cmd = path
            break
    
    if not soffice_cmd:
        print("  ERROR: Could not find soffice executable")
        return False
    
    cmd = f'{soffice_cmd} --headless --accept="socket,host=localhost,port={port};urp;StarOffice.ComponentContext" &'
    print(f"  Running: {cmd}")
    os.system(cmd)
    
    # Wait for LibreOffice to start
    print(f"  Waiting for LibreOffice to start (timeout: {timeout}s)...", end="", flush=True)
    for i in range(timeout):
        time.sleep(1)
        print(".", end="", flush=True)
        # Check if process is running
        result = os.popen("ps aux | grep soffice | grep -v grep").read()
        if result:
            print(" Started!")
            return True
    
    print(" Timeout!")
    return False


def load_spreadsheet(context: Any, file_path: str) -> Any:
    """
    Load a spreadsheet file using LibreOffice.
    
    Args:
        context: LibreOffice component context
        file_path: Path to the Excel file
    
    Returns:
        Spreadsheet document object
    """
    smgr = context.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    
    # Convert to absolute path and file URL
    abs_path = os.path.abspath(file_path)
    file_url = uno.systemPathToFileUrl(abs_path)
    
    # Load properties - Hidden=True for headless operation
    properties = (
        PropertyValue("Hidden", 0, True, 0),
        PropertyValue("ReadOnly", 0, True, 0),
    )
    
    document = desktop.loadComponentFromURL(file_url, "_blank", 0, properties)
    return document


def parse_cell_rich_text(cell: Any, cell_ref: str, show_progress: bool = True) -> Cell:
    """
    Parse rich text from a LibreOffice cell with character-level color extraction.
    
    Args:
        cell: LibreOffice cell object
        cell_ref: Cell reference (e.g., "A1")
        show_progress: Whether to show character-by-character progress
    
    Returns:
        Cell object with color segments
    """
    segments: list[Segment] = []
    
    # Get cell text
    cell_text = cell.getString()
    if not cell_text:
        return Cell(cell_number=cell_ref, color_groups=[])
    
    if show_progress:
        print(f"    Processing {len(cell_text)} characters...", end="", flush=True)
    
    # Access the cell's text cursor for character-level formatting
    text = cell.getText()
    cursor = text.createTextCursor()
    
    current_color: tuple[int, int, int] | None = None
    current_text = ""
    
    # Iterate through each character
    for i in range(len(cell_text)):
        # Show progress every 100 characters
        if show_progress and i > 0 and i % 100 == 0:
            print(f"\r    Processing {len(cell_text)} characters... {i}/{len(cell_text)} ({i*100//len(cell_text)}%)", end="", flush=True)
        
        # Move cursor to character position
        cursor.gotoStart(False)
        cursor.goRight(i, False)
        cursor.goRight(1, True)  # Select one character
        
        # Get character color
        try:
            char_color_long = cursor.getPropertyValue("CharColor")
            char_color = rgb_from_long(char_color_long)
            is_default = False
        except Exception:
            # If we can't get color, use black as default
            char_color = (0, 0, 0)
            is_default = True
        
        char = cell_text[i]
        
        # Check if color changed
        if current_color is None:
            current_color = char_color
            current_text = char
        elif current_color == char_color:
            current_text += char
        else:
            # Color changed, save current segment
            segments.append(Segment(
                r=current_color[0],
                g=current_color[1],
                b=current_color[2],
                text=current_text,
                is_default_color=is_default
            ))
            current_color = char_color
            current_text = char
    
    if show_progress:
        print(f"\r    Processing {len(cell_text)} characters... Done!     ")
    
    # Add final segment
    if current_text:
        segments.append(Segment(
            r=current_color[0] if current_color else 0,
            g=current_color[1] if current_color else 0,
            b=current_color[2] if current_color else 0,
            text=current_text,
            is_default_color=is_default
        ))
    
    return Cell(cell_number=cell_ref, color_groups=segments)


def get_cell_reference(col: int, row: int) -> str:
    """Convert column and row numbers to Excel-style reference (e.g., A1)."""
    col_letter = ""
    col_num = col + 1  # LibreOffice uses 0-based indexing
    
    while col_num > 0:
        col_num -= 1
        col_letter = chr(65 + (col_num % 26)) + col_letter
        col_num //= 26
    
    return f"{col_letter}{row + 1}"


def main(file_path: str = "Book.xlsx", auto_start: bool = True) -> list[Cell]:
    """
    Main function to parse Excel file using LibreOffice.
    
    Args:
        file_path: Path to Excel file
        auto_start: Whether to automatically start LibreOffice
    
    Returns:
        List of Cell objects with rich text segments
    """
    print(f"Parsing {file_path} using LibreOffice UNO...")
    
    # Start LibreOffice if requested
    if auto_start:
        print("Starting LibreOffice in headless mode...")
        if not start_libreoffice_headless():
            print("\nFailed to start LibreOffice automatically.")
            print("Please start it manually:")
            print('  soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" &')
            print("\nOr on macOS:")
            print('  /Applications/LibreOffice.app/Contents/MacOS/soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" &')
            raise RuntimeError("Failed to start LibreOffice")
    
    # Connect to LibreOffice
    print("\nConnecting to LibreOffice...")
    try:
        context = connect_to_libreoffice()
        print("  Connected successfully!")
    except Exception as e:
        print(f"  Connection failed: {e}")
        print("\nTroubleshooting:")
        print("1. Make sure LibreOffice is running:")
        print("   ps aux | grep soffice")
        print("2. Try starting it manually:")
        print('   /Applications/LibreOffice.app/Contents/MacOS/soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" &')
        raise
    
    # Load spreadsheet
    print(f"Loading {file_path}...")
    document = load_spreadsheet(context, file_path)
    
    # Get first sheet
    sheets = document.getSheets()
    sheet = sheets.getByIndex(0)
    
    print(f"\nWorksheet: {sheet.getName()}")
    print("=" * 80)
    
    cells: list[Cell] = []
    
    # Iterate through used range
    # Get the used area
    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)
    cursor.gotoEndOfUsedArea(True)
    
    used_range = sheet.getCellRangeByPosition(
        cursor.getRangeAddress().StartColumn,
        cursor.getRangeAddress().StartRow,
        cursor.getRangeAddress().EndColumn,
        cursor.getRangeAddress().EndRow
    )
    
    # Count total cells with text for progress tracking
    total_cells = 0
    for row_idx in range(used_range.getRangeAddress().StartRow, 
                         used_range.getRangeAddress().EndRow + 1):
        for col_idx in range(used_range.getRangeAddress().StartColumn,
                            used_range.getRangeAddress().EndColumn + 1):
            cell = sheet.getCellByPosition(col_idx, row_idx)
            if cell.getString():
                total_cells += 1
    
    print(f"\nFound {total_cells} cells with text to process")
    print("=" * 80)
    
    # Iterate through cells
    processed_cells = 0
    for row_idx in range(used_range.getRangeAddress().StartRow, 
                         used_range.getRangeAddress().EndRow + 1):
        for col_idx in range(used_range.getRangeAddress().StartColumn,
                            used_range.getRangeAddress().EndColumn + 1):
            cell = sheet.getCellByPosition(col_idx, row_idx)
            cell_text = cell.getString()
            
            if cell_text:
                processed_cells += 1
                cell_ref = get_cell_reference(col_idx, row_idx)
                
                print(f"\n[{processed_cells}/{total_cells}] Cell {cell_ref}:")
                cell_obj = parse_cell_rich_text(cell, cell_ref, show_progress=True)
                cells.append(cell_obj)
                
                # Print cell information
                print(f"  Total segments: {len(cell_obj.color_groups)}")
                
                for i, segment in enumerate(cell_obj.color_groups, 1):
                    print(f"\n  Segment {i}:")
                    print(f"    RGB: ({segment.r}, {segment.g}, {segment.b})")
                    print(f"    Text: {segment.text[:50]}{'...' if len(segment.text) > 50 else ''}")
                    print(f"    is_default_color: {segment.is_default_color}")
                    print(f"    is_black: {segment.is_black}")
                    print(f"    is_red: {segment.is_red}")
                    print(f"    is_blue: {segment.is_blue}")
                
                print()
    
    print("=" * 80)
    print(f"\nTotal cells processed: {len(cells)}")
    
    # Close document
    document.close(True)
    
    return cells


if __name__ == "__main__":
    import sys
    
    file_path = sys.argv[1] if len(sys.argv) > 1 else "Book.xlsx"
    
    try:
        cells = main(file_path, auto_start=True)
    except ConnectionError as e:
        print(f"\n{e}")
        print("\nTo manually start LibreOffice, run:")
        print('soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"')
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
