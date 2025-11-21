#!/usr/bin/env python3
"""
Robocorp task to parse Excel files with LibreOffice UNO.
"""

import os
import sys
from pathlib import Path

# Import the LibreOffice parser
from tech_libreoffice import main as parse_excel


def parse_excel_task():
    """
    Main Robocorp task to parse Excel file with rich text formatting.
    """
    print("=" * 80)
    print("Robocorp Task: Parse Excel with LibreOffice")
    print("=" * 80)
    
    # Detect if running in Docker
    in_docker = os.path.exists('/.dockerenv')
    
    # Configuration
    file_path = os.getenv("EXCEL_FILE", "Book.xlsx")
    
    if in_docker:
        print("\n🐳 Running inside Docker container")
        host = "127.0.0.1"
        auto_start = False  # LibreOffice already running via Dockerfile CMD
        print(f"   Connecting to existing LibreOffice at {host}:2002")
    else:
        print("\n💻 Running locally")
        host = "localhost"
        auto_start = True  # Start LibreOffice locally
        print(f"   Will start LibreOffice at {host}:2002")
    
    print(f"\n📄 Excel file: {file_path}")
    
    # Check if file exists
    if not Path(file_path).exists():
        print(f"\n❌ Error: File not found: {file_path}")
        sys.exit(1)
    
    try:
        # Parse the Excel file
        cells = parse_excel(
            file_path=file_path,
            auto_start=auto_start,
            host=host,
            port=2002
        )
        
        print("\n" + "=" * 80)
        print(f"✅ Successfully parsed {len(cells)} cells")
        print("=" * 80)
        
        # Save results to output directory
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        
        output_file = output_dir / "parsed_cells.txt"
        with open(output_file, "w") as f:
            f.write(f"Parsed {len(cells)} cells from {file_path}\n\n")
            for cell in cells:
                f.write(f"Cell {cell.cell_number}:\n")
                for segment in cell.color_groups:
                    f.write(f"  RGB({segment.r}, {segment.g}, {segment.b}): {segment.text}\n")
                f.write("\n")
        
        print(f"\n📝 Results saved to: {output_file}")
        
        return cells
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    parse_excel_task()
