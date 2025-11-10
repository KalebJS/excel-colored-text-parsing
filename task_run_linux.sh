#!/bin/bash
# Helper script to run LibreOffice UNO parser on Linux

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo "LibreOffice UNO Parser - Linux"
echo "==============================="
echo ""

# Check if LibreOffice is installed
if ! command -v soffice &> /dev/null; then
    echo -e "${RED}Error: LibreOffice not found${NC}"
    echo ""
    echo "Install with:"
    echo "  Ubuntu/Debian: sudo apt install libreoffice python3-uno"
    echo "  Arch: sudo pacman -S libreoffice-fresh"
    echo "  Fedora: sudo dnf install libreoffice python3-uno"
    exit 1
fi

# Check if python3-uno is available
if ! python3 -c "import uno" 2>/dev/null; then
    echo -e "${RED}Error: python3-uno not found${NC}"
    echo ""
    echo "Install with:"
    echo "  Ubuntu/Debian: sudo apt install python3-uno"
    echo "  Arch: sudo pacman -S libreoffice-fresh"
    echo "  Fedora: sudo dnf install python3-uno"
    exit 1
fi

echo -e "${GREEN}✓ LibreOffice and python3-uno found${NC}"

# Check if pydantic is installed
if ! python3 -c "import pydantic" 2>/dev/null; then
    echo -e "${YELLOW}Warning: pydantic not found, installing...${NC}"
    pip3 install pydantic
fi

# Check if LibreOffice is already running
if pgrep -f "soffice.*accept.*2002" > /dev/null; then
    echo -e "${GREEN}✓ LibreOffice is already running${NC}"
else
    echo "Starting LibreOffice in headless mode..."
    soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" --nofirststartwizard &
    
    # Wait for LibreOffice to start
    echo -n "Waiting for LibreOffice to start"
    for i in {1..10}; do
        sleep 1
        echo -n "."
        if pgrep -f "soffice.*accept.*2002" > /dev/null; then
            echo -e " ${GREEN}Started!${NC}"
            break
        fi
    done
    
    if ! pgrep -f "soffice.*accept.*2002" > /dev/null; then
        echo -e " ${RED}Failed!${NC}"
        echo "Could not start LibreOffice"
        exit 1
    fi
fi

echo ""

# Get file path from argument or use default
FILE_PATH="${1:-Book.xlsx}"

if [ ! -f "$FILE_PATH" ]; then
    echo -e "${RED}Error: File not found: $FILE_PATH${NC}"
    exit 1
fi

echo "Parsing file: $FILE_PATH"
echo ""

# Run the parser
python3 parse_rich_text_libreoffice.py "$FILE_PATH"

echo ""
echo "Done!"
echo ""
echo "To stop LibreOffice:"
echo "  killall soffice.bin"
