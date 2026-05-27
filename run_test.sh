#!/bin/bash
# One-shot runner for the beer-pour RFID verification test logger.
#
# Builds rfid_gc_live from rfid_gc_live.c + the SRC/ CAEN library, sanity-
# checks USB permissions on /dev/ttyACM0, then launches beer_pour_logger.py.

set -u

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo "===== Beer-Pour RFID Verification Test Logger ====="
echo ""

echo -e "${YELLOW}Checking for CAEN library files in SRC folder...${NC}"

if [ ! -d "SRC" ]; then
    echo -e "${RED}SRC folder not found!${NC}"
    echo "Please create SRC folder and copy CAEN library files there."
    exit 1
fi

REQUIRED_FILES=(
    "SRC/CAENRFIDLib_Light.c"
    "SRC/CAENRFIDLib_Light.h"
    "SRC/CAENRFIDTypes_Light.h"
    "SRC/IO_Light.c"
    "SRC/IO_Light.h"
    "SRC/Protocol_Light.h"
    "SRC/host.c"
    "SRC/host.h"
)

MISSING=()
for f in "${REQUIRED_FILES[@]}"; do
    [ -f "$f" ] || MISSING+=("$f")
done

if [ ${#MISSING[@]} -ne 0 ]; then
    echo -e "${RED}Missing required CAEN library files:${NC}"
    for f in "${MISSING[@]}"; do
        echo "  - $f"
    done
    exit 1
fi
echo -e "${GREEN}All required CAEN files found.${NC}"

echo -e "${YELLOW}Checking USB device access...${NC}"
if [ -e /dev/ttyACM0 ]; then
    if [ ! -r /dev/ttyACM0 ] || [ ! -w /dev/ttyACM0 ]; then
        echo -e "${YELLOW}USB device found but no read/write permission. Fixing...${NC}"
        sudo chmod 666 /dev/ttyACM0
    fi
    echo -e "${GREEN}USB device access OK (/dev/ttyACM0).${NC}"
else
    echo -e "${YELLOW}No CAEN RFID reader detected on /dev/ttyACM0.${NC}"
    echo "Plug in the CAEN reader (USB) and retry."
fi

echo -e "${YELLOW}Compiling rfid_gc_live...${NC}"
chmod +x compile_gc.sh 2>/dev/null
./compile_gc.sh
if [ $? -ne 0 ]; then
    echo -e "${RED}Compilation failed. See error messages above.${NC}"
    exit 1
fi
echo -e "${GREEN}Compilation successful.${NC}"

echo ""
echo -e "${YELLOW}Checking Python dependencies (openpyxl, Pillow)...${NC}"
if ! python3 -c "import openpyxl, PIL" 2>/dev/null; then
    echo -e "${YELLOW}Missing one of openpyxl / Pillow. Install with:${NC}"
    echo "  sudo apt install -y python3-openpyxl python3-pil"
    echo "  # or, on systems without those apt packages:"
    echo "  pip3 install --break-system-packages -r requirements.txt"
    exit 1
fi
echo -e "${GREEN}Python deps OK.${NC}"

echo ""
echo -e "${GREEN}Launching beer-pour test logger...${NC}"
echo ""

exec python3 beer_pour_logger.py
