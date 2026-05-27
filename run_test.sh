#!/bin/bash
# One-shot runner for the beer-pour RFID reliability test logger.
#
# Mirrors the antenna-position-test runner style: CAEN library check,
# USB permissions check, compile, then launch test_logger.py.

echo "===== Beer-pour RFID Test Logger ====="
echo ""

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo -e "${YELLOW}Checking for CAEN library files in SRC folder...${NC}"

if [ ! -d "SRC" ]; then
    echo -e "${RED}SRC folder not found!${NC}"
    echo "Please run this script from test/Test-2/, the folder that contains SRC/."
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

MISSING_FILES=()
for file in "${REQUIRED_FILES[@]}"; do
    if [ ! -f "$file" ]; then
        MISSING_FILES+=("$file")
    fi
done

if [ ${#MISSING_FILES[@]} -ne 0 ]; then
    echo -e "${RED}Missing required CAEN library files:${NC}"
    for file in "${MISSING_FILES[@]}"; do
        echo "  - $file"
    done
    exit 1
fi

echo -e "${GREEN}All required CAEN files found.${NC}"

echo -e "${YELLOW}Checking USB device access...${NC}"
if [ -e /dev/ttyACM0 ] || [ -e /dev/ttyUSB0 ]; then
    if [ ! -r /dev/ttyACM0 ] && [ ! -r /dev/ttyUSB0 ]; then
        echo -e "${YELLOW}USB device found but no read permission.${NC}"
        echo "Adding user to dialout group..."
        sudo usermod -a -G dialout "$USER"
        echo -e "${GREEN}User added to dialout group. Please log out and back in, then re-run.${NC}"
        exit 1
    else
        echo -e "${GREEN}USB device access OK.${NC}"
    fi
else
    echo -e "${YELLOW}No CAEN RFID reader detected on /dev/ttyACM0 or /dev/ttyUSB0.${NC}"
    echo "Plug it in and re-run."
fi

echo -e "${YELLOW}Setting permissions on shell scripts...${NC}"
chmod +x compile_gc.sh 2>/dev/null

echo -e "${YELLOW}Compiling RFID reader...${NC}"
./compile_gc.sh
if [ $? -ne 0 ]; then
    echo -e "${RED}Compilation failed. See errors above.${NC}"
    exit 1
fi
echo -e "${GREEN}Compilation successful!${NC}"

echo -e "${YELLOW}Checking Python dependencies (openpyxl, Pillow)...${NC}"
if ! python3 -c "import openpyxl, PIL" 2>/dev/null; then
    echo -e "${YELLOW}Installing python3-openpyxl and python3-pil via apt ...${NC}"
    sudo apt update
    sudo apt install -y python3-openpyxl python3-pil
    if ! python3 -c "import openpyxl, PIL" 2>/dev/null; then
        echo -e "${RED}Could not import openpyxl/PIL even after apt install.${NC}"
        echo "Try (last resort): sudo pip3 install --break-system-packages -r requirements.txt"
        exit 1
    fi
fi
echo -e "${GREEN}Python deps OK.${NC}"

echo ""
echo -e "${GREEN}Launching beer-pour test logger...${NC}"
echo ""
python3 test_logger.py
