#!/bin/bash

echo "[GC Build] Compiling rfid_standard..."

if [ ! -d "SRC" ]; then
    echo "[GC Build] ERROR: SRC directory not found!"
    exit 1
fi

if [ ! -f "rfid_standard.c" ]; then
    echo "[GC Build] ERROR: rfid_standard.c not found!"
    exit 1
fi

gcc \
  rfid_standard.c \
  SRC/host.c SRC/CAENRFIDLib_Light.c SRC/IO_Light.c \
  -ISRC \
  -o rfid_standard \
  -lpthread -lm \
  -Wall

if [ $? -eq 0 ]; then
    chmod +x rfid_standard
    echo "[GC Build] Success! Run with: ./rfid_standard"
else
    echo "[GC Build] Compilation failed!"
    exit 1
fi
