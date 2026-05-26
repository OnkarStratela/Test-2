#!/bin/bash

# Fix USB permissions if needed
if [ -e /dev/ttyACM0 ] && [ ! -r /dev/ttyACM0 ]; then
    echo "[GC] Setting USB permissions..."
    sudo chmod 666 /dev/ttyACM0
fi

chmod +x compile_standard.sh
./compile_standard.sh

if [ $? -eq 0 ]; then
    ./rfid_standard "$@"
fi
