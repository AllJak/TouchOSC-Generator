#!/bin/zsh

cd -- "$(dirname "$0")" || exit 1

PYTHON_BIN=$(command -v python3)

# prüfen ob python3 existiert
if [ -z "$PYTHON_BIN" ]; then
    echo "Python3 wurde nicht gefunden."
    exit 1
fi

# prüfen ob openpyxl installiert ist
$PYTHON_BIN -c "import openpyxl" 2>/dev/null

# wenn nicht installiert → installieren
if [ $? -ne 0 ]; then
    echo "Installiere openpyxl..."
    $PYTHON_BIN -m pip install openpyxl
fi

# Python Script starten
$PYTHON_BIN "Skript_V6-5.py"

exit
