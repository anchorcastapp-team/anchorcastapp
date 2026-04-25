#!/bin/bash
echo "SermonCast — Bible Data Downloader"
echo "===================================="
echo ""
echo "Downloading complete KJV Bible (31,102 verses)..."

URL="https://raw.githubusercontent.com/scrollmapper/bible_databases/master/json/t_kjv.json"

if command -v curl &> /dev/null; then
    curl -L -o kjv.json "$URL"
elif command -v wget &> /dev/null; then
    wget -O kjv.json "$URL"
else
    echo "ERROR: Neither curl nor wget found."
    echo "Please manually download from:"
    echo "https://github.com/scrollmapper/bible_databases/blob/master/json/t_kjv.json"
    echo "Save as kjv.json in this folder."
    exit 1
fi

if [ -f kjv.json ]; then
    LINES=$(python3 -c "import json; d=json.load(open('kjv.json')); print(len(d))" 2>/dev/null || echo "unknown")
    echo ""
    echo "SUCCESS! kjv.json downloaded ($LINES verses)."
    echo "Restart SermonCast to load the complete Bible."
else
    echo "ERROR: Download failed."
fi
