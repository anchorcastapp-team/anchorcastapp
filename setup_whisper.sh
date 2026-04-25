#!/bin/bash
echo ""
echo " ✝  SermonCast — Whisper Setup"
echo " --------------------------------"
echo " Installs local speech recognition (offline, no API costs)"
echo ""

# Check Python
if ! command -v python3 &>/dev/null; then
    echo " ERROR: Python 3 not found."
    echo " Install with: brew install python3  (macOS)"
    echo "               sudo apt install python3 python3-pip  (Ubuntu)"
    exit 1
fi

echo " Installing faster-whisper..."
pip3 install faster-whisper --quiet

if [ $? -ne 0 ]; then
    echo " ERROR: pip install failed."
    echo " Try: pip3 install faster-whisper"
    exit 1
fi

echo ""
echo " Downloading Whisper base model (~74 MB)..."
python3 -c "from faster_whisper import WhisperModel; WhisperModel('base', device='cpu', compute_type='int8'); print('Model ready!')"

if [ $? -ne 0 ]; then
    echo " ERROR: Model download failed."
    exit 1
fi

echo ""
echo " ✓ Whisper setup complete!"
echo " Restart SermonCast — transcription works offline now."
echo ""
