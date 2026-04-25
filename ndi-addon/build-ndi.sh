#!/bin/bash
echo ""
echo "  ✝  AnchorCast NDI Addon Builder (macOS/Linux)"
echo "  -----------------------------------------------"
echo ""

cd "$(dirname "$0")"

NDI_INCLUDE=""
NDI_LIB=""

if [ "$(uname)" = "Darwin" ]; then
    NDI_INCLUDE="/Library/NDI SDK for Apple/include"
    NDI_LIB="/Library/NDI SDK for Apple/lib/macOS"
    if [ ! -f "$NDI_INCLUDE/Processing.NDI.Lib.h" ]; then
        echo "  ERROR: NDI SDK for Apple not found at /Library/NDI SDK for Apple/"
        echo "  Download it free from: https://ndi.video/for-developers/ndi-sdk/"
        echo ""
        exit 1
    fi
    echo "  NDI SDK for Apple found."
    # Create symlink without spaces for node-gyp compatibility
    NDI_LINK="/tmp/ndi-sdk-apple"
    rm -f "$NDI_LINK"
    ln -sf "/Library/NDI SDK for Apple" "$NDI_LINK"
    echo "  Created build symlink at $NDI_LINK"
else
    if [ ! -f "/usr/include/Processing.NDI.Lib.h" ] && [ ! -f "/usr/local/include/Processing.NDI.Lib.h" ]; then
        echo "  WARNING: NDI SDK headers not found in /usr/include or /usr/local/include"
        echo "  Download from: https://ndi.video/for-developers/ndi-sdk/"
        echo ""
        exit 1
    fi
    echo "  NDI SDK found."
fi

echo "  Installing node-addon-api..."
npm install --ignore-scripts
if [ $? -ne 0 ]; then
    echo "  ERROR: npm install failed"
    exit 1
fi

ELECTRON_VER=$(node -e "try{console.log(require('../node_modules/electron/package.json').version)}catch(e){console.log('')}" 2>/dev/null)

if [ -n "$ELECTRON_VER" ]; then
    echo "  Building native addon for Electron $ELECTRON_VER..."
    node-gyp rebuild --target="$ELECTRON_VER" --arch=$(uname -m) --dist-url=https://electronjs.org/headers
else
    echo "  Building native addon for system Node.js..."
    node-gyp rebuild
fi

if [ $? -ne 0 ]; then
    echo ""
    echo "  Build failed. Common fixes:"
    echo "    1. Install Xcode Command Line Tools: xcode-select --install"
    echo "    2. Install node-gyp globally: npm install -g node-gyp"
    echo "    3. Check NDI SDK path in binding.gyp"
    exit 1
fi

echo ""
echo "  ✓ NDI addon built successfully!"
echo "  File: build/Release/ndi_sender.node"
echo ""
echo "  Restart AnchorCast — NDI output will now use the official NDI SDK."
echo "  In OBS/vMix: look for \"AnchorCast\" in NDI sources."
echo ""
