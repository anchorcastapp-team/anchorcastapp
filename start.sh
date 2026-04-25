#!/usr/bin/env bash
# SermonCast — Quick Start Script
# Run this once after unzipping: bash start.sh

set -e

echo ""
echo "  ✝  SermonCast — AI Sermon Display"
echo "  ─────────────────────────────────"
echo ""

# Check Node.js
if ! command -v node &> /dev/null; then
  echo "  ✗ Node.js not found."
  echo "  Please install Node.js 18+ from https://nodejs.org"
  echo "  Then run this script again."
  exit 1
fi

NODE_VER=$(node -v | cut -d'v' -f2 | cut -d'.' -f1)
if [ "$NODE_VER" -lt 18 ]; then
  echo "  ✗ Node.js version $NODE_VER detected. Version 18+ required."
  echo "  Please update at https://nodejs.org"
  exit 1
fi

echo "  ✓ Node.js $(node -v) detected"
echo ""

# Install dependencies
if [ ! -d "node_modules" ]; then
  echo "  Installing dependencies (first run only, may take 1-2 minutes)..."
  npm install --silent
  echo "  ✓ Dependencies installed"
else
  echo "  ✓ Dependencies already installed"
fi

echo ""
echo "  Launching SermonCast..."
echo ""
npm start
