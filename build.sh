#!/usr/bin/env bash
# Build script for Render - installs LibreOffice then Python packages
set -e

echo "Installing LibreOffice..."
apt-get update -qq && apt-get install -y -qq libreoffice libreoffice-calc libreoffice-writer

echo "Installing Python packages..."
pip install -r requirements.txt

echo "Build complete!"
