#!/bin/bash
echo "============================================"
echo "  PDF Splitter - Starting..."
echo "============================================"

# Install requirements
echo "Installing required packages..."
pip install -r requirements.txt

echo ""
echo "Starting server... Open http://127.0.0.1:5000 in your browser"
echo "Press CTRL+C to stop the server"
echo ""

python3.10 main.py
