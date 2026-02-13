#!/bin/bash
# Helper script to run the app with your Monday.com API token

# Usage:
#   ./run_with_token.sh YOUR_API_TOKEN
# Or edit this file and set TOKEN="..." then just run ./run_with_token.sh

if [ -z "$1" ]; then
    # Check if TOKEN is set in this script
    if [ -z "$TOKEN" ]; then
        echo "Usage: ./run_with_token.sh YOUR_API_TOKEN"
        echo ""
        echo "Or edit this file and set: TOKEN=\"your_token_here\""
        echo ""
        echo "Get your token from:"
        echo "https://monday.com → Profile → Developers → My Access Tokens"
        exit 1
    fi
else
    TOKEN="$1"
fi

export MONDAY_API_TOKEN="$TOKEN"
python3 run_app.py
