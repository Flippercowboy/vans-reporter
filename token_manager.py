#!/usr/bin/env python3
"""
Token Manager for Vans Reporter App

Handles secure storage and retrieval of Monday.com API token
in user's home directory (~/.vans_reporter_config)
"""

import os
import json
import tkinter as tk
from tkinter import messagebox, simpledialog
from pathlib import Path


CONFIG_FILE = Path.home() / ".vans_reporter_config"


def get_token():
    """
    Get the Monday.com API token.

    Checks the config file first. If not found, prompts the user
    to enter it via a dialog.

    Returns:
        str: The API token, or None if user cancels
    """
    # Try to load from config file
    token = _load_token_from_file()

    if token:
        return token

    # No token found, prompt user
    token = _prompt_for_token()

    if token:
        _save_token_to_file(token)

    return token


def _load_token_from_file():
    """Load token from config file if it exists."""
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                return config.get('monday_api_token')
    except Exception as e:
        print(f"Warning: Could not load token from config file: {e}")

    return None


def _save_token_to_file(token):
    """Save token to config file."""
    try:
        config = {'monday_api_token': token}
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)

        # Set file permissions to user-only (0600)
        os.chmod(CONFIG_FILE, 0o600)
        print(f"Token saved to {CONFIG_FILE}")
    except Exception as e:
        print(f"Warning: Could not save token to config file: {e}")


def _prompt_for_token():
    """Show dialog to prompt user for API token."""
    # Create hidden root window
    root = tk.Tk()
    root.withdraw()

    # Show instructions first
    message = """Welcome to Vans Reporter!

To use this app, you need a Monday.com API token.

How to get your token:
1. Go to https://monday.com
2. Click your profile picture (bottom left)
3. Select "Developers" → "My Access Tokens"
4. Click "Show" next to your token and copy it

Please paste your token in the next dialog."""

    messagebox.showinfo("Monday.com API Token Required", message)

    # Prompt for token
    token = simpledialog.askstring(
        "Enter API Token",
        "Paste your Monday.com API token:",
        show='*'  # Hide token characters
    )

    root.destroy()

    if token and token.strip():
        return token.strip()

    return None


def reset_token():
    """Delete the saved token (for testing or reconfiguration)."""
    try:
        if CONFIG_FILE.exists():
            CONFIG_FILE.unlink()
            print(f"Token removed from {CONFIG_FILE}")
            return True
    except Exception as e:
        print(f"Error removing token: {e}")

    return False


if __name__ == "__main__":
    # Test the token manager
    print("Testing token manager...")
    token = get_token()
    if token:
        print(f"Token retrieved: {token[:10]}...")
    else:
        print("No token provided")
