#!/usr/bin/env python3
"""
Vans Department Monthly Reporting App

A desktop application for automating monthly resource allocation reports
for the Vans department. Fetches data from Monday.com, calculates time
allocations with custom rules, and generates PowerPoint presentations.
"""

import sys
import tkinter as tk
from gui.main_window import launch_gui


def main():
    """Main entry point for the application."""
    try:
        launch_gui()
    except KeyboardInterrupt:
        print("\\nApplication terminated by user.")
        sys.exit(0)
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
