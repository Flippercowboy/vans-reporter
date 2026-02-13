#!/usr/bin/env python3
"""
Launch script for Vans Department Monthly Reporting App.

Run this script to start the application:
    python3 run_app.py
"""

import sys
from pathlib import Path

# Add parent directory to path so we can import the package
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))

# Import and launch the application
from vans_reporting_app.gui.main_window import launch_gui

if __name__ == "__main__":
    try:
        launch_gui()
    except KeyboardInterrupt:
        print("\nApplication terminated by user.")
        sys.exit(0)
    except Exception as e:
        print(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
