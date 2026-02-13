"""
Setup script for creating standalone macOS application using py2app.

Usage:
    python3 setup.py py2app
"""

from setuptools import setup

APP = ['run_app.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'packages': ['tkinter', 'urllib', 'json'],
    'includes': [
        'token_manager',
        'data.models',
        'data.monday_client',
        'data.calculator',
        'gui.main_window',
        'gui.data_preview',
        'powerpoint.generator',
        'powerpoint.slides',
    ],
    'excludes': ['matplotlib', 'numpy', 'pandas', 'scipy'],
    'plist': {
        'CFBundleName': 'Vans Reporter',
        'CFBundleDisplayName': 'Vans Reporter',
        'CFBundleGetInfoString': 'Generate monthly reports for Vans department',
        'CFBundleIdentifier': 'com.monday.vansreporter',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': '© 2026 Monday.com',
        'LSMinimumSystemVersion': '10.13.0',
    }
}

setup(
    name='Vans Reporter',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
