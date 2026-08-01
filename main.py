"""
VBOX Test Suite File Renamer
Main entry point for the application.

This application safely renames VBOX test files by updating their serial prefix
based on data from an Excel spreadsheet.

Usage:
    python main.py
"""

from ui import VBOXRenamerUI
from PyQt5.QtWidgets import QApplication
import sys


def main():
    """Launch the application."""
    app = QApplication(sys.argv)
    window = VBOXRenamerUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
