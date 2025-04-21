#!/usr/bin/env python
"""
Development mode installation script for DocxWriter.

This script installs DocxWriter in development mode, making it easier
to test changes without having to reinstall the package each time.
"""

import os
import sys
import subprocess
import platform


def get_python_executable():
    """Get the correct Python executable to use."""
    # If we're in a virtual environment, use that
    if hasattr(sys, "real_prefix") or (
        hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix
    ):
        if platform.system() == "Windows":
            return os.path.join(sys.prefix, "Scripts", "python.exe")
        else:
            return os.path.join(sys.prefix, "bin", "python")

    # Otherwise use the current Python executable
    return sys.executable


def install_dev():
    """Install the package in development mode."""
    python = get_python_executable()

    print(f"Installing DocxWriter in development mode using {python}...")

    # Install package in development mode
    try:
        subprocess.check_call([python, "-m", "pip", "install", "-e", "."])
        print("\nDocxWriter installed successfully in development mode!")
        print("\nYou can now use it as a module:")
        print("  python -m docxwriter -h")
        print("\nOr import it in your code:")
        print("  from docxwriter.document_creator import DocumentCreator")
    except subprocess.CalledProcessError as e:
        print(f"\nError installing package: {e}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(install_dev())
