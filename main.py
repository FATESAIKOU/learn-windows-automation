"""
Windows Automation Toolkit

This is the main entry point for the Windows automation toolkit.
Choose between win32 (low-level) or pywinauto (high-level) automation.
"""

import sys
import argparse
from learn_windows_automation.win32 import main as win32_main
from learn_windows_automation.pywinauto import main as pywinauto_main


def main():
    parser = argparse.ArgumentParser(description="Windows Automation Toolkit")
    parser.add_argument(
        "backend", 
        choices=["win32", "pywinauto"],
        help="Choose automation backend: win32 (low-level) or pywinauto (high-level)"
    )
    
    args = parser.parse_args()
    
    if args.backend == "win32":
        print("Starting Win32 automation...")
        win32_main()
    elif args.backend == "pywinauto":
        print("Starting pywinauto automation...")
        pywinauto_main()


if __name__ == "__main__":
    main()
