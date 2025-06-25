#!/usr/bin/env python3
"""Simple clipboard manipulation example script."""

import sys
from pathlib import Path

# Add the project source to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from win_automation.utils import WindowsUtils


def main() -> None:
    """Main function for clipboard manipulation."""
    if len(sys.argv) < 2:
        print("Usage: simple_clipboard.py <command> [text]")
        print("Commands:")
        print("  get - Get text from clipboard")
        print("  set <text> - Set text to clipboard")
        return

    command = sys.argv[1].lower()

    try:
        if command == "get":
            text = WindowsUtils.get_clipboard_text()
            if text:
                print(f"Clipboard content: {text}")
            else:
                print("Clipboard is empty or contains non-text data")

        elif command == "set":
            if len(sys.argv) < 3:
                print("Error: Please provide text to set")
                return

            text = " ".join(sys.argv[2:])
            success = WindowsUtils.set_clipboard_text(text)
            if success:
                print(f"Successfully set clipboard to: {text}")
            else:
                print("Failed to set clipboard text")

        else:
            print(f"Unknown command: {command}")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
