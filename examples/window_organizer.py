#!/usr/bin/env python3
"""Window organizer example script."""

import sys
from pathlib import Path

# Add the project source to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from win_automation.utils import WindowsUtils


def main() -> None:
    """Main function for window organization."""
    print("Window Organizer Example")
    print("This script would organize windows on the screen")
    print("Arguments received:", sys.argv[1:])
    
    # For now, just demonstrate clipboard functionality
    print("\nTesting clipboard functionality:")
    try:
        # Get current clipboard content
        current_text = WindowsUtils.get_clipboard_text()
        print(f"Current clipboard: {current_text}")
        
        # Set a test message
        test_message = "Window organizer was here!"
        success = WindowsUtils.set_clipboard_text(test_message)
        if success:
            print(f"Set clipboard to: {test_message}")
        else:
            print("Failed to set clipboard")
            
    except Exception as e:
        print(f"Error accessing clipboard: {e}")


if __name__ == "__main__":
    main()
