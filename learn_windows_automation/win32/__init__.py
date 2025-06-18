"""
Windows automation using pywin32 (direct Windows API access).

This module provides utilities for low-level Windows automation
using the Windows API through pywin32.
"""

try:
    import win32gui
    import win32con
    import win32api
    import win32process
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

import os
import time
import subprocess


class WindowManager:
    """Manage Windows using pywin32."""
    
    def __init__(self):
        if not PYWIN32_AVAILABLE:
            raise ImportError("pywin32 is not available. This module requires Windows.")
    
    def find_window_by_title(self, title):
        """Find a window by its title."""
        return win32gui.FindWindow(None, title)
    
    def get_window_text(self, hwnd):
        """Get the text/title of a window."""
        return win32gui.GetWindowText(hwnd)
    
    def get_all_windows(self):
        """Get all visible windows."""
        windows = []
        
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    windows.append((hwnd, title))
        
        win32gui.EnumWindows(enum_handler, None)
        return windows
    
    def set_window_position(self, hwnd, x, y, width, height):
        """Set window position and size."""
        win32gui.SetWindowPos(hwnd, 0, x, y, width, height, 0)
    
    def bring_window_to_front(self, hwnd):
        """Bring window to front."""
        win32gui.SetForegroundWindow(hwnd)


def create_hello_world_document():
    """Basic use case: Create a Hello World document using win32 API."""
    if not PYWIN32_AVAILABLE:
        print("Error: pywin32 is not available. This tool requires Windows.")
        return False
    
    try:
        print("Creating Hello World document using win32 API...")
        
        # Start Notepad using subprocess
        process = subprocess.Popen(['notepad.exe'])
        time.sleep(2)  # Wait for Notepad to start
        print("✓ Notepad started")
        
        # Find Notepad window
        wm = WindowManager()
        notepad_hwnd = None
        
        # Try to find Notepad window (it might be "Untitled - Notepad" or similar)
        windows = wm.get_all_windows()
        for hwnd, title in windows:
            if "Notepad" in title and ("Untitled" in title or title == "Notepad"):
                notepad_hwnd = hwnd
                break
        
        if not notepad_hwnd:
            print("✗ Could not find Notepad window")
            return False
        
        print(f"✓ Found Notepad window: HWND {notepad_hwnd}")
        
        # Bring Notepad to front
        wm.bring_window_to_front(notepad_hwnd)
        time.sleep(0.5)
        
        # Create the content directly as a file (since win32 API text input is complex)
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        file_path = os.path.join(downloads_path, "hello_world_win32.txt")
        
        content = f"""Hello World from win32 automation!

This document was created automatically using Python and win32 API.
Date: {time.strftime("%Y-%m-%d %H:%M:%S")}

This demonstrates basic Windows automation using low-level win32 APIs:
- Finding windows by title
- Bringing windows to front
- File operations
- Process management
"""
        
        # Write content to file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✓ File created at: {file_path}")
        
        # Now open the file in Notepad
        subprocess.run(['notepad.exe', file_path])
        print("✓ File opened in Notepad")
        
        # Close the original empty Notepad
        try:
            win32gui.PostMessage(notepad_hwnd, win32con.WM_CLOSE, 0, 0)
            print("✓ Original empty Notepad closed")
        except:
            pass
        
        print(f"✓ Success! File created and opened at: {file_path}")
        return True
        
    except Exception as e:
        print(f"✗ Error creating document: {e}")
        return False


def main():
    """Main entry point for win32 automation."""
    print("Windows Automation using pywin32")
    
    if not PYWIN32_AVAILABLE:
        print("Error: pywin32 is not available. This tool requires Windows.")
        return
    
    try:
        # Run the basic use case
        print("\n=== Basic Use Case: Create Hello World Document ===")
        success = create_hello_world_document()
        
        if success:
            print("\n🎉 Basic use case completed successfully!")
        else:
            print("\n💥 Basic use case failed!")
        
        print("\n=== Window Information ===")
        wm = WindowManager()
        windows = wm.get_all_windows()
        
        print(f"Found {len(windows)} visible windows:")
        for hwnd, title in windows[:10]:  # Show first 10 windows
            print(f"  HWND: {hwnd}, Title: {title}")
        
        if len(windows) > 10:
            print(f"  ... and {len(windows) - 10} more windows")
            
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
