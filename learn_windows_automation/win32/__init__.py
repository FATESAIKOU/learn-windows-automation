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


def main():
    """Main entry point for win32 automation."""
    print("Windows Automation using pywin32")
    
    if not PYWIN32_AVAILABLE:
        print("Error: pywin32 is not available. This tool requires Windows.")
        return
    
    try:
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
