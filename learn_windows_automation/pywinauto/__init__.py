"""
Windows automation using pywinauto (high-level automation).

This module provides utilities for high-level Windows automation
using pywinauto for easier interaction with desktop applications.
"""

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import find_elements
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False


class ApplicationManager:
    """Manage applications using pywinauto."""
    
    def __init__(self):
        if not PYWINAUTO_AVAILABLE:
            raise ImportError("pywinauto is not available.")
    
    def connect_by_title(self, title):
        """Connect to an application by window title."""
        app = Application().connect(title=title)
        return app
    
    def connect_by_process(self, process_id):
        """Connect to an application by process ID."""
        app = Application().connect(process=process_id)
        return app
    
    def start_application(self, path):
        """Start an application."""
        app = Application().start(path)
        return app
    
    def get_desktop_windows(self):
        """Get all desktop windows."""
        desktop = Desktop(backend="uia")
        windows = desktop.windows()
        return [(w.handle, w.window_text()) for w in windows if w.is_visible()]
    
    def find_window_by_class(self, class_name):
        """Find windows by class name."""
        elements = find_elements(class_name=class_name)
        return elements


class WindowController:
    """Control windows using pywinauto."""
    
    def __init__(self, app):
        self.app = app
    
    def get_main_window(self):
        """Get the main application window."""
        return self.app.top_window()
    
    def click_button(self, window, button_text):
        """Click a button by text."""
        button = window.child_window(title=button_text, control_type="Button")
        button.click()
    
    def type_text(self, window, control_id, text):
        """Type text into a control."""
        control = window.child_window(auto_id=control_id)
        control.type_keys(text)
    
    def get_text(self, window, control_id):
        """Get text from a control."""
        control = window.child_window(auto_id=control_id)
        return control.get_value()


def main():
    """Main entry point for pywinauto automation."""
    print("Windows Automation using pywinauto")
    
    if not PYWINAUTO_AVAILABLE:
        print("Error: pywinauto is not available.")
        return
    
    try:
        app_manager = ApplicationManager()
        windows = app_manager.get_desktop_windows()
        
        print(f"Found {len(windows)} desktop windows:")
        for handle, title in windows[:10]:  # Show first 10 windows
            print(f"  Handle: {handle}, Title: {title}")
        
        if len(windows) > 10:
            print(f"  ... and {len(windows) - 10} more windows")
            
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
