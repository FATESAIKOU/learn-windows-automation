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


def create_hello_world_document():
    """Basic use case: Create a Hello World document using Notepad."""
    if not PYWINAUTO_AVAILABLE:
        print("Error: pywinauto is not available.")
        return False
    
    try:
        print("Creating Hello World document using pywinauto...")
        
        # Start Notepad
        app = Application().start("notepad.exe")
        print("✓ Notepad started")
        
        # Get the main window
        notepad = app.UntitledNotepad  # Notepad's default window
        
        # Type Hello World text
        notepad.Edit.type_keys("Hello World from pywinauto automation!\n\n")
        notepad.Edit.type_keys("This document was created automatically using Python and pywinauto.\n")
        notepad.Edit.type_keys("Date: {}\n".format(time.strftime("%Y-%m-%d %H:%M:%S")))
        print("✓ Text typed into Notepad")
        
        # Save the file
        notepad.menu_select("File->Save As")
        time.sleep(1)  # Wait for Save As dialog
        
        # Get Save As dialog
        save_dialog = app.SaveAs
        
        # Get Downloads folder path
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        file_path = os.path.join(downloads_path, "hello_world_pywinauto.txt")
        
        # Type the file path
        save_dialog.ComboBox.type_keys(file_path)
        print(f"✓ File path set to: {file_path}")
        
        # Click Save button
        save_dialog.Save.click()
        time.sleep(1)  # Wait for save to complete
        print("✓ File saved successfully")
        
        # Close Notepad
        notepad.close()
        print("✓ Notepad closed")
        
        # Verify file was created
        if os.path.exists(file_path):
            print(f"✓ Success! File created at: {file_path}")
            return True
        else:
            print("✗ File was not created successfully")
            return False
            
    except Exception as e:
        print(f"✗ Error creating document: {e}")
        try:
            # Try to close Notepad if it's still open
            app = Application().connect(title_re=".*Notepad")
            app.top_window().close()
        except:
            pass
        return False


def main():
    """Main entry point for pywinauto automation."""
    print("Windows Automation using pywinauto")
    
    if not PYWINAUTO_AVAILABLE:
        print("Error: pywinauto is not available.")
        return
    
    try:
        # Run the basic use case
        print("\n=== Basic Use Case: Create Hello World Document ===")
        success = create_hello_world_document()
        
        if success:
            print("\n🎉 Basic use case completed successfully!")
        else:
            print("\n💥 Basic use case failed!")
        
        print("\n=== Desktop Windows Information ===")
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
