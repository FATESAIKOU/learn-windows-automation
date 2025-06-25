#!/usr/bin/env python3
"""
Multi-App Coordinator - Advanced pywinauto example
Coordinates data flow between multiple applications: Notepad, Calculator, Browser
"""

import sys
import time
from pathlib import Path

# Add the src directory to the path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))


try:
    import pywinauto
    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False


class MultiAppCoordinator:
    """Coordinates operations across multiple Windows applications."""

    def __init__(self):
        """Initialize multi-app coordinator."""
        if not PYWINAUTO_AVAILABLE:
            raise RuntimeError("pywinauto not available for multi-app automation")

        self.desktop = Desktop(backend="uia")
        self.apps = {}
        self.windows = {}

    def launch_application(self, app_name: str, exe_path: str, window_title: str = None) -> None:
        """Launch an application and store reference."""
        try:
            print(f"Launching {app_name}...")
            app = Application(backend="uia").start(exe_path)
            self.apps[app_name] = app

            # Wait for window to appear
            time.sleep(2)

            # Find the main window
            if window_title:
                window = app.window(title_re=window_title)
            else:
                window = app.top_window()

            window.wait("visible", timeout=10)
            self.windows[app_name] = window
            print(f"{app_name} launched successfully")

        except Exception as e:
            print(f"Failed to launch {app_name}: {e}")
            raise

    def connect_to_application(self, app_name: str, process_name: str, window_title: str = None) -> None:
        """Connect to existing application."""
        try:
            print(f"Connecting to {app_name}...")
            app = Application(backend="uia").connect(process=process_name)
            self.apps[app_name] = app

            # Find the window
            if window_title:
                window = app.window(title_re=window_title)
            else:
                window = app.top_window()

            self.windows[app_name] = window
            print(f"Connected to {app_name} successfully")

        except Exception as e:
            print(f"Failed to connect to {app_name}: {e}")
            raise

    def focus_window(self, app_name: str) -> None:
        """Bring application window to front."""
        if app_name not in self.windows:
            raise ValueError(f"Application {app_name} not found")

        window = self.windows[app_name]
        window.set_focus()
        time.sleep(0.5)
        print(f"Focused on {app_name}")

    def send_text_to_app(self, app_name: str, text: str, control_name: str = None) -> None:
        """Send text to specific application."""
        if app_name not in self.windows:
            raise ValueError(f"Application {app_name} not found")

        self.focus_window(app_name)

        if control_name:
            # Send to specific control
            window = self.windows[app_name]
            control = window.child_window(auto_id=control_name)
            control.set_text(text)
        else:
            # Send to focused window
            send_keys(text)

        print(f"Sent text to {app_name}: {text[:50]}{'...' if len(text) > 50 else ''}")

    def get_text_from_app(self, app_name: str, control_name: str = None) -> str:
        """Get text from specific application."""
        if app_name not in self.windows:
            raise ValueError(f"Application {app_name} not found")

        self.focus_window(app_name)
        window = self.windows[app_name]

        if control_name:
            control = window.child_window(auto_id=control_name)
            text = control.get_value()
        else:
            # Get all text from window
            text = window.window_text()

        print(f"Retrieved text from {app_name}: {text[:50]}{'...' if len(text) > 50 else ''}")
        return text

    def click_button(self, app_name: str, button_name: str) -> None:
        """Click a button in specific application."""
        if app_name not in self.windows:
            raise ValueError(f"Application {app_name} not found")

        self.focus_window(app_name)
        window = self.windows[app_name]

        try:
            # Try to find button by name/text
            button = window.child_window(title=button_name, control_type="Button")
            button.click()
            print(f"Clicked button '{button_name}' in {app_name}")
        except Exception:
            try:
                # Try to find by auto_id
                button = window.child_window(auto_id=button_name)
                button.click()
                print(f"Clicked button '{button_name}' in {app_name}")
            except Exception as e:
                print(f"Failed to click button '{button_name}' in {app_name}: {e}")
                raise

    def copy_data_between_apps(self, source_app: str, target_app: str,
                              source_control: str = None, target_control: str = None) -> None:
        """Copy data from source app to target app using clipboard."""
        # Get data from source
        self.focus_window(source_app)

        if source_control:
            window = self.windows[source_app]
            control = window.child_window(auto_id=source_control)
            control.select_all()
        else:
            send_keys("^a")  # Ctrl+A

        send_keys("^c")  # Ctrl+C
        time.sleep(0.5)

        # Paste to target
        self.focus_window(target_app)

        if target_control:
            window = self.windows[target_app]
            control = window.child_window(auto_id=target_control)
            control.click_input()

        send_keys("^v")  # Ctrl+V
        time.sleep(0.5)

        print(f"Copied data from {source_app} to {target_app}")

    def arrange_windows(self, layout: str = "side_by_side") -> None:
        """Arrange application windows on screen."""
        if layout == "side_by_side":
            # Arrange windows side by side
            screen_width = self.desktop.rectangle().width()
            screen_height = self.desktop.rectangle().height()

            window_width = screen_width // len(self.windows)

            for i, (app_name, window) in enumerate(self.windows.items()):
                x = i * window_width
                y = 0
                width = window_width
                height = screen_height - 100  # Leave space for taskbar

                window.move_window(x, y, width, height)
                print(f"Positioned {app_name} at ({x}, {y}) with size ({width}, {height})")

        elif layout == "cascade":
            # Cascade windows
            offset = 50
            for i, (app_name, window) in enumerate(self.windows.items()):
                x = i * offset
                y = i * offset
                width = 800
                height = 600

                window.move_window(x, y, width, height)
                print(f"Cascaded {app_name} at ({x}, {y})")

    def close_application(self, app_name: str) -> None:
        """Close specific application."""
        if app_name in self.windows:
            try:
                window = self.windows[app_name]
                window.close()
                del self.windows[app_name]
                del self.apps[app_name]
                print(f"Closed {app_name}")
            except Exception as e:
                print(f"Error closing {app_name}: {e}")

    def close_all_applications(self) -> None:
        """Close all managed applications."""
        for app_name in list(self.windows.keys()):
            self.close_application(app_name)


def demo_notepad_calculator_workflow(coordinator: MultiAppCoordinator) -> None:
    """Demonstrate workflow between Notepad and Calculator."""
    print("=== Notepad & Calculator Demo ===")

    try:
        # Launch applications
        coordinator.launch_application("notepad", "notepad.exe", ".*Notepad")
        coordinator.launch_application("calculator", "calc.exe", "Calculator")

        # Arrange windows side by side
        coordinator.arrange_windows("side_by_side")

        # Work with Calculator
        coordinator.focus_window("calculator")
        time.sleep(1)

        # Perform calculation: 15 + 25 * 2
        coordinator.click_button("calculator", "num1Button")
        coordinator.click_button("calculator", "num5Button")
        coordinator.click_button("calculator", "plusButton")
        coordinator.click_button("calculator", "num2Button")
        coordinator.click_button("calculator", "num5Button")
        coordinator.click_button("calculator", "multiplyButton")
        coordinator.click_button("calculator", "num2Button")
        coordinator.click_button("calculator", "equalButton")

        time.sleep(1)

        # Copy result to clipboard
        send_keys("^c")

        # Switch to Notepad and create report
        coordinator.focus_window("notepad")

        report_text = """
Calculation Report
==================
Date: 2025-06-26
Calculation: 15 + 25 * 2
Result: """

        coordinator.send_text_to_app("notepad", report_text)

        # Paste the calculation result
        send_keys("^v")

        # Add conclusion
        conclusion = """

Analysis:
- The calculation follows order of operations
- Multiplication performed before addition
- Final result is mathematically correct
"""

        coordinator.send_text_to_app("notepad", conclusion)

        print("Demo completed successfully!")

    except Exception as e:
        print(f"Demo failed: {e}")


def demo_multi_window_data_flow(coordinator: MultiAppCoordinator) -> None:
    """Demonstrate data flow between multiple applications."""
    print("=== Multi-Window Data Flow Demo ===")

    try:
        # Launch applications
        coordinator.launch_application("notepad", "notepad.exe", ".*Notepad")

        # Create sample data in Notepad
        sample_data = """Product Sales Data:
- Laptop: $1200
- Mouse: $25
- Keyboard: $75
- Monitor: $300
- Total: $1600"""

        coordinator.send_text_to_app("notepad", sample_data)

        # Launch second Notepad instance for processing
        coordinator.launch_application("notepad2", "notepad.exe", ".*Notepad")

        # Copy data between applications
        coordinator.copy_data_between_apps("notepad", "notepad2")

        # Process data in second window
        processed_data = """

PROCESSED REPORT:
=================
The above data has been imported and processed.
Processing timestamp: 2025-06-26
Status: Data validation complete
"""

        coordinator.send_text_to_app("notepad2", processed_data)

        # Arrange windows for comparison
        coordinator.arrange_windows("side_by_side")

        print("Multi-window data flow demo completed!")

    except Exception as e:
        print(f"Data flow demo failed: {e}")


def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        print("Usage: python multi_app_coordinator.py <demo_type>")
        print("Demo types:")
        print("  calc_notepad    - Notepad & Calculator workflow")
        print("  data_flow       - Multi-window data flow")
        print("  custom          - Custom coordination (interactive)")
        return 1

    demo_type = sys.argv[1]
    coordinator = MultiAppCoordinator()

    try:
        if demo_type == "calc_notepad":
            demo_notepad_calculator_workflow(coordinator)
        elif demo_type == "data_flow":
            demo_multi_window_data_flow(coordinator)
        elif demo_type == "custom":
            print("Custom coordination mode - implement your own workflow here")
            print("Available methods:", [method for method in dir(coordinator) if not method.startswith('_')])
        else:
            print(f"Unknown demo type: {demo_type}")
            return 1

        # Wait for user input before cleanup
        input("Press Enter to close all applications and exit...")

    except Exception as e:
        print(f"Error running demo: {e}")
        return 1
    finally:
        # Clean up
        coordinator.close_all_applications()

    return 0


if __name__ == "__main__":
    sys.exit(main())
