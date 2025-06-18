"""
Tests for pywinauto automation module.

These tests use mocking to avoid requiring a Windows environment.
"""

import unittest
from unittest.mock import Mock, patch, MagicMock
import sys


class TestApplicationManager(unittest.TestCase):
    """Test ApplicationManager class."""
    
    def setUp(self):
        # Mock pywinauto modules and classes
        self.pywinauto_mock = Mock()
        self.application_mock = Mock()
        self.desktop_mock = Mock()
        self.findwindows_mock = Mock()
        
        # Setup the mock structure
        self.application_class_mock = Mock(return_value=self.application_mock)
        self.desktop_class_mock = Mock(return_value=self.desktop_mock)
        
        # Create patches
        self.patches = [
            patch.dict('sys.modules', {
                'pywinauto': self.pywinauto_mock,
                'pywinauto.findwindows': self.findwindows_mock
            }),
            patch('learn_windows_automation.pywinauto.PYWINAUTO_AVAILABLE', True),
            patch('learn_windows_automation.pywinauto.Application', self.application_class_mock),
            patch('learn_windows_automation.pywinauto.Desktop', self.desktop_class_mock),
            patch('learn_windows_automation.pywinauto.find_elements', self.findwindows_mock.find_elements)
        ]
        
        # Start all patches
        for p in self.patches:
            p.start()
        
        # Import the module after mocking
        from learn_windows_automation.pywinauto import ApplicationManager
        self.ApplicationManager = ApplicationManager
    
    def tearDown(self):
        # Stop all patches
        for p in self.patches:
            p.stop()
    
    def test_application_manager_initialization(self):
        """Test ApplicationManager can be initialized when pywinauto is available."""
        am = self.ApplicationManager()
        self.assertIsInstance(am, self.ApplicationManager)
    
    def test_connect_by_title(self):
        """Test connecting to application by title."""
        # Setup mock
        mock_app = Mock()
        self.application_mock.connect.return_value = mock_app
        
        am = self.ApplicationManager()
        result = am.connect_by_title("Test Application")
        
        self.application_mock.connect.assert_called_once_with(title="Test Application")
        self.assertEqual(result, mock_app)
    
    def test_connect_by_process(self):
        """Test connecting to application by process ID."""
        # Setup mock
        mock_app = Mock()
        self.application_mock.connect.return_value = mock_app
        
        am = self.ApplicationManager()
        result = am.connect_by_process(1234)
        
        self.application_mock.connect.assert_called_once_with(process=1234)
        self.assertEqual(result, mock_app)
    
    def test_start_application(self):
        """Test starting an application."""
        # Setup mock
        mock_app = Mock()
        self.application_mock.start.return_value = mock_app
        
        am = self.ApplicationManager()
        result = am.start_application("notepad.exe")
        
        self.application_mock.start.assert_called_once_with("notepad.exe")
        self.assertEqual(result, mock_app)
    
    def test_get_desktop_windows(self):
        """Test getting desktop windows."""
        # Setup mock windows
        mock_window1 = Mock()
        mock_window1.handle = 12345
        mock_window1.window_text.return_value = "Window 1"
        mock_window1.is_visible.return_value = True
        
        mock_window2 = Mock()
        mock_window2.handle = 67890
        mock_window2.window_text.return_value = "Window 2"
        mock_window2.is_visible.return_value = True
        
        self.desktop_mock.windows.return_value = [mock_window1, mock_window2]
        
        am = self.ApplicationManager()
        result = am.get_desktop_windows()
        
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0], (12345, "Window 1"))
        self.assertEqual(result[1], (67890, "Window 2"))
    
    def test_find_window_by_class(self):
        """Test finding window by class name."""
        # Setup mock
        mock_elements = [Mock(), Mock()]
        self.findwindows_mock.find_elements.return_value = mock_elements
        
        am = self.ApplicationManager()
        result = am.find_window_by_class("TestClass")
        
        self.findwindows_mock.find_elements.assert_called_once_with(class_name="TestClass")
        self.assertEqual(result, mock_elements)


class TestWindowController(unittest.TestCase):
    """Test WindowController class."""
    
    def setUp(self):
        # Mock pywinauto modules
        self.pywinauto_mock = Mock()
        sys.modules['pywinauto'] = self.pywinauto_mock
        sys.modules['pywinauto.findwindows'] = Mock()
        
        # Import the module after mocking
        from learn_windows_automation.pywinauto import WindowController
        self.WindowController = WindowController
        
        # Create a mock application
        self.mock_app = Mock()
        self.controller = self.WindowController(self.mock_app)
    
    def tearDown(self):
        # Clean up mocks
        for module in ['pywinauto', 'pywinauto.findwindows']:
            if module in sys.modules:
                del sys.modules[module]
    
    def test_get_main_window(self):
        """Test getting main window."""
        mock_window = Mock()
        self.mock_app.top_window.return_value = mock_window
        
        result = self.controller.get_main_window()
        
        self.mock_app.top_window.assert_called_once()
        self.assertEqual(result, mock_window)
    
    def test_click_button(self):
        """Test clicking a button."""
        mock_window = Mock()
        mock_button = Mock()
        mock_window.child_window.return_value = mock_button
        
        self.controller.click_button(mock_window, "OK")
        
        mock_window.child_window.assert_called_once_with(
            title="OK", control_type="Button"
        )
        mock_button.click.assert_called_once()
    
    def test_type_text(self):
        """Test typing text into a control."""
        mock_window = Mock()
        mock_control = Mock()
        mock_window.child_window.return_value = mock_control
        
        self.controller.type_text(mock_window, "textbox1", "Hello World")
        
        mock_window.child_window.assert_called_once_with(auto_id="textbox1")
        mock_control.type_keys.assert_called_once_with("Hello World")
    
    def test_get_text(self):
        """Test getting text from a control."""
        mock_window = Mock()
        mock_control = Mock()
        mock_control.get_value.return_value = "Control Text"
        mock_window.child_window.return_value = mock_control
        
        result = self.controller.get_text(mock_window, "textbox1")
        
        mock_window.child_window.assert_called_once_with(auto_id="textbox1")
        mock_control.get_value.assert_called_once()
        self.assertEqual(result, "Control Text")


class TestPywinautoMain(unittest.TestCase):
    """Test main function."""
    
    @patch('learn_windows_automation.pywinauto.PYWINAUTO_AVAILABLE', False)
    @patch('builtins.print')
    def test_main_without_pywinauto(self, mock_print):
        """Test main function when pywinauto is not available."""
        from learn_windows_automation.pywinauto import main
        main()
        
        mock_print.assert_any_call("Windows Automation using pywinauto")
        mock_print.assert_any_call("Error: pywinauto is not available.")


if __name__ == '__main__':
    unittest.main()
