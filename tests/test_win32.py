"""
Tests for win32 automation module.

These tests use mocking to avoid requiring a Windows environment.
"""

import unittest
from unittest.mock import Mock, patch, MagicMock
import sys


class TestWindowManager(unittest.TestCase):
    """Test WindowManager class."""
    
    def setUp(self):
        # Mock pywin32 modules
        self.win32gui_mock = Mock()
        self.win32con_mock = Mock()
        self.win32api_mock = Mock()
        self.win32process_mock = Mock()
        
        # Create the module patches
        self.patches = [
            patch.dict('sys.modules', {
                'win32gui': self.win32gui_mock,
                'win32con': self.win32con_mock,
                'win32api': self.win32api_mock,
                'win32process': self.win32process_mock
            }),
            patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
        ]
        
        # Start all patches
        for p in self.patches:
            p.start()
        
        # Import the module after mocking
        from learn_windows_automation.win32 import WindowManager
        self.WindowManager = WindowManager
    
    def tearDown(self):
        # Stop all patches
        for p in self.patches:
            p.stop()
    
    def test_window_manager_initialization(self):
        """Test WindowManager can be initialized when pywin32 is available."""
        wm = self.WindowManager()
        self.assertIsInstance(wm, self.WindowManager)
    
    def test_find_window_by_title(self):
        """Test finding window by title."""
        # Setup mock
        self.win32gui_mock.FindWindow.return_value = 12345
        
        wm = self.WindowManager()
        result = wm.find_window_by_title("Test Window")
        
        self.win32gui_mock.FindWindow.assert_called_once_with(None, "Test Window")
        self.assertEqual(result, 12345)
    
    def test_get_window_text(self):
        """Test getting window text."""
        # Setup mock
        self.win32gui_mock.GetWindowText.return_value = "Test Window Title"
        
        wm = self.WindowManager()
        result = wm.get_window_text(12345)
        
        self.win32gui_mock.GetWindowText.assert_called_once_with(12345)
        self.assertEqual(result, "Test Window Title")
    
    def test_get_all_windows(self):
        """Test getting all windows."""
        # Setup mock for EnumWindows
        def mock_enum_windows(callback, ctx):
            # Simulate calling the callback with some window handles
            callback(12345, ctx)  # Visible window
            callback(67890, ctx)  # Another visible window
        
        self.win32gui_mock.EnumWindows.side_effect = mock_enum_windows
        self.win32gui_mock.IsWindowVisible.return_value = True
        self.win32gui_mock.GetWindowText.side_effect = ["Window 1", "Window 2"]
        
        wm = self.WindowManager()
        result = wm.get_all_windows()
        
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0], (12345, "Window 1"))
        self.assertEqual(result[1], (67890, "Window 2"))
    
    def test_set_window_position(self):
        """Test setting window position."""
        wm = self.WindowManager()
        wm.set_window_position(12345, 100, 200, 800, 600)
        
        self.win32gui_mock.SetWindowPos.assert_called_once_with(
            12345, 0, 100, 200, 800, 600, 0
        )
    
    def test_bring_window_to_front(self):
        """Test bringing window to front."""
        wm = self.WindowManager()
        wm.bring_window_to_front(12345)
        
        self.win32gui_mock.SetForegroundWindow.assert_called_once_with(12345)


class TestWin32Main(unittest.TestCase):
    """Test main function."""
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', False)
    @patch('builtins.print')
    def test_main_without_pywin32(self, mock_print):
        """Test main function when pywin32 is not available."""
        from learn_windows_automation.win32 import main
        main()
        
        mock_print.assert_any_call("Windows Automation using pywin32")
        mock_print.assert_any_call("Error: pywin32 is not available. This tool requires Windows.")


if __name__ == '__main__':
    unittest.main()
