"""
Tests for win32 automation module.

These tests use mocking to avoid requiring a Windows environment.
"""

import unittest
from unittest.mock import Mock, patch
import sys


class TestWindowManager(unittest.TestCase):
    """Test WindowManager class."""
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
    @patch('learn_windows_automation.win32.win32gui')
    @patch('learn_windows_automation.win32.win32con')
    @patch('learn_windows_automation.win32.win32api')
    @patch('learn_windows_automation.win32.win32process')
    def test_window_manager_initialization(self, mock_win32process, mock_win32api, mock_win32con, mock_win32gui):
        """Test WindowManager can be initialized when pywin32 is available."""
        from learn_windows_automation.win32 import WindowManager
        
        wm = WindowManager()
        self.assertIsInstance(wm, WindowManager)
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
    @patch('learn_windows_automation.win32.win32gui')
    def test_find_window_by_title(self, mock_win32gui):
        """Test finding window by title."""
        from learn_windows_automation.win32 import WindowManager
        
        # Setup mock
        mock_win32gui.FindWindow.return_value = 12345
        
        wm = WindowManager()
        result = wm.find_window_by_title("Test Window")
        
        mock_win32gui.FindWindow.assert_called_once_with(None, "Test Window")
        self.assertEqual(result, 12345)
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
    @patch('learn_windows_automation.win32.win32gui')
    def test_get_window_text(self, mock_win32gui):
        """Test getting window text."""
        from learn_windows_automation.win32 import WindowManager
        
        # Setup mock
        mock_win32gui.GetWindowText.return_value = "Test Window Title"
        
        wm = WindowManager()
        result = wm.get_window_text(12345)
        
        mock_win32gui.GetWindowText.assert_called_once_with(12345)
        self.assertEqual(result, "Test Window Title")
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
    @patch('learn_windows_automation.win32.win32gui')
    def test_get_all_windows(self, mock_win32gui):
        """Test getting all windows."""
        from learn_windows_automation.win32 import WindowManager
        
        # Setup mock for EnumWindows
        def mock_enum_windows(callback, ctx):
            # Simulate calling the callback with some window handles
            callback(12345, ctx)  # Visible window
            callback(67890, ctx)  # Another visible window
        
        mock_win32gui.EnumWindows.side_effect = mock_enum_windows
        mock_win32gui.IsWindowVisible.return_value = True
        mock_win32gui.GetWindowText.side_effect = ["Window 1", "Window 2"]
        
        wm = WindowManager()
        result = wm.get_all_windows()
        
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0], (12345, "Window 1"))
        self.assertEqual(result[1], (67890, "Window 2"))
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
    @patch('learn_windows_automation.win32.win32gui')
    def test_set_window_position(self, mock_win32gui):
        """Test setting window position."""
        from learn_windows_automation.win32 import WindowManager
        
        wm = WindowManager()
        wm.set_window_position(12345, 100, 200, 800, 600)
        
        mock_win32gui.SetWindowPos.assert_called_once_with(
            12345, 0, 100, 200, 800, 600, 0
        )
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
    @patch('learn_windows_automation.win32.win32gui')
    def test_bring_window_to_front(self, mock_win32gui):
        """Test bringing window to front."""
        from learn_windows_automation.win32 import WindowManager
        
        wm = WindowManager()
        wm.bring_window_to_front(12345)
        
        mock_win32gui.SetForegroundWindow.assert_called_once_with(12345)


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
    
    @patch('learn_windows_automation.win32.PYWIN32_AVAILABLE', True)
    @patch('learn_windows_automation.win32.WindowManager')
    @patch('builtins.print')
    def test_main_with_pywin32_available(self, mock_print, mock_wm_class):
        """Test main function when pywin32 is available."""
        from learn_windows_automation.win32 import main
        
        # Setup mock
        mock_wm = Mock()
        mock_wm.get_all_windows.return_value = [
            (12345, "Window 1"),
            (67890, "Window 2")
        ]
        mock_wm_class.return_value = mock_wm
        
        main()
        
        mock_print.assert_any_call("Windows Automation using pywin32")
        mock_print.assert_any_call("Found 2 visible windows:")
        mock_print.assert_any_call("  HWND: 12345, Title: Window 1")
        mock_print.assert_any_call("  HWND: 67890, Title: Window 2")


if __name__ == '__main__':
    unittest.main()
    
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
