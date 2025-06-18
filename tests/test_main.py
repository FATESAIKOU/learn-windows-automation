"""
Tests for main entry points.
"""

import unittest
from unittest.mock import patch, Mock
import sys
from io import StringIO


class TestMainEntryPoints(unittest.TestCase):
    """Test main entry points."""
    
    @patch('main.win32_main')
    def test_main_with_win32_backend(self, mock_win32_main):
        """Test main function with win32 backend."""
        with patch('sys.argv', ['main.py', 'win32']):
            from main import main
            
            # Capture output
            with patch('sys.stdout', new_callable=StringIO) as mock_stdout:
                main()
            
            # Check that win32 main was called
            mock_win32_main.assert_called_once()
            
            # Check output
            output = mock_stdout.getvalue()
            self.assertIn("Starting Win32 automation...", output)
    
    @patch('main.pywinauto_main')
    def test_main_with_pywinauto_backend(self, mock_pywinauto_main):
        """Test main function with pywinauto backend."""
        with patch('sys.argv', ['main.py', 'pywinauto']):
            from main import main
            
            # Capture output
            with patch('sys.stdout', new_callable=StringIO) as mock_stdout:
                main()
            
            # Check that pywinauto main was called
            mock_pywinauto_main.assert_called_once()
            
            # Check output
            output = mock_stdout.getvalue()
            self.assertIn("Starting pywinauto automation...", output)


class TestWin32Automation(unittest.TestCase):
    """Test win32_automation entry point."""
    
    @patch('learn_windows_automation.win32.main')
    def test_win32_automation_entry_point(self, mock_main):
        """Test win32_automation entry point."""
        from learn_windows_automation.win32_automation import main
        
        # This should just call the win32 main function
        main()
        mock_main.assert_called_once()


class TestPywinautoAutomation(unittest.TestCase):
    """Test pywinauto_automation entry point."""
    
    @patch('learn_windows_automation.pywinauto.main')
    def test_pywinauto_automation_entry_point(self, mock_main):
        """Test pywinauto_automation entry point."""
        from learn_windows_automation.pywinauto_automation import main
        
        # This should just call the pywinauto main function
        main()
        mock_main.assert_called_once()


if __name__ == '__main__':
    unittest.main()
