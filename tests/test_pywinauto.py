"""
Tests for pywinauto automation module.

These tests use mocking to avoid requiring a Windows environment.
"""

import unittest
from unittest.mock import Mock, patch
import sys


class TestApplicationManager(unittest.TestCase):
    """Test ApplicationManager class."""
    
    @patch('learn_windows_automation.pywinauto.PYWINAUTO_AVAILABLE', True)
    @patch('learn_windows_automation.pywinauto.Application')
    @patch('learn_windows_automation.pywinauto.Desktop')
    @patch('learn_windows_automation.pywinauto.find_elements')
    def test_application_manager_initialization(self, mock_find_elements, mock_desktop, mock_application):
        """Test ApplicationManager can be initialized when pywinauto is available."""
        from learn_windows_automation.pywinauto import ApplicationManager
        
        am = ApplicationManager()
        self.assertIsInstance(am, ApplicationManager)


class TestMainFunction(unittest.TestCase):
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
