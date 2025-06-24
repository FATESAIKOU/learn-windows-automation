"""Tests for Windows utilities."""

import pytest
from unittest.mock import patch, MagicMock
from win_automation.utils import WindowsUtils, AutomationUtils


class TestWindowsUtils:
    """Test cases for WindowsUtils class."""
    
    @patch('win_automation.utils.WIN32_AVAILABLE', True)
    @patch('win_automation.utils.win32clipboard')
    def test_get_clipboard_text_success(self, mock_clipboard: MagicMock) -> None:
        """Test successful clipboard text retrieval."""
        mock_clipboard.GetClipboardData.return_value = "test text"
        
        result = WindowsUtils.get_clipboard_text()
        
        assert result == "test text"
        mock_clipboard.OpenClipboard.assert_called_once()
        mock_clipboard.GetClipboardData.assert_called_once()
        mock_clipboard.CloseClipboard.assert_called_once()
    
    @patch('win_automation.utils.WIN32_AVAILABLE', False)
    def test_get_clipboard_text_no_win32(self) -> None:
        """Test clipboard access when win32 not available."""
        with pytest.raises(RuntimeError, match="pywin32 not available"):
            WindowsUtils.get_clipboard_text()
    
    @patch('win_automation.utils.WIN32_AVAILABLE', True)
    @patch('win_automation.utils.win32clipboard')
    def test_set_clipboard_text_success(self, mock_clipboard: MagicMock) -> None:
        """Test successful clipboard text setting."""
        result = WindowsUtils.set_clipboard_text("test text")
        
        assert result is True
        mock_clipboard.OpenClipboard.assert_called_once()
        mock_clipboard.EmptyClipboard.assert_called_once()
        mock_clipboard.SetClipboardData.assert_called_once()
        mock_clipboard.CloseClipboard.assert_called_once()


class TestAutomationUtils:
    """Test cases for AutomationUtils class."""
    
    @patch('win_automation.utils.PYWINAUTO_AVAILABLE', True)
    @patch('win_automation.utils.pywinauto')
    def test_find_window_success(self, mock_pywinauto: MagicMock) -> None:
        """Test successful window finding."""
        mock_window = MagicMock()
        mock_app = MagicMock()
        mock_app.connect.return_value = mock_window
        mock_pywinauto.Application.return_value = mock_app
        
        result = AutomationUtils.find_window("Test Window")
        
        assert result == mock_window
        mock_pywinauto.Application.assert_called_once()
        mock_app.connect.assert_called_once_with(title="Test Window")
    
    @patch('win_automation.utils.PYWINAUTO_AVAILABLE', False)
    def test_find_window_no_pywinauto(self) -> None:
        """Test window finding when pywinauto not available."""
        with pytest.raises(RuntimeError, match="pywinauto not available"):
            AutomationUtils.find_window("Test Window")
