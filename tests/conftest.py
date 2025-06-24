"""Test configuration and shared fixtures."""

import pytest
from unittest.mock import MagicMock


@pytest.fixture
def mock_win32_clipboard():
    """Mock win32clipboard module."""
    mock_clipboard = MagicMock()
    mock_clipboard.OpenClipboard = MagicMock()
    mock_clipboard.CloseClipboard = MagicMock()
    mock_clipboard.GetClipboardData = MagicMock(return_value="test data")
    mock_clipboard.SetClipboardData = MagicMock()
    mock_clipboard.EmptyClipboard = MagicMock()
    mock_clipboard.CF_TEXT = 1
    return mock_clipboard


@pytest.fixture
def mock_pywinauto():
    """Mock pywinauto module."""
    mock_pywinauto = MagicMock()
    mock_app = MagicMock()
    mock_window = MagicMock()
    mock_app.connect.return_value = mock_window
    mock_pywinauto.Application.return_value = mock_app
    return mock_pywinauto
