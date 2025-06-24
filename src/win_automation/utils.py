"""Windows API utilities and wrappers."""

from typing import Any

try:
    import win32clipboard

    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

try:
    import pywinauto

    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False


class WindowsUtils:
    """Utility class for common Windows operations."""

    @staticmethod
    def get_clipboard_text() -> str | None:
        """Get text from Windows clipboard.

        Returns:
            Clipboard text content or None if not available
        """
        if not WIN32_AVAILABLE:
            raise RuntimeError("pywin32 not available")

        try:
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            return data
        except Exception:
            win32clipboard.CloseClipboard()
            return None

    @staticmethod
    def set_clipboard_text(text: str) -> bool:
        """Set text to Windows clipboard.

        Args:
            text: Text to set in clipboard

        Returns:
            True if successful, False otherwise
        """
        if not WIN32_AVAILABLE:
            raise RuntimeError("pywin32 not available")

        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_TEXT, text)
            win32clipboard.CloseClipboard()
            return True
        except Exception:
            win32clipboard.CloseClipboard()
            return False


class AutomationUtils:
    """Utility class for UI automation using pywinauto."""

    @staticmethod
    def find_window(title: str) -> Any | None:
        """Find window by title.

        Args:
            title: Window title to search for

        Returns:
            Window object or None if not found
        """
        if not PYWINAUTO_AVAILABLE:
            raise RuntimeError("pywinauto not available")

        try:
            app = pywinauto.Application()
            window = app.connect(title=title)
            return window
        except Exception:
            return None
