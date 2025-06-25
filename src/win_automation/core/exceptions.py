"""Custom exceptions for Windows automation toolkit."""


class AutomationError(Exception):
    """Base exception for automation-related errors."""
    pass


class ScriptNotFoundError(AutomationError):
    """Raised when a requested script cannot be found."""
    pass


class ConfigurationError(AutomationError):
    """Raised when there's an issue with configuration."""
    pass


class ScriptExecutionError(AutomationError):
    """Raised when script execution fails."""
    pass


class WindowsAPIError(AutomationError):
    """Raised when Windows API calls fail."""
    pass
