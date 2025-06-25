"""Core package for Windows automation toolkit."""

from .config import ConfigManager
from .exceptions import AutomationError, ConfigurationError, ScriptNotFoundError
from .script_manager import ScriptManager

__all__ = [
    "ScriptManager",
    "ConfigManager",
    "AutomationError",
    "ScriptNotFoundError",
    "ConfigurationError",
]
