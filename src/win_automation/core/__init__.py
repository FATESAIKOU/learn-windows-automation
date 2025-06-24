"""Core package for Windows automation toolkit."""

from .script_manager import ScriptManager
from .config import ConfigManager
from .exceptions import AutomationError, ScriptNotFoundError, ConfigurationError

__all__ = [
    "ScriptManager",
    "ConfigManager", 
    "AutomationError",
    "ScriptNotFoundError",
    "ConfigurationError",
]
