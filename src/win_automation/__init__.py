"""Windows automation toolkit main package."""

from .core import ScriptManager, ConfigManager
from .utils import WindowsUtils, FileSystemUtils

__version__ = "0.1.0"

__all__ = [
    "ScriptManager",
    "ConfigManager", 
    "WindowsUtils",
    "FileSystemUtils",
]
