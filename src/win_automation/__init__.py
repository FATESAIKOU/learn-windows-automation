"""Windows automation toolkit main package."""

from .core import ConfigManager, ScriptManager
from .utils import FileSystemUtils, WindowsUtils

__version__ = "0.1.0"

__all__ = [
    "ScriptManager",
    "ConfigManager",
    "WindowsUtils",
    "FileSystemUtils",
]
