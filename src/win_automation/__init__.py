"""Windows automation toolkit main package."""

from .cli import main
from .core import ConfigManager, ScriptManager
from .utils import FileSystemUtils, WindowsUtils

__version__ = "0.1.0"

__all__ = [
    "main",
    "ScriptManager",
    "ConfigManager",
    "WindowsUtils",
    "FileSystemUtils",
]
