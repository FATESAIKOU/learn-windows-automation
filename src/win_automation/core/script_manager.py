"""Script management for Windows automation toolkit."""

from pathlib import Path
from typing import List, Optional, Dict, Any

from .config import ConfigManager
from .exceptions import ScriptNotFoundError


class ScriptManager:
    """Manages user scripts discovery and execution."""

    def __init__(
        self, 
        scripts_directory: str = "scripts",
        config_manager: Optional[ConfigManager] = None
    ) -> None:
        """Initialize script manager."""
        self.scripts_directory = Path(scripts_directory)
        self.config_manager = config_manager or ConfigManager()

    def discover_scripts(self) -> List[str]:
        """Discover available user scripts."""
        return self.config_manager.list_available_scripts()

    def execute_script(self, script_name: str, args: List[str]) -> int:
        """Execute a user script with arguments."""
        # For now, just return success - will implement in step 4
        print(f"Would execute script: {script_name} with args: {args}")
        return 0

    def validate_script(self, script_name: str) -> bool:
        """Validate that a script exists."""
        script_info = self.config_manager.get_script_info(script_name)
        return script_info is not None

    def get_script_info(self, script_name: str) -> Optional[Dict[str, Any]]:
        """Get detailed information about a script."""
        return self.config_manager.get_script_info(script_name)
