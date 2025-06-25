"""Script management for Windows automation toolkit."""

import subprocess
import sys
from pathlib import Path
from typing import Any

from .config import ConfigManager
from .exceptions import ScriptExecutionError, ScriptNotFoundError


class ScriptManager:
    """Manages user scripts discovery and execution."""

    def __init__(
        self,
        scripts_directory: str = "scripts",
        config_manager: ConfigManager | None = None
    ) -> None:
        """Initialize script manager."""
        self.scripts_directory = Path(scripts_directory)
        self.config_manager = config_manager or ConfigManager()

    def discover_scripts(self) -> list[str]:
        """Discover available user scripts."""
        return self.config_manager.list_available_scripts()

    def execute_script(self, script_name: str, args: list[str]) -> int:
        """Execute a user script with arguments."""
        script_info = self.get_script_info(script_name)
        if not script_info:
            raise ScriptNotFoundError(f"Script '{script_name}' not found")

        if not script_info.get("enabled", True):
            raise ScriptExecutionError(f"Script '{script_name}' is disabled")

        script_path = Path(script_info["path"])
        if not script_path.exists():
            raise ScriptExecutionError(
                f"Script file not found: {script_path}"
            )

        # Make path absolute if it's relative
        if not script_path.is_absolute():
            script_path = Path.cwd() / script_path

        try:
            # Execute the script using the current Python interpreter
            cmd = [sys.executable, str(script_path)] + args
            result = subprocess.run(
                cmd,
                capture_output=False,
                text=True,
                check=False
            )
            return result.returncode
        except Exception as e:
            raise ScriptExecutionError(
                f"Failed to execute script '{script_name}': {e}"
            ) from e

    def validate_script(self, script_name: str) -> bool:
        """Validate that a script exists."""
        script_info = self.config_manager.get_script_info(script_name)
        return script_info is not None

    def get_script_info(self, script_name: str) -> dict[str, Any] | None:
        """Get detailed information about a script."""
        return self.config_manager.get_script_info(script_name)
