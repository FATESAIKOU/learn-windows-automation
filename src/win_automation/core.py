"""Core utilities for Windows automation."""


class ScriptManager:
    """Manages user scripts discovery and execution."""

    def __init__(self, scripts_directory: str = "scripts") -> None:
        """Initialize script manager.

        Args:
            scripts_directory: Directory containing user scripts
        """
        self.scripts_directory = scripts_directory

    def discover_scripts(self) -> list[str]:
        """Discover available user scripts.

        Returns:
            List of script names
        """
        # TODO: Implement script discovery
        return []

    def execute_script(self, script_name: str, args: list[str]) -> int:
        """Execute a user script with arguments.

        Args:
            script_name: Name of the script to execute
            args: Arguments to pass to the script

        Returns:
            Exit code of the script
        """
        # TODO: Implement script execution
        return 0
