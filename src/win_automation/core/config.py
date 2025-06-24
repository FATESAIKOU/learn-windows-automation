"""Configuration management for Windows automation toolkit."""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional
import tomllib

from .exceptions import ConfigurationError


class ConfigManager:
    """Manages application and script configurations."""
    
    def __init__(self, config_dir: str = "config") -> None:
        """Initialize configuration manager.
        
        Args:
            config_dir: Directory containing configuration files
        """
        self.config_dir = Path(config_dir)
        self._scripts_config: Optional[Dict[str, Any]] = None
        self._settings_config: Optional[Dict[str, Any]] = None
    
    @property
    def scripts_config_path(self) -> Path:
        """Path to scripts configuration file."""
        return self.config_dir / "scripts.toml"
    
    @property
    def settings_config_path(self) -> Path:
        """Path to settings configuration file."""
        return self.config_dir / "settings.toml"
    
    def load_scripts_config(self) -> Dict[str, Any]:
        """Load scripts configuration from TOML file.
        
        Returns:
            Scripts configuration dictionary
            
        Raises:
            ConfigurationError: If configuration file cannot be loaded
        """
        if self._scripts_config is None:
            self._scripts_config = self._load_toml_file(self.scripts_config_path)
        return self._scripts_config
    
    def load_settings_config(self) -> Dict[str, Any]:
        """Load application settings from TOML file.
        
        Returns:
            Settings configuration dictionary
            
        Raises:
            ConfigurationError: If configuration file cannot be loaded
        """
        if self._settings_config is None:
            self._settings_config = self._load_toml_file(self.settings_config_path)
        return self._settings_config
    
    def get_script_info(self, script_name: str) -> Optional[Dict[str, Any]]:
        """Get information about a specific script.
        
        Args:
            script_name: Name of the script
            
        Returns:
            Script information dictionary or None if not found
        """
        scripts_config = self.load_scripts_config()
        return scripts_config.get("scripts", {}).get(script_name)
    
    def list_available_scripts(self) -> List[str]:
        """List all available script names.
        
        Returns:
            List of script names
        """
        scripts_config = self.load_scripts_config()
        return list(scripts_config.get("scripts", {}).keys())
    
    def _load_toml_file(self, file_path: Path) -> Dict[str, Any]:
        """Load a TOML file and return its contents.
        
        Args:
            file_path: Path to the TOML file
            
        Returns:
            Parsed TOML content
            
        Raises:
            ConfigurationError: If file cannot be loaded or parsed
        """
        if not file_path.exists():
            return {}
        
        try:
            with open(file_path, "rb") as f:
                return tomllib.load(f)
        except (OSError, tomllib.TOMLDecodeError) as e:
            raise ConfigurationError(f"Failed to load {file_path}: {e}") from e
