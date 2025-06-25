"""Tests for core functionality."""

import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from win_automation.core import ConfigManager, ScriptManager
from win_automation.core.exceptions import ScriptExecutionError, ScriptNotFoundError


class TestScriptManager:
    """Test cases for ScriptManager class."""

    def test_init_default_directory(self) -> None:
        """Test ScriptManager initialization with default directory."""
        manager = ScriptManager()
        assert manager.scripts_directory == Path("scripts")

    def test_init_custom_directory(self) -> None:
        """Test ScriptManager initialization with custom directory."""
        custom_dir = "custom_scripts"
        manager = ScriptManager(custom_dir)
        assert manager.scripts_directory == Path(custom_dir)

    def test_discover_scripts_with_mock_config(self) -> None:
        """Test script discovery with mocked config."""
        mock_config = MagicMock(spec=ConfigManager)
        mock_config.list_available_scripts.return_value = []

        manager = ScriptManager(config_manager=mock_config)
        scripts = manager.discover_scripts()
        assert scripts == []

    @patch("win_automation.core.script_manager.subprocess.run")
    def test_execute_script_success(self, mock_subprocess: MagicMock) -> None:
        """Test successful script execution."""
        mock_config = MagicMock(spec=ConfigManager)
        mock_config.get_script_info.return_value = {
            "path": "test_script.py",
            "enabled": True,
        }

        # Mock subprocess.run to return success
        mock_subprocess.return_value.returncode = 0

        # Mock Path.exists() to return True
        with patch("pathlib.Path.exists", return_value=True):
            manager = ScriptManager(config_manager=mock_config)
            exit_code = manager.execute_script("test_script", ["arg1", "arg2"])

        assert exit_code == 0
        mock_subprocess.assert_called_once()

    def test_execute_script_not_found(self) -> None:
        """Test script execution when script not found."""
        mock_config = MagicMock(spec=ConfigManager)
        mock_config.get_script_info.return_value = None

        manager = ScriptManager(config_manager=mock_config)

        with pytest.raises(
            ScriptNotFoundError, match="Script 'nonexistent' not found"
        ):
            manager.execute_script("nonexistent", [])

    def test_execute_script_disabled(self) -> None:
        """Test script execution when script is disabled."""
        mock_config = MagicMock(spec=ConfigManager)
        mock_config.get_script_info.return_value = {
            "path": "test_script.py",
            "enabled": False,
        }

        manager = ScriptManager(config_manager=mock_config)

        with pytest.raises(
            ScriptExecutionError, match="Script 'disabled_script' is disabled"
        ):
            manager.execute_script("disabled_script", [])
