"""Tests for core functionality."""

import pytest
from win_automation.core import ScriptManager


class TestScriptManager:
    """Test cases for ScriptManager class."""
    
    def test_init_default_directory(self) -> None:
        """Test ScriptManager initialization with default directory."""
        manager = ScriptManager()
        assert manager.scripts_directory == "scripts"
    
    def test_init_custom_directory(self) -> None:
        """Test ScriptManager initialization with custom directory."""
        custom_dir = "custom_scripts"
        manager = ScriptManager(custom_dir)
        assert manager.scripts_directory == custom_dir
    
    def test_discover_scripts_empty(self) -> None:
        """Test script discovery returns empty list when not implemented."""
        manager = ScriptManager()
        scripts = manager.discover_scripts()
        assert scripts == []
    
    def test_execute_script_placeholder(self) -> None:
        """Test script execution returns success code when not implemented."""
        manager = ScriptManager()
        exit_code = manager.execute_script("test_script", ["arg1", "arg2"])
        assert exit_code == 0
