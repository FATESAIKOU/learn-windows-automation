# Windows Automation Toolkit

A scalable, testable, and maintainable Windows automation toolkit in Python that provides a unified CLI interface for running automation scripts for Excel, PowerPoint, window management, and other system tasks.

## Features

- 🏗️ **Modular Architecture**: Clean separation of core functionality, utilities, and examples
- 🖥️ **Unified CLI**: Single command-line interface for all automation scripts
- 📊 **Office Automation**: Support for Excel and PowerPoint automation via pywin32
- 🪟 **Window Management**: Advanced window control and workspace organization via pywinauto
- ⚙️ **Configuration Management**: TOML-based configuration for scripts and settings
- 🧪 **Comprehensive Testing**: Full test coverage with mocked Windows APIs
- 🎯 **Type Safety**: Complete type annotations with mypy checking
- 🧹 **Code Quality**: Enforced with ruff linting and formatting

## Quick Start

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd learn-windows-automation-gen

# Install dependencies using uv
uv install

# Activate the virtual environment
source .venv/bin/activate  # Linux/Mac
# or
.venv\Scripts\activate     # Windows
```

### Usage

```bash
# List all available scripts
uv run win-automation list-scripts

# Get information about a specific script
uv run win-automation info <script-name>

# Run a script with arguments
uv run win-automation run <script-name> [args...]
```

### Example Commands

```bash
# Simple clipboard operations
uv run win-automation run simple_clipboard get
uv run win-automation run simple_clipboard set "Hello World"

# Test script execution
uv run win-automation run test_execution arg1 arg2

# Advanced Excel processing (requires Excel)
uv run win-automation run excel_data_processor sample.xlsx

# Window organization
uv run win-automation run window_organizer
```

## Project Structure

```
learn-windows-automation-gen/
├── src/win_automation/          # Main package
│   ├── __init__.py             # Package exports
│   ├── cli.py                  # CLI interface using Typer
│   ├── core/                   # Core functionality
│   │   ├── __init__.py
│   │   ├── config.py           # Configuration management
│   │   ├── exceptions.py       # Custom exceptions
│   │   └── script_manager.py   # Script discovery and execution
│   └── utils/                  # Utility modules
│       ├── __init__.py
│       ├── file_system.py      # File system operations
│       └── windows.py          # Windows-specific utilities
├── examples/                   # Automation scripts
│   ├── simple_clipboard.py     # Basic clipboard operations
│   ├── excel_data_processor.py # Advanced Excel automation
│   ├── powerpoint_batch_processor.py # PowerPoint automation
│   ├── multi_app_coordinator.py # Multi-application workflows
│   ├── advanced_window_manager.py # Window management
│   └── ...                     # Additional examples
├── config/                     # Configuration files
│   ├── scripts.toml           # Script registry
│   └── settings.toml          # Application settings
├── tests/                     # Test suite
├── docs/                      # Documentation
└── pyproject.toml            # Project configuration
```

## Available Scripts

### Basic Examples
- **simple_clipboard**: Text clipboard manipulation
- **test_execution**: Testing script execution with arguments
- **window_organizer**: Basic window organization demo

### Advanced Office Automation
- **excel_data_processor**: Advanced Excel processing with charts and formatting
- **powerpoint_batch_processor**: Batch PowerPoint presentation processing
- **multi_app_coordinator**: Coordinate workflows between multiple applications
- **advanced_window_manager**: Sophisticated window management and layouts

## Adding New Scripts

1. **Create your script** in the `examples/` directory
2. **Add script entry** to `config/scripts.toml`:
   ```toml
   [scripts.your_script]
   path = "examples/your_script.py"
   description = "Description of your script"
   category = "category"
   subcategory = "subcategory"
   enabled = true
   ```
3. **Write tests** in the `tests/` directory
4. **Update documentation** as needed

### Script Template

```python
"""Your automation script."""

import sys
from typing import Any

# Import required automation utilities
from win_automation.utils import WindowsUtils
# or other utilities as needed

def main() -> int:
    """Main function for your script."""
    try:
        # Your automation logic here
        print("Script running...")
        return 0
    except Exception as e:
        print(f"Error: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
```

## Dependencies

### Core Dependencies
- **Python 3.13+**: Modern Python with latest features
- **typer**: CLI framework
- **pywin32**: Windows API access (Excel, PowerPoint, etc.)
- **pywinauto**: GUI automation
- **tomllib/tomli**: TOML configuration parsing

### Development Dependencies
- **pytest**: Testing framework
- **pytest-mock**: Mocking for tests
- **ruff**: Linting and formatting
- **mypy**: Type checking
- **uv**: Fast package management

## Development

### Running Tests

```bash
# Run all tests
uv run pytest

# Run specific test file
uv run pytest tests/test_core.py

# Run with verbose output
uv run pytest -v

# Run tests with coverage
uv run pytest --cov=src/win_automation
```

### Code Quality

```bash
# Lint code
uv run ruff check

# Format code
uv run ruff format

# Type checking
uv run mypy .

# Fix auto-fixable issues
uv run ruff check --fix
```

### Project Management

```bash
# Install dependencies
uv install

# Add new dependency
uv add package-name

# Add development dependency
uv add --dev package-name

# Update dependencies
uv update
```

## Architecture

### Core Components

1. **ScriptManager**: Discovers, validates, and executes automation scripts
2. **ConfigManager**: Handles TOML configuration loading and validation
3. **CLI Module**: Provides the command-line interface using Typer
4. **Utilities**: Windows-specific helpers for clipboard, file system, etc.

### Error Handling

The toolkit uses custom exceptions for better error reporting:
- `AutomationError`: Base exception for all automation errors
- `ScriptNotFoundError`: When a requested script cannot be found
- `ScriptExecutionError`: When script execution fails
- `ConfigError`: Configuration-related errors

### Testing Strategy

- **Unit Tests**: Mock Windows APIs for reliable testing
- **Integration Tests**: Test CLI commands and script execution
- **Type Safety**: Full mypy coverage ensures type correctness
- **Code Quality**: Ruff ensures consistent code style

## Configuration

### Script Configuration (`config/scripts.toml`)

```toml
[scripts.script_name]
path = "examples/script_name.py"
description = "Script description"
category = "office|system|productivity|automation"
subcategory = "excel|powerpoint|clipboard|window"
enabled = true
```

### Application Settings (`config/settings.toml`)

```toml
[automation]
timeout = 30
retry_count = 3

[logging]
level = "INFO"
file = "automation.log"
```

## Examples and Use Cases

### Excel Automation
```bash
# Process Excel data with charts
uv run win-automation run excel_data_processor data.xlsx

# Generate reports
uv run win-automation run excel_report_generator
```

### PowerPoint Automation
```bash
# Batch process presentations
uv run win-automation run powerpoint_batch_processor batch config.json

# Create sample presentation
uv run win-automation run powerpoint_batch_processor create sample.pptx
```

### Multi-Application Workflows
```bash
# Coordinate Notepad and Calculator
uv run win-automation run multi_app_coordinator calc_notepad

# Data flow between applications
uv run win-automation run multi_app_coordinator data_flow
```

### Window Management
```bash
# Advanced window management
uv run win-automation run advanced_window_manager

# Organize windows
uv run win-automation run window_organizer
```

## Contributing

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature-name`
3. **Write tests** for your changes
4. **Ensure code quality**: Run `uv run ruff check` and `uv run mypy .`
5. **Run tests**: `uv run pytest`
6. **Submit a pull request**

### Code Style

- Follow PEP 8 style guidelines
- Use type hints for all functions
- Write docstrings for public APIs
- Keep line length under 88 characters
- Use meaningful variable and function names

## Troubleshooting

### Common Issues

1. **ImportError for win32com**: Install pywin32: `uv add pywin32`
2. **ImportError for pywinauto**: Install pywinauto: `uv add pywinauto`
3. **Script not found**: Check `config/scripts.toml` for correct path
4. **Permission denied**: Run as administrator for system-level operations

### Windows-Specific Notes

- Some operations require administrator privileges
- Excel/PowerPoint scripts need the respective applications installed
- Window management works best on Windows 10+ with UI Automation

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [uv](https://github.com/astral-sh/uv) for fast dependency management
- Uses [Typer](https://typer.tiangolo.com/) for the CLI interface
- Automation powered by [pywin32](https://github.com/mhammond/pywin32) and [pywinauto](https://github.com/pywinauto/pywinauto)
- Code quality ensured by [Ruff](https://github.com/astral-sh/ruff) and [mypy](https://mypy.readthedocs.io/)
