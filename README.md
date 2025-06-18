# Windows Desktop Automation Toolkit

A comprehensive Python project for automating Windows desktop applications using two different approaches:

- **win32**: Low-level Windows API access via pywin32 for direct system interaction
- **pywinauto**: High-level automation via pywinauto for easier application control

## Features

- 🔧 **Dual Backend Support**: Choose between low-level (win32) and high-level (pywinauto) automation
- 📦 **Modern Python Packaging**: Uses `uv` for fast dependency management
- 🧪 **Comprehensive Testing**: Full test suite with mocking for cross-platform development
- 🚀 **Easy to Use**: Command-line interface and console scripts for quick access
- 🔍 **Well Documented**: Clear separation of concerns and comprehensive documentation

## Requirements

- Python 3.8+
- Windows OS (for runtime, development can be done on any platform)
- `uv` package manager

## Installation

### Using uv (Recommended)

```bash
# Clone the repository
git clone <repository-url>
cd learn-windows-automation

# Install dependencies
uv sync

# For development with additional tools
uv sync --dev
```

### Using pip (Alternative)

```bash
pip install -e .
```

## Quick Start

### Command Line Interface

```bash
# Using win32 (low-level Windows API)
uv run python main.py win32

# Using pywinauto (high-level automation)
uv run python main.py pywinauto
```

### Console Scripts

After installation, you can use the console scripts directly:

```bash
# Win32 automation
win32-automation

# pywinauto automation
pywinauto-automation
```

### Programmatic Usage

```python
# Using win32 for low-level operations
from learn_windows_automation.win32 import WindowManager

wm = WindowManager()
windows = wm.get_all_windows()
for hwnd, title in windows:
    print(f"Window: {title}")

# Using pywinauto for high-level automation
from learn_windows_automation.pywinauto import ApplicationManager

am = ApplicationManager()
app = am.connect_by_title("Calculator")
```

## Development

### Using Makefile (Recommended)

```bash
# Install development dependencies
make install-dev

# Run tests
make test

# Run tests with coverage
make test-coverage

# Clean build artifacts
make clean

# See all available commands
make help
```

### Manual Commands

```bash
# Install dependencies
uv sync --dev

# Run tests
uv run pytest

# Run tests with coverage
uv run pytest --cov=learn_windows_automation --cov-report=html

# Run specific test file
uv run pytest tests/test_win32.py -v
```

## Project Structure

```
learn-windows-automation/
├── learn_windows_automation/          # Main package
│   ├── __init__.py
│   ├── win32_automation.py           # Win32 entry point
│   ├── pywinauto_automation.py       # pywinauto entry point
│   ├── win32/                        # Low-level Windows API automation
│   │   └── __init__.py
│   └── pywinauto/                    # High-level automation
│       └── __init__.py
├── tests/                            # Test suite
│   ├── __init__.py
│   ├── test_main.py                  # Main entry point tests
│   ├── test_win32.py                 # Win32 module tests
│   └── test_pywinauto.py             # pywinauto module tests
├── main.py                           # Main CLI entry point
├── pyproject.toml                    # Project configuration
├── Makefile                          # Development shortcuts
└── README.md                         # This file
```

## API Overview

### Win32 Module (Low-level)

The `win32` module provides direct access to Windows APIs:

```python
from learn_windows_automation.win32 import WindowManager

wm = WindowManager()

# Find windows
hwnd = wm.find_window_by_title("Notepad")
title = wm.get_window_text(hwnd)

# Manipulate windows
wm.set_window_position(hwnd, 100, 100, 800, 600)
wm.bring_window_to_front(hwnd)

# List all windows
windows = wm.get_all_windows()
```

### pywinauto Module (High-level)

The `pywinauto` module provides easier application automation:

```python
from learn_windows_automation.pywinauto import ApplicationManager, WindowController

# Connect to applications
am = ApplicationManager()
app = am.connect_by_title("Calculator")

# Control application
controller = WindowController(app)
main_window = controller.get_main_window()
controller.click_button(main_window, "1")
controller.click_button(main_window, "+")
controller.click_button(main_window, "2")
```

## Testing

The project includes comprehensive tests that work on any platform using mocking:

```bash
# Run all tests
make test

# Run with coverage
make test-coverage

# Run specific test categories
uv run pytest -m "not slow"          # Skip slow tests
uv run pytest tests/test_win32.py    # Test specific module
```

## Key Differences: win32 vs pywinauto

| Feature | win32 | pywinauto |
|---------|-------|-----------|
| **Level** | Low-level Windows API | High-level automation |
| **Learning Curve** | Steeper | Easier |
| **Performance** | Faster | Moderate |
| **Flexibility** | Very high | High |
| **Best For** | System-level operations | Application automation |
| **Dependencies** | pywin32 only | pywinauto + comtypes |

## Common Use Cases

### Win32 Module
- Window management and positioning
- System-level automation
- Performance-critical operations
- Custom window handling

### pywinauto Module
- Application testing and automation
- GUI interaction simulation
- Form filling and data entry
- Cross-application workflows

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Run the test suite: `make test`
6. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Troubleshooting

### Common Issues

1. **ImportError on non-Windows systems**: This is expected for runtime. Development and testing work on all platforms.

2. **Permission errors**: Some Windows automation operations require elevated privileges.

3. **Application not found**: Ensure the target application is running and visible.

### Getting Help

- Check the test files for usage examples
- Review the API documentation in the source code
- Create an issue for bugs or feature requests

## Changelog

### v0.1.0
- Initial release
- Basic win32 and pywinauto integration
- Command-line interface
- Comprehensive test suite
- Cross-platform development support

### Programmatic Usage

#### Using win32 (Low-level)

```python
from learn_windows_automation.win32 import WindowManager

# Create window manager
wm = WindowManager()

# Find a window by title
hwnd = wm.find_window_by_title("Notepad")

# Get all visible windows
windows = wm.get_all_windows()

# Bring window to front
wm.bring_window_to_front(hwnd)
```

#### Using pywinauto (High-level)

```python
from learn_windows_automation.pywinauto import ApplicationManager, WindowController

# Connect to an application
am = ApplicationManager()
app = am.connect_by_title("Calculator")

# Control the application
controller = WindowController(app)
main_window = controller.get_main_window()
controller.click_button(main_window, "1")
```

## Development

### Running Tests

```bash
# Run all tests
uv run python -m pytest tests/

# Run tests with unittest
uv run python -m unittest discover tests/

# Run specific test file
uv run python -m unittest tests.test_win32
```

### Project Structure

```
learn-windows-automation/
├── learn_windows_automation/
│   ├── __init__.py
│   ├── win32/
│   │   └── __init__.py          # Low-level Windows API automation
│   ├── pywinauto/
│   │   └── __init__.py          # High-level automation
│   ├── win32_automation.py      # Win32 entry point
│   └── pywinauto_automation.py  # pywinauto entry point
├── tests/
│   ├── __init__.py
│   ├── test_win32.py           # Tests for win32 module
│   ├── test_pywinauto.py       # Tests for pywinauto module
│   └── test_main.py            # Tests for main entry points
├── main.py                     # Main CLI entry point
├── pyproject.toml             # Project configuration
└── README.md
```

## Dependencies

### Core Dependencies
- `pywin32>=306`: Windows API access
- `pywinauto>=0.6.8`: High-level Windows automation

### Development Dependencies
- `pytest>=7.0.0`: Testing framework
- `pytest-mock>=3.10.0`: Mocking utilities

## Platform Support

This toolkit is designed specifically for Windows environments. However, the test suite uses mocking to allow development and testing on non-Windows platforms.

## License

This project is open source. Please check the license file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run the tests to ensure they pass
5. Submit a pull request

## Notes

- The win32 module provides direct access to Windows APIs for maximum control
- The pywinauto module offers a more user-friendly interface for common automation tasks
- Both modules include error handling for missing dependencies
- Tests use comprehensive mocking to work on any platform
