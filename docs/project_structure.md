# Project Structure Design

## Final Directory Structure

```
win-automation/
├── src/win_automation/           # Main package
│   ├── __init__.py               # Package version and exports
│   ├── cli.py                    # Main CLI entry point
│   ├── core/                     # Core functionality
│   │   ├── __init__.py
│   │   ├── script_manager.py     # Script discovery and execution
│   │   ├── config.py             # Configuration management
│   │   └── exceptions.py         # Custom exceptions
│   ├── utils/                    # Reusable utilities
│   │   ├── __init__.py
│   │   ├── windows.py            # Windows API wrappers
│   │   ├── office.py             # Office automation utilities
│   │   ├── web.py                # Web automation utilities
│   │   └── file_system.py        # File system operations
│   └── mocks/                    # Mock implementations for testing
│       ├── __init__.py
│       ├── win32_mock.py         # pywin32 mocks
│       └── pywinauto_mock.py     # pywinauto mocks
├── scripts/                      # User scripts directory
│   ├── office/                   # Office automation scripts
│   │   ├── excel/
│   │   ├── powerpoint/
│   │   ├── word/
│   │   └── outlook/
│   ├── productivity/             # Productivity automation
│   │   ├── file_management/
│   │   ├── text_processing/
│   │   └── data_conversion/
│   ├── system/                   # System automation
│   │   ├── window_management/
│   │   ├── process_control/
│   │   └── registry_operations/
│   └── web/                      # Web automation
│       ├── browser_automation/
│       └── web_scraping/
├── tests/                        # Test suite
│   ├── unit/                     # Unit tests with mocks
│   │   ├── test_core/
│   │   └── test_utils/
│   ├── integration/              # Integration tests (safe operations)
│   │   ├── test_script_manager/
│   │   └── test_config/
│   ├── manual/                   # Manual test instructions
│   │   ├── office_tests/
│   │   └── system_tests/
│   └── fixtures/                 # Test data and fixtures
├── config/                       # Configuration files
│   ├── scripts.toml              # Script registry
│   └── settings.toml             # Application settings
├── docs/                         # Documentation
│   ├── api/                      # API documentation
│   ├── user_guide/               # User guides
│   └── testing_scope.md          # Testing strategy
└── examples/                     # Example user scripts
    ├── simple_clipboard.py
    ├── excel_report_generator.py
    └── window_organizer.py
```

## Module Functionality Design

### Core Modules

#### 1. Script Manager (`core/script_manager.py`)
- **Responsibility**: Discover, validate, and execute user scripts
- **Features**:
  - Script discovery from `scripts/` directory
  - Script metadata parsing and validation
  - Argument passing and environment setup
  - Error handling and logging

#### 2. Configuration Manager (`core/config.py`)
- **Responsibility**: Manage application and script configurations
- **Features**:
  - TOML-based configuration parsing
  - Script registration and metadata storage
  - Settings validation and defaults
  - Environment-specific configurations

#### 3. Utils Package (`utils/`)
- **Purpose**: Provide high-level, reusable automation utilities
- **Modules**:
  - `windows.py`: Window management, clipboard, system info
  - `office.py`: Excel, Word, PowerPoint automation helpers
  - `web.py`: Browser automation utilities
  - `file_system.py`: File operations, path utilities

### Script Organization Strategy

#### By Use Case Categories:
1. **Office**: Excel reports, PowerPoint automation, Word processing
2. **Productivity**: File management, text processing, data conversion
3. **System**: Window management, process control, registry operations
4. **Web**: Browser automation, web scraping

#### Script Registration System:
Each script directory contains a `script.toml` file with metadata:
```toml
[script]
name = "excel_report_generator"
description = "Generate monthly reports from Excel data"
category = "office"
subcategory = "excel"

[arguments]
input_file = { type = "string", required = true, description = "Input Excel file path" }
output_dir = { type = "string", required = false, default = ".", description = "Output directory" }

[requirements]
dependencies = ["pywin32", "openpyxl"]
min_python_version = "3.13"
```

### Testing Architecture

#### Three-Tier Testing Strategy:
1. **Unit Tests**: 100% mocked, fast execution
2. **Integration Tests**: Safe Windows operations, CI-compatible
3. **Manual Tests**: Office applications, documented procedures

#### Mock Framework:
- Centralized mock implementations in `src/win_automation/mocks/`
- Reusable mock fixtures for common Windows operations
- Automatic mock injection for development on non-Windows platforms

### Configuration Management

#### TOML-Based Configuration:
- `config/scripts.toml`: Registry of all available scripts
- `config/settings.toml`: Application-wide settings
- Per-script `script.toml`: Individual script metadata

#### Example `config/scripts.toml`:
```toml
[scripts.excel_report_generator]
path = "scripts/office/excel/report_generator.py"
enabled = true
last_updated = "2025-06-25T08:00:00Z"

[scripts.window_organizer]
path = "scripts/system/window_management/organizer.py"
enabled = true
last_updated = "2025-06-25T08:00:00Z"
```

## Benefits of This Structure

1. **Scalability**: Easy to add new script categories and utilities
2. **Maintainability**: Clear separation of concerns and responsibilities
3. **Testability**: Comprehensive testing strategy with proper mocking
4. **User-Friendly**: Intuitive script organization by use case
5. **Flexibility**: Support for both simple scripts and complex automation workflows

## Migration Path

1. **Phase 1**: Implement core structure and basic script manager
2. **Phase 2**: Add utility modules and mock framework
3. **Phase 3**: Implement configuration management and script registry
4. **Phase 4**: Add example scripts and documentation
5. **Phase 5**: Extend with advanced features (interactive CLI, etc.)
